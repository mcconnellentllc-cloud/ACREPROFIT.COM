require('dotenv').config();
const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const Stripe = require('stripe');
const nodemailer = require('nodemailer');
const crypto = require('crypto');
const path = require('path');
const fs = require('fs');

// File upload handling
let multer;
try {
    multer = require('multer');
} catch (e) {
    console.log('Multer not installed. File uploads disabled.');
}

// SharePoint/Excel sync dependencies (optional - only load if configured)
let Client, ClientSecretCredential, cron, XLSX;
try {
    const graphClient = require('@microsoft/microsoft-graph-client');
    const azureIdentity = require('@azure/identity');
    Client = graphClient.Client;
    ClientSecretCredential = azureIdentity.ClientSecretCredential;
    cron = require('node-cron');
    XLSX = require('xlsx');
} catch (e) {
    console.log('SharePoint sync dependencies not installed. Run npm install to enable.')
}

const app = express();

// Initialize Stripe (will be configured per-request for Connect)
const stripe = process.env.STRIPE_SECRET_KEY ? new Stripe(process.env.STRIPE_SECRET_KEY) : null;
const PORT = process.env.PORT || 3001;

// S6: JWT_SECRET must be explicitly set - no silent fallback to a known string
const JWT_SECRET = process.env.JWT_SECRET;
if (!JWT_SECRET) {
    console.error('FATAL: JWT_SECRET environment variable is required. Refusing to start.');
    process.exit(1);
}

// S3: helmet sets common security headers (HSTS, X-Content-Type-Options,
// X-Frame-Options, Referrer-Policy, etc). CSP disabled because this server
// returns JSON not HTML - CSP is enforced by the static frontend host.
app.use(helmet({ contentSecurityPolicy: false }));

// S5: CORS whitelist. Browser requests from any other origin are rejected.
// No-origin requests (Stripe webhooks, curl, server-to-server) are allowed.
const allowedOrigins = [
    'https://acreprofit.com',
    'https://www.acreprofit.com',
    'https://acreprofit-com.onrender.com'
];
if (process.env.NODE_ENV !== 'production') {
    allowedOrigins.push(
        'http://localhost:3000',
        'http://localhost:3001',
        'http://localhost:8080',
        'http://127.0.0.1:3000',
        'http://127.0.0.1:8080'
    );
}
app.use(cors({
    origin: (origin, cb) => {
        if (!origin) return cb(null, true);                 // non-browser / same-origin
        if (allowedOrigins.includes(origin)) return cb(null, true);
        console.warn(`CORS blocked origin: ${origin}`);
        cb(new Error('Not allowed by CORS'));
    },
    credentials: true
}));

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));

// S4: rate limiters for authentication endpoints. Brute-force protection -
// applied per-IP. Webhook, order, and read routes are not limited.
const loginLimiter = rateLimit({
    windowMs: 15 * 60 * 1000,
    max: 10,                                                // 10 attempts / 15 min / IP
    message: { error: 'Too many login attempts. Please try again in 15 minutes.' },
    standardHeaders: true,
    legacyHeaders: false
});
const signupLimiter = rateLimit({
    windowMs: 60 * 60 * 1000,
    max: 5,                                                 // 5 signups / hour / IP
    message: { error: 'Too many signup attempts. Please try again in an hour.' },
    standardHeaders: true,
    legacyHeaders: false
});
const passwordResetLimiter = rateLimit({
    windowMs: 60 * 60 * 1000,
    max: 5,                                                 // 5 reset requests / hour / IP
    message: { error: 'Too many password reset attempts. Please try again in an hour.' },
    standardHeaders: true,
    legacyHeaders: false
});

// MongoDB Connection
const connectDB = async () => {
    try {
        if (process.env.MONGODB_URI) {
            await mongoose.connect(process.env.MONGODB_URI);
            console.log('MongoDB connected successfully');
            // Initialize admin users
            await initializeAdmins();
            // Seed supplier User docs (JABCO/SIMS/CORBET/CPD) + normalize
            // Chemical.sourceSupplier + backfill Chemical.supplierId
            const { initializeSuppliers } = require('./seedSuppliers');
            await initializeSuppliers(User, Chemical);
            // Seed atomic sequence counters from existing records (I8)
            await initializeCounters();
            // Seed initial inventory
            await seedJabcoInventory();
            // Seed March 2026 purchase orders
            await seedMarch2026PurchaseOrders();
            // Backfill inventory from any received POs missing inventory records
            await backfillInventoryFromPOs();
            // Seed Hydrovant inventory (360 gal, paid by Kyle)
            await seedHydrovantInventory();
            // NOTE: seedPOLedgerEntries() disabled - ledger entries now managed by
            // the clean loan migration below (see "CLEAN LEDGER" section)
            // await seedPOLedgerEntries();
            // Sync product catalog with inventory records
            await syncCatalogToInventory();
        } else {
            console.log('No MongoDB URI provided, running without database');
        }
    } catch (error) {
        console.error('MongoDB connection error:', error.message);
        process.exit(1);
    }
};

// ============ MODELS ============

// User Model
const userSchema = new mongoose.Schema({
    name: { type: String, required: true },
    email: { type: String, required: true, unique: true, lowercase: true },
    password: { type: String, required: true },
    phone: String,
    address: {
        street: String,
        city: String,
        state: String,
        zip: String
    },
    role: {
        type: String,
        enum: ['customer', 'admin', 'superadmin', 'supplier', 'distributor'],
        default: 'customer'
    },
    isActive: { type: Boolean, default: true },
    // Supplier-specific fields
    companyName: String, // For suppliers - company/business name
    bidEligible: { type: Boolean, default: true }, // Include in Price Mining bid sheets. false = direct-purchase only (e.g. Corbet/Hydrovant)
    notes: String, // Internal notes about this user (operational hints, preferences, etc.)
    supplierCode: String, // Unique code for supplier (e.g., "CPD", "AGRISTAR")
    representative: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }, // For customers - their rep
    representativeId: String, // kyle, ty, or chad - for quick lookup
    // Stripe Connect for representatives
    stripeAccountId: String, // Connected Stripe account ID
    stripeAccountStatus: { type: String, enum: ['pending', 'active', 'inactive'], default: 'pending' },
    // Stripe Customer for customers (to save payment methods)
    stripeCustomerId: String, // Stripe Customer ID for saving payment methods
    savePaymentMethod: { type: Boolean, default: false }, // User preference to save or not
    // Check payment info for representatives
    checkPayableTo: String,
    checkMailingAddress: {
        street: String,
        city: String,
        state: String,
        zip: String
    },
    farm: {
        name: String,
        acres: Number,
        state: String,
        county: String,
        address: String,
        zip: String
    },
    crops: [String],

    // ============ APPLICATOR COMPLIANCE FIELDS ============

    // Business Classification
    businessType: {
        type: String,
        enum: ['farm', 'commercial_applicator', 'dealer', 'other'],
        default: 'farm'
    },

    // Private Applicator License (for farmers buying RUPs)
    privateApplicatorLicense: {
        hasLicense: { type: Boolean, default: false },
        licenseNumber: String,
        state: String,                    // Issuing state (2-letter)
        expirationDate: Date,
        certificationCategories: [String], // e.g., ['01-Agricultural Plant', '10-Demonstration']
        licenseDocumentUrl: String,       // Uploaded license image/PDF
        verifiedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
        verifiedAt: Date,
        verificationStatus: {
            type: String,
            enum: ['pending', 'verified', 'expired', 'rejected'],
            default: 'pending'
        }
    },

    // Commercial Applicator License (for commercial applicators)
    commercialApplicatorLicense: {
        hasLicense: { type: Boolean, default: false },
        licenseNumber: String,
        state: String,
        businessName: String,
        expirationDate: Date,
        certificationCategories: [String],
        licenseDocumentUrl: String,
        verifiedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
        verifiedAt: Date,
        verificationStatus: {
            type: String,
            enum: ['pending', 'verified', 'expired', 'rejected'],
            default: 'pending'
        }
    },

    // Paraquat Training Certification (EPA-mandated)
    paraquatCertification: {
        completed: { type: Boolean, default: false },
        completionDate: Date,
        expirationDate: Date,            // Valid for 3 years
        certificateNumber: String,
        certificateUrl: String,          // Uploaded certificate
        verifiedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
        verifiedAt: Date
    },

    // Dicamba Training Certification (annual requirement)
    dicambaCertification: {
        completed: { type: Boolean, default: false },
        completionDate: Date,
        trainingYear: Number,            // e.g., 2026 - must be current year
        trainingProvider: String,        // Who provided the training
        certificateUrl: String,
        verifiedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
        verifiedAt: Date
    },

    // Compliance Agreement
    complianceAgreement: {
        agreedToTerms: { type: Boolean, default: false },
        agreementDate: Date,
        agreementVersion: String,        // Track which version they agreed to
        ipAddress: String               // Record IP for legal purposes
    },

    // RUP Purchase Eligibility (calculated field)
    canPurchaseRUP: { type: Boolean, default: false },
    rupEligibilityNotes: String,

    // S1: Force password rotation on first login after admin seeding.
    // Defaults to false so normal signups are unaffected. Set true by the
    // admin seeder (for fresh admins with random temp passwords) and by the
    // startup migration (for existing Farm2026! admins).
    mustChangePassword: { type: Boolean, default: false },

    createdAt: { type: Date, default: Date.now }
});

userSchema.pre('save', async function(next) {
    if (!this.isModified('password')) return next();
    this.password = await bcrypt.hash(this.password, 10);
    next();
});

userSchema.methods.comparePassword = async function(candidatePassword) {
    return bcrypt.compare(candidatePassword, this.password);
};

const User = mongoose.model('User', userSchema);

// Order Model
const orderSchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    representativeId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    crop: { type: String, required: true },
    program: { type: String },
    acres: { type: Number, required: true },
    gpa: { type: Number, default: 10 }, // Gallons per acre
    totalWaterVolume: { type: Number }, // acres * gpa
    year: { type: Number, default: () => new Date().getFullYear() },
    notes: { type: String },

    // Order lines from calculator (enhanced structure)
    orderLines: [{
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        productName: String,
        category: String,
        rate: Number,
        rateUnit: String,
        totalNeeded: Number,
        packageSize: Number,
        packSize: String,
        packageUnit: String,
        packagesNeeded: Number,
        onHandQuantity: Number,
        pricePerPackage: Number,
        lineTotal: Number,
        status: { type: String, enum: ['confirmed', 'needs_quote', 'error'], default: 'confirmed' },
        supplier: String,
        isAutoAdded: { type: Boolean, default: false }
    }],

    // Hydrovant auto-calculation
    hydrovant: {
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        productName: String,
        gallonsNeeded: Number,
        packageSize: Number,
        packagesNeeded: Number,
        pricePerPackage: Number,
        lineTotal: Number,
        status: String
    },

    // Legacy chemicals array (backward compatibility)
    chemicals: [{
        name: String,
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        rate: Number,
        rateUnit: String,
        totalAmount: Number,
        totalUnit: String,
        packageSize: Number,
        packageUnit: String,
        packagesNeeded: Number,
        pricePerPackage: Number,
        totalPrice: Number,
        status: String,
        isAutoAdded: Boolean
    }],
    seeds: [{
        name: String,
        crop: String,
        rate: Number,
        rateUnit: String, // 'seeds/acre' or 'lbs/acre'
        totalAmount: Number,
        bagsNeeded: Number,
        seedsPerBag: Number,
        pricePerBag: Number,
        totalPrice: Number
    }],
    pivotBio: [{
        product: String,
        acres: Number,
        rate: Number,
        totalAmount: Number,
        pricePerUnit: Number,
        totalPrice: Number
    }],
    additionalProducts: {
        hydrovant: { type: Number, default: 0 },
        multiseal: { type: Number, default: 0 },
        pump: { type: Number, default: 0 }
    },
    totalCost: Number,
    costPerAcre: Number,
    // Commission tracking (3-tier pricing)
    repCommission: { type: Number, default: 0 }, // Amount going to rep (sellPrice - adminPrice)
    adminRevenue: { type: Number, default: 0 }, // Acre Profit revenue (adminPrice - costPrice)
    totalCost_cost: { type: Number }, // Total at cost tier
    totalCost_admin: { type: Number }, // Total at admin tier
    // Payment information
    paymentMethod: {
        type: String,
        enum: ['stripe_ach', 'check', 'pending'],
        default: 'pending'
    },
    paymentStatus: {
        type: String,
        enum: ['pending', 'processing', 'paid', 'failed', 'refunded'],
        default: 'pending'
    },
    stripePaymentIntentId: String,
    checkNumber: String,
    checkReceivedDate: Date,
    paidAt: Date,

    // Order status
    status: {
        type: String,
        enum: ['draft', 'submitted', 'confirmed', 'ordered', 'shipped', 'delivered', 'archived', 'cancelled', 'quote_pending', 'quote_sent', 'payment_pending', 'payment_secured'],
        default: 'draft'
    },
    orderStatus: {
        type: String,
        enum: ['pending_quote', 'ready_for_checkout', 'confirmed', 'paid'],
        default: 'pending_quote'
    },

    // Calculator-specific fields
    valorWarning: { type: Boolean, default: false },
    totalConfirmedPrice: { type: Number, default: 0 },

    // Timestamps
    submittedAt: Date,
    quotedAt: Date, // When admin responded to quote request
    quotedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    quoteResponse: String, // Admin's quote notes

    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const Order = mongoose.model('Order', orderSchema);

// Rep Application Model
const repApplicationSchema = new mongoose.Schema({
    firstName: { type: String, required: true },
    lastName: { type: String, required: true },
    email: { type: String, required: true },
    phone: { type: String, required: true },
    city: { type: String, required: true },
    state: { type: String, required: true },
    farming: { type: String, required: true },
    acres: Number,
    experience: { type: String, required: true },
    network: String,
    why: { type: String, required: true },
    status: {
        type: String,
        enum: ['pending', 'approved', 'rejected'],
        default: 'pending'
    },
    createdAt: { type: Date, default: Date.now }
});

const RepApplication = mongoose.model('RepApplication', repApplicationSchema);

// Rep Commission Tracking Model
const repCommissionSchema = new mongoose.Schema({
    repId: { type: String, required: true }, // 'kyle', 'ty', 'chad' or ObjectId for user-reps
    repName: String,
    // Running totals
    totalCommissionEarned: { type: Number, default: 0 },
    totalCommissionPaid: { type: Number, default: 0 },
    commissionBalance: { type: Number, default: 0 }, // Earned - Paid
    // Commission history
    history: [{
        orderId: mongoose.Schema.Types.ObjectId,
        customerId: mongoose.Schema.Types.ObjectId,
        customerName: String,
        orderTotal: Number, // Retail total
        commissionAmount: Number,
        status: { type: String, enum: ['pending', 'paid'], default: 'pending' },
        date: { type: Date, default: Date.now },
        paidDate: Date
    }],
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const RepCommission = mongoose.model('RepCommission', repCommissionSchema);

// Chemical Pricing Model
const chemicalSchema = new mongoose.Schema({
    // Product info
    productName: { type: String, required: true }, // e.g., "Dicamba DMA", "LV 6"
    sourceSupplier: { type: String, required: true }, // Where we buy from: "CPD", "Agri-Star"
    manufacturer: { type: String }, // Who makes it (from label): "Red Eagle", "ADAMA Essentials"
    supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }, // Link to supplier user account

    // Category and crop info
    category: { type: String, enum: ['herbicide', 'fungicide', 'insecticide', 'adjuvant', 'fertilizer', 'other'], default: 'herbicide' },
    crops: [String], // Which crops this can be used on: ['corn', 'soybeans', 'wheat']

    // Packaging
    packSize: { type: String, required: true }, // e.g., "2x2.5", "Shuttle", "4x5", "20"
    unit: { type: String, required: true }, // e.g., "gl" (gallon), "oz", "lb"
    unitsPerPack: { type: Number }, // e.g., 250 for a Shuttle (250 gal)

    // Pricing - 3-tier pricing model with two dollar amount margins
    costPrice: { type: Number, required: true }, // Tier 1: What we pay the supplier (per unit)
    adminMarginDollars: { type: Number, default: 0 }, // Admin margin $ - dollar amount added to cost
    adminPrice: { type: Number }, // Tier 2: Cost + admin margin dollars (per unit)
    marginDollars: { type: Number, default: 0 }, // Rep margin $ - dollar amount added to admin price
    sellPrice: { type: Number, required: true }, // Tier 3: Retail price - what customer pays (per unit)
    priceIsSpeculated: { type: Boolean, default: false }, // true = estimated price, not confirmed by PO
    // Legacy fields (kept for backward compatibility)
    adminMargin: { type: Number, default: 0 }, // Legacy: Admin margin % (no longer used)
    regularMargin: { type: Number, default: 0 }, // Legacy: Regular margin %
    margin: { type: Number }, // Total margin: (sellPrice - costPrice) / sellPrice * 100

    // Application info (for program building)
    defaultRate: { type: Number }, // Default application rate
    rateUnit: { type: String }, // e.g., "oz/acre", "pt/acre", "qt/acre"
    minRate: { type: Number },
    maxRate: { type: Number },

    // Version/date tracking
    priceDate: { type: Date, default: Date.now },
    priceVersion: { type: String }, // Optional identifier like "2026-Q1" or "v1"

    // Comparison/equivalent data
    equivalentProduct: String, // Product name this is equivalent to
    equivalentSupplier: String, // Supplier of equivalent product
    notes: String, // e.g., "Formulation equiv -11%", "Need to get equivalents"

    // ============ REGULATORY COMPLIANCE FIELDS ============

    // EPA Registration (REQUIRED for all pesticides)
    epaRegistrationNumber: String, // e.g., "524-579", "100-1623"

    // Restriction Classification
    isRestrictedUse: { type: Boolean, default: false }, // RUP flag
    rupStates: [String], // States where this is classified as RUP (2-letter codes)

    // Signal Word (EPA mandated - appears on label)
    signalWord: {
        type: String,
        enum: ['DANGER', 'DANGER-POISON', 'WARNING', 'CAUTION', 'NONE'],
        default: 'CAUTION'
    },

    // Hazard Classifications
    hazardClassifications: [{
        type: String,
        enum: [
            'acute_oral_toxicity',
            'acute_dermal_toxicity',
            'acute_inhalation_toxicity',
            'eye_irritant',
            'skin_irritant',
            'skin_sensitizer',
            'carcinogen',
            'reproductive_toxin',
            'environmental_hazard_aquatic',
            'environmental_hazard_bees',
            'groundwater_advisory'
        ]
    }],

    // Required Certifications to Purchase
    requiredCertifications: [{
        type: String,
        enum: [
            'private_applicator',      // State private applicator license
            'commercial_applicator',   // State commercial applicator license
            'paraquat_training',       // EPA-mandated Paraquat training
            'dicamba_training',        // Annual Dicamba OTT training
            'fumigant_training'        // Soil fumigant training
        ]
    }],

    // Safety Data Sheet (SDS)
    sdsUrl: String,        // URL to SDS PDF
    sdsRevisionDate: Date, // Last SDS revision

    // EPA Label
    labelUrl: String,      // URL to EPA-approved label PDF
    labelRevisionDate: Date,

    // State Registrations (pesticides must be registered in each state)
    stateRegistrations: [{
        state: { type: String, maxlength: 2 }, // Two-letter state code
        registrationNumber: String,
        expirationDate: Date,
        isRestricted: { type: Boolean, default: false }, // RUP in this state
        restrictions: String // State-specific restrictions
    }],

    // Active Ingredients (for reporting and compliance)
    activeIngredients: [{
        name: String,              // e.g., "Glyphosate", "Atrazine"
        percentage: Number,        // e.g., 41.0
        poundsPerGallon: Number,   // e.g., 4.17 lb AE/gal
        casNumber: String          // Chemical Abstracts Service number
    }],

    // DOT Transportation / Storage
    dotHazClass: String,           // DOT hazardous materials class (e.g., "6.1", "8")
    unNumber: String,              // UN identification number (e.g., "UN2902")
    packingGroup: String,          // I, II, or III
    storageRequirements: String,   // Special storage instructions
    shelfLifeMonths: Number,       // Product shelf life

    // Manufacturer Information
    manufacturer: String,          // e.g., "BASF", "Bayer", "Syngenta"
    manufacturerAddress: String,
    manufacturerPhone: String,     // Emergency contact
    epaEstablishmentNumber: String, // EPA Est. No. on label

    // Additional Compliance Flags
    requiresApplicatorVerification: { type: Boolean, default: false }, // Must verify license before sale
    requiresAnnualTraining: { type: Boolean, default: false },         // Requires annual training (Dicamba)
    hasBuyerAgreement: { type: Boolean, default: false },              // Requires signed agreement
    isGroundwaterAdvisory: { type: Boolean, default: false },          // Has groundwater advisory
    hasBufferZoneRequirements: { type: Boolean, default: false },      // Has application buffer zones
    bufferZoneDetails: String,

    // Compliance Notes
    complianceNotes: String, // Internal notes about compliance requirements

    // Status
    isActive: { type: Boolean, default: true },
    availableForOrder: { type: Boolean, default: true },

    // Metadata
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

// Calculate prices from dollar margin amounts before save
chemicalSchema.pre('save', function(next) {
    if (this.costPrice) {
        // Use adminPrice if explicitly set, otherwise calculate from dollar margin
        if (!this.adminPrice && this.adminMarginDollars !== undefined) {
            this.adminPrice = Math.round((this.costPrice + (this.adminMarginDollars || 0)) * 100) / 100;
        } else if (!this.adminPrice) {
            // Fallback: admin price = cost price if no margin set
            this.adminPrice = this.costPrice;
        }

        // Calculate adminMarginDollars from adminPrice if not explicitly set
        if (this.adminMarginDollars === undefined || this.adminMarginDollars === null) {
            this.adminMarginDollars = Math.round((this.adminPrice - this.costPrice) * 100) / 100;
        }

        // Calculate marginDollars from sellPrice and adminPrice if not explicitly set
        if ((this.marginDollars === undefined || this.marginDollars === null) && this.sellPrice && this.adminPrice) {
            this.marginDollars = Math.round((this.sellPrice - this.adminPrice) * 100) / 100;
        }

        // Calculate total margin percentage for reference
        if (this.sellPrice && this.sellPrice > 0) {
            this.margin = Math.round(((this.sellPrice - this.costPrice) / this.sellPrice) * 100 * 100) / 100;
        }
    }
    next();
});

// Index for quick lookups
chemicalSchema.index({ productName: 1, sourceSupplier: 1, packSize: 1 });
chemicalSchema.index({ sourceSupplier: 1 });
chemicalSchema.index({ supplierId: 1 });
chemicalSchema.index({ category: 1 });
chemicalSchema.index({ crops: 1 });
chemicalSchema.index({ priceDate: -1 });
// Compliance indexes
chemicalSchema.index({ isRestrictedUse: 1 });
chemicalSchema.index({ epaRegistrationNumber: 1 });
chemicalSchema.index({ 'stateRegistrations.state': 1 });

const Chemical = mongoose.model('Chemical', chemicalSchema);

// Chemical Price History Model (for tracking price changes over time)
const chemicalPriceHistorySchema = new mongoose.Schema({
    chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical', required: true },
    productName: String,
    sourceSupplier: String,
    packSize: String,
    unit: String,
    costPrice: { type: Number },
    sellPrice: { type: Number },
    priceDate: { type: Date, default: Date.now },
    priceVersion: String,
    changedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now }
});

const ChemicalPriceHistory = mongoose.model('ChemicalPriceHistory', chemicalPriceHistorySchema);

// ============ AUDIT LOG MODEL ============
// Immutable record of sensitive actions: impersonation, price changes, role changes, etc.
const auditLogSchema = new mongoose.Schema({
    action: {
        type: String,
        required: true,
        enum: [
            'impersonation_start',
            'impersonation_end',
            'order_placed_as_customer',
            'price_change',
            'margin_change',
            'ledger_entry_edit',
            'password_reset',
            'role_change',
            'customer_created',
            'cash_deposit',
            'check_written',
            'inventory_adjustment',
            'rup_block',
            'license_change',
            'login_failure'
        ]
    },
    performedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    performedByName: String,
    performedByRole: String,
    targetUser: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    targetUserName: String,
    entityType: String,
    entityId: mongoose.Schema.Types.ObjectId,
    entityRef: String,
    before: mongoose.Schema.Types.Mixed,
    after: mongoose.Schema.Types.Mixed,
    amount: Number,
    reason: String,
    ipAddress: String,
    userAgent: String,
    createdAt: { type: Date, default: Date.now, immutable: true }
});
auditLogSchema.index({ performedBy: 1, createdAt: -1 });
auditLogSchema.index({ action: 1, createdAt: -1 });
auditLogSchema.index({ targetUser: 1, createdAt: -1 });
auditLogSchema.index({ entityType: 1, entityId: 1 });

// Enforce createdAt immutability (per Grok Round 3 hardening)
auditLogSchema.pre('save', function(next) {
    if (!this.isNew && this.isModified('createdAt')) {
        return next(new Error('AuditLog.createdAt is immutable'));
    }
    next();
});
// Block updates that would mutate an audit entry
auditLogSchema.pre('findOneAndUpdate', function(next) {
    const update = this.getUpdate() || {};
    if (update.createdAt || (update.$set && update.$set.createdAt)) {
        return next(new Error('AuditLog.createdAt is immutable'));
    }
    next();
});

const AuditLog = mongoose.model('AuditLog', auditLogSchema);

// ============ RUP COMPLIANCE VALIDATOR ============
// Validates that a customer has the required licenses/certifications
// to purchase any restricted-use pesticide in their order.
// Returns { ok: true } if all clear, or { ok: false, errors: [...] } if blocked.
async function validateRupCompliance({ customer, items, allChemicals }) {
    const errors = [];
    const now = new Date();

    // Helper: does the customer have a valid applicator license?
    const hasValidApplicatorLicense = () => {
        const priv = customer?.privateApplicatorLicense;
        const comm = customer?.commercialApplicatorLicense;

        const validPriv = priv?.hasLicense
            && priv?.verificationStatus === 'verified'
            && priv?.expirationDate
            && new Date(priv.expirationDate) > now;

        const validComm = comm?.hasLicense
            && comm?.verificationStatus === 'verified'
            && comm?.expirationDate
            && new Date(comm.expirationDate) > now;

        return validPriv || validComm;
    };

    const licenseExpDate = () => {
        const priv = customer?.privateApplicatorLicense;
        const comm = customer?.commercialApplicatorLicense;
        if (priv?.hasLicense && priv?.expirationDate) return new Date(priv.expirationDate);
        if (comm?.hasLicense && comm?.expirationDate) return new Date(comm.expirationDate);
        return null;
    };

    const hasValidParaquatCert = () => {
        const cert = customer?.paraquatCertification;
        return cert?.completed && cert?.expirationDate && new Date(cert.expirationDate) > now;
    };

    const hasValidDicambaCert = () => {
        const cert = customer?.dicambaCertification;
        return cert?.completed && cert?.expirationDate && new Date(cert.expirationDate) > now;
    };

    for (const item of items) {
        const productName = item.productName || item.name;
        if (!productName) continue;

        const chemical = allChemicals.find(c =>
            c.productName === productName ||
            c.productName?.toLowerCase() === productName.toLowerCase()
        );
        if (!chemical) continue;
        if (!chemical.isRestrictedUse) continue;

        // Restricted-use pesticide - need at minimum a valid applicator license
        if (!hasValidApplicatorLicense()) {
            const exp = licenseExpDate();
            if (exp && exp <= now) {
                errors.push({
                    product: productName,
                    reason: `${productName} is a Restricted Use Pesticide. Your applicator license expired on ${exp.toLocaleDateString()}. Please renew before ordering.`
                });
            } else {
                errors.push({
                    product: productName,
                    reason: `${productName} is a Restricted Use Pesticide. A verified Private or Commercial Applicator License is required. Upload yours at /compliance.html or contact your rep.`
                });
            }
            continue;
        }

        // Product-specific certifications
        const required = chemical.requiredCertifications || [];
        if (required.includes('paraquat_training') && !hasValidParaquatCert()) {
            errors.push({
                product: productName,
                reason: `${productName} contains Paraquat. EPA-mandated Paraquat Training certification (valid for 3 years) is required. Upload yours at /compliance.html.`
            });
        }
        if (required.includes('dicamba_training') && !hasValidDicambaCert()) {
            errors.push({
                product: productName,
                reason: `${productName} contains Dicamba. Annual Dicamba training certification is required. Upload yours at /compliance.html.`
            });
        }
    }

    return errors.length === 0 ? { ok: true } : { ok: false, errors };
}

// Helper: non-blocking audit log writes (failures logged but don't break the audited action)
async function logAudit({ action, req, targetUser, targetUserName, entityType, entityId, entityRef, before, after, amount, reason }) {
    try {
        await AuditLog.create({
            action,
            performedBy: req?.user?._id,
            performedByName: req?.user?.name,
            performedByRole: req?.user?.role,
            targetUser,
            targetUserName,
            entityType,
            entityId,
            entityRef,
            before,
            after,
            amount,
            reason,
            ipAddress: req?.ip || req?.headers?.['x-forwarded-for'],
            userAgent: req?.headers?.['user-agent']
        });
    } catch (e) {
        console.error('Audit log failed (non-blocking):', e.message);
    }
}

// ============ DISTRIBUTOR PRICING MODEL ============
// Allows each distributor to set their own retail prices
// Distributors cannot see wholesale/cost prices - only their retail price
const distributorPricingSchema = new mongoose.Schema({
    distributorId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical', required: true },

    // The retail price this distributor charges their customers
    retailPrice: { type: Number, required: true },

    // Optional markup percentage (for reference)
    markupPercent: { type: Number },

    // Whether this product is available from this distributor
    isAvailable: { type: Boolean, default: true },

    // Notes
    notes: String,

    updatedAt: { type: Date, default: Date.now },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }
});

distributorPricingSchema.index({ distributorId: 1, chemicalId: 1 }, { unique: true });
distributorPricingSchema.index({ distributorId: 1 });
distributorPricingSchema.index({ chemicalId: 1 });

const DistributorPricing = mongoose.model('DistributorPricing', distributorPricingSchema);

// ============ CHEMICAL QUOTE MODEL ============
// For comparing prices from different suppliers for the same product
// Allows tracking quotes over time to find best deals
const chemicalQuoteSchema = new mongoose.Schema({
    // Product identification (normalized name for comparison)
    productName: { type: String, required: true }, // Generic/common name: "Dicamba DMA", "Atrazine 4-L"
    brandName: String, // Brand-specific name if different

    // Supplier info
    supplier: { type: String, required: true }, // "Sims", "CPD", "Agri-Star", etc.
    supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },

    // Packaging
    packSize: { type: String, required: true }, // "265 gal", "2.5 gal", "16 oz"
    unit: { type: String, required: true }, // "gal", "oz", "lb"
    unitsPerPack: Number, // Total units in the pack

    // Pricing
    pricePerUnit: { type: Number, required: true }, // Price per gallon/oz/lb
    packPrice: Number, // Full pack price (calculated)

    // Volume discounts / special pricing
    volumeDiscount: {
        minQuantity: Number, // e.g., "10 totes"
        discountedPrice: Number, // e.g., "$12.75/gal for 10+ totes"
        notes: String // "for only 10 totes and no other product"
    },

    // Quote metadata
    quoteDate: { type: Date, default: Date.now },
    expirationDate: Date, // When quote expires
    quoteReference: String, // Quote number or reference from supplier

    // Status
    isActive: { type: Boolean, default: true },
    isPurchased: { type: Boolean, default: false }, // Did we buy at this price?
    purchaseDate: Date,
    purchaseOrderId: { type: mongoose.Schema.Types.ObjectId, ref: 'PurchaseOrder' },

    // Notes
    notes: String, // "same as Anthem Maxx - no longer making the Maxx"

    // Audit
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

// Indexes for quick lookups and comparisons
chemicalQuoteSchema.index({ productName: 1 });
chemicalQuoteSchema.index({ supplier: 1 });
chemicalQuoteSchema.index({ productName: 1, supplier: 1 });
chemicalQuoteSchema.index({ quoteDate: -1 });
chemicalQuoteSchema.index({ pricePerUnit: 1 });
chemicalQuoteSchema.index({ isActive: 1, productName: 1, pricePerUnit: 1 }); // For finding best active price

// Virtual to calculate savings vs other quotes
chemicalQuoteSchema.virtual('packPriceCalculated').get(function() {
    if (this.pricePerUnit && this.unitsPerPack) {
        return Math.round(this.pricePerUnit * this.unitsPerPack * 100) / 100;
    }
    return this.packPrice;
});

const ChemicalQuote = mongoose.model('ChemicalQuote', chemicalQuoteSchema);

// ============ RUP SALE RECORD MODEL ============
// Required by EPA/state law to maintain records of all Restricted Use Pesticide sales
// Must be kept for minimum 2 years (recommend 3 years)
const rupSaleRecordSchema = new mongoose.Schema({
    // Transaction Info
    saleDate: { type: Date, required: true, default: Date.now },
    orderNumber: String, // Reference to the order
    orderId: { type: mongoose.Schema.Types.ObjectId, ref: 'ChemicalOrder' },

    // Seller Info (Acre Profit / Rep)
    sellerId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    sellerName: String,
    dealerLicenseNumber: String,      // Your pesticide dealer license
    dealerLicenseState: String,

    // Purchaser Info (Customer)
    purchaserId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    purchaserName: { type: String, required: true },
    purchaserAddress: {
        street: String,
        city: String,
        state: String,
        zip: String
    },
    purchaserPhone: String,
    purchaserEmail: String,

    // Applicator License Info (REQUIRED for RUP sales)
    applicatorLicenseType: {
        type: String,
        enum: ['private', 'commercial'],
        required: true
    },
    applicatorLicenseNumber: { type: String, required: true },
    applicatorLicenseState: { type: String, required: true },
    applicatorLicenseExpiration: Date,
    applicatorCertificationCategories: [String],

    // License Verification
    licenseVerificationMethod: {
        type: String,
        enum: ['document_on_file', 'online_verification', 'phone_verification', 'in_person'],
        required: true
    },
    licenseVerifiedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    licenseVerifiedAt: Date,
    licenseDocumentUrl: String,       // Copy of license on file

    // Product Info
    chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
    productName: { type: String, required: true },
    epaRegistrationNumber: { type: String, required: true },
    activeIngredient: String,
    signalWord: String,

    // Quantity
    quantity: { type: Number, required: true },
    unit: String,
    packSize: String,
    totalAmount: Number,              // Total dollar amount

    // Additional Certifications (if required)
    paraquatCertRequired: { type: Boolean, default: false },
    paraquatCertVerified: { type: Boolean, default: false },
    paraquatCertNumber: String,
    paraquatCertDate: Date,

    dicambaCertRequired: { type: Boolean, default: false },
    dicambaCertVerified: { type: Boolean, default: false },
    dicambaCertYear: Number,

    // Buyer Acknowledgement
    buyerAcknowledgement: {
        acknowledged: { type: Boolean, default: false },
        acknowledgedAt: Date,
        ipAddress: String,
        statement: { type: String, default: 'I certify that I am a licensed applicator and will use this product in accordance with the label.' }
    },

    // Compliance Notes
    notes: String,

    // Record Status
    status: {
        type: String,
        enum: ['completed', 'pending_verification', 'cancelled', 'flagged'],
        default: 'completed'
    },

    // Audit Trail
    createdAt: { type: Date, default: Date.now },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedAt: { type: Date, default: Date.now },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }
});

// Indexes for compliance reporting
rupSaleRecordSchema.index({ saleDate: -1 });
rupSaleRecordSchema.index({ purchaserId: 1 });
rupSaleRecordSchema.index({ epaRegistrationNumber: 1 });
rupSaleRecordSchema.index({ applicatorLicenseNumber: 1 });
rupSaleRecordSchema.index({ status: 1 });

const RupSaleRecord = mongoose.model('RupSaleRecord', rupSaleRecordSchema);

// ============ COMPANY SETTINGS MODEL ============
// Stores business settings, licenses, and credentials
// Confidential fields are only accessible to superadmin
const companySettingsSchema = new mongoose.Schema({
    // Company Info (public)
    companyName: { type: String, default: 'AcreProfit, LLC' },
    companyAddress: {
        street: String,
        city: String,
        state: String,
        zip: String
    },
    companyPhone: String,
    companyEmail: String,

    // Colorado Pesticide Dealer License (public fields)
    pesticideDealerLicense: {
        licenseNumber: { type: String }, // CO Dealer ID: 90575
        state: { type: String, default: 'CO' },
        issuedDate: Date,
        expirationDate: Date, // December 31 of current year
        status: {
            type: String,
            enum: ['active', 'expired', 'pending', 'suspended'],
            default: 'active'
        },
        // Show on invoices and to distributors
        displayOnInvoices: { type: Boolean, default: true },
        displayToDistributors: { type: Boolean, default: true }
    },

    // Confidential License Details (superadmin only)
    confidentialLicenseInfo: {
        agLicenseId: String,  // AgLicense ID: 0050BD
        pin: String,         // Pin: 110081
        portalUrl: String,   // https://www.ag.state.co.us/elicense/SecurityLogin.aspx
        portalUsername: String,
        notes: String        // Any admin notes
    },

    // Reminders
    reminders: [{
        title: String,
        description: String,
        dueDate: Date,
        reminderType: {
            type: String,
            enum: ['license_renewal', 'compliance', 'payment', 'other'],
            default: 'other'
        },
        status: {
            type: String,
            enum: ['pending', 'completed', 'dismissed'],
            default: 'pending'
        },
        createdAt: { type: Date, default: Date.now },
        completedAt: Date,
        completedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }
    }],

    updatedAt: { type: Date, default: Date.now },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }
});

const CompanySettings = mongoose.model('CompanySettings', companySettingsSchema);

// Chemical Order Model
const chemicalOrderSchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    representativeId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },

    // Order details
    orderNumber: { type: String, unique: true },
    orderType: { type: String, enum: ['direct', 'program', 'custom'], default: 'direct' },

    // Items ordered
    items: [{
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        productName: String,
        packSize: String,
        unit: String,
        quantity: Number, // Number of packs
        unitPrice: Number, // Price per unit at time of order
        totalPrice: Number,
        timing: String, // Delivery timing (early-spring, pre-plant, etc.)
        isCustom: { type: Boolean, default: false },
        // For program orders
        acres: Number,
        rate: Number,
        rateUnit: String,
        calculatedAmount: Number // Total amount needed before rounding to packs
    }],

    // Program reference (if ordering from a program)
    programId: { type: mongoose.Schema.Types.ObjectId, ref: 'SprayProgram' },
    programName: String,
    totalAcres: Number,

    // Totals
    subtotal: Number,
    discount: { type: Number, default: 0 },
    discountReason: String,
    marginAdjustment: { type: Number, default: 0 },       // Per-unit margin adjustment ($ off or on)
    marginAdjustmentReason: String,                        // Why distributor adjusted margin
    marginAdjustedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }, // Who approved it
    freight: { type: Number, default: 0 },                 // Delivery/freight charge
    deliveryOption: { type: String, enum: ['pickup', 'delivery'], default: 'pickup' },
    deliveryAddress: String,
    processingFee: { type: Number, default: 0 },
    total: Number,

    // Pickup & Year
    pickupLocation: String,
    year: Number,

    // Contact info (for checkout)
    contactInfo: {
        name: String,
        email: String,
        phone: String,
        farm: String
    },

    // Status tracking
    status: {
        type: String,
        enum: ['draft', 'submitted', 'confirmed', 'ordered_from_supplier', 'received', 'ready_for_pickup', 'delivered', 'cancelled', 'archived', 'payment_pending', 'payment_secured'],
        default: 'draft'
    },

    // Payment
    paymentStatus: { type: String, enum: ['pending', 'processing', 'paid', 'partial', 'failed'], default: 'pending' },
    paymentMethod: { type: String, enum: ['ach', 'check', 'card', 'stripe_ach'], default: 'check' },

    // Dates
    submittedAt: Date,
    confirmedAt: Date,
    orderedFromSupplierAt: Date,
    receivedAt: Date,
    deliveredAt: Date,

    // Spray application parameters (from Build a Recipe)
    sprayParams: {
        gallonsPerAcre: Number,
        nozzlePressure: Number, // PSI
        nozzleType: String,     // e.g., AIXR, TTI, XR
        crop: String,
        acres: Number
    },

    // Notes
    customerNotes: String,
    internalNotes: String,

    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

// ============ ATOMIC SEQUENCE COUNTER (I8) ============
// Single Counter collection, one doc per (kind-year) sequence. nextSequence()
// uses findOneAndUpdate/$inc/upsert which is atomic under MongoDB, so two
// concurrent saves can never produce the same number. Replaces the old
// countDocuments()+1 pattern that was race-prone under concurrent order
// creation.

const counterSchema = new mongoose.Schema({
    _id: { type: String, required: true },  // e.g. 'chemicalOrder-2026'
    seq: { type: Number, default: 0 }
});
const Counter = mongoose.model('Counter', counterSchema);

async function nextSequence(name) {
    const result = await Counter.findOneAndUpdate(
        { _id: name },
        { $inc: { seq: 1 } },
        { new: true, upsert: true }
    );
    return result.seq;
}

// Seed counters from existing records. $setOnInsert means this is a no-op on
// redeploy - never stomps a live counter. Safe to run every startup.
// Parses the max numeric suffix from existing XXX-YYYY-NNNNN numbers so the
// first atomic increment after deploy returns a unique next number.
async function initializeCounters() {
    try {
        const year = new Date().getFullYear();

        const seedOne = async (counterName, Model, field, prefix) => {
            const escPrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const last = await Model.findOne({ [field]: { $regex: `^${escPrefix}` } })
                .sort({ [field]: -1 })
                .select(field)
                .lean();
            let seed = 0;
            if (last && last[field]) {
                const parts = last[field].split('-');
                const n = parseInt(parts[parts.length - 1], 10);
                if (!isNaN(n)) seed = n;
            }
            const result = await Counter.findOneAndUpdate(
                { _id: counterName },
                { $setOnInsert: { seq: seed } },
                { upsert: true, new: true }
            );
            console.log(`Counter ${counterName}: seq=${result.seq} (seeded from max=${seed})`);
        };

        await seedOne(`chemicalOrder-${year}`, mongoose.model('ChemicalOrder'), 'orderNumber', `CO-${year}-`);
        await seedOne(`invoice-${year}`, mongoose.model('Invoice'), 'invoiceNumber', `INV-${year}-`);
        await seedOne(`quoteRequest-${year}`, mongoose.model('QuoteRequest'), 'quoteNumber', `QR-${year}-`);
        await seedOne(`supplierBidSheet-${year}`, mongoose.model('SupplierBidSheet'), 'bidNumber', `BID-${year}-`);
    } catch (err) {
        console.error('initializeCounters error:', err.message);
    }
}

// Auto-generate order number
chemicalOrderSchema.pre('save', async function(next) {
    if (!this.orderNumber) {
        const year = new Date().getFullYear();
        const seq = await nextSequence(`chemicalOrder-${year}`);
        this.orderNumber = `CO-${year}-${String(seq).padStart(5, '0')}`;
    }
    next();
});

const ChemicalOrder = mongoose.model('ChemicalOrder', chemicalOrderSchema);

// Ledger Entry Model (Who Owes Who tracking)
const ledgerEntrySchema = new mongoose.Schema({
    representativeId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    date: { type: Date, default: Date.now },
    description: { type: String, required: true },
    amount: { type: Number, required: true }, // Always positive
    // debit = rep owes Acre Profit more, credit = rep's debt decreases
    type: { type: String, enum: ['debit', 'credit'], required: true },
    category: {
        type: String,
        enum: ['order', 'payment', 'supplier_payment', 'commission', 'adjustment', 'refund', 'cash_deposit', 'admin_withdrawal'],
        default: 'adjustment'
    },
    referenceType: { type: String, enum: ['ChemicalOrder', 'Order', 'PurchaseOrder', 'Manual'], default: 'Manual' },
    referenceId: { type: mongoose.Schema.Types.ObjectId },
    // Positive = rep owes Acre Profit, Negative = Acre Profit owes rep
    runningBalance: { type: Number, default: 0 },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    notes: String,
    createdAt: { type: Date, default: Date.now }
});

ledgerEntrySchema.index({ representativeId: 1, date: -1 });
ledgerEntrySchema.index({ referenceType: 1, referenceId: 1 });

const LedgerEntry = mongoose.model('LedgerEntry', ledgerEntrySchema);

// Helper: Create a ledger entry and compute running balance
async function createLedgerEntry({ representativeId, description, amount, type, category, referenceType, referenceId, createdBy, notes, session }) {
    const query = LedgerEntry.findOne({ representativeId }).sort({ date: -1, createdAt: -1 });
    if (session) query.session(session);
    const lastEntry = await query.lean();

    const previousBalance = lastEntry ? lastEntry.runningBalance : 0;
    const balanceChange = type === 'debit' ? amount : -amount;
    const newBalance = Math.round((previousBalance + balanceChange) * 100) / 100;

    const entry = new LedgerEntry({
        representativeId,
        description,
        amount,
        type,
        category: category || 'adjustment',
        referenceType: referenceType || 'Manual',
        referenceId,
        runningBalance: newBalance,
        createdBy,
        notes,
        date: new Date()
    });

    await entry.save({ session });
    return entry;
}

// Spray Program Model (saved custom programs)
// IMPORTANT: These are SUGGESTIONS only - each field requires its own evaluation
const sprayProgramSchema = new mongoose.Schema({
    name: { type: String, required: true }, // e.g., "Round 1 Corn Spray"
    description: String,
    roundNumber: { type: Number }, // Round 1, 2, 3, etc.

    // Program type - NOTE: "suggestion" not "recommendation" (legal)
    type: { type: String, enum: ['suggestion', 'custom', 'template'], default: 'suggestion' },
    isPublic: { type: Boolean, default: false }, // Public programs visible to customers

    // Target crop
    crop: { type: String, required: true }, // corn, soybeans, wheat, etc.

    // Disclaimer - required on all programs
    disclaimer: {
        type: String,
        default: 'This is a suggestion only. Each field requires its own evaluation to determine if this chemical program will work for your specific conditions.'
    },

    // Program passes/applications (can have multiple chemicals per round)
    applications: [{
        name: String, // e.g., "Burndown", "Pre-emergent", "Post-emergent"
        timing: String, // e.g., "14 days before planting", "At planting", "V4-V6"
        deliveryWindow: String, // When product needs to arrive at rep location (e.g., "Late March", "Early May")
        chemicals: [{
            chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
            productName: String,
            suggestedRate: Number, // Use "suggested" not "recommended"
            rateUnit: String, // oz/acre, pt/acre, qt/acre, gal/acre, lb/acre
            packSize: String,
            unit: String,
            isAdjuvant: { type: Boolean, default: false }, // Flag adjuvants for cost and auto-calc handling
            notes: String // e.g., "Adjust based on weed pressure"
        }]
    }],

    // Program-level safety callouts - distinct from the per-chemical-derived
    // rotationRestrictions / grazingRestrictions. Use for agronomic requirements
    // that apply to the program as a whole (seed treatment requirements, timing
    // windows, runoff advisories, etc). Rendered as a bulleted warning list on
    // the program detail view.
    precautions: [String],

    // Cost estimate per acre (calculated)
    estimatedCostPerAcre: Number,

    // Auto-generated restriction fields
    groundType: String, // Description of best ground/soil conditions for this program
    rotationRestrictions: String, // Combined crop rotation restrictions from all chemicals
    grazingRestrictions: String, // Combined grazing/forage restrictions from all chemicals

    // Status
    isActive: { type: Boolean, default: true },

    // Owner
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },

    // Metadata
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const SprayProgram = mongoose.model('SprayProgram', sprayProgramSchema);

// ============ CHEMICAL RESTRICTION DATA ============
// Maps active ingredients/product names to their known restrictions
const chemicalRestrictions = {
    'glyphosate': {
        rotation: 'No crop rotation restrictions',
        grazing: 'Do not graze treated areas for 7 days',
        groundNotes: 'Non-selective - kills all green vegetation on contact'
    },
    'flumioxazin': {
        rotation: 'Sorghum/Milo: 10 months after application at 2 oz/acre with 15+ inches rainfall (corn-to-milo spring rotation is typically fine). Wheat: 4 months. Soybeans: 12 months. Pre-plant to sorghum: 30 days minimum at max 2 oz/acre.',
        grazing: 'Do not graze or harvest forage for 30 days',
        groundNotes: 'Pre-emerge residual broadleaf control. Works best on medium-textured soils. Need rainfall/irrigation to activate.'
    },
    'dicamba': {
        rotation: 'Soybeans/sensitive broadleaf crops: 30 days minimum',
        grazing: 'Do not graze treated areas for 7 days',
        groundNotes: 'Broadleaf systemic herbicide. Watch for drift to sensitive crops'
    },
    'atrazine': {
        rotation: 'Soybeans: 12 months. Small grains: next season OK if under 1 lb ai/acre',
        grazing: 'Do not graze sorghum forage for 60 days. Corn silage: 21 days',
        groundNotes: 'Corn/sorghum residual. Restricted Use Pesticide (ground/surface water)'
    },
    'mesotrione': {
        rotation: 'Wheat: 10 months. Soybeans: 18 months. Sorghum: 18 months',
        grazing: 'Do not graze treated areas for 45 days',
        groundNotes: 'Group 27 bleaching herbicide for corn. Post-emerge broadleaf and grass control'
    },
    'sulfentrazone': {
        rotation: 'Wheat: 4 months. Corn: 18 months. Sorghum: 18 months. Soybeans: 12 months',
        grazing: 'Do not graze treated areas for 28 days',
        groundNotes: 'Pre-emerge residual. Avoid sandy soils or pH above 7.5'
    },
    'metribuzin': {
        rotation: 'Sorghum/oats: 12 months. Follow label for specific crops',
        grazing: 'Do not graze or harvest within 28 days',
        groundNotes: 'Group 5 herbicide. Effective on broadleaf weeds and some grasses'
    },
    '2,4-d': {
        rotation: 'No crop rotation restrictions',
        grazing: 'Do not graze dairy cattle for 7 days. Meat animals: no restriction after 3 days',
        groundNotes: 'Post-emerge broadleaf herbicide for wheat, corn, sorghum, pasture'
    },
    'metsulfuron': {
        rotation: 'Corn, sorghum, grass crops: 60 days. Wheat: no restriction',
        grazing: 'Do not graze or cut for hay within 28 days',
        groundNotes: 'Low-rate broadleaf control for wheat and pasture'
    },
    'pyroxasulfone': {
        rotation: 'Check label for specific crops. Generally 12-18 months for non-labeled crops',
        grazing: 'Do not graze treated areas for 30 days',
        groundNotes: 'Group 15 residual grass herbicide. Extended pre-emerge control'
    },
    'hydrovant': {
        rotation: 'No restrictions (adjuvant)',
        grazing: 'No restrictions (adjuvant)',
        groundNotes: 'Activator-sticker adjuvant - improves coverage and uptake'
    }
};

// Map product names to their active ingredient keys in chemicalRestrictions
function matchChemicalToRestriction(productName) {
    if (!productName) return null;
    const name = productName.toLowerCase();

    // Direct active ingredient matches
    if (name.includes('glyphosate') || name.includes('xsate') || name.includes('glystar')) return 'glyphosate';
    if (name.includes('flumioxazin') || name.includes('valor')) return 'flumioxazin';
    if (name.includes('dicamba')) return 'dicamba';
    if (name.includes('atrazine')) return 'atrazine';
    if (name.includes('mesotrione') || name.includes('meso 4sc') || name.includes('meso ')) return 'mesotrione';
    if (name.includes('sulfentrazone')) return 'sulfentrazone';
    if (name.includes('metribuzin') || name.includes('rancor')) return 'metribuzin';
    if (name.includes('2,4-d') || name.includes('2,4d') || name.includes('lv 6') || name.includes('lv-6') || name.includes('defy lv')) return '2,4-d';
    if (name.includes('metsulfuron') || name.includes('mivum')) return 'metsulfuron';
    if (name.includes('pyroxasulfone') || name.includes('anthem')) return 'pyroxasulfone';
    if (name.includes('hydrovant')) return 'hydrovant';

    return null;
}

// Extract the numeric day count from a grazing restriction string
function extractGrazingDays(grazingStr) {
    if (!grazingStr) return 0;
    const matches = grazingStr.match(/(\d+)\s*days?/gi);
    if (!matches) return 0;
    let maxDays = 0;
    for (const m of matches) {
        const num = parseInt(m);
        if (num > maxDays) maxDays = num;
    }
    return maxDays;
}

// Auto-generate restrictions from a list of chemicals
// chemicals: array of objects with productName (or name) field
function generateRecipeRestrictions(chemicals) {
    if (!chemicals || chemicals.length === 0) {
        return { groundType: null, rotationRestrictions: null, grazingRestrictions: null };
    }

    const rotationParts = [];
    const grazingEntries = [];
    const groundNotes = [];
    const seen = new Set();

    for (const chem of chemicals) {
        const productName = chem.productName || chem.name || '';
        const key = matchChemicalToRestriction(productName);
        if (!key || seen.has(key)) continue;
        seen.add(key);

        const data = chemicalRestrictions[key];
        if (!data) continue;

        // Skip adjuvants for rotation/grazing (they have no real restrictions)
        const isAdjuvant = key === 'hydrovant';

        if (!isAdjuvant && data.rotation && !data.rotation.toLowerCase().includes('no crop rotation restrictions') && !data.rotation.toLowerCase().includes('no restrictions')) {
            rotationParts.push(`${productName}: ${data.rotation}`);
        }

        if (!isAdjuvant && data.grazing && !data.grazing.toLowerCase().includes('no restrictions')) {
            grazingEntries.push({ text: data.grazing, days: extractGrazingDays(data.grazing), productName });
        }

        if (data.groundNotes && !isAdjuvant) {
            groundNotes.push(data.groundNotes);
        }
    }

    // Build rotation restrictions - combine all
    const rotationRestrictions = rotationParts.length > 0
        ? rotationParts.join('. ')
        : 'No specific crop rotation restrictions for this mix';

    // Build grazing restrictions - use the most restrictive (longest wait)
    let grazingRestrictions;
    if (grazingEntries.length > 0) {
        // Sort by days descending to get the most restrictive first
        grazingEntries.sort((a, b) => b.days - a.days);
        const mostRestrictive = grazingEntries[0];
        if (grazingEntries.length === 1) {
            grazingRestrictions = mostRestrictive.text;
        } else {
            // Show the most restrictive, then note other restrictions
            const otherNotes = grazingEntries.slice(1)
                .filter(e => e.days > 0 && e.days !== mostRestrictive.days)
                .map(e => `${e.productName}: ${e.days} days`);
            grazingRestrictions = mostRestrictive.text;
            if (otherNotes.length > 0) {
                grazingRestrictions += '. Other components: ' + otherNotes.join('; ');
            }
        }
    } else {
        grazingRestrictions = 'No specific grazing restrictions for this mix';
    }

    // Build ground type description from combined notes
    const groundType = groundNotes.length > 0
        ? groundNotes.join('. ')
        : null;

    return { groundType, rotationRestrictions, grazingRestrictions };
}

// Merch Order Model - REMOVED (Printify integration was never completed)

// Purchase Order Model (Orders FROM suppliers - what Acre Profit buys)
const purchaseOrderSchema = new mongoose.Schema({
    // Auto-generated PO number: PO-2026-00001
    poNumber: { type: String, unique: true },

    // Supplier information
    supplier: {
        name: { type: String, required: true },
        contact: String,
        phone: String,
        email: String
    },

    // Line items - products ordered with quantities and pricing
    items: [{
        productName: { type: String, required: true },
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        description: String,
        packSize: String,
        unit: String, // gal, case, bag, etc.

        // Quantity and pricing (prices can be edited)
        quantityOrdered: { type: Number, required: true },
        pricePerUnit: { type: Number, required: true }, // Cost price from supplier
        totalPrice: Number, // Calculated: quantity * price

        // For tracking receives and splits
        quantityReceived: { type: Number, default: 0 }, // Total received so far
        quantityAllocated: { type: Number, default: 0 }, // Total allocated to distributors
        quantityRemaining: Number // Calculated: ordered - allocated
    }],

    // Totals
    subtotal: Number,
    freight: { type: Number, default: 0 },
    otherFees: { type: Number, default: 0 },
    superAdminFee: { type: Number, default: 0 }, // Admin markup/fee - only visible to superadmins
    totalCost: Number,

    // Status workflow
    status: {
        type: String,
        enum: ['draft', 'submitted', 'confirmed', 'partial_received', 'received', 'closed', 'cancelled'],
        default: 'draft'
    },

    // Payment tracking
    paidBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }, // Who paid the supplier
    paidDate: Date,
    paymentStatus: {
        type: String,
        enum: ['unpaid', 'paid', 'partial'],
        default: 'unpaid'
    },
    paymentMethod: String, // 'check', 'ach', 'wire', etc.
    checkNumber: String,

    // Dates
    orderDate: { type: Date, default: Date.now },
    expectedDeliveryDate: Date,
    receivedDate: Date,

    // Shipping/delivery info
    deliveryLocation: String,
    bolNumber: String, // Bill of Lading
    trackingInfo: String,

    // Notes
    notes: String,
    internalNotes: String,

    // Attached documents (invoices, receipts, BOLs)
    documents: [{
        fileName: String,           // Original filename
        storedName: String,         // Filename in storage (PO-2026-00001-invoice.pdf)
        fileType: String,           // MIME type
        fileSize: Number,           // Size in bytes
        documentType: {             // Type of document
            type: String,
            enum: ['invoice', 'receipt', 'bol', 'packing_slip', 'other'],
            default: 'invoice'
        },
        uploadedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
        uploadedAt: { type: Date, default: Date.now },
        notes: String
    }],

    // Audit trail
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

purchaseOrderSchema.index({ poNumber: 1 });
purchaseOrderSchema.index({ 'supplier.name': 1 });
purchaseOrderSchema.index({ status: 1 });
purchaseOrderSchema.index({ orderDate: -1 });

const PurchaseOrder = mongoose.model('PurchaseOrder', purchaseOrderSchema);

// Supplier Model - Save supplier information for reuse
// Suppliers live in the User collection with role:'supplier'. The legacy
// Supplier/supplierSchema model (previously defined here) was removed along
// with Block B's dead routes. See the seeder at server/seedSuppliers.js.

// Quote Request Model - Customers submit quantity needed, admin provides pricing
const quoteRequestSchema = new mongoose.Schema({
    // Auto-generated quote number: QR-2026-00001
    quoteNumber: { type: String, unique: true },

    // Customer info
    customerId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    customerName: String,
    customerEmail: String,
    customerPhone: String,

    // Representative (if any)
    representativeId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },

    // Spray program context
    crop: String,              // corn, wheat, soybeans, sorghum, fallow, etc.
    timing: String,            // fall, early-spring, pre-emerge, post-emerge, post-harvest
    acres: Number,             // How many acres for this program
    gallonsPerAcre: { type: Number, default: 15 },

    // Items requested - customer specifies product and quantity needed
    items: [{
        productName: { type: String, required: true },
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        category: String,
        packSize: String,
        unit: String,
        quantityNeeded: { type: Number, required: true },

        // Admin fills in pricing after quote review
        costPrice: Number,        // Supplier cost (admin enters)
        adminPrice: Number,       // Admin price
        sellPrice: Number,        // Customer price (what they'll pay)
        totalPrice: Number,       // sellPrice * quantityNeeded
        priceNotes: String,       // Notes about pricing (e.g., "bulk discount applied")
        isPriced: { type: Boolean, default: false }
    }],

    // Totals (calculated after pricing)
    estimatedTotal: Number,

    // Status workflow
    status: {
        type: String,
        enum: ['submitted', 'pricing', 'quoted', 'accepted', 'declined', 'expired', 'converted'],
        default: 'submitted'
    },

    // Important dates
    submittedAt: { type: Date, default: Date.now },
    pricedAt: Date,           // When admin added pricing
    quotedAt: Date,           // When quote was sent to customer
    expiresAt: Date,          // Quote expiration (e.g., 7 days from quoted)
    respondedAt: Date,        // When customer accepted/declined

    // If converted to order
    convertedToOrderId: { type: mongoose.Schema.Types.ObjectId, ref: 'Order' },
    convertedAt: Date,

    // Notes
    customerNotes: String,    // Notes from customer about their request
    adminNotes: String,       // Internal admin notes

    // Delivery info for quote
    deliveryLocation: String,
    preferredDeliveryDate: Date,

    // Metadata
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

// Auto-generate quote number
quoteRequestSchema.pre('save', async function(next) {
    if (!this.quoteNumber) {
        const year = new Date().getFullYear();
        const seq = await nextSequence(`quoteRequest-${year}`);
        this.quoteNumber = `QR-${year}-${String(seq).padStart(5, '0')}`;
    }
    next();
});

quoteRequestSchema.index({ customerId: 1 });
quoteRequestSchema.index({ representativeId: 1 });
quoteRequestSchema.index({ status: 1 });
quoteRequestSchema.index({ submittedAt: -1 });

const QuoteRequest = mongoose.model('QuoteRequest', quoteRequestSchema);

// Supplier Bid Sheet Model - Send quantity needs to suppliers, compare prices
const supplierBidSheetSchema = new mongoose.Schema({
    // Auto-generated bid number: BID-2026-00001
    bidNumber: { type: String, unique: true },

    // Title/description for this bid
    title: { type: String, required: true },
    description: String,

    // Items we need quotes for
    items: [{
        productName: { type: String, required: true },
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        category: String,
        packSize: String,
        unit: String,
        quantityNeeded: { type: Number, required: true },
        notes: String // Any special requirements
    }],

    // Suppliers invited to bid
    invitedSuppliers: [{
        supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
        supplierName: String,
        contactEmail: String,
        contactPhone: String,
        invitedAt: Date,
        emailedAt: Date,              // Set when bid invitation email actually sent
        lastEmailError: String,       // Populated on send failure or placeholder-email skip
        status: {
            type: String,
            enum: ['invited', 'viewed', 'responded', 'declined', 'no_response'],
            default: 'invited'
        }
    }],

    // Supplier responses/bids
    supplierBids: [{
        supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
        supplierName: String,
        receivedAt: { type: Date, default: Date.now },

        // Their pricing for each item
        itemPricing: [{
            productName: String,
            itemIndex: Number, // Reference to items array
            pricePerUnit: Number,
            totalPrice: Number,
            availableQuantity: Number, // How much they can supply
            leadTimeDays: Number,
            notes: String
        }],

        // Totals
        subtotal: Number,
        freight: Number,
        totalBid: Number,

        // Validity
        validUntil: Date,
        paymentTerms: String,
        deliveryTerms: String,
        bidNotes: String,

        // Selection
        isSelected: { type: Boolean, default: false },
        selectedAt: Date
    }],

    // Status workflow
    status: {
        type: String,
        enum: ['draft', 'sent', 'responses_received', 'evaluating', 'awarded', 'po_created', 'cancelled', 'expired'],
        default: 'draft'
    },

    // Dates
    createdAt: { type: Date, default: Date.now },
    sentAt: Date,
    responseDueDate: Date,
    awardedAt: Date,
    expiresAt: Date,

    // If converted to PO
    awardedSupplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    awardedSupplierName: String,
    purchaseOrderId: { type: mongoose.Schema.Types.ObjectId, ref: 'PurchaseOrder' },

    // Metadata
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    updatedAt: { type: Date, default: Date.now }
});

// Auto-generate bid number
supplierBidSheetSchema.pre('save', async function(next) {
    if (!this.bidNumber) {
        const year = new Date().getFullYear();
        const seq = await nextSequence(`supplierBidSheet-${year}`);
        this.bidNumber = `BID-${year}-${String(seq).padStart(5, '0')}`;
    }
    next();
});

supplierBidSheetSchema.index({ status: 1 });
supplierBidSheetSchema.index({ createdAt: -1 });
supplierBidSheetSchema.index({ 'invitedSuppliers.supplierId': 1 });

const SupplierBidSheet = mongoose.model('SupplierBidSheet', supplierBidSheetSchema);

// Purchase Order Split Model (How a PO is split between distributors)
const purchaseOrderSplitSchema = new mongoose.Schema({
    // Link to parent purchase order
    purchaseOrderId: { type: mongoose.Schema.Types.ObjectId, ref: 'PurchaseOrder', required: true },
    poNumber: String, // Denormalized for quick display

    // Distributor receiving this portion
    distributorId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    distributorName: String, // Denormalized

    // Split number for this PO (e.g., PO-2026-00001-A, PO-2026-00001-B)
    splitCode: String, // A, B, C, etc.

    // Items allocated to this distributor
    items: [{
        productName: String,
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        packSize: String,
        unit: String,

        // Allocation
        quantityAllocated: { type: Number, required: true },
        pricePerUnit: Number, // May be different if distributor pricing differs
        totalPrice: Number,

        // Reference to original PO item index
        originalItemIndex: Number
    }],

    // Totals for this split
    subtotal: Number,
    freightAllocation: { type: Number, default: 0 },
    totalCost: Number,

    // Status
    status: {
        type: String,
        enum: ['allocated', 'shipped', 'delivered', 'invoiced', 'paid'],
        default: 'allocated'
    },

    // Delivery tracking
    deliveryDate: Date,
    deliveryLocation: String,
    receivedBy: String,

    // Notes
    notes: String,

    // Audit
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

purchaseOrderSplitSchema.index({ purchaseOrderId: 1 });
purchaseOrderSplitSchema.index({ distributorId: 1 });
purchaseOrderSplitSchema.index({ status: 1 });

const PurchaseOrderSplit = mongoose.model('PurchaseOrderSplit', purchaseOrderSplitSchema);

// ============ INVENTORY MODEL ============
// Tracks actual stock levels by product and location
const inventorySchema = new mongoose.Schema({
    // Product reference
    chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical', required: true },
    productName: { type: String, required: true },
    packSize: String,
    unit: String, // gal, case, bag, lb, oz

    // Stock levels
    quantityOnHand: { type: Number, default: 0 }, // Current available stock
    quantityReserved: { type: Number, default: 0 }, // Reserved for pending orders
    quantityAvailable: { type: Number, default: 0 }, // onHand - reserved

    // Reorder tracking
    reorderPoint: { type: Number, default: 0 }, // Alert when stock falls below
    reorderQuantity: { type: Number, default: 0 }, // Suggested reorder amount

    // Location (for multi-warehouse)
    location: { type: String, default: 'main' }, // main, haxtun, otis, etc.
    distributorId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' }, // Which distributor holds this

    // Cost tracking (FIFO/average cost)
    averageCost: { type: Number, default: 0 }, // Weighted average cost
    lastCost: { type: Number, default: 0 }, // Most recent purchase cost

    // Timestamps
    lastReceivedDate: Date,
    lastSoldDate: Date,

    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

inventorySchema.index({ chemicalId: 1, location: 1 }, { unique: true });
inventorySchema.index({ productName: 1 });
inventorySchema.index({ distributorId: 1 });
inventorySchema.index({ quantityOnHand: 1 });

const Inventory = mongoose.model('Inventory', inventorySchema);

// ============ INVENTORY TRANSACTION MODEL ============
// Audit trail for all inventory movements
const inventoryTransactionSchema = new mongoose.Schema({
    inventoryId: { type: mongoose.Schema.Types.ObjectId, ref: 'Inventory', required: true },
    chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
    productName: String,

    // Transaction type
    type: {
        type: String,
        enum: ['receive', 'sale', 'adjustment', 'transfer', 'return', 'damage', 'expired', 'reserve', 'release'],
        required: true
    },

    // Quantity change (positive for additions, negative for removals)
    quantityChange: { type: Number, required: true },
    previousQuantity: Number,
    newQuantity: Number,

    // Cost info
    unitCost: Number,
    totalCost: Number,

    // Reference documents
    referenceType: { type: String, enum: ['PurchaseOrder', 'ChemicalOrder', 'Manual', 'Transfer'] },
    referenceId: { type: mongoose.Schema.Types.ObjectId },
    referenceNumber: String, // PO number, Order number, etc.

    // Location
    location: String,
    fromLocation: String, // For transfers
    toLocation: String, // For transfers

    // Notes
    notes: String,
    reason: String, // For adjustments

    // Audit
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now }
});

inventoryTransactionSchema.index({ inventoryId: 1, createdAt: -1 });
inventoryTransactionSchema.index({ referenceType: 1, referenceId: 1 });
inventoryTransactionSchema.index({ type: 1 });
inventoryTransactionSchema.index({ createdAt: -1 });

const InventoryTransaction = mongoose.model('InventoryTransaction', inventoryTransactionSchema);

// ============ INVENTORY BATCH/LOT MODEL ============
// Tracks inventory by PO - each batch has its own cost and quantity
const inventoryBatchSchema = new mongoose.Schema({
    // Product reference
    chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical', required: true },
    productName: { type: String, required: true },
    packSize: String,
    unit: String,

    // Batch identification - PO is the primary identifier
    poNumber: { type: String, required: true }, // e.g., "JABCO-2026-001"
    poId: { type: mongoose.Schema.Types.ObjectId, ref: 'PurchaseOrder' },
    lotNumber: String, // Supplier's lot number from label (e.g., "05826M254M")

    // Quantities (in base units - gallons, lbs, etc.)
    quantityReceived: { type: Number, required: true }, // Original amount from this PO
    quantityRemaining: { type: Number, required: true }, // What's left to sell
    quantitySold: { type: Number, default: 0 }, // Sold from this batch
    quantityAdjusted: { type: Number, default: 0 }, // Manual adjustments (+/-)

    // Cost for THIS specific batch
    costPerUnit: { type: Number, required: true }, // What we paid per unit for this PO
    totalCost: { type: Number }, // Total cost for this batch

    // Location
    location: { type: String, default: 'main' },

    // Status
    status: {
        type: String,
        enum: ['active', 'depleted', 'expired', 'damaged', 'returned'],
        default: 'active'
    },

    // Dates
    receivedDate: { type: Date, default: Date.now },
    expirationDate: Date,

    // Supplier info
    supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    supplierName: String,

    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

inventoryBatchSchema.index({ chemicalId: 1, status: 1 });
inventoryBatchSchema.index({ poNumber: 1 });
inventoryBatchSchema.index({ productName: 1 });
inventoryBatchSchema.index({ status: 1, quantityRemaining: 1 });

const InventoryBatch = mongoose.model('InventoryBatch', inventoryBatchSchema);

// ============ CUSTOMER INVOICE MODEL ============
// Generated invoices for customer orders
const invoiceSchema = new mongoose.Schema({
    // Invoice number: INV-2026-00001
    invoiceNumber: { type: String, unique: true },

    // Customer info
    customerId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    customerName: String,
    customerEmail: String,
    customerPhone: String,
    customerAddress: {
        street: String,
        city: String,
        state: String,
        zip: String
    },

    // Order reference
    orderId: { type: mongoose.Schema.Types.ObjectId, ref: 'ChemicalOrder' },
    orderNumber: String,

    // Representative
    representativeId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    representativeName: String,

    // Line items
    items: [{
        productName: String,
        description: String,
        packSize: String,
        unit: String,
        unitsPerPack: Number,
        quantity: Number,     // Total units (e.g., 265 gal)
        packQuantity: Number, // Number of packs ordered (e.g., 1 tote)
        unitPrice: Number,    // Sell price per unit (customer pays)
        costPrice: Number,    // Cost price per unit (what we paid) - admin only
        adminPrice: Number,   // Admin price per unit - admin only
        totalPrice: Number,   // quantity * unitPrice
        margin: Number        // (unitPrice - costPrice) * quantity - admin only
    }],

    // Totals
    subtotal: Number,
    discount: { type: Number, default: 0 },
    discountReason: String,
    tax: { type: Number, default: 0 },
    shipping: { type: Number, default: 0 },
    total: Number,

    // Payment tracking
    amountPaid: { type: Number, default: 0 },
    amountDue: Number,
    paymentStatus: {
        type: String,
        enum: ['unpaid', 'partial', 'paid', 'refunded'],
        default: 'unpaid'
    },
    paymentMethod: String,
    paymentDate: Date,
    stripePaymentIntentId: String,

    // Dates
    invoiceDate: { type: Date, default: Date.now },
    dueDate: Date,

    // Status
    status: {
        type: String,
        enum: ['draft', 'sent', 'viewed', 'paid', 'overdue', 'cancelled'],
        default: 'draft'
    },

    // Delivery info
    deliveryStatus: {
        type: String,
        enum: ['pending', 'ready', 'shipped', 'delivered', 'signed'],
        default: 'pending'
    },
    deliveryDate: Date,
    deliveryLocation: String,
    deliverySignature: String, // Base64 signature image
    deliverySignedBy: String,
    deliverySignedAt: Date,
    deliveryNotes: String,

    // Notes
    notes: String,
    internalNotes: String,

    // Audit
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    sentAt: Date,
    sentBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

invoiceSchema.index({ invoiceNumber: 1 });
invoiceSchema.index({ customerId: 1 });
invoiceSchema.index({ orderId: 1 });
invoiceSchema.index({ status: 1 });
invoiceSchema.index({ invoiceDate: -1 });

const Invoice = mongoose.model('Invoice', invoiceSchema);

// ============ PROGRAM QUOTE MODEL ============
// Snapshot of a spray program priced for a specific customer at a specific
// acreage. Sent via email as a conversation-starter — farmer reviews, logs
// in to build their own order OR replies and admin clicks "Create Invoice
// & Collect Payment" on the customer card. Distinct from QuoteRequest
// (price-mining off-catalog requests) and PriceMiningQuote (supplier bids).
const programQuoteSchema = new mongoose.Schema({
    quoteNumber: { type: String, unique: true }, // e.g., PQ-2026-00001
    customerId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    representativeId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },

    programId: { type: mongoose.Schema.Types.ObjectId, ref: 'SprayProgram' },
    programName: String,
    crop: String,
    totalAcres: Number,

    // Snapshot of the items at quote time. Prices locked here so the farmer
    // sees the same numbers in the email and on the customer card even if
    // the catalog moves. Convert-to-Invoice re-reads these directly.
    items: [{
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        productName: String,
        packSize: String,
        unit: String,
        quantity: Number,
        unitPrice: Number,
        totalPrice: Number,
        acres: Number,
        rate: Number,
        rateUnit: String,
        calculatedAmount: Number
    }],

    subtotal: Number,
    total: Number,
    costPerAcre: Number,

    sprayParams: {
        gallonsPerAcre: Number,
        tankSize: Number
    },

    status: {
        type: String,
        enum: ['draft', 'sent', 'converted', 'expired', 'declined'],
        default: 'draft'
    },
    sentAt: Date,
    convertedAt: Date,
    convertedToInvoiceId: { type: mongoose.Schema.Types.ObjectId, ref: 'Invoice' },
    // Expiry = createdAt + 30 days unless overridden. Status transitions to
    // 'expired' are a separate concern (cron/on-read check); not enforced here.
    expiresAt: Date,

    notes: String,

    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

programQuoteSchema.index({ quoteNumber: 1 });
programQuoteSchema.index({ customerId: 1 });
programQuoteSchema.index({ status: 1 });
programQuoteSchema.index({ createdAt: -1 });

const ProgramQuote = mongoose.model('ProgramQuote', programQuoteSchema);

// ============ CHEMICAL MIX RECIPE BOOK ============

// Chemical Mix Model - "Recipe Book" for custom chemical cocktails
const chemicalMixSchema = new mongoose.Schema({
    // Mix identification
    name: { type: String, required: true }, // "Dad's Bean Blend", "Kyle's Burndown Special"
    slug: { type: String, unique: true }, // URL-friendly version of name

    // Creator info
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    isAnonymous: { type: Boolean, default: true }, // Hide creator identity
    creatorDisplayName: String, // Optional: show this name instead (e.g., "McConnell Farms")

    // Crop and timing
    crop: { type: String, required: true }, // corn, soybeans, wheat, etc.
    timing: {
        type: String,
        enum: [
            'burndown',           // Pre-plant burndown
            'pre-emerge',         // After planting, before emergence
            'early-post',         // V1-V3 / VC-V2 beans
            'post-emerge',        // General post-emerge
            'v4-v6',              // Corn V4-V6 window
            'v6-plus',            // Corn V6+ / late post
            'r1-r3',              // Reproductive stages (beans)
            'layby',              // Last application before canopy
            'tassel',             // At/around tasseling
            'harvest-aid',        // Pre-harvest / desiccant
            'fall-application',   // Post-harvest fall app
            'cover-crop',         // Cover crop termination
            'other'
        ],
        required: true
    },
    maxGrowthStage: String, // e.g., "V6", "R2" - max crop height/stage for this mix
    timingNotes: String, // e.g., "Apply before 12 inch weeds", "Best in morning"

    // The recipe - ingredients list
    ingredients: [{
        chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
        productName: String, // Stored for reference even if chemical deleted
        category: String, // herbicide, fungicide, etc.
        rate: { type: Number, required: true }, // Application rate
        rateUnit: { type: String, required: true }, // oz/acre, pt/acre, qt/acre, lb/acre
        notes: String // e.g., "Can sub with generic glyphosate"
    }],

    // Total water/carrier volume
    gallonsPerAcre: { type: Number, default: 15 }, // GPA for the mix

    // Story and description
    description: String, // Brief description of what it does
    story: String, // The full story - why they developed it, how it came to be
    tips: String, // Application tips, what to watch for

    // Media
    images: [{
        url: String,
        caption: String,
        uploadedAt: { type: Date, default: Date.now }
    }],

    // Yield map / success data
    yieldData: [{
        year: Number,
        crop: String,
        acres: Number,
        yieldPerAcre: Number, // bu/acre, tons/acre, etc.
        yieldUnit: { type: String, default: 'bu/acre' },
        location: String, // General area, county, etc.
        notes: String,
        imageUrl: String // Yield map image
    }],

    // AI Analysis (from Claude)
    aiAnalysis: {
        lastAnalyzedAt: Date,
        compatibility: String, // Tank mix compatibility notes
        restrictions: String, // Label restrictions summary
        rotationalCrops: String, // Rotational restrictions
        bufferZones: String, // Buffer zone requirements
        ppeRequired: String, // PPE requirements
        modesOfAction: [String], // List of MOA groups
        resistanceManagement: String, // Resistance management notes
        warnings: [String], // Any warnings or cautions
        fullAnalysis: String // Complete analysis text
    },

    // Auto-generated restriction fields
    groundType: String, // Description of best ground/soil conditions for this mix
    rotationRestrictions: String, // Combined crop rotation restrictions from all chemicals
    grazingRestrictions: String, // Combined grazing/forage restrictions from all chemicals

    // Ratings and popularity
    totalRatings: { type: Number, default: 0 },
    averageRating: { type: Number, default: 0 },
    totalViews: { type: Number, default: 0 },
    totalOrders: { type: Number, default: 0 }, // Times ingredients were ordered

    // Tags for searchability
    tags: [String], // e.g., ["waterhemp", "marestail", "burndown", "no-till"]

    // Status
    status: {
        type: String,
        enum: ['draft', 'published', 'archived'],
        default: 'draft'
    },
    isPublic: { type: Boolean, default: true }, // Visible to other farmers
    isFeatured: { type: Boolean, default: false }, // Admin can feature top recipes

    // Timestamps
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now },
    publishedAt: Date
});

// Indexes for searching
chemicalMixSchema.index({ crop: 1, timing: 1 });
chemicalMixSchema.index({ status: 1, isPublic: 1 });
chemicalMixSchema.index({ averageRating: -1 });
chemicalMixSchema.index({ totalOrders: -1 });
chemicalMixSchema.index({ tags: 1 });
chemicalMixSchema.index({ createdBy: 1 });
chemicalMixSchema.index({ slug: 1 });
chemicalMixSchema.index({ name: 'text', description: 'text', story: 'text', tags: 'text' });

// Generate slug before save
chemicalMixSchema.pre('save', function(next) {
    if (this.isModified('name') || !this.slug) {
        // Create URL-friendly slug
        let slug = this.name
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, '-')
            .replace(/^-|-$/g, '');
        // Add random suffix for uniqueness
        slug = `${slug}-${Math.random().toString(36).substring(2, 8)}`;
        this.slug = slug;
    }
    this.updatedAt = new Date();
    next();
});

const ChemicalMix = mongoose.model('ChemicalMix', chemicalMixSchema);

// Chemical Mix Rating Model
const chemicalMixRatingSchema = new mongoose.Schema({
    mixId: { type: mongoose.Schema.Types.ObjectId, ref: 'ChemicalMix', required: true },
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },

    // Rating (1-5 stars)
    rating: { type: Number, required: true, min: 1, max: 5 },

    // Optional review
    review: String,

    // Their results
    usedOnCrop: String,
    usedOnAcres: Number,
    yieldResult: Number,
    yieldUnit: String,
    wouldRecommend: { type: Boolean, default: true },

    // Did they modify it?
    madeModifications: { type: Boolean, default: false },
    modifications: String, // What they changed

    // Photos of their results
    images: [{
        url: String,
        caption: String
    }],

    // Helpful votes
    helpfulVotes: { type: Number, default: 0 },

    createdAt: { type: Date, default: Date.now }
});

// One rating per user per mix
chemicalMixRatingSchema.index({ mixId: 1, userId: 1 }, { unique: true });
chemicalMixRatingSchema.index({ mixId: 1, rating: -1 });

const ChemicalMixRating = mongoose.model('ChemicalMixRating', chemicalMixRatingSchema);

// Helper: Update mix rating averages
async function updateMixRatingStats(mixId) {
    const stats = await ChemicalMixRating.aggregate([
        { $match: { mixId: new mongoose.Types.ObjectId(mixId) } },
        { $group: {
            _id: '$mixId',
            averageRating: { $avg: '$rating' },
            totalRatings: { $sum: 1 }
        }}
    ]);

    if (stats.length > 0) {
        await ChemicalMix.findByIdAndUpdate(mixId, {
            averageRating: Math.round(stats[0].averageRating * 10) / 10,
            totalRatings: stats[0].totalRatings
        });
    }
}

// Helper: Generate invoice number
async function generateInvoiceNumber() {
    const year = new Date().getFullYear();
    const seq = await nextSequence(`invoice-${year}`);
    return `INV-${year}-${String(seq).padStart(5, '0')}`;
}

// Helper: Generate program-quote number. Same atomic sequence pattern as
// invoices so concurrent quote creations never collide on the unique index.
async function generateQuoteNumber() {
    const year = new Date().getFullYear();
    const seq = await nextSequence(`programQuote-${year}`);
    return `PQ-${year}-${String(seq).padStart(5, '0')}`;
}

// Auto-generate an Invoice record when an order transitions to paid.
// Fires from the Stripe webhook (payment_intent.succeeded / charge.succeeded)
// and from the admin check-received route. Idempotent - one invoice per order.
// Never throws - logs and returns null on failure so it can't break the
// payment confirmation flow.
async function autoGenerateInvoiceForPaidOrder(order) {
    try {
        if (!order) return null;
        if (order.paymentStatus !== 'paid') {
            console.log(`autoGenerateInvoice: skipping ${order._id} - paymentStatus=${order.paymentStatus}`);
            return null;
        }

        // Idempotency - one invoice per order
        const existing = await Invoice.findOne({ orderId: order._id });
        if (existing) return existing;

        // Detect which order collection this is (ChemicalOrder vs legacy Order)
        const isChemicalOrder = order.constructor?.modelName === 'ChemicalOrder'
            || (typeof order.orderNumber === 'string' && order.orderNumber.startsWith('CO-'));

        const customer = await User.findById(order.userId).lean();
        if (!customer) {
            console.error(`autoGenerateInvoice: customer not found for order ${order._id}`);
            return null;
        }

        const invoiceNumber = await generateInvoiceNumber();

        let invoiceItems = [];
        let subtotal = 0;
        let total = 0;
        let orderNumber = '';

        if (isChemicalOrder) {
            // Enrich items with catalog cost/margin data
            const chemicalIds = (order.items || []).map(i => i.chemicalId).filter(Boolean);
            const chemicalsMap = {};
            if (chemicalIds.length > 0) {
                const chems = await Chemical.find({ _id: { $in: chemicalIds } }).lean();
                chems.forEach(c => { chemicalsMap[c._id.toString()] = c; });
            }

            invoiceItems = (order.items || []).map(item => {
                const chem = item.chemicalId ? chemicalsMap[item.chemicalId.toString()] : null;
                const costPrice = chem?.costPrice || 0;
                const adminPrice = chem?.adminPrice || 0;
                const unitPrice = item.unitPrice || item.pricePerUnit || chem?.sellPrice || 0;
                const qty = item.quantity || 0;
                return {
                    productName: item.productName,
                    description: `${item.packSize || ''} ${item.unit || ''}`.trim(),
                    packSize: item.packSize,
                    unit: item.unit,
                    unitsPerPack: chem?.unitsPerPack || 1,
                    quantity: qty,
                    packQuantity: item.packQuantity || (chem?.unitsPerPack ? Math.ceil(qty / chem.unitsPerPack) : qty),
                    unitPrice,
                    costPrice,
                    adminPrice,
                    totalPrice: qty * unitPrice,
                    margin: (unitPrice - costPrice) * qty
                };
            });
            subtotal = order.subtotal || 0;
            total = order.total || subtotal;
            orderNumber = order.orderNumber || '';
        } else {
            // Legacy Order model has chemicals/seeds/pivotBio arrays
            (order.chemicals || []).forEach(chem => {
                invoiceItems.push({
                    productName: chem.name,
                    description: `${chem.packageSize} ${chem.packageUnit}`,
                    packSize: `${chem.packageSize}`,
                    unit: chem.packageUnit,
                    quantity: chem.packagesNeeded,
                    unitPrice: chem.pricePerPackage,
                    totalPrice: chem.totalPrice
                });
            });
            (order.seeds || []).forEach(seed => {
                invoiceItems.push({
                    productName: seed.name,
                    description: `${seed.crop} seed`,
                    packSize: 'bag',
                    unit: 'bags',
                    quantity: seed.bagsNeeded,
                    unitPrice: seed.pricePerBag,
                    totalPrice: seed.totalPrice
                });
            });
            (order.pivotBio || []).forEach(pb => {
                invoiceItems.push({
                    productName: pb.product,
                    description: 'PivotBio',
                    packSize: 'unit',
                    unit: 'units',
                    quantity: Math.ceil(pb.totalAmount),
                    unitPrice: pb.pricePerUnit,
                    totalPrice: pb.totalPrice
                });
            });
            subtotal = order.totalCost || invoiceItems.reduce((sum, i) => sum + (i.totalPrice || 0), 0);
            total = subtotal;
            orderNumber = `ORD-${order._id.toString().slice(-8).toUpperCase()}`;
        }

        const invoice = new Invoice({
            invoiceNumber,
            customerId: customer._id,
            customerName: customer.name,
            customerEmail: customer.email,
            customerPhone: customer.phone,
            customerAddress: customer.address,
            orderId: order._id,
            orderNumber,
            representativeId: order.representativeId,
            items: invoiceItems,
            subtotal,
            discount: order.discount || 0,
            total,
            amountPaid: total,
            amountDue: 0,
            paymentStatus: 'paid',
            paymentMethod: order.paymentMethod,
            paymentDate: order.paidAt || new Date(),
            stripePaymentIntentId: order.stripePaymentIntentId || null,
            status: 'paid',
            invoiceDate: new Date(),
            dueDate: new Date(),
            createdBy: order.createdBy || order.representativeId || null
        });

        await invoice.save();
        console.log(`Invoice ${invoice.invoiceNumber} auto-generated for order ${order._id}`);
        return invoice;
    } catch (err) {
        console.error(`autoGenerateInvoice error for order ${order?._id}:`, err.message);
        return null;
    }
}

// ============ RUP SALE RECORD AUTO-CREATION (I5) ============
// Colorado + FIFRA require a record of every Restricted Use Pesticide sale to a
// licensed applicator, maintained for at least 2 years. Lifecycle:
//   order placed with RUP item -> RupSaleRecord(status='pending_verification')
//   payment clears            -> status='completed', saleDate = payment date
//   order cancelled/archived  -> status='cancelled' (never deleted - audit trail)
// Detection is always via the Chemical.isRestrictedUse schema field.

// Helper: return [{item, chemical}] pairs for every RUP item on an order
async function getRupItemsForOrder(order) {
    if (!order?.items?.length) return [];
    const chemicalIds = order.items.map(i => i.chemicalId).filter(Boolean);
    if (chemicalIds.length === 0) return [];
    const rupChems = await Chemical.find({
        _id: { $in: chemicalIds },
        isRestrictedUse: true
    }).lean();
    if (rupChems.length === 0) return [];
    const rupChemMap = {};
    rupChems.forEach(c => { rupChemMap[c._id.toString()] = c; });
    return order.items
        .filter(i => i.chemicalId && rupChemMap[i.chemicalId.toString()])
        .map(i => ({ item: i, chemical: rupChemMap[i.chemicalId.toString()] }));
}

// Auto-create RupSaleRecord(s) for every RUP item on an order.
// Idempotent - skips any (orderId, chemicalId) pair that already has a non-cancelled record.
// Never throws - logs and returns on failure.
async function autoCreateRupRecordsForOrder(order, opts = {}) {
    try {
        if (!order) return [];
        const status = opts.status || 'pending_verification';
        const rupPairs = await getRupItemsForOrder(order);
        if (rupPairs.length === 0) return [];

        const customer = await User.findById(order.userId).lean();
        if (!customer) {
            console.error(`autoCreateRupRecords: customer not found for order ${order._id}`);
            return [];
        }

        const created = [];
        for (const { item, chemical } of rupPairs) {
            // Idempotency - one active record per (orderId, chemicalId)
            const existing = await RupSaleRecord.findOne({
                orderId: order._id,
                chemicalId: chemical._id,
                status: { $ne: 'cancelled' }
            });
            if (existing) continue;

            // Pick whichever license the customer has verified (private or commercial)
            const hasPrivate = customer.privateApplicatorLicense?.hasLicense;
            const license = hasPrivate
                ? customer.privateApplicatorLicense
                : customer.commercialApplicatorLicense;
            const licenseType = hasPrivate ? 'private' : 'commercial';

            if (!license?.licenseNumber || !license?.state) {
                // Should not happen - validateRupCompliance blocks orders without a valid license.
                // If we're here, something upstream is broken - log and skip this item.
                console.warn(`autoCreateRupRecords: customer ${customer._id} missing license data - skipping ${chemical.productName}`);
                continue;
            }

            const totalUnits = (item.quantity || 0) * (chemical.unitsPerPack || 1);
            const unitPrice = item.unitPrice || chemical.sellPrice || 0;
            const totalAmount = Math.round(totalUnits * unitPrice * 100) / 100;
            const requiredCerts = chemical.requiredCertifications || [];

            const record = new RupSaleRecord({
                saleDate: new Date(),
                orderNumber: order.orderNumber,
                orderId: order._id,
                sellerId: order.representativeId || order.createdBy,
                sellerName: order.representativeName || '',
                purchaserId: customer._id,
                purchaserName: customer.name,
                purchaserAddress: customer.address || {},
                purchaserPhone: customer.phone,
                purchaserEmail: customer.email,
                applicatorLicenseType: licenseType,
                applicatorLicenseNumber: license.licenseNumber,
                applicatorLicenseState: license.state,
                applicatorLicenseExpiration: license.expirationDate,
                applicatorCertificationCategories: license.certificationCategories || [],
                licenseVerificationMethod: 'document_on_file',
                licenseVerifiedBy: license.verifiedBy,
                licenseVerifiedAt: license.verifiedAt,
                licenseDocumentUrl: license.licenseDocumentUrl,
                chemicalId: chemical._id,
                productName: chemical.productName,
                epaRegistrationNumber: chemical.epaRegistrationNumber || 'UNKNOWN',
                activeIngredient: Array.isArray(chemical.activeIngredients) ? chemical.activeIngredients[0] : '',
                signalWord: chemical.signalWord,
                quantity: totalUnits,
                unit: chemical.unit,
                packSize: chemical.packSize,
                totalAmount,
                paraquatCertRequired: requiredCerts.includes('paraquat_training'),
                paraquatCertVerified: customer.paraquatCertification?.completed === true,
                paraquatCertNumber: customer.paraquatCertification?.certificateNumber,
                paraquatCertDate: customer.paraquatCertification?.completionDate,
                dicambaCertRequired: requiredCerts.includes('dicamba_training'),
                dicambaCertVerified: customer.dicambaCertification?.completed === true,
                dicambaCertYear: customer.dicambaCertification?.trainingYear,
                status,
                createdBy: order.createdBy || order.representativeId || null
            });

            await record.save();
            created.push(record);
            console.log(`RupSaleRecord created (${status}) for order ${order._id} product ${chemical.productName}`);
        }
        return created;
    } catch (err) {
        console.error(`autoCreateRupRecords error for order ${order?._id}:`, err.message);
        return [];
    }
}

// Promote pending_verification records to completed once payment clears.
// If no pending records exist but the order has RUP items (e.g., record creation
// previously failed), create them directly as 'completed' so we never ship RUP
// product without a record.
async function promoteRupRecordsToCompleted(order) {
    try {
        if (!order?._id) return;
        const result = await RupSaleRecord.updateMany(
            { orderId: order._id, status: 'pending_verification' },
            { $set: { status: 'completed', saleDate: order.paidAt || new Date(), updatedAt: new Date() } }
        );
        if (result.modifiedCount > 0) {
            console.log(`RupSaleRecords promoted to completed for order ${order._id}: ${result.modifiedCount}`);
            return;
        }
        // No pending records - fallback: create as completed if RUP items exist on this order
        const rupPairs = await getRupItemsForOrder(order);
        if (rupPairs.length > 0) {
            await autoCreateRupRecordsForOrder(order, { status: 'completed' });
        }
    } catch (err) {
        console.error(`promoteRupRecords error for order ${order?._id}:`, err.message);
    }
}

// Cancel any non-cancelled records tied to this order. Never delete - audit trail.
async function cancelRupRecordsForOrder(order) {
    try {
        if (!order?._id) return;
        const result = await RupSaleRecord.updateMany(
            { orderId: order._id, status: { $ne: 'cancelled' } },
            { $set: { status: 'cancelled', updatedAt: new Date() } }
        );
        if (result.modifiedCount > 0) {
            console.log(`RupSaleRecords cancelled for order ${order._id}: ${result.modifiedCount}`);
        }
    } catch (err) {
        console.error(`cancelRupRecords error for order ${order?._id}:`, err.message);
    }
}

// Helper: Update inventory when receiving a PO
async function receiveInventory({ chemicalId, productName, packSize, unit, quantity, unitCost, location, purchaseOrderId, poNumber, lotNumber, supplierName, userId, session }) {
    const invQuery = Inventory.findOne({ chemicalId, location: location || 'main' });
    if (session) invQuery.session(session);
    let inventory = await invQuery;

    if (!inventory) {
        inventory = new Inventory({
            chemicalId,
            productName,
            packSize,
            unit,
            location: location || 'main',
            quantityOnHand: 0,
            quantityReserved: 0,
            quantityAvailable: 0,
            averageCost: unitCost,
            lastCost: unitCost
        });
    }

    const previousQuantity = inventory.quantityOnHand;
    const newQuantity = previousQuantity + quantity;

    if (previousQuantity > 0 && inventory.averageCost > 0) {
        const totalOldValue = previousQuantity * inventory.averageCost;
        const totalNewValue = quantity * unitCost;
        inventory.averageCost = (totalOldValue + totalNewValue) / newQuantity;
    } else {
        inventory.averageCost = unitCost;
    }

    inventory.quantityOnHand = newQuantity;
    inventory.quantityAvailable = newQuantity - inventory.quantityReserved;
    inventory.lastCost = unitCost;
    inventory.lastReceivedDate = new Date();
    inventory.updatedAt = new Date();

    await inventory.save({ session });

    const batch = new InventoryBatch({
        chemicalId,
        productName,
        packSize,
        unit,
        poNumber: poNumber || 'UNTRACKED',
        poId: purchaseOrderId,
        lotNumber: lotNumber || '',
        quantityReceived: quantity,
        quantityRemaining: quantity,
        costPerUnit: unitCost,
        totalCost: quantity * unitCost,
        location: location || 'main',
        supplierName: supplierName || '',
        status: 'active',
        receivedDate: new Date()
    });

    await batch.save({ session });

    const transaction = new InventoryTransaction({
        inventoryId: inventory._id,
        chemicalId,
        productName,
        type: 'receive',
        quantityChange: quantity,
        previousQuantity,
        newQuantity,
        unitCost,
        totalCost: quantity * unitCost,
        referenceType: 'PurchaseOrder',
        referenceId: purchaseOrderId,
        referenceNumber: poNumber,
        location: location || 'main',
        createdBy: userId
    });

    await transaction.save({ session });

    // Auto-update product cost when new batch arrives at a different price
    const chemQuery = Chemical.findById(chemicalId);
    if (session) chemQuery.session(session);
    const chemical = await chemQuery;
    if (chemical) {
        // Clear speculation flag on any PO receive — price is now confirmed
        if (chemical.priceIsSpeculated === true) {
            chemical.priceIsSpeculated = false;
            chemical.priceDate = new Date();
            chemical.updatedAt = new Date();
            await chemical.save({ session });
            console.log(`Speculation cleared: ${productName} — PO ${poNumber || 'manual'} landed`);
        }

        if (unitCost !== chemical.costPrice) {
            const oldCost = chemical.costPrice;
            chemical.costPrice = unitCost;
            chemical.adminPrice = Math.round((unitCost + (chemical.adminMarginDollars || 0)) * 100) / 100;
            chemical.sellPrice = Math.round((chemical.adminPrice + (chemical.marginDollars || 0)) * 100) / 100;
            chemical.margin = chemical.sellPrice > 0 ? Math.round(((chemical.sellPrice - chemical.costPrice) / chemical.sellPrice) * 10000) / 100 : 0;
            chemical.priceDate = new Date();
            chemical.updatedAt = new Date();
            await chemical.save({ session });

            await ChemicalPriceHistory.create([{
                chemicalId: chemical._id,
                productName: chemical.productName,
                sourceSupplier: chemical.sourceSupplier,
                packSize: chemical.packSize,
                unit: chemical.unit,
                costPrice: unitCost,
                previousCostPrice: oldCost,
                adminPrice: chemical.adminPrice,
                sellPrice: chemical.sellPrice,
                priceVersion: `PO-${poNumber || 'manual'}`,
                notes: `Cost $${oldCost.toFixed(2)} → $${unitCost.toFixed(2)} via ${poNumber || 'receive'}. Margins unchanged.`
            }], { session });

            console.log(`Price auto-updated: ${productName} cost $${oldCost.toFixed(2)} → $${unitCost.toFixed(2)}, sell $${chemical.sellPrice.toFixed(2)} (margins unchanged)`);
        }
    }

    return { inventory, transaction, batch };
}

// Helper: Deduct inventory when fulfilling an order
async function deductInventory({ chemicalId, quantity, location, orderId, orderNumber, userId, notes }) {
    const inventory = await Inventory.findOne({ chemicalId, location: location || 'main' });

    if (!inventory) {
        throw new Error(`No inventory found for product at location ${location || 'main'}`);
    }

    if (inventory.quantityAvailable < quantity) {
        throw new Error(`Insufficient inventory. Available: ${inventory.quantityAvailable}, Requested: ${quantity}`);
    }

    const previousQuantity = inventory.quantityOnHand;
    const newQuantity = previousQuantity - quantity;

    inventory.quantityOnHand = newQuantity;
    inventory.quantityAvailable = newQuantity - inventory.quantityReserved;
    inventory.lastSoldDate = new Date();
    inventory.updatedAt = new Date();

    await inventory.save();

    // Create transaction record
    const transaction = new InventoryTransaction({
        inventoryId: inventory._id,
        chemicalId,
        productName: inventory.productName,
        type: 'sale',
        quantityChange: -quantity,
        previousQuantity,
        newQuantity,
        unitCost: inventory.averageCost,
        totalCost: quantity * inventory.averageCost,
        referenceType: 'ChemicalOrder',
        referenceId: orderId,
        referenceNumber: orderNumber,
        location: location || 'main',
        notes,
        createdBy: userId
    });

    await transaction.save();

    return { inventory, transaction };
}

// Helper: Reserve inventory when an order is placed/updated
async function reserveInventory({ chemicalId, quantity, location, orderId, orderNumber, userId, notes, session }) {
    const invQuery = Inventory.findOne({ chemicalId, location: location || 'main' });
    if (session) invQuery.session(session);
    let inventory = await invQuery;

    if (!inventory) {
        const chemQuery = Chemical.findById(chemicalId);
        if (session) chemQuery.session(session);
        const chemical = await chemQuery;
        inventory = new Inventory({
            chemicalId,
            productName: chemical?.productName || 'Unknown Product',
            packSize: chemical?.packSize || '',
            unit: chemical?.unit || 'units',
            location: location || 'main',
            quantityOnHand: 0,
            quantityReserved: 0,
            quantityAvailable: 0,
            averageCost: 0,
            lastCost: 0
        });
    }

    const previousReserved = inventory.quantityReserved || 0;
    inventory.quantityReserved = previousReserved + quantity;
    inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
    inventory.updatedAt = new Date();

    await inventory.save({ session });

    const transaction = new InventoryTransaction({
        inventoryId: inventory._id,
        chemicalId,
        productName: inventory.productName,
        type: 'reserve',
        quantityChange: quantity,
        previousQuantity: previousReserved,
        newQuantity: inventory.quantityReserved,
        unitCost: inventory.averageCost || 0,
        totalCost: quantity * (inventory.averageCost || 0),
        referenceType: 'ChemicalOrder',
        referenceId: orderId,
        referenceNumber: orderNumber,
        location: location || 'main',
        notes: notes || 'Inventory reserved for order',
        createdBy: userId
    });

    await transaction.save({ session });

    return { inventory, transaction };
}

// Helper: Release reserved inventory (order cancelled/modified)
async function releaseInventory({ chemicalId, quantity, location, orderId, orderNumber, userId, notes }) {
    const inventory = await Inventory.findOne({ chemicalId, location: location || 'main' });

    if (!inventory) {
        // No inventory to release from
        return null;
    }

    const previousReserved = inventory.quantityReserved || 0;
    const releaseQty = Math.min(quantity, previousReserved); // Can't release more than reserved

    inventory.quantityReserved = previousReserved - releaseQty;
    inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
    inventory.updatedAt = new Date();

    await inventory.save();

    // Create transaction record for audit trail
    const transaction = new InventoryTransaction({
        inventoryId: inventory._id,
        chemicalId,
        productName: inventory.productName,
        type: 'release',
        quantityChange: -releaseQty,
        previousQuantity: previousReserved,
        newQuantity: inventory.quantityReserved,
        unitCost: inventory.averageCost || 0,
        totalCost: releaseQty * (inventory.averageCost || 0),
        referenceType: 'ChemicalOrder',
        referenceId: orderId,
        referenceNumber: orderNumber,
        location: location || 'main',
        notes: notes || 'Inventory released from order',
        createdBy: userId
    });

    await transaction.save();

    return { inventory, transaction };
}

// Helper: Update order inventory reservations (handles item changes)
async function updateOrderInventory({ oldItems, newItems, orderId, orderNumber, location, userId }) {
    const changes = [];

    // Build maps of old and new quantities by chemicalId
    const oldQtyMap = new Map();
    const newQtyMap = new Map();

    (oldItems || []).forEach(item => {
        if (item.chemicalId) {
            const key = item.chemicalId.toString();
            oldQtyMap.set(key, (oldQtyMap.get(key) || 0) + (item.quantity || 0));
        }
    });

    (newItems || []).forEach(item => {
        if (item.chemicalId) {
            const key = item.chemicalId.toString();
            newQtyMap.set(key, (newQtyMap.get(key) || 0) + (item.quantity || 0));
        }
    });

    // Get all unique chemicalIds
    const allChemicalIds = new Set([...oldQtyMap.keys(), ...newQtyMap.keys()]);

    for (const chemicalId of allChemicalIds) {
        const oldQty = oldQtyMap.get(chemicalId) || 0;
        const newQty = newQtyMap.get(chemicalId) || 0;
        const diff = newQty - oldQty;

        if (diff > 0) {
            // Need to reserve more
            const result = await reserveInventory({
                chemicalId,
                quantity: diff,
                location,
                orderId,
                orderNumber,
                userId,
                notes: `Order updated: reserved ${diff} more units`
            });
            changes.push({ chemicalId, action: 'reserve', quantity: diff, result });
        } else if (diff < 0) {
            // Need to release some
            const result = await releaseInventory({
                chemicalId,
                quantity: Math.abs(diff),
                location,
                orderId,
                orderNumber,
                userId,
                notes: `Order updated: released ${Math.abs(diff)} units`
            });
            changes.push({ chemicalId, action: 'release', quantity: Math.abs(diff), result });
        }
    }

    return changes;
}

// Helper: Generate next PO number
async function generatePONumber() {
    const year = new Date().getFullYear();
    const prefix = `PO-${year}-`;

    const lastPO = await PurchaseOrder.findOne({ poNumber: { $regex: `^${prefix}` } })
        .sort({ poNumber: -1 })
        .lean();

    let nextNum = 1;
    if (lastPO && lastPO.poNumber) {
        const lastNum = parseInt(lastPO.poNumber.split('-')[2], 10);
        if (!isNaN(lastNum)) {
            nextNum = lastNum + 1;
        }
    }

    return `${prefix}${String(nextNum).padStart(5, '0')}`;
}

// ============ INITIALIZE ADMIN USERS ============

async function initializeAdmins() {
    const admins = [
        { name: 'Acre Profit Admin', email: 'contact@acreprofit.com', phone: '970-571-1015', role: 'superadmin' },
        { name: 'Kyle McConnell',    email: 'office@togoag.com',       phone: '970-571-1015', role: 'distributor' },
        { name: 'Ty Mollohan',       email: 'tymollohan77@gmail.com',  phone: '970-520-2340', role: 'distributor' },
        { name: 'Chad Bamford',      email: 'ckbamford@yahoo.com',     phone: '970-520-3716', role: 'distributor' },
        { name: 'Seth Rolfs',        email: 'seth@acreprofit.com',     phone: '785-531-0680', role: 'distributor' },
        { name: 'Tyson',             email: 'fyeagllc@gmail.com',                             role: 'distributor' }
    ];

    for (const admin of admins) {
        try {
            const existing = await User.findOne({ email: admin.email });
            if (!existing) {
                // New admin: random 12-char temp password, must-change flag on.
                // Temp is logged to the Render console - Kyle grabs it once and
                // distributes out-of-band. After first login the user changes it.
                const tempPassword = crypto.randomBytes(9).toString('base64').replace(/[+/=]/g, '').slice(0, 12);
                await User.create({
                    ...admin,
                    password: tempPassword,
                    mustChangePassword: true
                });
                console.log(`[INITIAL PASSWORD] ${admin.email}: ${tempPassword} — MUST CHANGE ON FIRST LOGIN`);
                continue;
            }

            // Existing admin: sync the role, never touch the password or the
            // mustChangePassword flag. The initial flag migration was a one-shot
            // that's already run in production - running it every startup would
            // flip freshly-rotated admins BACK to mustChangePassword=true on
            // every Render restart, trapping them in a loop. Kyle confirmed
            // this bug in production. Future admins who need to be flagged
            // (e.g. compromised account) use POST /api/admin/users/:id/
            // force-password-change instead.
            const before = {
                role: existing.role,
                mustChangePassword: existing.mustChangePassword
            };
            let changed = false;

            if (existing.role !== admin.role) {
                existing.role = admin.role;
                changed = true;
            }

            if (changed) {
                await existing.save();
                const after = await User.findOne({ email: admin.email })
                    .select('role mustChangePassword')
                    .lean();
                console.log(`Admin ${admin.email}: role synced before=${JSON.stringify(before)} after=${JSON.stringify(after)}`);
            } else {
                console.log(`Admin ${admin.email}: OK (role=${existing.role}, mustChangePassword=${existing.mustChangePassword})`);
            }
        } catch (error) {
            console.error(`Admin ${admin.email} migration error:`, error.message);
        }
    }
}

// Seed Jabco Glyphosate Inventory (PO#2119 - 4,240 gal @ $13.25)
async function seedJabcoInventory() {
    try {
        // Check if batch already exists
        const existingBatch = await InventoryBatch.findOne({ lotNumber: 'jabco2119' });
        if (existingBatch) {
            console.log('Jabco inventory (jabco2119) already exists');
            return;
        }

        // Find or create the product
        let product = await Chemical.findOne({
            productName: 'XSATE Glyphosate 53.8%',
            packSize: '265 gal'
        });

        if (!product) {
            product = await Chemical.create({
                productName: 'XSATE Glyphosate 53.8%',
                packSize: '265 gal',
                unit: 'gal',
                unitsPerPack: 265,
                costPrice: 13.25,
                adminMarginDollars: 0,
                adminPrice: 13.25,
                marginDollars: 0,
                sellPrice: 13.25,
                category: 'herbicide',
                sourceSupplier: 'Jabco',
                epaRegistrationNumber: '89343-5',
                signalWord: 'CAUTION',
                notes: '5.4 lb/gal glyphosate - Xingfa USA',
                activeIngredients: [{ name: 'XSATE Glyphosate 53.8%', percentage: 53.8, poundsPerGallon: 5.4 }]
            });
            console.log('Created XSATE Glyphosate 53.8% product');
        } else {
            // Don't overwrite prices - set them in admin panel
            console.log('XSATE Glyphosate already exists - set margins in admin panel');
        }

        // Create inventory record
        let inventory = await Inventory.findOne({ chemicalId: product._id, location: 'main' });
        if (!inventory) {
            inventory = await Inventory.create({
                chemicalId: product._id,
                productName: 'XSATE Glyphosate 53.8%',
                packSize: '265 gal',
                unit: 'gal',
                location: 'main',
                quantityOnHand: 4240,
                quantityReserved: 0,
                quantityAvailable: 4240,
                averageCost: 13.25,
                lastCost: 13.25,
                lastReceivedDate: new Date('2026-03-09')
            });
        } else {
            // Update existing inventory
            inventory.quantityOnHand += 4240;
            inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
            inventory.lastCost = 13.25;
            inventory.lastReceivedDate = new Date('2026-03-09');
            await inventory.save();
        }

        // Create batch record
        await InventoryBatch.create({
            chemicalId: product._id,
            productName: 'XSATE Glyphosate 53.8%',
            packSize: '265 gal',
            unit: 'gal',
            poNumber: 'JABCO-2119',
            lotNumber: 'jabco2119',
            quantityReceived: 4240,
            quantityRemaining: 4240,
            costPerUnit: 13.25,
            totalCost: 56180.00,
            location: 'main',
            supplierName: 'Jabco',
            status: 'active',
            receivedDate: new Date('2026-03-09')
        });

        // Create transaction record
        await InventoryTransaction.create({
            inventoryId: inventory._id,
            chemicalId: product._id,
            productName: 'XSATE Glyphosate 53.8%',
            type: 'receive',
            quantityChange: 4240,
            previousQuantity: 0,
            newQuantity: 4240,
            unitCost: 13.25,
            totalCost: 56180.00,
            referenceType: 'PurchaseOrder',
            referenceNumber: 'JABCO-2119',
            location: 'main',
            notes: 'Initial Jabco inventory - Sales Order #2119'
        });

        console.log('Seeded Jabco inventory: 4,240 gal XSATE Glyphosate @ $13.25/gal (lot: jabco2119)');
    } catch (error) {
        console.error('Error seeding Jabco inventory:', error.message);
    }
}

// Seed March 30, 2026 Purchase Orders (JABCO SO2129, Sims #103850, Sims #103849)
async function seedMarch2026PurchaseOrders() {
    try {
        // Check if already seeded
        const existingJabco = await PurchaseOrder.findOne({ poNumber: 'JABCO-SO2129' });
        const existingSims850 = await PurchaseOrder.findOne({ poNumber: 'SIMS-103850' });
        const existingSims849 = await PurchaseOrder.findOne({ poNumber: 'SIMS-103849' });

        if (existingJabco && existingSims850 && existingSims849) {
            console.log('March 2026 purchase orders already seeded');
            return;
        }

        // ============ JABCO SO2129 - Acre Profit LLC - 3/30/2026 ============
        if (!existingJabco) {
            // Find or create products for JABCO order
            const jabcoProducts = [
                {
                    productName: 'Meso 4SC',
                    packSize: '2x2.5 GL Case',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 720,
                    unitPrice: 45.75,
                    lineTotal: 32940.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Flumioxazin 51% WDG',
                    packSize: '4x5 Lb Case',
                    unit: 'lb',
                    unitsPerPack: 20,
                    qty: 1440,
                    unitPrice: 14.00,
                    lineTotal: 20160.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Sulfentrazone 39.6% SC',
                    packSize: '2x2.5 Gl Case',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 180,
                    unitPrice: 71.50,
                    lineTotal: 12870.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Dicamba 49.8% SL',
                    packSize: '2x2.5 Gl Case',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 180,
                    unitPrice: 30.25,
                    lineTotal: 5445.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Defy LV-6',
                    packSize: '2x2.5 Gl Case',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 180,
                    unitPrice: 28.90,
                    lineTotal: 5202.00,
                    category: 'herbicide'
                }
            ];

            const jabcoItems = [];
            for (const prod of jabcoProducts) {
                // Find or create the product
                let chemical = await Chemical.findOne({
                    productName: prod.productName,
                    packSize: prod.packSize
                });

                if (!chemical) {
                    chemical = await Chemical.create({
                        productName: prod.productName,
                        packSize: prod.packSize,
                        unit: prod.unit,
                        unitsPerPack: prod.unitsPerPack,
                        costPrice: prod.unitPrice,
                        adminMarginDollars: 0,
                        marginDollars: 0,
                        adminPrice: prod.unitPrice, // No margin until set by admin
                        sellPrice: prod.unitPrice, // No margin until set by admin
                        category: prod.category,
                        sourceSupplier: 'JABCO',
                        signalWord: 'CAUTION'
                    });
                }

                jabcoItems.push({
                    productName: prod.productName,
                    chemicalId: chemical._id,
                    packSize: prod.packSize,
                    unit: prod.unit,
                    quantityOrdered: prod.qty,
                    pricePerUnit: prod.unitPrice,
                    totalPrice: prod.lineTotal,
                    quantityAllocated: 0,
                    quantityRemaining: prod.qty
                });
            }

            await PurchaseOrder.create({
                poNumber: 'JABCO-SO2129',
                supplier: {
                    name: 'JABCO LLC',
                    contact: 'JABCO Sales'
                },
                items: jabcoItems,
                subtotal: 76617.00,
                freight: 0,
                totalCost: 76617.00,
                status: 'received',
                orderDate: new Date('2026-03-30'),
                notes: 'Billed to: Kyle McConnell'
            });

            console.log('Created JABCO SO2129 - $76,617.00 (5 products)');
        }

        // ============ Sims #103850 - Ty Mollohan - 3/30/2026 ============
        if (!existingSims850) {
            const simsProducts850 = [
                {
                    productName: 'Dicamba DMA (Tigris)',
                    packSize: '265 gal',
                    unit: 'gal',
                    unitsPerPack: 265,
                    qty: 1060,
                    unitPrice: 25.50,
                    lineTotal: 27030.00,
                    category: 'herbicide'
                },
                {
                    productName: 'LV6 De-Ester LV6 (Drexel)',
                    packSize: '265 gal',
                    unit: 'gal',
                    unitsPerPack: 265,
                    qty: 1060,
                    unitPrice: 22.75,
                    lineTotal: 24115.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Anthem NXT',
                    packSize: '2x2.5 gal',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 90,
                    unitPrice: 430.00,
                    lineTotal: 38700.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Mivum',
                    packSize: '8x16 oz',
                    unit: 'oz',
                    unitsPerPack: 128,
                    qty: 1600,
                    unitPrice: 2.25,
                    lineTotal: 3600.00,
                    category: 'herbicide'
                }
            ];

            const simsItems850 = [];
            for (const prod of simsProducts850) {
                let chemical = await Chemical.findOne({
                    productName: prod.productName,
                    packSize: prod.packSize
                });

                if (!chemical) {
                    chemical = await Chemical.create({
                        productName: prod.productName,
                        packSize: prod.packSize,
                        unit: prod.unit,
                        unitsPerPack: prod.unitsPerPack,
                        costPrice: prod.unitPrice,
                        adminMarginDollars: 0,
                        marginDollars: 0,
                        adminPrice: prod.unitPrice, // No margin until set by admin
                        sellPrice: prod.unitPrice, // No margin until set by admin
                        category: prod.category,
                        sourceSupplier: 'Sims Fertilizer & Chemical',
                        signalWord: 'CAUTION'
                    });
                }

                simsItems850.push({
                    productName: prod.productName,
                    chemicalId: chemical._id,
                    packSize: prod.packSize,
                    unit: prod.unit,
                    quantityOrdered: prod.qty,
                    pricePerUnit: prod.unitPrice,
                    totalPrice: prod.lineTotal,
                    quantityAllocated: 0,
                    quantityRemaining: prod.qty
                });
            }

            await PurchaseOrder.create({
                poNumber: 'SIMS-103850',
                supplier: {
                    name: 'Sims Fertilizer & Chemical',
                    contact: 'Sims Sales'
                },
                items: simsItems850,
                subtotal: 93445.00,
                freight: 0,
                totalCost: 93445.00,
                status: 'received',
                orderDate: new Date('2026-03-30'),
                notes: 'Billed to: Ty Mollohan'
            });

            console.log('Created Sims #103850 - $93,445.00 (4 products)');
        }

        // ============ Sims #103849 - Ty Mollohan - 3/30/2026 ============
        if (!existingSims849) {
            // Find or create Atrazine 4L Tote
            let atrazineTote = await Chemical.findOne({
                productName: 'Atrazine 4L',
                packSize: 'Tote'
            });

            if (!atrazineTote) {
                atrazineTote = await Chemical.create({
                    productName: 'Atrazine 4L',
                    packSize: 'Tote',
                    unit: 'gal',
                    unitsPerPack: 265, // Standard tote size
                    costPrice: 12.50,
                    adminMarginDollars: 0,
                    adminPrice: 12.50,
                    marginDollars: 0,
                    sellPrice: 12.50,
                    category: 'herbicide',
                    sourceSupplier: 'Sims Fertilizer & Chemical',
                    signalWord: 'CAUTION',
                    isRestrictedUse: true,
                    notes: 'Restricted Use Pesticide'
                });
            }

            await PurchaseOrder.create({
                poNumber: 'SIMS-103849',
                supplier: {
                    name: 'Sims Fertilizer & Chemical',
                    contact: 'Sims Sales'
                },
                items: [{
                    productName: 'Atrazine 4L',
                    chemicalId: atrazineTote._id,
                    packSize: 'Tote',
                    unit: 'gal',
                    quantityOrdered: 4240,
                    pricePerUnit: 12.50,
                    totalPrice: 53000.00,
                    quantityAllocated: 0,
                    quantityRemaining: 4240
                }],
                subtotal: 53000.00,
                freight: 0,
                totalCost: 53000.00,
                status: 'received',
                orderDate: new Date('2026-03-30'),
                notes: 'Billed to: Ty Mollohan'
            });

            console.log('Created Sims #103849 - $53,000.00 (Atrazine 4L Tote)');
        }

        console.log('March 2026 purchase orders seeded successfully. Total: $223,062.00');
    } catch (error) {
        console.error('Error seeding March 2026 purchase orders:', error.message);
    }
}

// Backfill inventory from all received POs that don't have inventory records yet
async function backfillInventoryFromPOs() {
    try {
        // Find all POs that are received or partially received
        const receivedPOs = await PurchaseOrder.find({
            status: { $in: ['received', 'partial_received'] }
        });

        if (receivedPOs.length === 0) {
            console.log('No received POs to backfill');
            return;
        }

        let totalBackfilled = 0;
        let ledgerEntriesCreated = 0;

        for (const po of receivedPOs) {
            // Check if this PO already has inventory batches
            const existingBatches = await InventoryBatch.find({ poId: po._id });
            if (existingBatches.length > 0) {
                continue; // Already has inventory records
            }

            // Also check by PO number
            const batchByNumber = await InventoryBatch.findOne({ poNumber: po.poNumber });
            if (batchByNumber) {
                continue;
            }

            console.log(`Backfilling inventory for PO ${po.poNumber}...`);

            for (const item of po.items) {
                let chemicalId = item.chemicalId;

                // Try to find chemical if not linked
                if (!chemicalId && item.productName) {
                    const chemical = await Chemical.findOne({
                        productName: { $regex: new RegExp(`^${item.productName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i') }
                    });
                    if (chemical) {
                        chemicalId = chemical._id;
                        item.chemicalId = chemicalId;
                    }
                }

                if (!chemicalId) {
                    console.warn(`  Skipping item "${item.productName}" - no chemical ID found`);
                    continue;
                }

                const quantity = item.quantityOrdered;

                await receiveInventory({
                    chemicalId,
                    productName: item.productName,
                    packSize: item.packSize,
                    unit: item.unit,
                    quantity,
                    unitCost: item.pricePerUnit,
                    location: 'main',
                    purchaseOrderId: po._id,
                    poNumber: po.poNumber,
                    lotNumber: `backfill-${po.poNumber}`,
                    supplierName: po.supplier?.name || '',
                    userId: po.createdBy
                });

                totalBackfilled++;
                console.log(`  + ${item.productName}: ${quantity} ${item.unit}`);
            }

            // Update PO items with quantityReceived
            for (const item of po.items) {
                item.quantityReceived = item.quantityOrdered;
            }
            po.receivedDate = po.receivedDate || po.updatedAt || new Date();
            await po.save();

            // Create ledger entry for PO payment - figure out who paid
            const billedTo = po.notes || '';
            const billedToMatch = billedTo.match(/billed\s*to:\s*(.+)/i);
            if (billedToMatch) {
                const personName = billedToMatch[1].trim();
                // Find the rep by name
                const rep = await User.findOne({
                    name: { $regex: new RegExp(personName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i') },
                    role: { $in: ['distributor', 'admin', 'superadmin'] }
                });

                if (rep) {
                    // Check if ledger entry already exists for this PO
                    const existingLedger = await LedgerEntry.findOne({
                        referenceType: 'PurchaseOrder',
                        referenceId: po._id
                    });

                    if (!existingLedger) {
                        await createLedgerEntry({
                            representativeId: rep._id,
                            description: `PO ${po.poNumber} - ${po.supplier?.name || 'Supplier'} (${po.items.length} items)`,
                            amount: po.totalCost || po.subtotal || 0,
                            type: 'credit',
                            category: 'supplier_payment',
                            referenceType: 'PurchaseOrder',
                            referenceId: po._id,
                            notes: `${personName} paid supplier for inventory. AP owes ${personName}. ${billedTo}`
                        });
                        ledgerEntriesCreated++;
                        console.log(`  Ledger: AP owes ${personName} $${po.totalCost} for PO ${po.poNumber}`);
                    }
                }
            }
        }

        if (totalBackfilled > 0) {
            console.log(`Inventory backfill complete: ${totalBackfilled} items, ${ledgerEntriesCreated} ledger entries`);
        }
    } catch (error) {
        console.error('Error backfilling inventory:', error.message);
    }
}

// Seed Hydrovant inventory - 360 gal at $75/gal cost, paid by Kyle
async function seedHydrovantInventory() {
    try {
        // Check if already seeded
        const existingBatch = await InventoryBatch.findOne({ lotNumber: 'hydrovant-kyle-360' });
        if (existingBatch) {
            console.log('Hydrovant inventory already seeded');
            return;
        }

        // Find or create Hydrovant product
        let hydrovant = await Chemical.findOne({
            productName: { $regex: /^hydrovant$/i }
        });

        if (!hydrovant) {
            hydrovant = await Chemical.create({
                productName: 'Hydrovant fA',
                packSize: '2x2.5 gal',
                unit: 'gal',
                unitsPerPack: 5,
                costPrice: 75.00,
                adminMarginDollars: 0,
                adminPrice: 75.00,
                marginDollars: 0,
                sellPrice: 75.00,
                category: 'adjuvant',
                sourceSupplier: 'JABCO',
                signalWord: 'CAUTION',
                notes: 'Premium NIS, water conditioner & drift control',
                defaultRate: 0.1,
                rateUnit: '% v/v',
                availableForOrder: true
            });
            console.log('Created Hydrovant product');
        } else {
            // Update cost price if needed, but don't overwrite margins
            if (hydrovant.costPrice !== 75.00) {
                hydrovant.costPrice = 75.00;
                await hydrovant.save();
                console.log('Updated Hydrovant cost price');
            }
        }

        // Add to inventory
        await receiveInventory({
            chemicalId: hydrovant._id,
            productName: 'Hydrovant fA',
            packSize: hydrovant.packSize || '2x2.5 gal',
            unit: 'gal',
            quantity: 360,
            unitCost: 75.00,
            location: 'main',
            purchaseOrderId: null,
            poNumber: 'DIRECT-HYDROVANT',
            lotNumber: 'hydrovant-kyle-360',
            supplierName: 'JABCO',
            userId: null
        });

        console.log('Added Hydrovant inventory: 360 gal at $75/gal ($27,000 total)');

        // Create ledger entry for Kyle
        const kyle = await User.findOne({ email: 'office@togoag.com' });
        if (kyle) {
            const existingLedger = await LedgerEntry.findOne({
                referenceType: 'Manual',
                notes: { $regex: /hydrovant.*360/i }
            });

            if (!existingLedger) {
                await createLedgerEntry({
                    representativeId: kyle._id,
                    description: 'Hydrovant fA - 360 gal at $75.00/gal (direct purchase)',
                    amount: 27000.00,
                    type: 'credit',
                    category: 'supplier_payment',
                    referenceType: 'Manual',
                    notes: 'Kyle paid $27,000 for Hydrovant fA. AP owes Kyle.'
                });
                console.log('Ledger: AP owes Kyle $27,000 for Hydrovant fA');
            }
        } else {
            console.log('Warning: Could not find Kyle McConnell account for ledger entry');
        }
    } catch (error) {
        console.error('Error seeding Hydrovant inventory:', error.message);
    }
}

// Ensure correct ledger entries for all PO purchases
async function seedPOLedgerEntries() {
    try {
        // Map of PO numbers to who paid
        const poBilling = {
            'JABCO-SO2129': { email: 'office@togoag.com', name: 'Kyle McConnell' },
            'SIMS-103850': { email: 'tymollohan77@gmail.com', name: 'Ty Mollohan' },
            'SIMS-103849': { email: 'tymollohan77@gmail.com', name: 'Ty Mollohan' }
        };

        for (const [poNumber, billing] of Object.entries(poBilling)) {
            // Check if ledger entry already exists for this PO
            const existing = await LedgerEntry.findOne({
                description: { $regex: new RegExp(poNumber) }
            });
            if (existing) continue;

            const po = await PurchaseOrder.findOne({ poNumber });
            if (!po) continue;

            const rep = await User.findOne({ email: billing.email });
            if (!rep) {
                console.log(`Could not find ${billing.name} for PO ${poNumber} ledger entry`);
                continue;
            }

            // Also check if there's a PurchaseOrder reference entry from backfill
            const existingPORef = await LedgerEntry.findOne({
                referenceType: 'PurchaseOrder',
                referenceId: po._id
            });
            if (existingPORef) continue;

            const total = po.totalCost || po.subtotal || 0;
            if (total <= 0) continue;

            const itemSummary = po.items.map(i => `${i.productName} (${i.quantityOrdered} ${i.unit || ''})`).join(', ');

            await createLedgerEntry({
                representativeId: rep._id,
                description: `PO ${poNumber} - ${po.supplier?.name || 'Supplier'}`,
                amount: total,
                type: 'credit',
                category: 'supplier_payment',
                referenceType: 'PurchaseOrder',
                referenceId: po._id,
                notes: `${billing.name} paid supplier. AP owes ${billing.name}. Items: ${itemSummary}`
            });

            console.log(`Ledger: AP owes ${billing.name} $${total.toLocaleString()} for PO ${poNumber}`);
        }
    } catch (error) {
        console.error('Error seeding PO ledger entries:', error.message);
    }
}

// ============ SPRAY PROGRAMS DATA ============

const sprayPrograms = {
    corn: {
        '2-pass': {
            name: '2-Pass Corn Program',
            description: 'Pre-emergent + Post-emergent application',
            chemicals: [
                { name: 'XSATE Glyphosate 53.8%', defaultRate: 32, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Atrazine 4L', defaultRate: 1.5, rateUnit: 'qt/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Metolachlor', defaultRate: 1.3, rateUnit: 'pt/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 }
            ]
        },
        '3-pass': {
            name: '3-Pass Corn Program',
            description: 'Burndown + Pre-emergent + Post-emergent application',
            chemicals: [
                { name: 'XSATE Glyphosate 53.8%', defaultRate: 32, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Atrazine 4L', defaultRate: 2, rateUnit: 'qt/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Metolachlor', defaultRate: 1.5, rateUnit: 'pt/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: '2,4-D Amine', defaultRate: 1, rateUnit: 'pt/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 }
            ]
        }
    },
    soybeans: {
        '2-pass': {
            name: '2-Pass Soybean Program',
            description: 'Pre-emergent + Post-emergent application',
            chemicals: [
                { name: 'XSATE Glyphosate 53.8%', defaultRate: 32, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Metribuzin', defaultRate: 0.5, rateUnit: 'lb/acre', packageSize: 50, packageUnit: 'lb', pricePerPackage: 0 }
            ]
        }
    },
    'dryland-corn': {
        'preplant': {
            name: 'Option 1 - Corn Preplant',
            description: '30 days pre-plant - wheat stubble with atrazine & valor',
            chemicals: [
                { name: 'XSATE Glyphosate 53.8%', defaultRate: 22, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Flumioxazin 51% WDG', defaultRate: 3, rateUnit: 'oz/acre', packageSize: 5, packageUnit: 'lb', pricePerPackage: 0 },
                { name: 'Atrazine 4L', defaultRate: 1, rateUnit: 'lb/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Hydrovant fA', defaultRate: 0.1, rateUnit: '% v/v', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0, isAdjuvant: true }
            ]
        },
        'post-plant-pre-emerge': {
            name: 'Option 2 - Corn Post Plant Pre-Emerge',
            description: 'Post plant pre-emerge - fall atrazine already applied',
            chemicals: [
                { name: 'XSATE Glyphosate 53.8%', defaultRate: 22, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Meso 4SC', defaultRate: 6, rateUnit: 'oz/acre', packageSize: 1, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Dicamba 49.8% SL', defaultRate: 4, rateUnit: 'oz/acre', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Anthem Max', defaultRate: 3, rateUnit: 'oz/acre', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Hydrovant fA', defaultRate: 0.1, rateUnit: '% v/v', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0, isAdjuvant: true }
            ]
        }
    }
};

// Seed Products
const seedProducts = {
    corn: {
        name: 'Corn Seed',
        rateUnit: 'seeds/acre',
        defaultRate: 32000,
        seedsPerBag: 80000,
        pricePerBag: 0
    },
    sorghum: {
        name: 'Sorghum Seed',
        rateUnit: 'seeds/acre',
        defaultRate: 50000,
        seedsPerBag: 600000,
        pricePerBag: 0
    },
    wheat: {
        name: 'Wheat Seed',
        rateUnit: 'lbs/acre',
        defaultRate: 90,
        lbsPerBag: 50,
        pricePerBag: 0
    },
    millet: {
        name: 'Millet Seed',
        rateUnit: 'lbs/acre',
        defaultRate: 25,
        lbsPerBag: 50,
        pricePerBag: 0
    },
    sorghum_sudan: {
        name: 'Sorghum Sudan Seed',
        rateUnit: 'lbs/acre',
        defaultRate: 30,
        lbsPerBag: 50,
        pricePerBag: 0
    }
};

// Pivot Bio Products
const pivotBioProducts = [
    {
        id: 'proven40',
        name: 'PROVEN 40',
        description: 'Nitrogen-fixing microbe for corn',
        infoUrl: 'https://www.pivotbio.com/proven-40',
        rateUnit: 'oz/acre',
        defaultRate: 3,
        pricePerUnit: 0
    },
    {
        id: 'return',
        name: 'RETURN',
        description: 'For small grains and cereals',
        infoUrl: 'https://www.pivotbio.com/return',
        rateUnit: 'oz/acre',
        defaultRate: 3,
        pricePerUnit: 0
    }
];

// Additional Products (standalone items)
const additionalProducts = {
    hydrovant: {
        id: 'hydrovant',
        name: 'Hydrovant',
        description: 'Drift reduction/deposition aid adjuvant',
        unit: 'gal',
        costPerUnit: 95,
        pricing: [
            { minQty: 1, maxQty: 9, pricePerUnit: 145 },
            { minQty: 10, maxQty: 60, pricePerUnit: 145 },
            { minQty: 61, maxQty: 180, pricePerUnit: 145 }
        ]
    },
    multiseal: {
        id: 'multiseal',
        name: 'Multi Seal',
        description: 'Sealant product',
        unit: 'bucket',
        pricePerUnit: 300
    },
    pump: {
        id: 'pump',
        name: 'Pump',
        description: 'Transfer pump',
        unit: 'each',
        pricePerUnit: 130
    }
};

// Helper function to calculate Hydrovant price
function calculateHydrovantPrice(quantity) {
    const pricing = additionalProducts.hydrovant.pricing;
    for (const tier of pricing) {
        if (quantity >= tier.minQty && quantity <= tier.maxQty) {
            return {
                pricePerUnit: tier.pricePerUnit,
                totalPrice: quantity * tier.pricePerUnit,
                tier: `${tier.minQty}-${tier.maxQty} gal`
            };
        }
    }
    // Default to highest tier for quantities over 180
    const lastTier = pricing[pricing.length - 1];
    return {
        pricePerUnit: lastTier.pricePerUnit,
        totalPrice: quantity * lastTier.pricePerUnit,
        tier: `${lastTier.minQty}+ gal`
    };
}

// ============ AUTH MIDDLEWARE ============

// Routes that a user with mustChangePassword=true is still allowed to hit.
// Anything else returns 403 so a determined admin who ignores the frontend
// redirect can't keep using their temp-password session.
const MUST_CHANGE_PASSWORD_ALLOWED_PATHS = new Set([
    '/api/auth/change-password',
    '/api/auth/me',
    '/api/auth/logout'
]);

const authMiddleware = async (req, res, next) => {
    try {
        const token = req.header('Authorization')?.replace('Bearer ', '');
        if (!token) {
            return res.status(401).json({ error: 'Authentication required' });
        }
        const decoded = jwt.verify(token, JWT_SECRET);
        const user = await User.findById(decoded.userId);
        if (!user) {
            return res.status(401).json({ error: 'User not found' });
        }

        // S1: block everything except the change-password / me routes while
        // the user is still on a temp password.
        if (user.mustChangePassword && !MUST_CHANGE_PASSWORD_ALLOWED_PATHS.has(req.path)) {
            return res.status(403).json({
                error: 'Password change required',
                passwordChangeRequired: true
            });
        }

        req.user = user;
        req.token = token;
        next();
    } catch (error) {
        res.status(401).json({ error: 'Invalid token' });
    }
};

const adminMiddleware = async (req, res, next) => {
    // Allow admin, distributor, and superadmin roles
    if (req.user.role !== 'admin' && req.user.role !== 'distributor' && req.user.role !== 'superadmin') {
        return res.status(403).json({ error: 'Admin access required' });
    }
    next();
};

// Helper function to check if user has admin-level access
const isAdminLevel = (user) => {
    return user.role === 'admin' || user.role === 'distributor' || user.role === 'superadmin';
};

// Helper function to check if user is a non-superadmin staff member (distributor/admin)
const isDistributor = (user) => {
    return user.role === 'admin' || user.role === 'distributor';
};

const superAdminMiddleware = async (req, res, next) => {
    if (req.user.role !== 'superadmin') {
        return res.status(403).json({ error: 'Super admin access required' });
    }
    next();
};

const supplierMiddleware = async (req, res, next) => {
    if (req.user.role !== 'supplier') {
        return res.status(403).json({ error: 'Supplier access required' });
    }
    next();
};

// ============ ROUTES ============

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', message: 'Acre Profit API is running' });
});

// Reset/reinitialize admin users (use this if login fails)
// S2: /api/admin/reset-admins was removed. It allowed anyone with the
// hardcoded secret 'acreprofit2026reset' (public in this repo) to delete
// and recreate every admin account. Recovery path now goes through Atlas
// directly - same place the JWT_SECRET and connection string live, so
// recovery requires the same level of access as everything else.

// ---- AUTH ROUTES ----

app.post('/api/auth/signup', signupLimiter, async (req, res) => {
    try {
        const { name, email, password, phone, address, farm, crops, representativeId } = req.body;

        const existingUser = await User.findOne({ email: email.toLowerCase() });
        if (existingUser) {
            return res.status(400).json({ error: 'Email already registered' });
        }

        // Look up the distributor ObjectId based on representativeId
        const repEmails = {
            kyle: 'office@togoag.com',
            ty: 'tymollohan77@gmail.com',
            chad: 'ckbamford@yahoo.com',
            seth: 'seth@acreprofit.com'
        };
        let representativeObjectId = null;
        if (representativeId && repEmails[representativeId]) {
            const distributor = await User.findOne({ email: repEmails[representativeId] });
            if (distributor) {
                representativeObjectId = distributor._id;
            }
        }

        const user = new User({
            name,
            email,
            password,
            phone,
            address,
            farm,
            crops,
            representativeId,
            representative: representativeObjectId, // Set the ObjectId for permission checks
            role: 'customer'
        });
        await user.save();

        const token = jwt.sign({ userId: user._id }, JWT_SECRET, { expiresIn: '30d' });

        res.status(201).json({
            user: {
                id: user._id,
                name: user.name,
                email: user.email,
                phone: user.phone,
                address: user.address,
                role: user.role,
                representativeId: user.representativeId,
                farm: user.farm
            },
            token
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.post('/api/auth/login', loginLimiter, async (req, res) => {
    try {
        const { email, password } = req.body;

        const user = await User.findOne({ email: email.toLowerCase() });
        if (!user) {
            return res.status(401).json({ error: 'Invalid email or password' });
        }

        const isMatch = await user.comparePassword(password);
        if (!isMatch) {
            return res.status(401).json({ error: 'Invalid email or password' });
        }

        // isActive === false check (not !isActive) so existing docs
        // without the field (undefined) are not blocked
        if (user.isActive === false) {
            return res.status(403).json({
                error: 'This account has been archived. Contact your rep or support to reactivate.'
            });
        }

        const token = jwt.sign({ userId: user._id }, JWT_SECRET, { expiresIn: '30d' });

        // Build user response based on role
        const userResponse = {
            id: user._id,
            name: user.name,
            email: user.email,
            role: user.role,
            farm: user.farm
        };

        // Add supplier-specific fields
        if (user.role === 'supplier') {
            userResponse.companyName = user.companyName;
            userResponse.supplierCode = user.supplierCode;
        }

        res.json({
            user: userResponse,
            token,
            passwordChangeRequired: user.mustChangePassword === true
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// S1: change password for an authenticated user. Used by the forced-rotation
// flow (mustChangePassword=true) and available any time after that. Verifies
// the old password, hashes the new via the userSchema.pre('save') hook,
// clears the flag, and returns a fresh JWT so the client doesn't need to
// re-login.
app.post('/api/auth/change-password', authMiddleware, async (req, res) => {
    try {
        const { currentPassword, newPassword } = req.body;
        if (!currentPassword || !newPassword) {
            return res.status(400).json({ error: 'Current and new password are required' });
        }
        if (newPassword.length < 8) {
            return res.status(400).json({ error: 'New password must be at least 8 characters' });
        }
        if (newPassword === currentPassword) {
            return res.status(400).json({ error: 'New password must be different from current password' });
        }

        const user = await User.findById(req.user._id);
        const isMatch = await user.comparePassword(currentPassword);
        if (!isMatch) {
            return res.status(401).json({ error: 'Current password is incorrect' });
        }

        user.password = newPassword;
        user.mustChangePassword = false;
        // Force dirty flag in case Mongoose's change detection skips a boolean
        // field whose new value matches the schema default (false).
        user.markModified('mustChangePassword');
        await user.save();

        // Issue a fresh token so the session continues seamlessly
        const token = jwt.sign({ userId: user._id }, JWT_SECRET, { expiresIn: '30d' });
        res.json({ message: 'Password changed successfully', token });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- PASSWORD RESET ROUTES ----

// Create email transporter
const createEmailTransporter = () => {
    // Use environment variables for email configuration
    // Supports Microsoft 365/Outlook, Gmail, SendGrid, or any SMTP service
    if (process.env.SMTP_HOST) {
        const config = {
            host: process.env.SMTP_HOST,
            port: parseInt(process.env.SMTP_PORT) || 587,
            secure: process.env.SMTP_SECURE === 'true',
            auth: {
                user: process.env.SMTP_USER,
                pass: process.env.SMTP_PASS
            }
        };
        // Office 365 requires TLS
        if (process.env.SMTP_HOST.includes('office365') || process.env.SMTP_HOST.includes('outlook')) {
            config.tls = {
                ciphers: 'SSLv3',
                rejectUnauthorized: false
            };
        }
        return nodemailer.createTransport(config);
    }
    // Default to Gmail if GMAIL_USER and GMAIL_APP_PASSWORD are set
    if (process.env.GMAIL_USER && process.env.GMAIL_APP_PASSWORD) {
        return nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: process.env.GMAIL_USER,
                pass: process.env.GMAIL_APP_PASSWORD
            }
        });
    }
    return null;
};

// Request password reset
app.post('/api/auth/forgot-password', passwordResetLimiter, async (req, res) => {
    try {
        const { email } = req.body;

        if (!email) {
            return res.status(400).json({ error: 'Email is required' });
        }

        const user = await User.findOne({ email: email.toLowerCase() });

        // Always return success to prevent email enumeration
        if (!user) {
            return res.json({ message: 'If an account exists with this email, a password reset link has been sent.' });
        }

        // Generate a secure token
        const token = crypto.randomBytes(32).toString('hex');
        const expiresAt = new Date(Date.now() + 60 * 60 * 1000); // 1 hour from now

        // Invalidate any existing tokens for this user
        await PasswordResetToken.updateMany(
            { userId: user._id, used: false },
            { used: true }
        );

        // Create new token
        await PasswordResetToken.create({
            userId: user._id,
            token,
            expiresAt
        });

        // Send email
        const transporter = createEmailTransporter();
        const resetUrl = `${process.env.FRONTEND_URL || 'https://acreprofit.com'}/reset-password.html?token=${token}`;

        if (transporter) {
            await transporter.sendMail({
                from: process.env.EMAIL_FROM || '"Acre Profit" <noreply@acreprofit.com>',
                to: user.email,
                subject: 'Reset Your Acre Profit Password',
                html: `
                    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                        <div style="background-color: #2d5a27; padding: 20px; text-align: center;">
                            <h1 style="color: white; margin: 0;">Acre Profit</h1>
                        </div>
                        <div style="padding: 30px; background-color: #f9f9f9;">
                            <h2 style="color: #333;">Reset Your Password</h2>
                            <p style="color: #666; line-height: 1.6;">
                                Hi ${user.name},
                            </p>
                            <p style="color: #666; line-height: 1.6;">
                                We received a request to reset your password. Click the button below to create a new password:
                            </p>
                            <div style="text-align: center; margin: 30px 0;">
                                <a href="${resetUrl}" style="background-color: #d4a017; color: white; padding: 14px 28px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">
                                    Reset Password
                                </a>
                            </div>
                            <p style="color: #666; line-height: 1.6;">
                                This link will expire in 1 hour. If you didn't request this, you can safely ignore this email.
                            </p>
                            <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
                            <p style="color: #999; font-size: 12px;">
                                If the button doesn't work, copy and paste this link into your browser:<br>
                                <a href="${resetUrl}" style="color: #2d5a27;">${resetUrl}</a>
                            </p>
                        </div>
                        <div style="padding: 20px; text-align: center; background-color: #333;">
                            <p style="color: #999; font-size: 12px; margin: 0;">
                                &copy; 2026 Acre Profit. Farmers Helping Farmers.
                            </p>
                        </div>
                    </div>
                `
            });
            console.log(`Password reset email sent to ${user.email}`);
        } else {
            console.log(`Password reset requested for ${user.email}. Token: ${token}`);
            console.log(`Reset URL: ${resetUrl}`);
            console.log('Note: Email not sent - no email service configured. Set SMTP or Gmail environment variables.');
        }

        res.json({ message: 'If an account exists with this email, a password reset link has been sent.' });
    } catch (error) {
        console.error('Forgot password error:', error);
        res.status(500).json({ error: 'Failed to process request' });
    }
});

// Reset password with token
app.post('/api/auth/reset-password', passwordResetLimiter, async (req, res) => {
    try {
        const { token, password } = req.body;

        if (!token || !password) {
            return res.status(400).json({ error: 'Token and password are required' });
        }

        if (password.length < 8) {
            return res.status(400).json({ error: 'Password must be at least 8 characters' });
        }

        // Find valid token
        const resetToken = await PasswordResetToken.findOne({
            token,
            used: false,
            expiresAt: { $gt: new Date() }
        });

        if (!resetToken) {
            return res.status(400).json({ error: 'Invalid or expired reset link' });
        }

        // Find user and update password
        const user = await User.findById(resetToken.userId);
        if (!user) {
            return res.status(400).json({ error: 'User not found' });
        }

        // Update password (the pre-save hook will hash it)
        user.password = password;
        await user.save();

        // Mark token as used
        resetToken.used = true;
        await resetToken.save();

        console.log(`Password reset successful for ${user.email}`);

        res.json({ message: 'Password reset successfully' });
    } catch (error) {
        console.error('Reset password error:', error);
        res.status(500).json({ error: 'Failed to reset password' });
    }
});

// Admin: Reset any user's password (superadmin only)
// Get users by role (for rep/distributor dropdowns)
app.get('/api/admin/users', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { role } = req.query;
        const query = {};
        if (role) query.role = role;
        const users = await User.find(query).select('name email role phone').sort({ name: 1 });
        res.json(users);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.post('/api/admin/users/:id/reset-password', authMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can reset passwords' });
        }
        const { newPassword } = req.body;
        if (!newPassword || newPassword.length < 6) {
            return res.status(400).json({ error: 'Password must be at least 6 characters' });
        }
        const user = await User.findById(req.params.id);
        if (!user) {
            return res.status(404).json({ error: 'User not found' });
        }
        user.password = newPassword;
        await user.save();

        await logAudit({
            action: 'password_reset',
            req,
            targetUser: user._id,
            targetUserName: user.name,
            entityType: 'User',
            entityId: user._id,
            reason: `Password reset by ${req.user.name} (${req.user.role})`
        });

        res.json({ message: `Password reset for ${user.email}` });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Superadmin-only: flip any user's mustChangePassword flag on without rotating
// their password. Forces them to rotate on next login. Used to patch accounts
// that slipped past initializeAdmins() migration, and as an operational tool
// when a password is suspected compromised but hasn't been confirmed.
app.post('/api/admin/users/:id/force-password-change', authMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can force password changes' });
        }
        const user = await User.findById(req.params.id);
        if (!user) {
            return res.status(404).json({ error: 'User not found' });
        }
        const before = { mustChangePassword: user.mustChangePassword };
        user.mustChangePassword = true;
        user.markModified('mustChangePassword');
        await user.save();

        await logAudit({
            action: 'force_password_change',
            req,
            targetUser: user._id,
            targetUserName: user.name,
            entityType: 'User',
            entityId: user._id,
            before,
            after: { mustChangePassword: true },
            reason: req.body.reason || `Forced by ${req.user.name} (${req.user.role})`
        });

        res.json({
            message: `${user.email} will be required to change password on next login`,
            user: { id: user._id, email: user.email, mustChangePassword: true }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.get('/api/auth/me', authMiddleware, async (req, res) => {
    const user = await User.findById(req.user._id).populate('representative', 'name email phone');
    res.json({
        user: {
            id: user._id,
            name: user.name,
            email: user.email,
            phone: user.phone,
            role: user.role,
            farm: user.farm,
            crops: user.crops,
            representative: user.representative,
            privateApplicatorLicense: user.privateApplicatorLicense
        }
    });
});

// Get user profile (alias for /api/auth/me)
app.get('/api/user/profile', authMiddleware, async (req, res) => {
    const user = await User.findById(req.user._id).populate('representative', 'name email phone');
    res.json({
        id: user._id,
        name: user.name,
        email: user.email,
        phone: user.phone,
        role: user.role,
        farm: user.farm,
        crops: user.crops,
        representative: user.representative,
        privateApplicatorLicense: user.privateApplicatorLicense,
        commercialApplicatorLicense: user.commercialApplicatorLicense
    });
});

// Update user's applicator license
app.put('/api/user/license', authMiddleware, async (req, res) => {
    try {
        const { privateApplicatorLicense } = req.body;

        if (!privateApplicatorLicense) {
            return res.status(400).json({ error: 'License data required' });
        }

        // User can submit license, but verification status defaults to pending
        const updatedLicense = {
            hasLicense: privateApplicatorLicense.hasLicense,
            licenseNumber: privateApplicatorLicense.licenseNumber,
            state: privateApplicatorLicense.state,
            expirationDate: privateApplicatorLicense.expirationDate,
            verificationStatus: 'pending' // Always pending until admin verifies
        };

        req.user.privateApplicatorLicense = {
            ...req.user.privateApplicatorLicense,
            ...updatedLicense
        };

        await req.user.save();

        res.json({
            message: 'License submitted for verification',
            privateApplicatorLicense: req.user.privateApplicatorLicense
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.put('/api/auth/me', authMiddleware, async (req, res) => {
    try {
        const updates = req.body;
        const allowedUpdates = ['name', 'phone', 'farm', 'crops'];

        Object.keys(updates).forEach(key => {
            if (allowedUpdates.includes(key)) {
                req.user[key] = updates[key];
            }
        });

        // Handle license / certification updates
        const licenseUpdates = {};
        if (updates.privateApplicatorLicense) {
            const incoming = updates.privateApplicatorLicense;
            const existing = req.user.privateApplicatorLicense?.toObject?.() || req.user.privateApplicatorLicense || {};
            req.user.privateApplicatorLicense = {
                ...existing,
                hasLicense: true,
                licenseNumber: incoming.licenseNumber || existing.licenseNumber,
                state: incoming.state || existing.state,
                expirationDate: incoming.expirationDate || existing.expirationDate,
                certificationCategories: incoming.certificationCategories || existing.certificationCategories || [],
                verificationStatus: 'pending' // reset verification on customer self-update
            };
            licenseUpdates.privateApplicatorLicense = req.user.privateApplicatorLicense;
        }
        if (updates.commercialApplicatorLicense) {
            const incoming = updates.commercialApplicatorLicense;
            const existing = req.user.commercialApplicatorLicense?.toObject?.() || req.user.commercialApplicatorLicense || {};
            req.user.commercialApplicatorLicense = {
                ...existing,
                hasLicense: true,
                licenseNumber: incoming.licenseNumber || existing.licenseNumber,
                state: incoming.state || existing.state,
                businessName: incoming.businessName || existing.businessName,
                expirationDate: incoming.expirationDate || existing.expirationDate,
                verificationStatus: 'pending'
            };
            licenseUpdates.commercialApplicatorLicense = req.user.commercialApplicatorLicense;
        }
        if (updates.paraquatCertification) {
            const incoming = updates.paraquatCertification;
            req.user.paraquatCertification = {
                ...(req.user.paraquatCertification?.toObject?.() || req.user.paraquatCertification || {}),
                completed: true,
                certificateNumber: incoming.certificateNumber,
                completionDate: incoming.completionDate,
                expirationDate: incoming.expirationDate
            };
            licenseUpdates.paraquatCertification = req.user.paraquatCertification;
        }
        if (updates.dicambaCertification) {
            const incoming = updates.dicambaCertification;
            req.user.dicambaCertification = {
                ...(req.user.dicambaCertification?.toObject?.() || req.user.dicambaCertification || {}),
                completed: true,
                certificateNumber: incoming.certificateNumber,
                completionDate: incoming.completionDate,
                expirationDate: incoming.expirationDate
            };
            licenseUpdates.dicambaCertification = req.user.dicambaCertification;
        }

        await req.user.save();

        // Audit log license changes
        if (Object.keys(licenseUpdates).length > 0) {
            await logAudit({
                action: 'license_change',
                req,
                targetUser: req.user._id,
                targetUserName: req.user.name,
                entityType: 'User',
                entityId: req.user._id,
                after: licenseUpdates,
                reason: 'Customer self-update via /compliance.html'
            });
        }

        res.json({ user: req.user });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- REPRESENTATIVES ROUTES ----

app.get('/api/representatives', async (req, res) => {
    try {
        const reps = await User.find({ role: { $in: ['admin', 'distributor', 'superadmin'] } })
            .select('name email phone');
        res.json(reps);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Public endpoint: Get active distributors for checkout/pickup location selection
// Price Mining: customer-submitted competitor quotes
const priceMiningQuoteSchema = new mongoose.Schema({
    lines: [{
        product: String,
        quantity: String
    }],
    whenNeeded: String,
    neededBy: Date,      // Machine-filterable needed-by date; whenNeeded stays for soft context (growth stage)
    notes: String,
    // Legacy fields
    product: String,
    supplier: String,
    price: String,
    submittedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    submittedByName: String,
    status: { type: String, enum: ['open', 'quoted', 'ordered', 'closed'], default: 'open' },
    createdAt: { type: Date, default: Date.now }
});
const PriceMiningQuote = mongoose.models.PriceMiningQuote || mongoose.model('PriceMiningQuote', priceMiningQuoteSchema);

app.post('/api/price-mining/submit', async (req, res) => {
    try {
        const { lines, whenNeeded, neededBy, notes, product, supplier, price } = req.body;

        // Support both multi-line and legacy single-line submissions
        const hasLines = lines && Array.isArray(lines) && lines.some(l => l.product);
        const hasLegacy = product;
        if (!hasLines && !hasLegacy) return res.status(400).json({ error: 'Add at least one product' });

        let submittedBy, submittedByName;
        const authHeader = req.header('Authorization');
        if (authHeader) {
            try {
                const decoded = jwt.verify(authHeader.replace('Bearer ', ''), JWT_SECRET);
                const user = await User.findById(decoded.userId);
                if (user) {
                    submittedBy = user._id;
                    submittedByName = user.name;
                }
            } catch (e) { /* anonymous ok */ }
        }

        const quote = await PriceMiningQuote.create({
            lines: hasLines ? lines.filter(l => l.product) : [{ product, quantity: supplier || '' }],
            whenNeeded,
            neededBy: neededBy ? new Date(neededBy) : null,
            notes: notes || price,
            submittedBy,
            submittedByName
        });
        res.json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.get('/api/price-mining/quotes', async (req, res) => {
    try {
        const quotes = await PriceMiningQuote.find().sort({ createdAt: -1 }).limit(100);
        res.json(quotes);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.get('/api/distributors', async (req, res) => {
    try {
        const distributors = await User.find({ role: { $in: ['admin', 'distributor', 'superadmin'] } })
            .select('name email phone')
            .sort({ name: 1 });

        // Return distributors with derived pickup location key (last name lowercase)
        const result = distributors.map(d => ({
            _id: d._id,
            name: d.name,
            email: d.email,
            phone: d.phone,
            locationKey: d.name.split(' ').pop().toLowerCase()
        }));

        res.json(result);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- SPRAY PROGRAM ROUTES ----

app.get('/api/crops', (req, res) => {
    const crops = Object.keys(sprayPrograms).map(crop => ({
        id: crop,
        name: crop.charAt(0).toUpperCase() + crop.slice(1),
        programs: Object.keys(sprayPrograms[crop])
    }));
    res.json(crops);
});

app.get('/api/crops/:crop/programs', (req, res) => {
    const { crop } = req.params;
    const programs = sprayPrograms[crop.toLowerCase()];

    if (!programs) {
        return res.status(404).json({ error: 'Crop not found' });
    }

    const programList = Object.entries(programs).map(([id, program]) => ({
        id,
        name: program.name,
        description: program.description,
        chemicalCount: program.chemicals.length
    }));

    res.json(programList);
});

app.get('/api/crops/:crop/programs/:programId', authMiddleware, (req, res) => {
    const { crop, programId } = req.params;
    const program = sprayPrograms[crop.toLowerCase()]?.[programId];

    if (!program) {
        return res.status(404).json({ error: 'Program not found' });
    }

    res.json(program);
});

// ---- SEED ROUTES ----

app.get('/api/seeds', (req, res) => {
    res.json(seedProducts);
});

app.post('/api/seeds/calculate', authMiddleware, (req, res) => {
    const { seedType, acres, customRate } = req.body;
    const seed = seedProducts[seedType];

    if (!seed) {
        return res.status(404).json({ error: 'Seed type not found' });
    }

    const rate = customRate || seed.defaultRate;
    let bagsNeeded, totalAmount;

    if (seed.rateUnit === 'seeds/acre') {
        totalAmount = rate * acres;
        bagsNeeded = Math.ceil(totalAmount / seed.seedsPerBag);
    } else {
        totalAmount = rate * acres;
        bagsNeeded = Math.ceil(totalAmount / seed.lbsPerBag);
    }

    res.json({
        seed: seed.name,
        acres,
        rate,
        rateUnit: seed.rateUnit,
        totalAmount,
        bagsNeeded,
        unitsPerBag: seed.seedsPerBag || seed.lbsPerBag,
        pricePerBag: seed.pricePerBag,
        totalPrice: bagsNeeded * seed.pricePerBag
    });
});

// ---- PIVOT BIO ROUTES ----

app.get('/api/pivot-bio', (req, res) => {
    res.json(pivotBioProducts);
});

// ---- ADDITIONAL PRODUCTS ROUTES ----

app.get('/api/products', (req, res) => {
    res.json(additionalProducts);
});

app.post('/api/products/hydrovant/calculate', (req, res) => {
    const { quantity } = req.body;
    if (!quantity || quantity < 1) {
        return res.status(400).json({ error: 'Quantity must be at least 1' });
    }
    const result = calculateHydrovantPrice(quantity);
    res.json({
        product: 'Hydrovant',
        quantity,
        ...result
    });
});

app.post('/api/products/calculate', (req, res) => {
    const { productId, quantity } = req.body;
    const product = additionalProducts[productId];

    if (!product) {
        return res.status(404).json({ error: 'Product not found' });
    }

    if (productId === 'hydrovant') {
        const result = calculateHydrovantPrice(quantity);
        return res.json({
            product: product.name,
            quantity,
            unit: product.unit,
            ...result
        });
    }

    res.json({
        product: product.name,
        quantity,
        unit: product.unit,
        pricePerUnit: product.pricePerUnit,
        totalPrice: quantity * product.pricePerUnit
    });
});

// ---- ORDER/CALCULATION ROUTES ----

app.post('/api/calculate', authMiddleware, (req, res) => {
    try {
        const { crop, programId, acres, customRates } = req.body;

        const program = sprayPrograms[crop.toLowerCase()]?.[programId];
        if (!program) {
            return res.status(404).json({ error: 'Program not found' });
        }

        const calculations = program.chemicals.map(chemical => {
            const rate = customRates?.[chemical.name] ?? chemical.defaultRate;

            let totalGallons;
            switch (chemical.rateUnit) {
                case 'oz/acre':
                    totalGallons = (rate * acres) / 128;
                    break;
                case 'pt/acre':
                    totalGallons = (rate * acres) / 8;
                    break;
                case 'qt/acre':
                    totalGallons = (rate * acres) / 4;
                    break;
                case 'gal/acre':
                    totalGallons = rate * acres;
                    break;
                case 'lb/acre':
                    totalGallons = rate * acres;
                    break;
                default:
                    totalGallons = rate * acres;
            }

            const packagesNeeded = Math.ceil(totalGallons / chemical.packageSize);
            const totalPrice = packagesNeeded * (chemical.pricePerPackage || 0);

            return {
                name: chemical.name,
                rate,
                rateUnit: chemical.rateUnit,
                totalAmount: Math.round(totalGallons * 100) / 100,
                totalUnit: chemical.rateUnit.includes('lb') ? 'lb' : 'gal',
                packageSize: chemical.packageSize,
                packageUnit: chemical.packageUnit,
                packagesNeeded,
                pricePerPackage: chemical.pricePerPackage || 0,
                totalPrice
            };
        });

        const totalPrice = calculations.reduce((sum, c) => sum + c.totalPrice, 0);
        const costPerAcre = acres > 0 ? Math.round((totalPrice / acres) * 100) / 100 : 0;

        res.json({
            crop,
            program: program.name,
            acres,
            chemicals: calculations,
            totalPrice,
            costPerAcre
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Save order
app.post('/api/orders', authMiddleware, async (req, res) => {
    try {
        const { crop, programId, acres, chemicals, seeds, pivotBio, representativeId, totalPrice } = req.body;

        const costPerAcre = acres > 0 ? Math.round((totalPrice / acres) * 100) / 100 : 0;

        const order = new Order({
            userId: req.user._id,
            representativeId: representativeId || req.user.representative,
            crop,
            program: programId,
            acres,
            chemicals,
            seeds,
            pivotBio,
            totalCost: totalPrice,
            costPerAcre,
            status: 'draft'
        });

        await order.save();
        res.status(201).json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get user's orders
app.get('/api/orders', authMiddleware, async (req, res) => {
    try {
        const orders = await Order.find({ userId: req.user._id })
            .populate('representativeId', 'name email phone')
            .sort({ createdAt: -1 });
        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single order (for spray sheet)
app.get('/api/orders/:orderId', authMiddleware, async (req, res) => {
    try {
        const order = await Order.findOne({ _id: req.params.orderId, userId: req.user._id })
            .populate('representativeId', 'name email phone')
            .populate('userId', 'name email phone farm');

        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Submit order
app.put('/api/orders/:orderId/submit', authMiddleware, async (req, res) => {
    try {
        const order = await Order.findOne({ _id: req.params.orderId, userId: req.user._id });
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        order.status = 'submitted';
        order.updatedAt = new Date();
        await order.save();

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- ADMIN ROUTES ----

// Get all customers (admin only) - with search support
// Create a new customer (admin only)
app.post('/api/admin/customers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { name, email, password, phone, farm, crops, state, acres, representativeId } = req.body;

        if (!name || !email) {
            return res.status(400).json({ error: 'Name and email are required' });
        }

        const existingUser = await User.findOne({ email: email.toLowerCase() });
        if (existingUser) {
            return res.status(400).json({ error: 'Email already registered' });
        }

        const user = new User({
            name,
            email: email.toLowerCase(),
            password: password || 'Farm2026!',
            phone,
            farm: {
                name: farm || '',
                acres: acres || 0,
                state: state || ''
            },
            crops: crops || [],
            representativeId: representativeId || 'kyle', // String rep ID for pickup location
            representative: req.user._id, // ObjectId of admin who created
            role: 'customer'
        });

        await user.save();

        const savedUser = await User.findById(user._id)
            .select('-password')
            .populate('representative', 'name email');

        res.status(201).json(savedUser);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.get('/api/admin/customers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { search } = req.query;
        let query = { role: 'customer' };

        // Hide archived customers by default.
        // Superadmin can pass ?includeArchived=true to see them.
        if (req.query.includeArchived !== 'true' || req.user.role !== 'superadmin') {
            query.isActive = { $ne: false };
        }

        // If not superadmin, only show their own customers
        if (isDistributor(req.user)) {
            query.representative = req.user._id;
        }

        // Add search filter if provided
        if (search) {
            const searchRegex = new RegExp(search, 'i');
            query.$or = [
                { name: searchRegex },
                { email: searchRegex },
                { 'farm.name': searchRegex },
                { 'farm.state': searchRegex }
            ];
        }

        const customers = await User.find(query)
            .select('-password')
            .populate('representative', 'name email')
            .sort({ name: 1 });
        res.json(customers);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ MULTI-COLLECTION ORDER READ HELPERS ============
// Admin-created orders land in the legacy Order collection (POST /api/admin/
// orders/for-customer at line 5296 writes new Order({...})). Customer-placed
// orders land in ChemicalOrder. These helpers merge both so admin read paths
// surface a complete picture without the frontend needing to know about two
// collections. Write paths are unchanged - cleanup of the legacy write is a
// separate migration conversation.

// Normalize a doc for list views. Frontend modal card reads crop/acres/
// totalCost; ChemicalOrder uses sprayParams.crop/totalAcres/total, so we
// project onto the legacy names. Detail view already handles both shapes.
function normalizeListOrder(doc, source) {
    const o = doc && doc.toObject ? doc.toObject() : (doc || {});
    if (source === 'chemical') {
        if (o.crop === undefined || o.crop === null) {
            o.crop = o.sprayParams?.crop || o.orderType || 'General';
        }
        if (o.acres === undefined || o.acres === null) {
            o.acres = o.totalAcres || 0;
        }
        if (o.totalCost === undefined || o.totalCost === null) {
            o.totalCost = o.total || 0;
        }
    }
    o._sourceCollection = source;
    return o;
}

// List fetch across both collections. Merges, sorts by createdAt desc,
// returns optional slice. totalCount reflects the full union, not the slice.
async function findOrdersInBothCollections(userQuery, opts = {}) {
    const { limit = null } = opts;
    const [legacy, chem] = await Promise.all([
        Order.find(userQuery)
            .populate('userId', 'name email phone farm')
            .populate('representativeId', 'name email')
            .sort({ createdAt: -1 })
            .lean(),
        ChemicalOrder.find(userQuery)
            .populate('userId', 'name email phone farm')
            .populate('representativeId', 'name email')
            .sort({ createdAt: -1 })
            .lean()
    ]);
    const merged = [
        ...legacy.map(o => normalizeListOrder(o, 'legacy')),
        ...chem.map(o => normalizeListOrder(o, 'chemical'))
    ].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    const totalCount = legacy.length + chem.length;
    const orders = limit ? merged.slice(0, limit) : merged;
    return { orders, totalCount };
}

// Detail fetch: try Order first, fall back to ChemicalOrder on miss.
// No normalization - displayOrderDetail() already handles both shapes.
// Stamps _sourceCollection so future frontend code can branch if needed.
async function findOrderInEitherCollection(query) {
    const legacyDoc = await Order.findOne(query)
        .populate('userId', 'name email phone farm')
        .populate('representativeId', 'name email');
    if (legacyDoc) {
        const obj = legacyDoc.toObject();
        obj._sourceCollection = 'legacy';
        return obj;
    }
    const chemDoc = await ChemicalOrder.findOne(query)
        .populate('userId', 'name email phone farm')
        .populate('representativeId', 'name email');
    if (chemDoc) {
        const obj = chemDoc.toObject();
        obj._sourceCollection = 'chemical';
        return obj;
    }
    return null;
}

// Archive / unarchive a customer (superadmin only)
app.put('/api/admin/customers/:id/archive', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const { archive } = req.body; // true = archive, false = unarchive
        const customer = await User.findOne({ _id: req.params.id, role: 'customer' });
        if (!customer) return res.status(404).json({ error: 'Customer not found' });

        customer.isActive = archive === true ? false : true;
        await customer.save();

        res.json({
            message: archive ? 'Customer archived' : 'Customer reactivated',
            customer: { _id: customer._id, name: customer.name, isActive: customer.isActive }
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single customer details (admin only)
app.get('/api/admin/customers/:customerId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const customer = await User.findById(req.params.customerId)
            .select('-password')
            .populate('representative', 'name email phone');

        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Check access for non-superadmin (distributors can only see their customers)
        if (isDistributor(req.user) &&
            customer.representative?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Access denied' });
        }

        // Get order count and recent orders from both collections
        const { orders, totalCount } = await findOrdersInBothCollections(
            { userId: customer._id },
            { limit: 10 }
        );

        // Get recent invoices built for this customer (by any distributor)
        const invoices = await Invoice.find({ customerId: customer._id })
            .populate('representativeId', 'name email')
            .sort({ invoiceDate: -1, createdAt: -1 })
            .limit(10)
            .lean();
        const invoiceCount = await Invoice.countDocuments({ customerId: customer._id });

        res.json({
            ...customer.toObject(),
            orders,
            orderCount: totalCount,
            invoices,
            invoiceCount
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update customer details (admin only)
app.put('/api/admin/customers/:customerId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const customer = await User.findById(req.params.customerId);

        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Check access for non-superadmin (distributors can only edit their customers)
        if (isDistributor(req.user) &&
            customer.representative?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Access denied' });
        }

        const { name, email, phone, farm, state, acres, crops, privateApplicatorLicense } = req.body;

        // Update fields if provided
        if (name !== undefined) customer.name = name;
        if (email !== undefined) customer.email = email.toLowerCase();
        if (phone !== undefined) customer.phone = phone;

        // Update farm as an object
        if (farm !== undefined || state !== undefined || acres !== undefined) {
            customer.farm = {
                ...customer.farm,
                name: farm !== undefined ? farm : customer.farm?.name,
                state: state !== undefined ? state : customer.farm?.state,
                acres: acres !== undefined ? acres : customer.farm?.acres
            };
        }

        if (crops !== undefined) customer.crops = crops;

        // Update private applicator license if provided
        if (privateApplicatorLicense !== undefined) {
            customer.privateApplicatorLicense = {
                ...(customer.privateApplicatorLicense?.toObject?.() || customer.privateApplicatorLicense || {}),
                ...privateApplicatorLicense,
                verifiedBy: privateApplicatorLicense.hasLicense ? req.user._id : undefined,
                verifiedAt: privateApplicatorLicense.hasLicense ? new Date() : undefined
            };
        }

        customer.updatedAt = new Date();
        await customer.save();

        res.json(customer);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Reassign customer to a different distributor (superadmin only)
app.put('/api/admin/customers/:customerId/assign', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const { distributorId } = req.body;
        const customer = await User.findById(req.params.customerId);

        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Find the distributor
        const distributor = await User.findById(distributorId);
        if (!distributor || !['admin', 'distributor'].includes(distributor.role)) {
            return res.status(404).json({ error: 'Distributor not found' });
        }

        // Update customer's representative
        customer.representative = distributor._id;

        // Also update representativeId string based on distributor email
        const emailToRepId = {
            'office@togoag.com': 'kyle',
            'tymollohan77@gmail.com': 'ty',
            'ckbamford@yahoo.com': 'chad',
            'seth@acreprofit.com': 'seth'
        };
        customer.representativeId = emailToRepId[distributor.email] || customer.representativeId;

        await customer.save();

        const updated = await User.findById(customer._id)
            .select('-password')
            .populate('representative', 'name email');

        res.json(updated);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all orders for a specific customer (admin only)
app.get('/api/admin/customers/:customerId/orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const customer = await User.findById(req.params.customerId);

        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Check access for non-superadmin (distributors can only see their customers' orders)
        if (isDistributor(req.user) &&
            customer.representative?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Access denied' });
        }

        const { orders } = await findOrdersInBothCollections({ userId: customer._id });

        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create order on behalf of a customer (admin only)
app.post('/api/admin/orders/for-customer', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { customerId, crop, programId, acres, gpa, year, chemicals, seeds, pivotBio, totalPrice, status, notes, discountCode } = req.body;

        // Validate customer exists and admin has access
        const customer = await User.findById(customerId);
        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Check access for non-superadmin (distributors can only create orders for their customers)
        if (isDistributor(req.user) &&
            customer.representative?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Access denied - not your customer' });
        }

        // RUP COMPLIANCE CHECK - validates against the CUSTOMER's licenses, not the distributor's
        const allChemicalsForRupCheck = await Chemical.find({ isActive: true })
            .select('productName isRestrictedUse requiredCertifications').lean();
        const itemsForRupCheck = (chemicals || []).map(c => ({
            productName: c.name || c.productName
        }));
        const rupCompliance = await validateRupCompliance({
            customer,
            items: itemsForRupCheck,
            allChemicals: allChemicalsForRupCheck
        });
        if (!rupCompliance.ok) {
            await logAudit({
                action: 'rup_block',
                req,
                targetUser: customer._id,
                targetUserName: customer.name,
                entityType: 'Order',
                reason: 'RUP compliance check failed - blocked at admin order creation',
                after: { errors: rupCompliance.errors }
            });
            return res.status(403).json({
                error: 'Restricted Use Pesticide compliance check failed for this customer',
                rupErrors: rupCompliance.errors,
                userMessage: `${customer.name} does not have valid licenses for one or more products in this order.`
            });
        }

        // Validate discount code
        let discountType = null;
        let discountDescription = '';
        if (discountCode) {
            const code = discountCode.trim().toUpperCase();
            if (code === 'NODISTMARG') {
                discountType = 'no_dist_margin';
                discountDescription = 'No distributor margin (admin price)';
            } else if (code === 'ATCOSTAP') {
                discountType = 'at_cost';
                discountDescription = 'At cost (no markup)';
            }
        }

        // Look up chemical pricing for discount code application
        const allChemicals = discountType ? await Chemical.find({ isActive: true }).lean() : [];

        const costPerAcre = acres > 0 ? Math.round((totalPrice / acres) * 100) / 100 : 0;

        // Normalize chemicals array and apply discount code pricing
        let totalDiscount = 0;
        const normalizedChemicals = (chemicals || []).map(c => {
            const name = c.name || c.productName;
            const qty = c.qty || c.quantity || 0;
            let pricePerUnit = c.pricePerUnit || c.price || 0;

            // Apply discount code pricing from server-side chemical data
            if (discountType && name) {
                const chem = allChemicals.find(ch =>
                    ch.productName === name || ch.productName.toLowerCase() === name.toLowerCase()
                );
                if (chem) {
                    const fullPrice = chem.sellPrice || pricePerUnit;
                    if (discountType === 'at_cost') {
                        pricePerUnit = chem.costPrice || fullPrice;
                    } else if (discountType === 'no_dist_margin') {
                        pricePerUnit = chem.adminPrice || fullPrice;
                    }
                    totalDiscount += (fullPrice - pricePerUnit) * qty;
                }
            }

            return {
                name,
                qty,
                unit: c.unit || 'gal',
                pricePerUnit,
                totalPrice: Math.round(qty * pricePerUnit * 100) / 100,
                chemicalId: c.chemicalId,
                packSize: c.packSize,
                sourceSupplier: c.sourceSupplier
            };
        });

        // Recalculate total from adjusted line items if discount applied
        const adjustedTotal = discountType
            ? normalizedChemicals.reduce((sum, c) => sum + (c.totalPrice || 0), 0)
            : totalPrice;
        const adjustedCostPerAcre = acres > 0 ? Math.round((adjustedTotal / acres) * 100) / 100 : 0;

        // Atomic transaction: order + inventory reservations must succeed together or not at all
        let order;
        const dbSession = await mongoose.startSession();
        try {
            await dbSession.withTransaction(async () => {
                order = new Order({
                    userId: customerId,
                    representativeId: req.user._id,
                    crop,
                    program: programId,
                    acres,
                    gpa,
                    year: year || new Date().getFullYear(),
                    chemicals: normalizedChemicals,
                    seeds,
                    pivotBio,
                    totalCost: adjustedTotal,
                    costPerAcre: adjustedCostPerAcre,
                    discount: Math.round(totalDiscount * 100) / 100,
                    discountReason: discountDescription || undefined,
                    status: status || 'draft',
                    notes,
                    createdBy: req.user._id
                });

                await order.save({ session: dbSession });

                // Reserve inventory + calculate commissions atomically
                const chemicalIds = (chemicals || []).map(c => c.chemicalId).filter(Boolean);
                if (chemicalIds.length > 0) {
                    const chemicalPricing = await Chemical.find({ _id: { $in: chemicalIds } })
                        .select('productName costPrice adminPrice sellPrice marginDollars adminMarginDollars')
                        .session(dbSession);

                    const pricingMap = {};
                    chemicalPricing.forEach(c => { pricingMap[c._id.toString()] = c; });

                    let totalRepCommission = 0;
                    let totalAdminRevenue = 0;

                    for (const chem of normalizedChemicals) {
                        const qty = chem.qty || chem.quantity || 0;
                        if (chem.chemicalId && qty > 0) {
                            await reserveInventory({
                                chemicalId: chem.chemicalId,
                                quantity: qty,
                                location: 'main',
                                orderId: order._id,
                                orderNumber: `ORD-${order._id.toString().slice(-8).toUpperCase()}`,
                                userId: req.user._id,
                                notes: `Reserved for order - ${crop}`,
                                session: dbSession
                            });

                            const pricing = pricingMap[chem.chemicalId.toString()];
                            if (pricing) {
                                totalRepCommission += (pricing.marginDollars || 0) * qty;
                                totalAdminRevenue += (pricing.adminMarginDollars || 0) * qty;
                            }
                        }
                    }

                    if (totalRepCommission > 0 || totalAdminRevenue > 0) {
                        order.repCommission = totalRepCommission;
                        order.adminRevenue = totalAdminRevenue;
                        await order.save({ session: dbSession });
                    }
                }
            });
        } catch (txErr) {
            await dbSession.endSession();
            console.error('Order transaction failed - rolled back:', txErr.message);
            return res.status(500).json({ error: 'Order could not be completed. No changes saved. Please try again.' });
        }
        await dbSession.endSession();

        // Send order confirmation email to customer
        const transporter = createEmailTransporter();
        if (transporter && customer.email) {
            try {
                // Fetch chemical details to get label URLs (already have chemicalIds from above)
                const chemicalDetails = chemicalIds.length > 0
                    ? await Chemical.find({ _id: { $in: chemicalIds } }).select('productName labelUrl sdsUrl')
                    : [];

                const chemicalLabelMap = {};
                chemicalDetails.forEach(c => {
                    chemicalLabelMap[c._id.toString()] = { labelUrl: c.labelUrl, sdsUrl: c.sdsUrl };
                });

                const chemicalsList = (chemicals || []).map(c => {
                    const details = chemicalLabelMap[c.chemicalId] || {};
                    const labelLink = details.labelUrl
                        ? `<a href="${details.labelUrl}" style="color: #4a7c59; margin-left: 8px;">View Label</a>`
                        : '';
                    return `<li>${c.productName || c.name} - ${c.qty || c.quantity} ${c.unit || 'units'}${labelLink}</li>`;
                }).join('');

                const frontendUrl = process.env.FRONTEND_URL || 'https://acreprofit.com';

                await transporter.sendMail({
                    from: process.env.EMAIL_FROM || '"Acre Profit" <noreply@acreprofit.com>',
                    to: customer.email,
                    subject: `Order Created - ${crop} (${acres} acres)`,
                    html: `
                        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                            <div style="background: #4a7c59; color: white; padding: 20px; text-align: center;">
                                <h1 style="margin: 0;">Acre Profit</h1>
                            </div>
                            <div style="padding: 20px; background: #f9f9f9;">
                                <h2>Order Created</h2>
                                <p>Hi ${customer.name},</p>
                                <p>An order has been created on your behalf by your representative.</p>

                                <div style="background: white; padding: 15px; border-radius: 5px; margin: 15px 0;">
                                    <h3 style="margin-top: 0;">Order Details</h3>
                                    <p><strong>Crop:</strong> ${crop}</p>
                                    <p><strong>Acres:</strong> ${acres.toLocaleString()}</p>
                                    <p><strong>Year:</strong> ${year || new Date().getFullYear()}</p>
                                    <p><strong>Total:</strong> $${totalPrice.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})}</p>
                                    <p><strong>Status:</strong> ${status || 'Draft'}</p>
                                    ${chemicals && chemicals.length > 0 ? `
                                        <h4>Products:</h4>
                                        <ul>${chemicalsList}</ul>
                                    ` : ''}
                                    ${notes ? `<p><strong>Notes:</strong> ${notes}</p>` : ''}
                                </div>

                                <div style="background: #d1fae5; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #10b981;">
                                    <h4 style="margin-top: 0; color: #065f46;">Ready to Pay?</h4>
                                    <p style="margin-bottom: 10px; color: #065f46;">Pay securely via bank transfer (ACH) - no fees, fast processing.</p>
                                    <a href="${frontendUrl}/my-orders.html" style="background: #10b981; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; font-weight: 600;">Pay with ACH Bank Transfer</a>
                                    <p style="margin-top: 10px; font-size: 12px; color: #065f46;">Or pay by check - contact your representative for details.</p>
                                </div>

                                <div style="background: #fff3cd; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #ffc107;">
                                    <h4 style="margin-top: 0; color: #856404;">Product Labels & Safety Data Sheets</h4>
                                    <p style="margin-bottom: 10px; color: #856404;">Access EPA-approved labels and SDS documents for all products in your order:</p>
                                    <a href="${frontendUrl}/chemical-docs.html" style="background: #4a7c59; color: white; padding: 8px 16px; text-decoration: none; border-radius: 5px; display: inline-block;">View All Labels & SDS</a>
                                </div>

                                <p>Log in to your dashboard to view and manage your order:</p>
                                <p><a href="${frontendUrl}/dashboard.html" style="background: #4a7c59; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">View Dashboard</a></p>

                                <p>If you have any questions, please contact your representative.</p>
                            </div>
                            <div style="padding: 15px; text-align: center; color: #666; font-size: 12px;">
                                <p>&copy; ${new Date().getFullYear()} Acre Profit. All rights reserved.</p>
                            </div>
                        </div>
                    `
                });
                console.log(`Order confirmation email sent to ${customer.email}`);
            } catch (emailError) {
                console.error('Failed to send order confirmation email:', emailError.message);
                // Don't fail the order creation if email fails
            }
        }

        res.status(201).json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Edit an order (admin only)
app.put('/api/admin/orders/:orderId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { crop, acres, chemicals, seeds, pivotBio, totalPrice, totalCost, status } = req.body;

        let query = { _id: req.params.orderId };

        // If not superadmin, can only edit their own customers' orders
        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        const order = await Order.findOne(query);
        if (!order) {
            return res.status(404).json({ error: 'Order not found or access denied' });
        }

        // Track inventory changes if chemicals are being modified
        if (chemicals !== undefined && order.status !== 'cancelled' && order.status !== 'archived') {
            // Convert Order chemicals format to the format updateOrderInventory expects
            const oldItems = (order.chemicals || []).map(c => ({
                chemicalId: c.chemicalId,
                quantity: c.packagesNeeded || c.qty || 0
            }));
            const newItems = chemicals.map(c => ({
                chemicalId: c.chemicalId,
                quantity: c.packagesNeeded || c.qty || c.quantity || 0
            }));

            try {
                await updateOrderInventory({
                    oldItems,
                    newItems,
                    orderId: order._id,
                    orderNumber: `ORD-${order._id.toString().slice(-8).toUpperCase()}`,
                    location: 'main',
                    userId: req.user._id
                });
            } catch (invError) {
                console.error('Inventory update warning:', invError.message);
                // Continue with order update even if inventory tracking fails
            }

            // Recalculate commissions when chemicals change
            try {
                const chemicalIds = chemicals.map(c => c.chemicalId).filter(Boolean);
                if (chemicalIds.length > 0) {
                    const chemicalPricing = await Chemical.find({ _id: { $in: chemicalIds } })
                        .select('productName costPrice adminPrice sellPrice marginDollars adminMarginDollars');

                    const pricingMap = {};
                    chemicalPricing.forEach(c => {
                        pricingMap[c._id.toString()] = c;
                    });

                    let totalRepCommission = 0;
                    let totalAdminRevenue = 0;

                    for (const chem of chemicals) {
                        const qty = chem.packagesNeeded || chem.qty || chem.quantity || 0;
                        if (chem.chemicalId && qty > 0) {
                            const pricing = pricingMap[chem.chemicalId.toString()];
                            if (pricing) {
                                totalRepCommission += (pricing.marginDollars || 0) * qty;
                                totalAdminRevenue += (pricing.adminMarginDollars || 0) * qty;
                            }
                        }
                    }

                    order.repCommission = totalRepCommission;
                    order.adminRevenue = totalAdminRevenue;
                }
            } catch (commErr) {
                console.error('Commission calculation warning:', commErr.message);
            }
        }

        // Update fields
        if (crop !== undefined) order.crop = crop;
        if (acres !== undefined) order.acres = acres;
        if (chemicals !== undefined) order.chemicals = chemicals;
        if (seeds !== undefined) order.seeds = seeds;
        if (pivotBio !== undefined) order.pivotBio = pivotBio;

        // Handle both totalPrice and totalCost (schema uses totalCost)
        const newTotal = totalCost !== undefined ? totalCost : totalPrice;
        if (newTotal !== undefined) {
            order.totalCost = newTotal;
            order.costPerAcre = order.acres > 0 ? Math.round((newTotal / order.acres) * 100) / 100 : 0;
        }
        if (status !== undefined) order.status = status;

        order.updatedAt = new Date();
        order.lastEditedBy = req.user._id;

        await order.save();
        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single order by ID (admin only)
app.get('/api/admin/orders/:orderId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = { _id: req.params.orderId };

        // If not superadmin, can only view their own customers' orders
        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        const order = await findOrderInEitherCollection(query);

        if (!order) {
            return res.status(404).json({ error: 'Order not found or access denied' });
        }

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all orders (admin only)
app.get('/api/admin/orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = {};

        // If not superadmin, only show orders for their customers
        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        const { orders } = await findOrdersInBothCollections(query);
        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update order status (admin only)
app.put('/api/admin/orders/:orderId/status', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status } = req.body;
        let query = { _id: req.params.orderId };

        // If not superadmin, can only update their own customers' orders
        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        // Get order first to check status and handle inventory
        const order = await Order.findOne(query);
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        const oldStatus = order.status;

        // Release inventory if order is being cancelled or archived
        if ((status === 'cancelled' || status === 'archived') &&
            oldStatus !== 'cancelled' && oldStatus !== 'archived') {
            try {
                for (const chem of order.chemicals || []) {
                    if (chem.chemicalId && (chem.packagesNeeded || chem.qty) > 0) {
                        await releaseInventory({
                            chemicalId: chem.chemicalId,
                            quantity: chem.packagesNeeded || chem.qty,
                            location: 'main',
                            orderId: order._id,
                            orderNumber: `ORD-${order._id.toString().slice(-8).toUpperCase()}`,
                            userId: req.user._id,
                            notes: `Order ${status}: inventory released`
                        });
                    }
                }
            } catch (invError) {
                console.error('Inventory release warning:', invError.message);
            }
        }

        // Update the order
        order.status = status;
        order.updatedAt = new Date();
        await order.save();

        // Auto-create ledger entry when regular order is delivered
        if (status === 'delivered' && order.representativeId) {
            const existingEntry = await LedgerEntry.findOne({
                referenceType: 'Order',
                referenceId: order._id,
                category: 'order'
            });

            if (!existingEntry) {
                await createLedgerEntry({
                    representativeId: order.representativeId,
                    description: `Order delivered - ${order.crop} program, ${order.acres} acres`,
                    amount: order.totalCost || 0,
                    type: 'debit',
                    category: 'order',
                    referenceType: 'Order',
                    referenceId: order._id,
                    createdBy: req.user._id,
                    notes: 'Auto-created on delivery'
                });
            }
        }

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Bulk update order statuses (admin only)
app.put('/api/admin/orders/bulk-status', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { orderIds, status } = req.body;

        if (!orderIds || !Array.isArray(orderIds) || orderIds.length === 0) {
            return res.status(400).json({ error: 'No orders selected' });
        }

        if (!status) {
            return res.status(400).json({ error: 'Status is required' });
        }

        const results = { success: [], failed: [] };

        for (const orderId of orderIds) {
            try {
                let query = { _id: orderId };

                // If not superadmin, can only update their own customers' orders
                if (isDistributor(req.user)) {
                    query.representativeId = req.user._id;
                }

                const order = await Order.findOne(query);
                if (!order) {
                    results.failed.push({ orderId, error: 'Order not found' });
                    continue;
                }

                const oldStatus = order.status;

                // Release inventory if order is being cancelled or archived
                if ((status === 'cancelled' || status === 'archived') &&
                    oldStatus !== 'cancelled' && oldStatus !== 'archived') {
                    try {
                        for (const chem of order.chemicals || []) {
                            if (chem.chemicalId && (chem.packagesNeeded || chem.qty) > 0) {
                                await releaseInventory({
                                    chemicalId: chem.chemicalId,
                                    quantity: chem.packagesNeeded || chem.qty,
                                    location: 'main',
                                    orderId: order._id,
                                    orderNumber: `ORD-${order._id.toString().slice(-8).toUpperCase()}`,
                                    userId: req.user._id,
                                    notes: `Bulk ${status}: inventory released`
                                });
                            }
                        }
                    } catch (invError) {
                        console.error('Inventory release warning:', invError.message);
                    }
                }

                // Update the order
                order.status = status;
                order.updatedAt = new Date();
                await order.save();

                // Auto-create ledger entry when regular order is delivered
                if (status === 'delivered' && order.representativeId) {
                    const existingEntry = await LedgerEntry.findOne({
                        referenceType: 'Order',
                        referenceId: order._id,
                        category: 'order'
                    });

                    if (!existingEntry) {
                        await createLedgerEntry({
                            representativeId: order.representativeId,
                            description: `Order delivered - ${order.crop} program, ${order.acres} acres`,
                            amount: order.totalCost || 0,
                            type: 'debit',
                            category: 'order',
                            referenceType: 'Order',
                            referenceId: order._id,
                            createdBy: req.user._id,
                            notes: 'Bulk update - auto-created on delivery'
                        });
                    }
                }

                results.success.push(orderId);
            } catch (err) {
                results.failed.push({ orderId, error: err.message });
            }
        }

        res.json({
            message: `Updated ${results.success.length} orders`,
            success: results.success,
            failed: results.failed
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Archive order (admin only) - soft delete for order issues
app.put('/api/admin/orders/:orderId/archive', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = { _id: req.params.orderId };

        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        const order = await Order.findOne(query);
        if (!order) {
            return res.status(404).json({ error: 'Order not found or access denied' });
        }

        // Release inventory if not already cancelled/archived
        if (order.status !== 'cancelled' && order.status !== 'archived') {
            try {
                for (const chem of order.chemicals || []) {
                    if (chem.chemicalId && (chem.packagesNeeded || chem.qty) > 0) {
                        await releaseInventory({
                            chemicalId: chem.chemicalId,
                            quantity: chem.packagesNeeded || chem.qty,
                            location: 'main',
                            orderId: order._id,
                            orderNumber: `ORD-${order._id.toString().slice(-8).toUpperCase()}`,
                            userId: req.user._id,
                            notes: 'Order archived: inventory released'
                        });
                    }
                }
            } catch (invError) {
                console.error('Inventory release warning:', invError.message);
            }
        }

        order.status = 'archived';
        order.updatedAt = new Date();
        order.lastEditedBy = req.user._id;
        await order.save();

        res.json({ message: 'Order archived successfully', order });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete order permanently (admin only) - only for superadmin or owner
app.delete('/api/admin/orders/:orderId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = { _id: req.params.orderId };

        // Distributors can only delete their own orders
        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        const order = await Order.findOne(query);
        if (!order) {
            return res.status(404).json({ error: 'Order not found or access denied' });
        }

        // Only allow deletion if order is in draft, archived, or cancelled state
        if (!['draft', 'archived', 'cancelled'].includes(order.status)) {
            return res.status(400).json({ error: 'Can only delete orders that are draft, archived, or cancelled. Please archive the order first.' });
        }

        await Order.deleteOne({ _id: req.params.orderId });

        res.json({ message: 'Order deleted successfully' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Send payment request email to customer for an existing order
app.post('/api/admin/orders/:orderId/send-payment-request', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Try both Order and ChemicalOrder
        let order = await Order.findById(req.params.orderId).populate('userId');
        let isChemOrder = false;
        if (!order) {
            order = await ChemicalOrder.findById(req.params.orderId).populate('userId');
            isChemOrder = true;
        }
        if (!order) return res.status(404).json({ error: 'Order not found' });

        const customer = order.userId;
        if (!customer?.email) return res.status(400).json({ error: 'Customer has no email address' });

        const transporter = createEmailTransporter();
        if (!transporter) return res.status(500).json({ error: 'Email not configured' });

        const total = order.totalCost || order.total || 0;
        const items = (order.chemicals || order.items || []).map(c =>
            `<li>${c.name || c.productName} - ${c.qty || c.quantity || 0} ${c.unit || 'units'} - $${((c.totalPrice || c.total || 0)).toFixed(2)}</li>`
        ).join('');

        const frontendUrl = process.env.FRONTEND_URL || 'https://acreprofit.com';

        await transporter.sendMail({
            from: process.env.EMAIL_FROM || '"Acre Profit" <noreply@acreprofit.com>',
            to: customer.email,
            subject: `Payment Request - $${total.toLocaleString(undefined, {minimumFractionDigits: 2})} - Acre Profit`,
            html: `
                <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                    <div style="background: #2d5a27; color: white; padding: 20px; text-align: center;">
                        <h1 style="margin: 0;">Acre Profit</h1>
                    </div>
                    <div style="padding: 20px; background: #f9f9f9;">
                        <h2>Payment Request</h2>
                        <p>Hi ${customer.name},</p>
                        <p>Your order is ready. Please review and submit payment at your convenience.</p>
                        <div style="background: white; padding: 15px; border-radius: 5px; margin: 15px 0;">
                            <h3 style="margin-top: 0;">Order Summary</h3>
                            ${items ? `<ul>${items}</ul>` : ''}
                            <p style="font-size: 1.2rem; font-weight: 700; color: #2d5a27;">Total: $${total.toLocaleString(undefined, {minimumFractionDigits: 2})}</p>
                        </div>
                        <div style="background: #d1fae5; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #10b981;">
                            <h4 style="margin-top: 0; color: #065f46;">Payment Options</h4>
                            <p style="color: #065f46;"><strong>Option 1 - ACH Bank Transfer (Preferred)</strong><br>Pay securely online. Fast processing, low fees.</p>
                            <a href="${frontendUrl}/my-orders.html" style="background: #10b981; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; font-weight: 600;">Pay Now - ACH Bank Transfer</a>
                            <p style="margin-top: 12px; color: #065f46;"><strong>Option 2 - Check</strong><br>Make check payable to: <strong>Acre Profit LLC</strong><br>Mail to: 34549 Highway 59, Haxtun, CO 80731<br>Include your name and order number on the memo line.</p>
                        </div>
                        <p>Questions? Reply to this email or contact your representative.</p>
                    </div>
                </div>
            `
        });

        res.json({ message: `Payment request sent to ${customer.email}` });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get order stats (admin only)
// Cash position: customer payments received minus checks written to distributors
app.get('/api/admin/stats/sales-total', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const paidStatuses = ['paid', 'confirmed', 'ready', 'delivered', 'completed', 'payment_secured'];

        const chemOrderSales = await ChemicalOrder.aggregate([
            { $match: { $or: [
                { paymentStatus: 'paid' },
                { status: { $in: paidStatuses } }
            ]}},
            { $group: { _id: null, total: { $sum: '$total' } } }
        ]);

        const orderSales = await Order.aggregate([
            { $match: { status: { $in: paidStatuses } } },
            { $group: { _id: null, total: { $sum: '$totalCost' } } }
        ]);

        // Checks written to distributors (category = payment, debit type = money out of AP)
        const checksWritten = await LedgerEntry.aggregate([
            { $match: { category: 'payment', type: 'debit' } },
            { $group: { _id: null, total: { $sum: '$amount' } } }
        ]);

        // Commissions earned by distributors (marginDollars × qty on paid orders)
        const commissionsByRep = await Order.aggregate([
            { $match: { status: { $in: paidStatuses }, repCommission: { $gt: 0 } } },
            { $group: { _id: '$representativeId', total: { $sum: '$repCommission' } } }
        ]);

        // Admin margin accumulated (adminMarginDollars × qty on paid orders)
        const adminRevenue = await Order.aggregate([
            { $match: { status: { $in: paidStatuses }, adminRevenue: { $gt: 0 } } },
            { $group: { _id: null, total: { $sum: '$adminRevenue' } } }
        ]);

        // Commissions paid out (category = commission, debit type)
        const commissionsPaid = await LedgerEntry.aggregate([
            { $match: { category: 'commission', type: 'debit' } },
            { $group: { _id: '$representativeId', total: { $sum: '$amount' } } }
        ]);
        const commissionsPaidMap = {};
        commissionsPaid.forEach(c => { if (c._id) commissionsPaidMap[c._id.toString()] = c.total; });

        // Admin margin withdrawn (category = admin_withdrawal or similar)
        const adminWithdrawn = await LedgerEntry.aggregate([
            { $match: { category: 'admin_withdrawal', type: 'debit' } },
            { $group: { _id: null, total: { $sum: '$amount' } } }
        ]);

        // Build commission owed per rep
        const repIds = commissionsByRep.map(c => c._id).filter(Boolean);
        const reps = await User.find({ _id: { $in: repIds } }).select('name email');
        const repMap = {};
        reps.forEach(r => { repMap[r._id.toString()] = r.name; });

        const commissionsOwed = commissionsByRep.map(c => {
            const earned = c.total || 0;
            const paid = commissionsPaidMap[c._id?.toString()] || 0;
            return {
                repId: c._id,
                repName: repMap[c._id?.toString()] || 'Unknown',
                earned,
                paid,
                owed: earned - paid
            };
        }).filter(c => c.owed > 0);

        const totalCommissionsOwed = commissionsOwed.reduce((sum, c) => sum + c.owed, 0);
        const adminMarginEarned = adminRevenue[0]?.total || 0;
        const adminMarginWithdrawn = adminWithdrawn[0]?.total || 0;
        const adminMarginBanked = adminMarginEarned - adminMarginWithdrawn;

        // Cash deposits / capital contributions (category = cash_deposit, credit = money into bank)
        const cashDeposits = await LedgerEntry.aggregate([
            { $match: { category: 'cash_deposit', type: 'credit' } },
            { $group: { _id: null, total: { $sum: '$amount' } } }
        ]);

        const customerPaymentsIn = (chemOrderSales[0]?.total || 0) + (orderSales[0]?.total || 0);
        const paidOutToDistributors = checksWritten[0]?.total || 0;
        const capitalIn = cashDeposits[0]?.total || 0;
        const cashInBank = capitalIn + customerPaymentsIn - paidOutToDistributors - (commissionsPaid.reduce((s, c) => s + c.total, 0)) - adminMarginWithdrawn;

        res.json({
            totalSales: customerPaymentsIn,
            cashInBank,
            paidOutToDistributors,
            customerPaymentsIn,
            commissionsOwed,
            totalCommissionsOwed,
            adminMarginEarned,
            adminMarginWithdrawn,
            adminMarginBanked
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Audit log (superadmin only - shows sensitive actions)
app.get('/api/admin/audit-log', authMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Superadmin only' });
        }
        const { action, targetUser, entityType, limit = 200 } = req.query;
        const query = {};
        if (action) query.action = action;
        if (targetUser) query.targetUser = targetUser;
        if (entityType) query.entityType = entityType;

        const entries = await AuditLog.find(query)
            .populate('performedBy', 'name email role')
            .populate('targetUser', 'name email')
            .sort({ createdAt: -1 })
            .limit(parseInt(limit));
        res.json(entries);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.get('/api/admin/stats', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = {};
        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        const totalOrders = await Order.countDocuments(query);
        const submittedOrders = await Order.countDocuments({ ...query, status: 'submitted' });
        const totalAcres = await Order.aggregate([
            { $match: query },
            { $group: { _id: null, total: { $sum: '$acres' } } }
        ]);

        let customerQuery = { role: 'customer' };
        if (isDistributor(req.user)) {
            customerQuery.representative = req.user._id;
        }
        const totalCustomers = await User.countDocuments(customerQuery);

        res.json({
            totalOrders,
            submittedOrders,
            totalAcres: totalAcres[0]?.total || 0,
            totalCustomers
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- REP APPLICATION ROUTES ----

// Submit rep application (public)
app.post('/api/rep-applications', async (req, res) => {
    try {
        const application = new RepApplication(req.body);
        await application.save();
        res.status(201).json({ message: 'Application submitted successfully', id: application._id });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all rep applications (superadmin only)
app.get('/api/rep-applications', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const applications = await RepApplication.find().sort({ createdAt: -1 });
        res.json(applications);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update rep application status (superadmin only)
app.put('/api/rep-applications/:id', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const { status } = req.body;
        const application = await RepApplication.findByIdAndUpdate(
            req.params.id,
            { status },
            { new: true }
        );

        if (!application) {
            return res.status(404).json({ error: 'Application not found' });
        }

        // If approved, create distributor user
        if (status === 'approved') {
            const existingUser = await User.findOne({ email: application.email.toLowerCase() });
            if (!existingUser) {
                const tempPassword = 'Farm2026!'; // They should change this
                await User.create({
                    name: `${application.firstName} ${application.lastName}`,
                    email: application.email,
                    password: tempPassword,
                    phone: application.phone,
                    role: 'distributor'
                });
            }
        }

        res.json(application);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- MERCH ORDER ROUTES REMOVED ----
// Printify/merch integration was never completed and has been removed.
// merch.html still exists as a static page but has no backend routes.

// ---- STRIPE PAYMENT ROUTES ----

// Get representative's payment info (for check payments)
app.get('/api/representatives/:repId/payment-info', async (req, res) => {
    try {
        const { repId } = req.params;

        // Map rep IDs to emails
        const repEmails = {
            kyle: 'office@togoag.com',
            ty: 'tymollohan77@gmail.com',
            chad: 'ckbamford@yahoo.com',
            seth: 'seth@acreprofit.com'
        };

        const rep = await User.findOne({ email: repEmails[repId] });
        if (!rep) {
            return res.status(404).json({ error: 'Representative not found' });
        }

        res.json({
            name: rep.name,
            phone: rep.phone,
            checkPayableTo: rep.checkPayableTo || rep.name,
            checkMailingAddress: rep.checkMailingAddress || null,
            stripeEnabled: !!rep.stripeAccountId && rep.stripeAccountStatus === 'active'
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create Stripe Connect onboarding link for representative
app.post('/api/stripe/connect/onboard', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (!stripe) {
            return res.status(400).json({ error: 'Stripe not configured' });
        }

        // Check if rep already has a Stripe account
        if (req.user.stripeAccountId) {
            // Create login link for existing account
            const loginLink = await stripe.accounts.createLoginLink(req.user.stripeAccountId);
            return res.json({ url: loginLink.url });
        }

        // Create new Connect account
        const account = await stripe.accounts.create({
            type: 'express',
            country: 'US',
            email: req.user.email,
            capabilities: {
                card_payments: { requested: true },
                transfers: { requested: true },
                us_bank_account_ach_payments: { requested: true }
            },
            business_type: 'individual',
            business_profile: {
                name: req.user.name,
                product_description: 'Agricultural products and chemicals'
            }
        });

        // Save account ID
        req.user.stripeAccountId = account.id;
        req.user.stripeAccountStatus = 'pending';
        await req.user.save();

        // Create onboarding link
        const accountLink = await stripe.accountLinks.create({
            account: account.id,
            refresh_url: `${process.env.FRONTEND_URL || 'https://acreprofit.com'}/admin.html?stripe=refresh`,
            return_url: `${process.env.FRONTEND_URL || 'https://acreprofit.com'}/admin.html?stripe=success`,
            type: 'account_onboarding'
        });

        res.json({ url: accountLink.url });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Check Stripe Connect account status
app.get('/api/stripe/connect/status', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (!stripe || !req.user.stripeAccountId) {
            return res.json({ connected: false, status: 'not_started' });
        }

        const account = await stripe.accounts.retrieve(req.user.stripeAccountId);

        // Update status in DB
        if (account.charges_enabled && account.payouts_enabled) {
            req.user.stripeAccountStatus = 'active';
        } else {
            req.user.stripeAccountStatus = 'pending';
        }
        await req.user.save();

        res.json({
            connected: true,
            status: req.user.stripeAccountStatus,
            chargesEnabled: account.charges_enabled,
            payoutsEnabled: account.payouts_enabled,
            detailsSubmitted: account.details_submitted
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- CUSTOMER PAYMENT METHOD ROUTES ----

// Update customer's save payment preference
app.put('/api/payments/save-preference', authMiddleware, async (req, res) => {
    try {
        const { savePaymentMethod } = req.body;

        req.user.savePaymentMethod = savePaymentMethod;
        await req.user.save();

        res.json({
            message: 'Preference updated',
            savePaymentMethod: req.user.savePaymentMethod
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get customer's payment methods
app.get('/api/payments/methods', authMiddleware, async (req, res) => {
    try {
        if (!stripe) {
            return res.json({ paymentMethods: [], savePreference: req.user.savePaymentMethod });
        }

        // If no Stripe customer, return empty
        if (!req.user.stripeCustomerId) {
            return res.json({
                paymentMethods: [],
                savePreference: req.user.savePaymentMethod
            });
        }

        // Get saved payment methods
        const paymentMethods = await stripe.paymentMethods.list({
            customer: req.user.stripeCustomerId,
            type: 'us_bank_account'
        });

        res.json({
            paymentMethods: paymentMethods.data.map(pm => ({
                id: pm.id,
                last4: pm.us_bank_account.last4,
                bankName: pm.us_bank_account.bank_name,
                accountType: pm.us_bank_account.account_type
            })),
            savePreference: req.user.savePaymentMethod
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create setup intent for adding payment method
app.post('/api/payments/setup-intent', authMiddleware, async (req, res) => {
    try {
        if (!stripe) {
            return res.status(400).json({ error: 'Stripe not configured' });
        }

        // Create or get Stripe Customer
        let customerId = req.user.stripeCustomerId;

        if (!customerId) {
            const customer = await stripe.customers.create({
                email: req.user.email,
                name: req.user.name,
                metadata: {
                    userId: req.user._id.toString()
                }
            });
            customerId = customer.id;
            req.user.stripeCustomerId = customerId;
            await req.user.save();
        }

        // Create setup intent for bank account
        const setupIntent = await stripe.setupIntents.create({
            customer: customerId,
            payment_method_types: ['us_bank_account'],
            payment_method_options: {
                us_bank_account: {
                    financial_connections: {
                        permissions: ['payment_method', 'balances']
                    }
                }
            }
        });

        res.json({
            clientSecret: setupIntent.client_secret,
            customerId: customerId
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete a saved payment method
app.delete('/api/payments/methods/:paymentMethodId', authMiddleware, async (req, res) => {
    try {
        if (!stripe) {
            return res.status(400).json({ error: 'Stripe not configured' });
        }

        const { paymentMethodId } = req.params;

        // Verify the payment method belongs to this customer
        const paymentMethod = await stripe.paymentMethods.retrieve(paymentMethodId);

        if (paymentMethod.customer !== req.user.stripeCustomerId) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        // Detach the payment method
        await stripe.paymentMethods.detach(paymentMethodId);

        res.json({ message: 'Payment method removed' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete all saved payment methods
app.delete('/api/payments/methods', authMiddleware, async (req, res) => {
    try {
        if (!stripe || !req.user.stripeCustomerId) {
            return res.json({ message: 'No payment methods to remove' });
        }

        // Get all payment methods
        const paymentMethods = await stripe.paymentMethods.list({
            customer: req.user.stripeCustomerId,
            type: 'us_bank_account'
        });

        // Detach all
        for (const pm of paymentMethods.data) {
            await stripe.paymentMethods.detach(pm.id);
        }

        res.json({ message: 'All payment methods removed', count: paymentMethods.data.length });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create payment intent for an order (ACH or Card)
app.post('/api/payments/create-intent', authMiddleware, async (req, res) => {
    try {
        const { orderId, paymentMethod, amount } = req.body;

        if (!stripe) {
            return res.status(400).json({ error: 'Stripe not configured' });
        }

        // Try to find the order in both Order and ChemicalOrder collections
        let order = await Order.findById(orderId).populate('representativeId');
        let isChemicalOrder = false;

        if (!order) {
            order = await ChemicalOrder.findById(orderId).populate('representativeId');
            isChemicalOrder = true;
        }

        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        if (order.userId.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        // Get the order amount
        const orderAmount = amount || order.total || order.totalCost || 0;
        if (orderAmount <= 0) {
            return res.status(400).json({ error: 'Invalid order amount' });
        }

        // Get representative for Stripe Connect (optional - if not set, payment goes to platform)
        let rep = order.representativeId;
        let useStripeConnect = rep && rep.stripeAccountId && rep.stripeAccountStatus === 'active';

        // Calculate platform fee (optional - 0% for now, can add later)
        const platformFeePercent = 0;
        const applicationFee = Math.round(orderAmount * 100 * platformFeePercent);

        // Create payment intent params
        const paymentIntentParams = {
            amount: Math.round(orderAmount * 100), // Convert to cents
            currency: 'usd',
            payment_method_types: paymentMethod === 'ach' ? ['us_bank_account'] : ['card'],
            metadata: {
                orderId: order._id.toString(),
                orderType: isChemicalOrder ? 'chemical' : 'regular',
                customerId: req.user._id.toString(),
                customerName: req.user.name
            }
        };

        // If rep has Stripe Connect, use transfer (sends funds to their account)
        if (useStripeConnect) {
            paymentIntentParams.transfer_data = {
                destination: rep.stripeAccountId
            };
            paymentIntentParams.metadata.repName = rep.name;

            if (applicationFee > 0) {
                paymentIntentParams.application_fee_amount = applicationFee;
            }
        }

        // For ACH, add specific options
        if (paymentMethod === 'ach') {
            paymentIntentParams.payment_method_options = {
                us_bank_account: {
                    financial_connections: {
                        permissions: ['payment_method', 'balances']
                    }
                }
            };
        }

        const paymentIntent = await stripe.paymentIntents.create(paymentIntentParams);

        // Update order with payment intent
        order.stripePaymentIntentId = paymentIntent.id;
        order.paymentMethod = paymentMethod === 'ach' ? 'stripe_ach' : 'stripe_card';
        order.paymentStatus = 'processing';
        await order.save();

        res.json({
            clientSecret: paymentIntent.client_secret,
            paymentIntentId: paymentIntent.id
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Mark order as paid by check
app.post('/api/payments/check-received', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { orderId, checkNumber } = req.body;

        // Try ChemicalOrder first (active system), then fall back to Order
        let order = await ChemicalOrder.findById(orderId);
        if (!order) {
            order = await Order.findById(orderId);
        }
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        // Verify the rep owns this order (unless superadmin)
        if (req.user.role !== 'superadmin' &&
            order.representativeId && order.representativeId.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        order.paymentMethod = 'check';
        order.paymentStatus = 'paid';
        if (order.checkNumber !== undefined) order.checkNumber = checkNumber;
        if (order.checkReceivedDate !== undefined) order.checkReceivedDate = new Date();
        order.paidAt = new Date();
        order.status = 'confirmed';
        order.updatedAt = new Date();
        await order.save();

        // Auto-generate invoice now that payment has cleared (idempotent)
        await autoGenerateInvoiceForPaidOrder(order);
        await promoteRupRecordsToCompleted(order);

        res.json({ message: 'Payment recorded', order });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Stripe webhook for payment confirmations (including ACH)
app.post('/api/webhooks/stripe', express.raw({ type: 'application/json' }), async (req, res) => {
    const sig = req.headers['stripe-signature'];
    const webhookSecret = process.env.STRIPE_WEBHOOK_SECRET;

    let event;
    try {
        event = stripe.webhooks.constructEvent(req.body, sig, webhookSecret);
    } catch (err) {
        return res.status(400).send(`Webhook Error: ${err.message}`);
    }

    // Helper to find order by ID
    async function findOrder(orderId, orderType) {
        if (orderType === 'chemical') {
            return await ChemicalOrder.findById(orderId);
        }
        let order = await Order.findById(orderId);
        if (!order) {
            order = await ChemicalOrder.findById(orderId);
        }
        return order;
    }

    // ACH payments go through processing state before succeeding
    if (event.type === 'payment_intent.processing') {
        const paymentIntent = event.data.object;
        const orderId = paymentIntent.metadata.orderId;
        const orderType = paymentIntent.metadata.orderType;

        if (orderId) {
            const order = await findOrder(orderId, orderType);
            if (order) {
                order.paymentStatus = 'processing';
                order.paymentMethod = paymentIntent.payment_method_types?.includes('us_bank_account') ? 'stripe_ach' : order.paymentMethod;
                order.updatedAt = new Date();
                await order.save();
                console.log(`ACH payment processing for order ${orderId}`);
            }
        }
    }

    if (event.type === 'payment_intent.succeeded') {
        const paymentIntent = event.data.object;
        const orderId = paymentIntent.metadata.orderId;
        const orderType = paymentIntent.metadata.orderType;
        const isACH = paymentIntent.payment_method_types?.includes('us_bank_account');

        if (orderId) {
            const order = await findOrder(orderId, orderType);

            if (order) {
                order.paymentStatus = 'paid';
                order.paidAt = new Date();
                // Update status to payment_secured (matches admin workflow)
                order.status = 'payment_secured';
                order.paymentMethod = isACH ? 'stripe_ach' : (order.paymentMethod || 'stripe');
                order.updatedAt = new Date();
                await order.save();
                console.log(`Payment ${isACH ? '(ACH) ' : ''}confirmed for order ${orderId} - status set to payment_secured`);

                // Auto-generate invoice + promote pending RUP records (both idempotent)
                await autoGenerateInvoiceForPaidOrder(order);
                await promoteRupRecordsToCompleted(order);

                // Send payment confirmation email to customer
                try {
                    const customer = await User.findById(order.userId);
                    const transporter = createTransporter();
                    if (transporter && customer?.email) {
                        const amount = (paymentIntent.amount / 100).toFixed(2);
                        const invoiceUrl = `${process.env.FRONTEND_URL || 'https://acreprofit.com'}/my-orders.html`;
                        await transporter.sendMail({
                            from: process.env.EMAIL_FROM || process.env.SMTP_USER,
                            to: customer.email,
                            subject: `Payment Confirmed - Acre Profit`,
                            html: `<div style="font-family: Arial, sans-serif; max-width: 600px;">
                                <h2 style="color: #2d5a27;">Payment Confirmed</h2>
                                <p>Hi ${customer.name || 'Farmer'},</p>
                                <p>Your payment of <strong>$${amount}</strong> has been received via ${isACH ? 'ACH bank transfer' : 'card'}.</p>
                                <p>Your order is now being processed. We'll notify you when your products are ready for pickup.</p>
                                <p style="margin-top:18px;">View your invoice and order details: <a href="${invoiceUrl}" style="color:#2d5a27; font-weight:600;">${invoiceUrl}</a></p>
                                <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
                                <p style="color: #888; font-size: 0.9em;">Thank you for choosing Acre Profit.<br>Questions? Contact your local representative.</p>
                            </div>`
                        });
                        console.log(`Payment confirmation email sent to ${customer.email}`);
                    }
                } catch (emailErr) {
                    console.error('Failed to send payment confirmation email:', emailErr.message);
                }
            }
        }
    }

    if (event.type === 'payment_intent.payment_failed') {
        const paymentIntent = event.data.object;
        const orderId = paymentIntent.metadata.orderId;
        const orderType = paymentIntent.metadata.orderType;

        if (orderId) {
            const order = await findOrder(orderId, orderType);

            if (order) {
                order.paymentStatus = 'failed';
                order.status = 'payment_pending'; // Reset to payment pending
                order.updatedAt = new Date();
                await order.save();
                console.log(`Payment failed for order ${orderId}`);
            }
        }
    }

    // Handle charge events for ACH (backup confirmation)
    if (event.type === 'charge.succeeded') {
        const charge = event.data.object;
        const orderId = charge.metadata?.orderId;
        const orderType = charge.metadata?.orderType;
        const isACH = charge.payment_method_details?.type === 'us_bank_account';

        if (orderId && isACH) {
            const order = await findOrder(orderId, orderType);
            if (order && order.paymentStatus !== 'paid') {
                order.paymentStatus = 'paid';
                order.paidAt = new Date();
                order.status = 'payment_secured';
                order.paymentMethod = 'stripe_ach';
                order.updatedAt = new Date();
                await order.save();
                console.log(`ACH charge confirmed for order ${orderId}`);

                // Auto-generate invoice + promote pending RUP records (both idempotent -
                // payment_intent.succeeded usually fires first)
                await autoGenerateInvoiceForPaidOrder(order);
                await promoteRupRecordsToCompleted(order);
            }
        }
    }

    res.json({ received: true });
});

// Update representative's check payment info
app.put('/api/representatives/check-info', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { checkPayableTo, checkMailingAddress } = req.body;

        req.user.checkPayableTo = checkPayableTo;
        req.user.checkMailingAddress = checkMailingAddress;
        await req.user.save();

        res.json({ message: 'Check payment info updated' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- CHEMICAL PRICING ROUTES ----

// Seed initial chemical pricing data
app.post('/api/chemicals/seed', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { secretKey } = req.body;

        // Protection to prevent accidental re-seeding
        if (secretKey !== 'acreprofit2026seed') {
            return res.status(403).json({ error: 'Invalid secret key' });
        }

        const priceVersion = '2026-03-06';

        // Actual CPD Products with real prices (costPrice = CPD price, sellPrice = TBD by admin)
        // sellPrice set to 0 initially - update in Pricing tab
        const cpdProducts = [
            // Dicamba products
            { productName: 'Dicamba 49.8% SL', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 30.25, sellPrice: 0, category: 'herbicide' },
            { productName: 'Dicamba 49.8% SL', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 28.25, sellPrice: 0, category: 'herbicide' },
            { productName: 'Dicamba HD', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 30.57, sellPrice: 0, category: 'herbicide' },

            // LV 6 (2,4-D)
            { productName: 'LV 6', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 29.90, sellPrice: 0, category: 'herbicide' },
            { productName: 'LV 6', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 27.90, sellPrice: 0, category: 'herbicide' },

            // Glyphosate - AgSaver is CPD equiv for RT3/Glystar Supreme
            { productName: 'AgSaver', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 13.32, sellPrice: 0, category: 'herbicide', equivalentProduct: 'RT3, Glystar Supreme', notes: 'Glyphosate' },

            // Glyphosate 5.4 lb - from JABCO/CPD
            { productName: 'Glyphosate 5.4', packSize: 'Tote', unit: 'gal', unitsPerPack: 265, costPrice: 13.25, adminMarginDollars: 0.07, adminPrice: 13.32, marginDollars: 0.77, sellPrice: 14.09, category: 'herbicide', notes: '5.4 lb/gal glyphosate' },

            // XSATE Glyphosate 53.8% - Xingfa USA via Jabco (EPA 89343-5) - $13.25/gal cost, $16/gal retail
            { productName: 'XSATE Glyphosate 53.8%', packSize: '265 gal', unit: 'gal', unitsPerPack: 265, costPrice: 13.25, adminMarginDollars: 1.25, adminPrice: 14.50, marginDollars: 1.50, sellPrice: 16.00, category: 'herbicide', sourceSupplier: 'Jabco', epaRegistrationNumber: '89343-5', signalWord: 'CAUTION', notes: '5.4 lb/gal glyphosate - Xingfa USA', activeIngredients: [{ name: 'XSATE Glyphosate 53.8%', percentage: 53.8, poundsPerGallon: 5.4 }] },
            { productName: 'XSATE Glyphosate 53.8%', packSize: '30 gal', unit: 'gal', unitsPerPack: 30, costPrice: 13.25, adminMarginDollars: 1.25, adminPrice: 14.50, marginDollars: 1.50, sellPrice: 16.00, category: 'herbicide', sourceSupplier: 'Jabco', epaRegistrationNumber: '89343-5', signalWord: 'CAUTION', notes: '5.4 lb/gal glyphosate - Xingfa USA', activeIngredients: [{ name: 'XSATE Glyphosate 53.8%', percentage: 53.8, poundsPerGallon: 5.4 }] },
            { productName: 'XSATE Glyphosate 53.8%', packSize: '2.5 gal', unit: 'gal', unitsPerPack: 2.5, costPrice: 13.25, adminMarginDollars: 1.25, adminPrice: 14.50, marginDollars: 1.50, sellPrice: 16.00, category: 'herbicide', sourceSupplier: 'Jabco', epaRegistrationNumber: '89343-5', signalWord: 'CAUTION', notes: '5.4 lb/gal glyphosate - Xingfa USA', activeIngredients: [{ name: 'XSATE Glyphosate 53.8%', percentage: 53.8, poundsPerGallon: 5.4 }] },

            // Atrazine
            { productName: 'Aatrex', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 13.35, sellPrice: 0, category: 'herbicide' },

            // Agri-Star is CPD equiv for Level Best Pro
            { productName: 'Agri-Star', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 39.25, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Level Best Pro' },
            { productName: 'Agri-Star', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 38.13, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Level Best Pro' },

            // Agri-Star Tapran is CPD equiv for Tapran
            { productName: 'Agri-Star Tapran', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 19.00, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Tapran' },
            { productName: 'Agri-Star Tapran', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 18.00, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Tapran' },

            // Aggrestrol
            { productName: 'Aggrestrol', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 21.00, sellPrice: 0, category: 'herbicide' },
            { productName: 'Aggrestrol', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 20.00, sellPrice: 0, category: 'herbicide' },

            // Sulfentrazone
            { productName: 'Sulfentrazone', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 70.50, sellPrice: 0, category: 'herbicide' },

            // Valor SX
            { productName: 'Valor SX', packSize: '4x5', unit: 'lb', unitsPerPack: 20, costPrice: 14.25, sellPrice: 0, category: 'herbicide' },

            // CPD-only products
            { productName: 'Glufosinate', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 16.00, sellPrice: 0, category: 'herbicide' },
            { productName: 'Paraquat', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 15.25, sellPrice: 0, category: 'herbicide', isRestrictedUse: true, requiredCertifications: ['private_applicator', 'paraquat_training'], notes: 'Restricted Use Pesticide - requires certification + Paraquat training' },
            { productName: 'Mesotrione', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 48.25, sellPrice: 0, category: 'herbicide' },
            { productName: 'Clethodim', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 33.50, sellPrice: 0, category: 'herbicide' },
            { productName: 'Clethodim', packSize: '135', unit: 'gal', unitsPerPack: 135, costPrice: 33.00, sellPrice: 0, category: 'herbicide' },

            // AMS (Ammonium Sulfate)
            { productName: 'AMS', packSize: '24 lb', unit: 'lb', unitsPerPack: 24, costPrice: 1.45, sellPrice: 0, category: 'adjuvant', notes: 'Ammonium Sulfate - water conditioner/adjuvant' },
            { productName: 'AMS x5', packSize: '24', unit: 'lb', unitsPerPack: 120, costPrice: 1.45, sellPrice: 0, category: 'adjuvant', notes: 'Ammonium Sulfate - 5 bag bundle' },

            // Hydrovant
            { productName: 'Hydrovant fA', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 95.00, sellPrice: 0, category: 'adjuvant', notes: 'Drift reduction/deposition aid adjuvant' },
            { productName: 'Hydrovant fA', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 265, costPrice: 95.00, sellPrice: 0, category: 'adjuvant', notes: 'Drift reduction/deposition aid adjuvant' },

            // Rancor 4F (Metribuzin) - JABCO SO# 2131
            { productName: 'Rancor 4F', packSize: '2x2.5 gal', unit: 'gal', unitsPerPack: 5, costPrice: 45.50, sellPrice: 0, category: 'herbicide', notes: 'Metribuzin 4F herbicide' }
        ];

        const results = { created: [], existing: [] };

        // Insert CPD products
        for (const chem of cpdProducts) {
            const existing = await Chemical.findOne({
                productName: chem.productName,
                sourceSupplier: 'JABCO',
                packSize: chem.packSize
            });

            if (!existing) {
                // Use explicit prices if set, otherwise default to cost (no auto-margin - set on site)
                const cost = chem.costPrice || 0;
                const admin = chem.adminPrice || cost;
                const sell = chem.sellPrice || admin;

                const newChem = await Chemical.create({
                    ...chem,
                    costPrice: cost,
                    adminPrice: admin,
                    sellPrice: sell,
                    margin: sell > 0 ? Math.round(((sell - cost) / sell) * 100) : 0,
                    sourceSupplier: 'JABCO',
                    priceVersion,
                    isActive: true,
                    availableForOrder: true,
                    createdBy: req.user._id
                });
                await ChemicalPriceHistory.create({
                    chemicalId: newChem._id,
                    productName: chem.productName,
                    sourceSupplier: 'JABCO',
                    packSize: chem.packSize,
                    unit: chem.unit,
                    costPrice: chem.costPrice,
                    sellPrice: chem.sellPrice,
                    priceVersion,
                    changedBy: req.user._id
                });
                results.created.push(`${chem.productName} (${chem.packSize})`);
            } else {
                // Update existing product if prices are 0
                if (existing.costPrice === 0 || existing.sellPrice === 0) {
                    const cost = chem.costPrice || existing.costPrice || 0;
                    const admin = chem.adminPrice || cost;
                    const sell = chem.sellPrice || admin;

                    existing.costPrice = cost;
                    existing.adminPrice = admin;
                    existing.sellPrice = sell;
                    existing.margin = sell > 0 ? Math.round(((sell - cost) / sell) * 100) : 0;
                    existing.priceVersion = priceVersion;
                    await existing.save();
                    results.existing.push(`${chem.productName} (${chem.packSize}) - UPDATED`);
                } else {
                    results.existing.push(`${chem.productName} (${chem.packSize})`);
                }
            }
        }

        res.json({
            message: 'CPD products loaded/updated successfully',
            summary: {
                created: results.created.length,
                alreadyExisted: results.existing.length,
                totalProducts: cpdProducts.length
            },
            products: results.created
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all chemicals (with optional filters) - PUBLIC for customer ordering
app.get('/api/chemicals', async (req, res) => {
    try {
        const { sourceSupplier, productName, category, crop, activeOnly, availableOnly } = req.query;
        let query = {};

        if (sourceSupplier) query.sourceSupplier = new RegExp(sourceSupplier, 'i');
        if (productName) query.productName = new RegExp(productName, 'i');
        if (category) query.category = category;
        if (crop) query.crops = crop;
        if (activeOnly === 'true') query.isActive = true;
        if (availableOnly === 'true') query.availableForOrder = true;

        const chemicals = await Chemical.find(query)
            .sort({ productName: 1, packSize: 1 });

        // Get inventory levels for all chemicals
        const inventoryRecords = await Inventory.find({ location: 'main' }).lean();
        const inventoryMap = {};
        inventoryRecords.forEach(inv => {
            const chemId = inv.chemicalId?.toString();
            if (chemId) {
                inventoryMap[chemId] = {
                    quantityAvailable: inv.quantityAvailable || 0,
                    quantityOnHand: inv.quantityOnHand || 0
                };
            }
        });

        // For public view - ONLY show products with valid retail pricing
        // NEVER expose wholesale/cost pricing to public
        const publicChemicals = chemicals
            .filter(c => c.sellPrice > 0) // Only show products with retail price set
            .map(c => {
                const inv = inventoryMap[c._id.toString()] || {};
                return {
                    _id: c._id,
                    productName: c.productName,
                    category: c.category,
                    crops: c.crops,
                    packSize: c.packSize,
                    unit: c.unit,
                    unitsPerPack: c.unitsPerPack,
                    price: c.sellPrice, // Only the retail price
                    sellPrice: c.sellPrice,
                    defaultRate: c.defaultRate,
                    rateUnit: c.rateUnit,
                    notes: c.notes,
                    isActive: c.isActive,
                    availableForOrder: c.availableForOrder,
                    inStock: (inv.quantityAvailable || 0) > 0,
                    quantityAvailable: inv.quantityAvailable || 0
                };
            });

        res.json(publicChemicals);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get chemicals with customer's distributor pricing (authenticated customers)
app.get('/api/chemicals/for-customer', authMiddleware, async (req, res) => {
    try {
        const { category, crop } = req.query;
        let query = { isActive: true, availableForOrder: true };

        if (category) query.category = category;
        if (crop) query.crops = crop;

        const chemicals = await Chemical.find(query).sort({ productName: 1, packSize: 1 });

        // Get the customer's representative/distributor
        const repId = req.user.representative || req.user.representativeId;

        // Get distributor pricing for this rep
        let distributorPricing = [];
        if (repId) {
            distributorPricing = await DistributorPricing.find({
                distributorId: repId,
                isAvailable: true
            });
        }

        // Create a map for quick lookup
        const pricingMap = {};
        distributorPricing.forEach(dp => {
            pricingMap[dp.chemicalId.toString()] = dp.retailPrice;
        });

        // Return ONLY products that have a retail price set by distributor
        // NEVER expose wholesale/cost pricing to customers
        const customerChemicals = chemicals
            .map(c => {
                const distributorPrice = pricingMap[c._id.toString()];
                // Must have a valid retail price from distributor
                const retailPrice = distributorPrice || 0;

                if (retailPrice <= 0) {
                    return null; // Don't show products without retail pricing
                }

                return {
                    _id: c._id,
                    productName: c.productName,
                    sourceSupplier: c.sourceSupplier,
                    category: c.category,
                    crops: c.crops,
                    packSize: c.packSize,
                    unit: c.unit,
                    unitsPerPack: c.unitsPerPack,
                    sellPrice: retailPrice,
                    price: retailPrice,
                    defaultRate: c.defaultRate,
                    rateUnit: c.rateUnit,
                    equivalentProduct: c.equivalentProduct,
                    notes: c.notes
                };
            })
            .filter(c => c !== null); // Remove products without retail pricing

        res.json(customerChemicals);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Customer: Get available products with inventory status and sell price ONLY
// Only shows products we have on hand or on order - no cost/admin pricing exposed
app.get('/api/chemicals/available', async (req, res) => {
    try {
        const { category } = req.query;

        // Get all inventory with stock or on order (across ALL locations, not just 'main')
        const inventory = await Inventory.find({
            $or: [
                { quantityOnHand: { $gt: 0 } },
                { quantityAvailable: { $gt: 0 } }
            ]
        }).populate('distributorId', 'name').lean();

        // Get chemicals on pending POs (on order but not yet received)
        const pendingPOs = await PurchaseOrder.find({
            status: { $in: ['draft', 'submitted', 'confirmed', 'partial_received'] }
        }).select('items').lean();

        const onOrderChemIds = new Set();
        const onOrderQty = {};
        for (const po of pendingPOs) {
            for (const item of po.items || []) {
                if (item.chemicalId) {
                    const key = item.chemicalId.toString();
                    onOrderChemIds.add(key);
                    onOrderQty[key] = (onOrderQty[key] || 0) + (item.quantityOrdered || 0);
                }
            }
        }

        // Combine: chemicals that have inventory OR are on order
        const inStockChemIds = new Set(inventory.map(i => i.chemicalId?.toString()).filter(Boolean));
        const allAvailableIds = new Set([...inStockChemIds, ...onOrderChemIds]);

        if (allAvailableIds.size === 0) {
            return res.json([]);
        }

        // Fetch chemical details - only products with sell price set above cost (margins configured)
        const query = {
            _id: { $in: Array.from(allAvailableIds) },
            sellPrice: { $gt: 0 },
            $expr: { $gt: ['$sellPrice', '$costPrice'] }
        };
        if (category) query.category = category;

        const chemicals = await Chemical.find(query)
            .sort({ category: 1, productName: 1 })
            .lean();

        // Build inventory map: aggregate across all locations per chemical
        const invMap = {}; // chemicalId -> { totalOnHand, locations: [{name, location, qty}] }
        inventory.forEach(inv => {
            const key = inv.chemicalId?.toString();
            if (!key) return;
            if (!invMap[key]) {
                invMap[key] = { totalOnHand: 0, locations: [] };
            }
            invMap[key].totalOnHand += (inv.quantityOnHand || 0);
            // Build location info: distributor first name + location name
            const distributorName = inv.distributorId?.name?.split(' ')[0] || '';
            const locName = inv.location && inv.location !== 'main'
                ? inv.location.charAt(0).toUpperCase() + inv.location.slice(1)
                : '';
            const label = distributorName && locName
                ? `${distributorName} - ${locName}`
                : distributorName || locName || 'Main';
            if (inv.quantityOnHand > 0) {
                invMap[key].locations.push({ name: label, qty: inv.quantityOnHand });
            }
        });

        // Return ONLY customer-safe data - NO cost, NO admin price, NO margins
        const available = chemicals.map(c => {
            const invData = invMap[c._id.toString()];
            const onHand = invData?.totalOnHand || 0;
            const onOrder = onOrderQty[c._id.toString()] || 0;

            let availability;
            if (onHand > 0) {
                availability = 'in_stock';
            } else if (onOrder > 0) {
                availability = 'on_order';
            } else {
                availability = 'limited';
            }

            return {
                _id: c._id,
                productName: c.productName,
                sourceSupplier: c.sourceSupplier,
                activeIngredients: (c.activeIngredients || []).map(ai => ({
                    name: ai.name,
                    percentage: ai.percentage
                })),
                category: c.category,
                packSize: c.packSize,
                unit: c.unit,
                unitsPerPack: c.unitsPerPack,
                sellPrice: c.sellPrice,
                priceIsSpeculated: c.priceIsSpeculated === true,
                defaultRate: c.defaultRate,
                rateUnit: c.rateUnit,
                signalWord: c.signalWord,
                epaRegistrationNumber: c.epaRegistrationNumber,
                labelUrl: c.labelUrl,
                sdsUrl: c.sdsUrl,
                availability,
                quantityOnHand: onHand,
                quantityOnOrder: onOrder,
                isRestrictedUse: c.isRestrictedUse,
                locations: invData?.locations || []
            };
        });

        res.json(available);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});
app.get('/api/chemicals/admin', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { sourceSupplier, productName, category } = req.query;
        let query = {};

        if (sourceSupplier) query.sourceSupplier = new RegExp(sourceSupplier, 'i');
        if (productName) query.productName = new RegExp(productName, 'i');
        if (category) query.category = category;

        const chemicals = await Chemical.find(query)
            .sort({ productName: 1, packSize: 1 });

        // Superadmin and distributors see everything including cost
        res.json(chemicals);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Search chemicals with inventory status (for calculator)
// Auth: all logged-in roles
app.get('/api/chemicals/search', authMiddleware, async (req, res) => {
    try {
        const { q, category, crop, limit = 20 } = req.query;

        // Build query
        let query = { isActive: true, availableForOrder: true };
        if (q) {
            query.productName = new RegExp(q, 'i');
        }
        if (category) {
            query.category = category;
        }
        if (crop) {
            query.crops = crop;
        }

        // Get chemicals
        const chemicals = await Chemical.find(query)
            .select('productName category crops packSize unit unitsPerPack sellPrice defaultRate rateUnit sourceSupplier')
            .limit(parseInt(limit))
            .lean();

        // Get inventory for all chemicals
        const chemicalIds = chemicals.map(c => c._id);
        const inventories = await Inventory.find({
            chemicalId: { $in: chemicalIds },
            location: 'main'
        }).lean();

        // Create inventory map
        const inventoryMap = {};
        inventories.forEach(inv => {
            inventoryMap[inv.chemicalId.toString()] = inv.quantityAvailable || 0;
        });

        // Map chemicals with inventory status
        const results = chemicals.map(c => ({
            _id: c._id,
            productName: c.productName,
            category: c.category,
            crops: c.crops,
            packSize: c.packSize,
            unit: c.unit,
            unitsPerPack: c.unitsPerPack,
            sellPrice: c.sellPrice,
            defaultRate: c.defaultRate,
            rateUnit: c.rateUnit,
            supplier: c.sourceSupplier,
            onHandQuantity: inventoryMap[c._id.toString()] || 0,
            inStock: (inventoryMap[c._id.toString()] || 0) > 0
        }));

        // Sort by in-stock first, then by name
        results.sort((a, b) => {
            if (a.inStock && !b.inStock) return -1;
            if (!a.inStock && b.inStock) return 1;
            return a.productName.localeCompare(b.productName);
        });

        res.json(results);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Distributor: Update distributor margin on a product (any distributor can do this, changes for all)
app.put('/api/chemicals/:id/distributor-margin', authMiddleware, async (req, res) => {
    try {
        // Any distributor, admin, or superadmin can set distributor margin
        if (!['distributor', 'admin', 'superadmin'].includes(req.user.role)) {
            return res.status(403).json({ error: 'Only distributors and admins can set distributor margin' });
        }

        const { marginDollars } = req.body;
        if (marginDollars === undefined || marginDollars < 0) {
            return res.status(400).json({ error: 'marginDollars is required and must be >= 0' });
        }

        const chemical = await Chemical.findById(req.params.id);
        if (!chemical) return res.status(404).json({ error: 'Product not found' });

        chemical.marginDollars = marginDollars;
        chemical.sellPrice = Math.round((chemical.adminPrice + marginDollars) * 100) / 100;
        chemical.margin = chemical.sellPrice > 0 ? Math.round(((chemical.sellPrice - chemical.costPrice) / chemical.sellPrice) * 10000) / 100 : 0;
        chemical.updatedAt = new Date();
        await chemical.save();

        // Sync inventory record
        const inv = await Inventory.findOne({ chemicalId: chemical._id, location: 'main' });
        if (inv) { inv.productName = chemical.productName; inv.updatedAt = new Date(); await inv.save(); }

        res.json({
            _id: chemical._id,
            productName: chemical.productName,
            adminPrice: chemical.adminPrice,
            marginDollars: chemical.marginDollars,
            sellPrice: chemical.sellPrice,
            margin: chemical.margin
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Superadmin: Update admin margin on a product (recalculates admin price and sell price)
app.put('/api/chemicals/:id/admin-margin', authMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can set admin margin' });
        }

        const { adminMarginDollars, costPrice } = req.body;
        const chemical = await Chemical.findById(req.params.id);
        if (!chemical) return res.status(404).json({ error: 'Product not found' });

        if (costPrice !== undefined) chemical.costPrice = costPrice;
        if (adminMarginDollars !== undefined) chemical.adminMarginDollars = adminMarginDollars;

        // Recalculate prices
        chemical.adminPrice = Math.round((chemical.costPrice + chemical.adminMarginDollars) * 100) / 100;
        chemical.sellPrice = Math.round((chemical.adminPrice + (chemical.marginDollars || 0)) * 100) / 100;
        chemical.margin = chemical.sellPrice > 0 ? Math.round(((chemical.sellPrice - chemical.costPrice) / chemical.sellPrice) * 10000) / 100 : 0;
        chemical.updatedAt = new Date();

        await chemical.save();

        res.json(chemical);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get unique suppliers list
app.get('/api/chemicals/suppliers', async (req, res) => {
    try {
        const suppliers = await Chemical.distinct('sourceSupplier');
        res.json(suppliers.sort());
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get unique product names list
app.get('/api/chemicals/products', async (req, res) => {
    try {
        const products = await Chemical.distinct('productName');
        res.json(products.sort());
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get chemicals by category
app.get('/api/chemicals/category/:category', async (req, res) => {
    try {
        const chemicals = await Chemical.find({
            category: req.params.category,
            isActive: true,
            availableForOrder: true
        }).sort({ productName: 1 });

        res.json(chemicals.map(c => ({
            _id: c._id,
            productName: c.productName,
            packSize: c.packSize,
            unit: c.unit,
            price: c.sellPrice,
            defaultRate: c.defaultRate,
            rateUnit: c.rateUnit
        })));
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get margin report (admin only)
app.get('/api/chemicals/report/margins', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const chemicals = await Chemical.find({ isActive: true }).sort({ margin: -1 });

        const report = chemicals.map(c => ({
            productName: c.productName,
            packSize: c.packSize,
            costPrice: c.costPrice,
            sellPrice: c.sellPrice,
            margin: c.margin,
            profitPerUnit: Math.round((c.sellPrice - c.costPrice) * 100) / 100
        }));

        res.json(report);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get price history for a chemical
app.get('/api/chemicals/:id/history', async (req, res) => {
    try {
        const history = await ChemicalPriceHistory.find({ chemicalId: req.params.id })
            .sort({ priceDate: -1 })
            .populate('changedBy', 'name email');

        res.json(history);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Add new chemical (admin only)
app.post('/api/chemicals', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { productName, sourceSupplier, packSize, unit, unitsPerPack, costPrice, sellPrice,
                category, crops, defaultRate, rateUnit, priceVersion, equivalentProduct, notes } = req.body;

        const chemical = new Chemical({
            productName,
            sourceSupplier,
            packSize,
            unit,
            unitsPerPack,
            costPrice,
            sellPrice,
            category: category || 'herbicide',
            crops: crops || [],
            defaultRate,
            rateUnit,
            priceVersion: priceVersion || new Date().toISOString().slice(0, 10),
            equivalentProduct,
            notes,
            createdBy: req.user._id
        });

        await chemical.save();

        // Add to price history
        await ChemicalPriceHistory.create({
            chemicalId: chemical._id,
            productName,
            sourceSupplier,
            packSize,
            unit,
            costPrice,
            sellPrice,
            priceVersion: chemical.priceVersion,
            changedBy: req.user._id
        });

        // Auto-create inventory record for new product
        const existingInv = await Inventory.findOne({ chemicalId: chemical._id, location: 'main' });
        if (!existingInv) {
            await new Inventory({
                chemicalId: chemical._id,
                productName: chemical.productName,
                packSize: chemical.packSize,
                unit: chemical.unit,
                location: 'main',
                quantityOnHand: 0, quantityReserved: 0, quantityAvailable: 0,
                averageCost: chemical.costPrice || 0,
                lastCost: chemical.costPrice || 0
            }).save();
        }

        res.status(201).json(chemical);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update chemical price (admin only)
app.put('/api/chemicals/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { productName, sourceSupplier, manufacturer, costPrice, adminPrice, sellPrice, adminMarginDollars, marginDollars, priceVersion, notes, equivalentProduct, isActive, availableForOrder,
                category, crops, defaultRate, rateUnit, unitsPerPack, packSize, unit, epaRegistrationNumber, signalWord, priceIsSpeculated } = req.body;

        const chemical = await Chemical.findById(req.params.id);
        if (!chemical) {
            return res.status(404).json({ error: 'Chemical not found' });
        }

        // If price changed, save to history
        if ((costPrice !== undefined && costPrice !== chemical.costPrice) ||
            (adminPrice !== undefined && adminPrice !== chemical.adminPrice) ||
            (sellPrice !== undefined && sellPrice !== chemical.sellPrice)) {
            await ChemicalPriceHistory.create({
                chemicalId: chemical._id,
                productName: chemical.productName,
                sourceSupplier: chemical.sourceSupplier,
                packSize: chemical.packSize,
                unit: chemical.unit,
                costPrice: costPrice || chemical.costPrice,
                sellPrice: sellPrice || chemical.sellPrice,
                priceVersion: priceVersion || new Date().toISOString().slice(0, 10),
                changedBy: req.user._id
            });

            if (costPrice !== undefined) chemical.costPrice = costPrice;
            if (adminPrice !== undefined) chemical.adminPrice = adminPrice;
            if (sellPrice !== undefined) chemical.sellPrice = sellPrice;
            chemical.priceDate = new Date();
        }

        // Handle margin dollar amounts (these recalculate prices)
        if (adminMarginDollars !== undefined) {
            chemical.adminMarginDollars = adminMarginDollars;
            // Recalculate admin price from cost + admin margin
            chemical.adminPrice = Math.round((chemical.costPrice + adminMarginDollars) * 100) / 100;
        }
        if (marginDollars !== undefined) {
            chemical.marginDollars = marginDollars;
            // Recalculate sell price from admin price + margin
            chemical.sellPrice = Math.round((chemical.adminPrice + marginDollars) * 100) / 100;
        }

        if (priceVersion !== undefined) chemical.priceVersion = priceVersion;
        if (notes !== undefined) chemical.notes = notes;
        if (equivalentProduct !== undefined) chemical.equivalentProduct = equivalentProduct;
        if (isActive !== undefined) chemical.isActive = isActive;
        if (availableForOrder !== undefined) chemical.availableForOrder = availableForOrder;
        if (productName !== undefined) chemical.productName = productName;
        if (sourceSupplier !== undefined) chemical.sourceSupplier = sourceSupplier;
        if (manufacturer !== undefined) chemical.manufacturer = manufacturer;
        if (category !== undefined) chemical.category = category;
        if (crops !== undefined) chemical.crops = crops;
        if (defaultRate !== undefined) chemical.defaultRate = defaultRate;
        if (rateUnit !== undefined) chemical.rateUnit = rateUnit;
        if (unitsPerPack !== undefined) chemical.unitsPerPack = unitsPerPack;
        if (packSize !== undefined) chemical.packSize = packSize;
        if (unit !== undefined) chemical.unit = unit;
        if (epaRegistrationNumber !== undefined) chemical.epaRegistrationNumber = epaRegistrationNumber;
        if (signalWord !== undefined) chemical.signalWord = signalWord;
        if (priceIsSpeculated !== undefined) chemical.priceIsSpeculated = priceIsSpeculated;

        chemical.updatedAt = new Date();
        await chemical.save();

        // Keep inventory record in sync with product catalog
        const inv = await Inventory.findOne({ chemicalId: chemical._id, location: 'main' });
        if (inv) {
            inv.productName = chemical.productName;
            inv.packSize = chemical.packSize;
            inv.unit = chemical.unit;
            inv.updatedAt = new Date();
            await inv.save();
        } else {
            // Create inventory record if missing
            await new Inventory({
                chemicalId: chemical._id,
                productName: chemical.productName,
                packSize: chemical.packSize,
                unit: chemical.unit,
                location: 'main',
                quantityOnHand: 0, quantityReserved: 0, quantityAvailable: 0,
                averageCost: chemical.costPrice || 0,
                lastCost: chemical.costPrice || 0
            }).save().catch(() => {}); // Ignore duplicate errors
        }

        res.json(chemical);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Bulk import chemicals (admin only)
app.post('/api/chemicals/bulk', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { chemicals, sourceSupplier, priceVersion, adminMargin = 10, retailMargin = 15 } = req.body;

        if (!Array.isArray(chemicals)) {
            return res.status(400).json({ error: 'chemicals must be an array' });
        }

        const results = [];
        const version = priceVersion || new Date().toISOString().slice(0, 10);

        for (const chem of chemicals) {
            // Calculate 3-tier pricing
            const costPrice = chem.costPrice || 0;
            const adminPrice = chem.adminPrice || costPrice * (1 + adminMargin / 100);
            const sellPrice = chem.sellPrice && chem.sellPrice !== costPrice
                ? chem.sellPrice
                : adminPrice * (1 + retailMargin / 100);

            let existing = await Chemical.findOne({
                productName: chem.productName,
                sourceSupplier: sourceSupplier || chem.sourceSupplier,
                packSize: chem.packSize
            });

            if (existing) {
                // Always update if costPrice changed OR if adminPrice is missing
                const needsUpdate = existing.costPrice !== costPrice ||
                                   !existing.adminPrice ||
                                   existing.adminPrice === 0;

                if (needsUpdate) {
                    await ChemicalPriceHistory.create({
                        chemicalId: existing._id,
                        productName: existing.productName,
                        sourceSupplier: existing.sourceSupplier,
                        packSize: existing.packSize,
                        unit: existing.unit,
                        costPrice: costPrice,
                        sellPrice: sellPrice,
                        priceVersion: version,
                        changedBy: req.user._id
                    });

                    existing.costPrice = costPrice;
                    existing.adminPrice = adminPrice;
                    existing.sellPrice = sellPrice;
                    existing.priceDate = new Date();
                    existing.priceVersion = version;
                    existing.updatedAt = new Date();
                    await existing.save();

                    results.push({ action: 'updated', chemical: existing });
                } else {
                    results.push({ action: 'unchanged', chemical: existing });
                }
            } else {
                const newChemical = new Chemical({
                    productName: chem.productName,
                    sourceSupplier: sourceSupplier || chem.sourceSupplier,
                    packSize: chem.packSize,
                    unit: chem.unit,
                    unitsPerPack: chem.unitsPerPack,
                    costPrice: costPrice,
                    adminPrice: adminPrice,
                    sellPrice: sellPrice,
                    category: chem.category || 'herbicide',
                    priceVersion: version,
                    equivalentProduct: chem.equivalentProduct,
                    notes: chem.notes,
                    createdBy: req.user._id
                });

                await newChemical.save();

                await ChemicalPriceHistory.create({
                    chemicalId: newChemical._id,
                    productName: newChemical.productName,
                    sourceSupplier: newChemical.sourceSupplier,
                    packSize: newChemical.packSize,
                    unit: newChemical.unit,
                    costPrice: newChemical.costPrice,
                    sellPrice: newChemical.sellPrice,
                    priceVersion: version,
                    changedBy: req.user._id
                });

                results.push({ action: 'created', chemical: newChemical });
            }
        }

        const summary = {
            total: results.length,
            created: results.filter(r => r.action === 'created').length,
            updated: results.filter(r => r.action === 'updated').length,
            unchanged: results.filter(r => r.action === 'unchanged').length
        };

        res.json({ summary, results });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete chemical (admin only)
app.delete('/api/chemicals/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Guard: must be archived (isActive=false) before hard delete. This
        // enforces the archive->delete two-step from the admin UI, so a single
        // misclick on a live product can't nuke it. Admin archives first,
        // confirms the product is actually dead, THEN hard-deletes.
        const existing = await Chemical.findById(req.params.id);
        if (!existing) {
            return res.status(404).json({ error: 'Chemical not found' });
        }
        if (existing.isActive !== false) {
            return res.status(400).json({
                error: 'Product must be archived before it can be deleted. Archive it first (isActive=false), then retry the delete.'
            });
        }
        await Chemical.findByIdAndDelete(req.params.id);
        res.json({ message: 'Chemical deleted', chemical: existing });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Archive (soft-delete) a product - hides it from customer views but keeps all
// order/invoice/inventory references intact. Reversible via /unarchive.
app.post('/api/chemicals/:id/archive', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const chemical = await Chemical.findById(req.params.id);
        if (!chemical) return res.status(404).json({ error: 'Chemical not found' });
        chemical.isActive = false;
        chemical.updatedAt = new Date();
        await chemical.save();
        res.json({ message: 'Chemical archived', chemical });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.post('/api/chemicals/:id/unarchive', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const chemical = await Chemical.findById(req.params.id);
        if (!chemical) return res.status(404).json({ error: 'Chemical not found' });
        chemical.isActive = true;
        chemical.updatedAt = new Date();
        await chemical.save();
        res.json({ message: 'Chemical unarchived', chemical });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Diagnostic: find every Chemical doc matching a productName and return the
// signals Kyle needs to decide which one is real vs duplicate. Case-insensitive
// substring match by default. Non-cancelled/non-draft order count per doc.
// Aggregate inventory across all locations per doc.
app.get('/api/admin/chemicals/duplicates-audit', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { productName } = req.query;
        if (!productName) {
            return res.status(400).json({ error: 'productName query parameter is required' });
        }

        const chemicals = await Chemical.find({
            productName: { $regex: productName, $options: 'i' }
        })
            .sort({ productName: 1, packSize: 1 })
            .lean();

        if (chemicals.length === 0) {
            return res.json({ query: productName, count: 0, matches: [] });
        }

        // Aggregate inventory totals per chemicalId across every location
        const chemIds = chemicals.map(c => c._id);
        const invAgg = await Inventory.aggregate([
            { $match: { chemicalId: { $in: chemIds } } },
            { $group: {
                _id: '$chemicalId',
                quantityOnHand: { $sum: { $ifNull: ['$quantityOnHand', 0] } },
                quantityReserved: { $sum: { $ifNull: ['$quantityReserved', 0] } },
                locationCount: { $sum: 1 }
            }}
        ]);
        const invMap = {};
        invAgg.forEach(r => { invMap[r._id.toString()] = r; });

        // Order count - orders that actually matter (not draft/cancelled)
        const matches = [];
        for (const chem of chemicals) {
            const inv = invMap[chem._id.toString()] || { quantityOnHand: 0, quantityReserved: 0, locationCount: 0 };
            const orderCount = await ChemicalOrder.countDocuments({
                'items.chemicalId': chem._id,
                status: { $nin: ['cancelled', 'draft'] }
            });
            // Also check if any SprayProgram references this chemicalId (so Kyle
            // knows the one linked to the milo program, etc.)
            const programRefCount = await SprayProgram.countDocuments({
                'applications.chemicals.chemicalId': chem._id
            });

            matches.push({
                _id: chem._id,
                productName: chem.productName,
                packSize: chem.packSize,
                unit: chem.unit,
                sourceSupplier: chem.sourceSupplier,
                isActive: chem.isActive !== false,
                isRestrictedUse: chem.isRestrictedUse === true,
                costPrice: chem.costPrice || 0,
                sellPrice: chem.sellPrice || 0,
                priceDate: chem.priceDate || null,
                quantityOnHand: inv.quantityOnHand,
                quantityReserved: inv.quantityReserved,
                inventoryLocations: inv.locationCount,
                orderCount,
                programRefCount,
                // Human-readable verdict hint
                recommendation: (inv.quantityOnHand > 0 || orderCount > 0 || programRefCount > 0)
                    ? 'KEEP - has inventory, orders, or program linkage'
                    : 'SAFE TO ARCHIVE - no inventory, no orders, no program refs'
            });
        }

        // Sort: active+linked first, dead rows last
        matches.sort((a, b) => {
            const score = m => (m.isActive ? 1 : 0) * 8 + (m.quantityOnHand > 0 ? 4 : 0) + (m.orderCount > 0 ? 2 : 0) + (m.programRefCount > 0 ? 1 : 0);
            return score(b) - score(a);
        });

        res.json({ query: productName, count: matches.length, matches });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ============ CHEMICAL QUOTE / PRICE COMPARISON API ROUTES ============

// Get all quotes with optional filters
app.get('/api/quotes', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { productName, supplier, isActive, sortBy } = req.query;
        const filter = {};

        if (productName) {
            filter.productName = { $regex: productName, $options: 'i' };
        }
        if (supplier) {
            filter.supplier = { $regex: supplier, $options: 'i' };
        }
        if (isActive !== undefined) {
            filter.isActive = isActive === 'true';
        }

        let sort = { quoteDate: -1 }; // Default: newest first
        if (sortBy === 'price') sort = { pricePerUnit: 1 };
        if (sortBy === 'product') sort = { productName: 1 };
        if (sortBy === 'supplier') sort = { supplier: 1 };

        const quotes = await ChemicalQuote.find(filter).sort(sort);
        res.json(quotes);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get unique product names for dropdown
app.get('/api/quotes/products', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const products = await ChemicalQuote.distinct('productName');
        res.json(products.sort());
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get unique suppliers for dropdown
app.get('/api/quotes/suppliers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const suppliers = await ChemicalQuote.distinct('supplier');
        res.json(suppliers.sort());
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Compare prices for a specific product across all suppliers
app.get('/api/quotes/compare/:productName', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const productName = decodeURIComponent(req.params.productName);

        // Get all active quotes for this product
        const quotes = await ChemicalQuote.find({
            productName: { $regex: `^${productName}$`, $options: 'i' },
            isActive: true
        }).sort({ pricePerUnit: 1 });

        if (quotes.length === 0) {
            return res.json({
                productName,
                quotes: [],
                bestDeal: null,
                comparison: null
            });
        }

        const bestDeal = quotes[0]; // Lowest price
        const highestPrice = quotes[quotes.length - 1];

        const comparison = {
            lowestPrice: bestDeal.pricePerUnit,
            highestPrice: highestPrice.pricePerUnit,
            priceDifference: Math.round((highestPrice.pricePerUnit - bestDeal.pricePerUnit) * 100) / 100,
            savingsPercent: Math.round(((highestPrice.pricePerUnit - bestDeal.pricePerUnit) / highestPrice.pricePerUnit) * 100 * 10) / 10,
            totalQuotes: quotes.length
        };

        res.json({
            productName,
            quotes,
            bestDeal,
            comparison
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get best deals across all products (lowest price per product)
app.get('/api/quotes/best-deals', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Aggregate to find best price for each product
        const bestDeals = await ChemicalQuote.aggregate([
            { $match: { isActive: true } },
            { $sort: { pricePerUnit: 1 } },
            {
                $group: {
                    _id: '$productName',
                    bestQuote: { $first: '$$ROOT' },
                    allPrices: { $push: { supplier: '$supplier', price: '$pricePerUnit', packSize: '$packSize' } },
                    count: { $sum: 1 }
                }
            },
            { $sort: { '_id': 1 } }
        ]);

        // Format results
        const results = bestDeals.map(item => ({
            productName: item._id,
            bestPrice: item.bestQuote.pricePerUnit,
            bestSupplier: item.bestQuote.supplier,
            packSize: item.bestQuote.packSize,
            unit: item.bestQuote.unit,
            quoteDate: item.bestQuote.quoteDate,
            notes: item.bestQuote.notes,
            alternativeCount: item.count - 1,
            alternatives: item.allPrices.slice(1) // Other options
        }));

        res.json(results);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Create a new quote
app.post('/api/quotes', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const quote = new ChemicalQuote({
            ...req.body,
            createdBy: req.user._id
        });

        // Calculate pack price if units provided
        if (quote.pricePerUnit && quote.unitsPerPack) {
            quote.packPrice = Math.round(quote.pricePerUnit * quote.unitsPerPack * 100) / 100;
        }

        await quote.save();
        res.status(201).json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Bulk import quotes (for entering multiple quotes at once)
app.post('/api/quotes/bulk', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { quotes } = req.body;
        if (!Array.isArray(quotes) || quotes.length === 0) {
            return res.status(400).json({ error: 'Quotes array is required' });
        }

        const results = {
            created: [],
            errors: []
        };

        for (const quoteData of quotes) {
            try {
                const quote = new ChemicalQuote({
                    ...quoteData,
                    createdBy: req.user._id
                });

                // Calculate pack price if units provided
                if (quote.pricePerUnit && quote.unitsPerPack) {
                    quote.packPrice = Math.round(quote.pricePerUnit * quote.unitsPerPack * 100) / 100;
                }

                await quote.save();
                results.created.push(quote);
            } catch (err) {
                results.errors.push({
                    data: quoteData,
                    error: err.message
                });
            }
        }

        res.status(201).json({
            message: `Created ${results.created.length} quotes, ${results.errors.length} errors`,
            ...results
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update a quote
app.put('/api/quotes/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const quote = await ChemicalQuote.findByIdAndUpdate(
            req.params.id,
            { ...req.body, updatedAt: new Date() },
            { new: true, runValidators: true }
        );

        if (!quote) {
            return res.status(404).json({ error: 'Quote not found' });
        }

        res.json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete a quote
app.delete('/api/quotes/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const quote = await ChemicalQuote.findByIdAndDelete(req.params.id);
        if (!quote) {
            return res.status(404).json({ error: 'Quote not found' });
        }
        res.json({ message: 'Quote deleted', quote });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Mark quote as purchased
app.put('/api/quotes/:id/purchased', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { purchaseOrderId } = req.body;
        const quote = await ChemicalQuote.findByIdAndUpdate(
            req.params.id,
            {
                isPurchased: true,
                purchaseDate: new Date(),
                purchaseOrderId,
                updatedAt: new Date()
            },
            { new: true }
        );

        if (!quote) {
            return res.status(404).json({ error: 'Quote not found' });
        }

        res.json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Deactivate old quotes (utility endpoint)
app.put('/api/quotes/deactivate-expired', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const result = await ChemicalQuote.updateMany(
            {
                expirationDate: { $lt: new Date() },
                isActive: true
            },
            {
                isActive: false,
                updatedAt: new Date()
            }
        );

        res.json({
            message: `Deactivated ${result.modifiedCount} expired quotes`
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ============ COMPLIANCE API ROUTES ============

// Get all RUP (Restricted Use Pesticide) products
app.get('/api/compliance/rup-products', authMiddleware, async (req, res) => {
    try {
        const rupProducts = await Chemical.find({ isRestrictedUse: true, isActive: true })
            .select('productName epaRegistrationNumber signalWord requiredCertifications category')
            .sort({ productName: 1 });
        res.json(rupProducts);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get products requiring special certifications
app.get('/api/compliance/products-by-certification/:certType', authMiddleware, async (req, res) => {
    try {
        const products = await Chemical.find({
            requiredCertifications: req.params.certType,
            isActive: true
        }).sort({ productName: 1 });
        res.json(products);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Check if user can purchase RUP products
app.get('/api/compliance/check-rup-eligibility/:userId', authMiddleware, async (req, res) => {
    try {
        const user = await User.findById(req.params.userId);
        if (!user) {
            return res.status(404).json({ error: 'User not found' });
        }

        const eligibility = {
            canPurchaseRUP: false,
            hasPrivateLicense: false,
            hasCommercialLicense: false,
            hasParaquatCert: false,
            hasDicambaCert: false,
            issues: []
        };

        const now = new Date();
        const currentYear = now.getFullYear();

        // Check private applicator license
        if (user.privateApplicatorLicense?.hasLicense) {
            if (user.privateApplicatorLicense.verificationStatus === 'verified') {
                if (user.privateApplicatorLicense.expirationDate > now) {
                    eligibility.hasPrivateLicense = true;
                } else {
                    eligibility.issues.push('Private applicator license is expired');
                }
            } else {
                eligibility.issues.push('Private applicator license pending verification');
            }
        }

        // Check commercial applicator license
        if (user.commercialApplicatorLicense?.hasLicense) {
            if (user.commercialApplicatorLicense.verificationStatus === 'verified') {
                if (user.commercialApplicatorLicense.expirationDate > now) {
                    eligibility.hasCommercialLicense = true;
                } else {
                    eligibility.issues.push('Commercial applicator license is expired');
                }
            } else {
                eligibility.issues.push('Commercial applicator license pending verification');
            }
        }

        // Check Paraquat certification (valid for 3 years)
        if (user.paraquatCertification?.completed) {
            if (user.paraquatCertification.expirationDate > now) {
                eligibility.hasParaquatCert = true;
            } else {
                eligibility.issues.push('Paraquat certification has expired');
            }
        }

        // Check Dicamba certification (must be current year)
        if (user.dicambaCertification?.completed) {
            if (user.dicambaCertification.trainingYear === currentYear) {
                eligibility.hasDicambaCert = true;
            } else {
                eligibility.issues.push(`Dicamba training is from ${user.dicambaCertification.trainingYear}, needs ${currentYear} training`);
            }
        }

        // User can purchase RUP if they have either a valid private or commercial license
        eligibility.canPurchaseRUP = eligibility.hasPrivateLicense || eligibility.hasCommercialLicense;

        res.json(eligibility);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Update user compliance/license info (admin or user themselves)
app.put('/api/compliance/user/:userId/license', authMiddleware, async (req, res) => {
    try {
        const { userId } = req.params;
        const updateData = req.body;

        // Only admin/distributor or the user themselves can update
        if (!isAdminLevel(req.user) && req.user._id.toString() !== userId) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        const user = await User.findById(userId);
        if (!user) {
            return res.status(404).json({ error: 'User not found' });
        }

        // Update license fields
        if (updateData.privateApplicatorLicense) {
            user.privateApplicatorLicense = { ...user.privateApplicatorLicense?.toObject(), ...updateData.privateApplicatorLicense };
            if (isAdminLevel(req.user)) {
                user.privateApplicatorLicense.verifiedBy = req.user._id;
                user.privateApplicatorLicense.verifiedAt = new Date();
            }
        }

        if (updateData.commercialApplicatorLicense) {
            user.commercialApplicatorLicense = { ...user.commercialApplicatorLicense?.toObject(), ...updateData.commercialApplicatorLicense };
            if (isAdminLevel(req.user)) {
                user.commercialApplicatorLicense.verifiedBy = req.user._id;
                user.commercialApplicatorLicense.verifiedAt = new Date();
            }
        }

        if (updateData.paraquatCertification) {
            user.paraquatCertification = { ...user.paraquatCertification?.toObject(), ...updateData.paraquatCertification };
            if (isAdminLevel(req.user)) {
                user.paraquatCertification.verifiedBy = req.user._id;
                user.paraquatCertification.verifiedAt = new Date();
            }
        }

        if (updateData.dicambaCertification) {
            user.dicambaCertification = { ...user.dicambaCertification?.toObject(), ...updateData.dicambaCertification };
            if (isAdminLevel(req.user)) {
                user.dicambaCertification.verifiedBy = req.user._id;
                user.dicambaCertification.verifiedAt = new Date();
            }
        }

        // Recalculate RUP eligibility
        const now = new Date();
        user.canPurchaseRUP = (
            (user.privateApplicatorLicense?.verificationStatus === 'verified' && user.privateApplicatorLicense?.expirationDate > now) ||
            (user.commercialApplicatorLicense?.verificationStatus === 'verified' && user.commercialApplicatorLicense?.expirationDate > now)
        );

        await user.save();
        res.json({ message: 'Compliance info updated', user });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Verify user license (admin only)
app.post('/api/compliance/verify-license/:userId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { userId } = req.params;
        const { licenseType, status, notes } = req.body;

        const user = await User.findById(userId);
        if (!user) {
            return res.status(404).json({ error: 'User not found' });
        }

        if (licenseType === 'private' && user.privateApplicatorLicense) {
            user.privateApplicatorLicense.verificationStatus = status;
            user.privateApplicatorLicense.verifiedBy = req.user._id;
            user.privateApplicatorLicense.verifiedAt = new Date();
        } else if (licenseType === 'commercial' && user.commercialApplicatorLicense) {
            user.commercialApplicatorLicense.verificationStatus = status;
            user.commercialApplicatorLicense.verifiedBy = req.user._id;
            user.commercialApplicatorLicense.verifiedAt = new Date();
        }

        if (notes) {
            user.rupEligibilityNotes = notes;
        }

        // Recalculate RUP eligibility
        const now = new Date();
        user.canPurchaseRUP = (
            (user.privateApplicatorLicense?.verificationStatus === 'verified' && user.privateApplicatorLicense?.expirationDate > now) ||
            (user.commercialApplicatorLicense?.verificationStatus === 'verified' && user.commercialApplicatorLicense?.expirationDate > now)
        );

        await user.save();
        res.json({ message: 'License verified', user });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Create RUP Sale Record (called during checkout for RUP products)
app.post('/api/compliance/rup-sale-record', authMiddleware, async (req, res) => {
    try {
        const recordData = req.body;

        // Get user compliance info
        const user = await User.findById(recordData.purchaserId);
        if (!user) {
            return res.status(404).json({ error: 'User not found' });
        }

        // Verify user can purchase RUP
        if (!user.canPurchaseRUP) {
            return res.status(403).json({ error: 'User is not eligible to purchase Restricted Use Pesticides' });
        }

        // Get chemical info
        const chemical = await Chemical.findById(recordData.chemicalId);

        // Build the record
        const record = new RupSaleRecord({
            ...recordData,
            sellerId: req.user._id,
            sellerName: req.user.name,
            purchaserName: user.name,
            purchaserAddress: {
                street: user.farm?.address || '',
                city: '',
                state: user.farm?.state || '',
                zip: user.farm?.zip || ''
            },
            purchaserPhone: user.phone,
            purchaserEmail: user.email,
            applicatorLicenseType: user.privateApplicatorLicense?.hasLicense ? 'private' : 'commercial',
            applicatorLicenseNumber: user.privateApplicatorLicense?.licenseNumber || user.commercialApplicatorLicense?.licenseNumber,
            applicatorLicenseState: user.privateApplicatorLicense?.state || user.commercialApplicatorLicense?.state,
            applicatorLicenseExpiration: user.privateApplicatorLicense?.expirationDate || user.commercialApplicatorLicense?.expirationDate,
            licenseVerificationMethod: 'document_on_file',
            licenseVerifiedBy: user.privateApplicatorLicense?.verifiedBy || user.commercialApplicatorLicense?.verifiedBy,
            licenseVerifiedAt: user.privateApplicatorLicense?.verifiedAt || user.commercialApplicatorLicense?.verifiedAt,
            epaRegistrationNumber: chemical?.epaRegistrationNumber || recordData.epaRegistrationNumber,
            signalWord: chemical?.signalWord,
            paraquatCertRequired: chemical?.requiredCertifications?.includes('paraquat_training'),
            paraquatCertVerified: user.paraquatCertification?.completed,
            paraquatCertNumber: user.paraquatCertification?.certificateNumber,
            paraquatCertDate: user.paraquatCertification?.completionDate,
            dicambaCertRequired: chemical?.requiredCertifications?.includes('dicamba_training'),
            dicambaCertVerified: user.dicambaCertification?.completed,
            dicambaCertYear: user.dicambaCertification?.trainingYear,
            createdBy: req.user._id
        });

        await record.save();
        res.status(201).json({ message: 'RUP sale record created', record });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get RUP Sale Records (admin - for compliance reporting)
app.get('/api/compliance/rup-sale-records', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { startDate, endDate, productName, purchaserId, limit = 100 } = req.query;

        const query = {};
        if (startDate || endDate) {
            query.saleDate = {};
            if (startDate) query.saleDate.$gte = new Date(startDate);
            if (endDate) query.saleDate.$lte = new Date(endDate);
        }
        if (productName) query.productName = new RegExp(productName, 'i');
        if (purchaserId) query.purchaserId = purchaserId;

        const records = await RupSaleRecord.find(query)
            .populate('purchaserId', 'name email')
            .populate('sellerId', 'name')
            .sort({ saleDate: -1 })
            .limit(parseInt(limit));

        res.json(records);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get compliance summary/dashboard (admin)
app.get('/api/compliance/dashboard', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const now = new Date();
        const currentYear = now.getFullYear();
        const thirtyDaysAgo = new Date(now - 30 * 24 * 60 * 60 * 1000);

        // Count RUP products
        const rupProductCount = await Chemical.countDocuments({ isRestrictedUse: true, isActive: true });

        // Count users with valid licenses
        const usersWithPrivateLicense = await User.countDocuments({
            'privateApplicatorLicense.verificationStatus': 'verified',
            'privateApplicatorLicense.expirationDate': { $gt: now }
        });

        const usersWithCommercialLicense = await User.countDocuments({
            'commercialApplicatorLicense.verificationStatus': 'verified',
            'commercialApplicatorLicense.expirationDate': { $gt: now }
        });

        // Count users pending verification
        const pendingVerification = await User.countDocuments({
            $or: [
                { 'privateApplicatorLicense.verificationStatus': 'pending' },
                { 'commercialApplicatorLicense.verificationStatus': 'pending' }
            ]
        });

        // Count users with expiring licenses (within 30 days)
        const expiringLicenses = await User.countDocuments({
            $or: [
                {
                    'privateApplicatorLicense.verificationStatus': 'verified',
                    'privateApplicatorLicense.expirationDate': { $gt: now, $lt: new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000) }
                },
                {
                    'commercialApplicatorLicense.verificationStatus': 'verified',
                    'commercialApplicatorLicense.expirationDate': { $gt: now, $lt: new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000) }
                }
            ]
        });

        // RUP sales in last 30 days
        const recentRupSales = await RupSaleRecord.countDocuments({
            saleDate: { $gte: thirtyDaysAgo }
        });

        // Products missing EPA registration
        const productsMissingEPA = await Chemical.countDocuments({
            isActive: true,
            category: { $in: ['herbicide', 'fungicide', 'insecticide'] },
            $or: [
                { epaRegistrationNumber: { $exists: false } },
                { epaRegistrationNumber: '' },
                { epaRegistrationNumber: null }
            ]
        });

        res.json({
            rupProductCount,
            usersWithPrivateLicense,
            usersWithCommercialLicense,
            totalLicensedUsers: usersWithPrivateLicense + usersWithCommercialLicense,
            pendingVerification,
            expiringLicenses,
            recentRupSales,
            productsMissingEPA,
            currentYear
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get users pending license verification (admin)
app.get('/api/compliance/pending-verification', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const users = await User.find({
            $or: [
                { 'privateApplicatorLicense.verificationStatus': 'pending', 'privateApplicatorLicense.hasLicense': true },
                { 'commercialApplicatorLicense.verificationStatus': 'pending', 'commercialApplicatorLicense.hasLicense': true }
            ]
        }).select('name email phone farm privateApplicatorLicense commercialApplicatorLicense createdAt');

        res.json(users);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get users with expiring licenses (admin)
app.get('/api/compliance/expiring-licenses', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const now = new Date();
        const sixtyDaysFromNow = new Date(now.getTime() + 60 * 24 * 60 * 60 * 1000);

        const users = await User.find({
            $or: [
                {
                    'privateApplicatorLicense.verificationStatus': 'verified',
                    'privateApplicatorLicense.expirationDate': { $gt: now, $lt: sixtyDaysFromNow }
                },
                {
                    'commercialApplicatorLicense.verificationStatus': 'verified',
                    'commercialApplicatorLicense.expirationDate': { $gt: now, $lt: sixtyDaysFromNow }
                }
            ]
        }).select('name email phone farm privateApplicatorLicense commercialApplicatorLicense');

        res.json(users);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ============ COMPANY SETTINGS ROUTES ============

// Get public company settings (license info for invoices/distributors)
app.get('/api/company/license', async (req, res) => {
    try {
        let settings = await CompanySettings.findOne();

        if (!settings) {
            // Return default if no settings exist
            return res.json({
                companyName: 'AcreProfit, LLC',
                pesticideDealerLicense: null
            });
        }

        // Only return public license information
        const publicInfo = {
            companyName: settings.companyName,
            pesticideDealerLicense: settings.pesticideDealerLicense?.displayOnInvoices ? {
                licenseNumber: settings.pesticideDealerLicense.licenseNumber,
                state: settings.pesticideDealerLicense.state,
                status: settings.pesticideDealerLicense.status
            } : null
        };

        res.json(publicInfo);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get full company settings (superadmin only)
app.get('/api/admin/company-settings', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        let settings = await CompanySettings.findOne();

        if (!settings) {
            // Create default settings if none exist
            settings = new CompanySettings({
                companyName: 'AcreProfit, LLC',
                pesticideDealerLicense: {
                    licenseNumber: '90575',
                    state: 'CO',
                    expirationDate: new Date('2026-12-31'),
                    status: 'active',
                    displayOnInvoices: true,
                    displayToDistributors: true
                },
                confidentialLicenseInfo: {
                    agLicenseId: '0050BD',
                    pin: '110081',
                    portalUrl: 'https://www.ag.state.co.us/elicense/SecurityLogin.aspx'
                },
                reminders: [{
                    title: 'Pesticide Dealer License Renewal',
                    description: 'Renew Colorado Pesticide Dealer License before December 31',
                    dueDate: new Date('2026-12-01'),
                    reminderType: 'license_renewal',
                    status: 'pending'
                }]
            });
            await settings.save();
        }

        res.json(settings);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Update company settings (superadmin only)
app.put('/api/admin/company-settings', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const updates = req.body;

        let settings = await CompanySettings.findOne();

        if (!settings) {
            settings = new CompanySettings(updates);
        } else {
            // Update allowed fields
            if (updates.companyName) settings.companyName = updates.companyName;
            if (updates.companyAddress) settings.companyAddress = updates.companyAddress;
            if (updates.companyPhone) settings.companyPhone = updates.companyPhone;
            if (updates.companyEmail) settings.companyEmail = updates.companyEmail;
            if (updates.pesticideDealerLicense) {
                settings.pesticideDealerLicense = {
                    ...settings.pesticideDealerLicense,
                    ...updates.pesticideDealerLicense
                };
            }
            if (updates.confidentialLicenseInfo) {
                settings.confidentialLicenseInfo = {
                    ...settings.confidentialLicenseInfo,
                    ...updates.confidentialLicenseInfo
                };
            }
        }

        settings.updatedAt = new Date();
        settings.updatedBy = req.user._id;

        await settings.save();
        res.json({ message: 'Settings updated successfully', settings });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Add reminder (admin/superadmin)
app.post('/api/admin/company-settings/reminders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { title, description, dueDate, reminderType } = req.body;

        let settings = await CompanySettings.findOne();
        if (!settings) {
            settings = new CompanySettings();
        }

        settings.reminders.push({
            title,
            description,
            dueDate: new Date(dueDate),
            reminderType: reminderType || 'other',
            status: 'pending'
        });

        settings.updatedAt = new Date();
        await settings.save();

        res.json({ message: 'Reminder added', reminders: settings.reminders });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Update reminder status (admin/superadmin)
app.put('/api/admin/company-settings/reminders/:reminderId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { reminderId } = req.params;
        const { status } = req.body;

        const settings = await CompanySettings.findOne();
        if (!settings) {
            return res.status(404).json({ error: 'Settings not found' });
        }

        const reminder = settings.reminders.id(reminderId);
        if (!reminder) {
            return res.status(404).json({ error: 'Reminder not found' });
        }

        reminder.status = status;
        if (status === 'completed') {
            reminder.completedAt = new Date();
            reminder.completedBy = req.user._id;
        }

        settings.updatedAt = new Date();
        await settings.save();

        res.json({ message: 'Reminder updated', reminder });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get active reminders (for admin dashboard)
app.get('/api/admin/reminders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const settings = await CompanySettings.findOne();
        if (!settings) {
            return res.json({ reminders: [] });
        }

        const activeReminders = settings.reminders.filter(r =>
            r.status === 'pending' && new Date(r.dueDate) >= new Date()
        ).sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate));

        res.json({ reminders: activeReminders });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ---- SUPPLIER ROUTES ----

// Get supplier profile
app.get('/api/supplier/profile', authMiddleware, supplierMiddleware, async (req, res) => {
    try {
        res.json({
            id: req.user._id,
            name: req.user.name,
            email: req.user.email,
            companyName: req.user.companyName,
            supplierCode: req.user.supplierCode,
            phone: req.user.phone
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Update supplier profile
app.put('/api/supplier/profile', authMiddleware, supplierMiddleware, async (req, res) => {
    try {
        const { companyName, phone, address } = req.body;

        if (companyName) req.user.companyName = companyName;
        if (phone) req.user.phone = phone;
        if (address) req.user.address = address;

        await req.user.save();

        res.json({
            message: 'Profile updated successfully',
            user: {
                id: req.user._id,
                name: req.user.name,
                email: req.user.email,
                companyName: req.user.companyName,
                supplierCode: req.user.supplierCode
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get supplier's products (chemicals linked to this supplier)
app.get('/api/supplier/products', authMiddleware, supplierMiddleware, async (req, res) => {
    try {
        // Find chemicals by supplierId OR by sourceSupplier matching supplierCode
        const chemicals = await Chemical.find({
            $or: [
                { supplierId: req.user._id },
                { sourceSupplier: req.user.supplierCode }
            ],
            isActive: true
        }).sort({ productName: 1 });

        res.json(chemicals);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Update product cost price (supplier sets their price)
app.put('/api/supplier/products/:id/price', authMiddleware, supplierMiddleware, async (req, res) => {
    try {
        const { costPrice } = req.body;

        if (costPrice === undefined || costPrice < 0) {
            return res.status(400).json({ error: 'Valid cost price is required' });
        }

        const chemical = await Chemical.findOne({
            _id: req.params.id,
            $or: [
                { supplierId: req.user._id },
                { sourceSupplier: req.user.supplierCode }
            ]
        });

        if (!chemical) {
            return res.status(404).json({ error: 'Product not found or not authorized' });
        }

        const oldCostPrice = chemical.costPrice;
        chemical.costPrice = costPrice;
        chemical.updatedAt = new Date();
        await chemical.save();

        // Record price history
        await ChemicalPriceHistory.create({
            chemicalId: chemical._id,
            productName: chemical.productName,
            sourceSupplier: chemical.sourceSupplier,
            packSize: chemical.packSize,
            unit: chemical.unit,
            costPrice: costPrice,
            sellPrice: chemical.sellPrice,
            priceVersion: `supplier-update-${new Date().toISOString().split('T')[0]}`,
            changedBy: req.user._id
        });

        res.json({
            message: 'Price updated successfully',
            product: {
                id: chemical._id,
                productName: chemical.productName,
                packSize: chemical.packSize,
                oldCostPrice,
                newCostPrice: costPrice
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Bulk update product prices (supplier)
app.put('/api/supplier/products/bulk-price', authMiddleware, supplierMiddleware, async (req, res) => {
    try {
        const { updates } = req.body; // Array of { productId, costPrice }

        if (!Array.isArray(updates) || updates.length === 0) {
            return res.status(400).json({ error: 'Updates array is required' });
        }

        const results = [];

        for (const update of updates) {
            const chemical = await Chemical.findOne({
                _id: update.productId,
                $or: [
                    { supplierId: req.user._id },
                    { sourceSupplier: req.user.supplierCode }
                ]
            });

            if (chemical && update.costPrice >= 0) {
                const oldCostPrice = chemical.costPrice;
                chemical.costPrice = update.costPrice;
                chemical.updatedAt = new Date();
                await chemical.save();

                await ChemicalPriceHistory.create({
                    chemicalId: chemical._id,
                    productName: chemical.productName,
                    sourceSupplier: chemical.sourceSupplier,
                    packSize: chemical.packSize,
                    unit: chemical.unit,
                    costPrice: update.costPrice,
                    sellPrice: chemical.sellPrice,
                    priceVersion: `supplier-bulk-${new Date().toISOString().split('T')[0]}`,
                    changedBy: req.user._id
                });

                results.push({
                    productId: chemical._id,
                    productName: chemical.productName,
                    oldCostPrice,
                    newCostPrice: update.costPrice,
                    status: 'updated'
                });
            } else {
                results.push({
                    productId: update.productId,
                    status: 'skipped',
                    reason: chemical ? 'Invalid price' : 'Not found or not authorized'
                });
            }
        }

        res.json({
            message: 'Bulk price update completed',
            results
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get supplier's order history (orders containing their products)
app.get('/api/supplier/orders', authMiddleware, supplierMiddleware, async (req, res) => {
    try {
        // Get all chemicals for this supplier
        const supplierChemicals = await Chemical.find({
            $or: [
                { supplierId: req.user._id },
                { sourceSupplier: req.user.supplierCode }
            ]
        }).select('productName');

        const chemicalNames = supplierChemicals.map(c => c.productName);

        // Find chemical orders that include products from this supplier
        const orders = await ChemicalOrder.find({
            'items.productName': { $in: chemicalNames },
            status: { $nin: ['draft', 'cancelled'] }
        })
        .populate('userId', 'name farm')
        .sort({ createdAt: -1 })
        .limit(100);

        // Filter items to only show this supplier's products
        const ordersWithSupplierItems = orders.map(order => {
            const supplierItems = order.items.filter(item =>
                chemicalNames.includes(item.productName)
            );
            return {
                _id: order._id,
                orderNumber: order.orderNumber,
                status: order.status,
                createdAt: order.createdAt,
                customer: order.userId ? {
                    name: order.userId.name,
                    farm: order.userId.farm?.name
                } : null,
                items: supplierItems,
                supplierTotal: supplierItems.reduce((sum, item) => sum + (item.quantity * item.unitPrice), 0)
            };
        });

        res.json(ordersWithSupplierItems);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Admin: Create supplier account
app.post('/api/admin/suppliers', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const { name, email, password, companyName, supplierCode, phone } = req.body;

        if (!name || !email || !password || !companyName || !supplierCode) {
            return res.status(400).json({
                error: 'Name, email, password, company name, and supplier code are required'
            });
        }

        const existingUser = await User.findOne({ email: email.toLowerCase() });
        if (existingUser) {
            return res.status(400).json({ error: 'Email already registered' });
        }

        const existingCode = await User.findOne({ supplierCode: supplierCode.toUpperCase() });
        if (existingCode) {
            return res.status(400).json({ error: 'Supplier code already in use' });
        }

        const supplier = new User({
            name,
            email: email.toLowerCase(),
            password,
            role: 'supplier',
            companyName,
            supplierCode: supplierCode.toUpperCase(),
            phone
        });

        await supplier.save();

        // Link existing chemicals with this supplier code to the new supplier account
        await Chemical.updateMany(
            { sourceSupplier: supplierCode.toUpperCase(), supplierId: { $exists: false } },
            { supplierId: supplier._id }
        );

        res.status(201).json({
            message: 'Supplier created successfully',
            supplier: {
                id: supplier._id,
                name: supplier.name,
                email: supplier.email,
                companyName: supplier.companyName,
                supplierCode: supplier.supplierCode
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Admin: Get all suppliers (read is available to any admin tier so distributors can
// populate the bid-sheet supplier checkbox list; writes below stay on superadmin).
app.get('/api/admin/suppliers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Optional ?bidEligible=true filter for the Price Mining selection modal
        const query = { role: 'supplier' };
        if (req.query.bidEligible === 'true') query.bidEligible = true;
        if (req.query.bidEligible === 'false') query.bidEligible = false;

        const suppliers = await User.find(query)
            .select('name email companyName supplierCode phone bidEligible createdAt')
            .sort({ companyName: 1 });

        // Get product count for each supplier
        const suppliersWithCounts = await Promise.all(suppliers.map(async (supplier) => {
            const productCount = await Chemical.countDocuments({
                $or: [
                    { supplierId: supplier._id },
                    { sourceSupplier: supplier.supplierCode }
                ],
                isActive: true
            });
            return {
                ...supplier.toObject(),
                productCount
            };
        }));

        res.json(suppliersWithCounts);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Admin: Update supplier
app.put('/api/admin/suppliers/:id', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const { name, companyName, phone, email, address, notes, bidEligible } = req.body;

        const supplier = await User.findOne({ _id: req.params.id, role: 'supplier' });
        if (!supplier) {
            return res.status(404).json({ error: 'Supplier not found' });
        }

        if (name !== undefined) supplier.name = name;
        if (companyName !== undefined) supplier.companyName = companyName;
        if (phone !== undefined) supplier.phone = phone;
        if (email !== undefined) supplier.email = email.toLowerCase();
        if (address !== undefined) supplier.address = address;
        if (notes !== undefined) supplier.notes = notes;
        if (bidEligible !== undefined) {
            supplier.bidEligible = bidEligible;
            supplier.markModified('bidEligible');
        }

        await supplier.save();

        res.json({
            message: 'Supplier updated successfully',
            supplier: {
                id: supplier._id,
                name: supplier.name,
                email: supplier.email,
                companyName: supplier.companyName,
                supplierCode: supplier.supplierCode,
                phone: supplier.phone,
                address: supplier.address,
                notes: supplier.notes,
                bidEligible: supplier.bidEligible
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Admin: Delete a supplier (superadmin only, scorched-earth - unlinks chemicals)
app.delete('/api/admin/suppliers/:id', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const supplier = await User.findOne({ _id: req.params.id, role: 'supplier' });
        if (!supplier) {
            return res.status(404).json({ error: 'Supplier not found' });
        }

        // Unlink any Chemical docs pointing at this supplier - don't orphan
        // the reference, just clear it. sourceSupplier string stays for history.
        const unlinkResult = await Chemical.updateMany(
            { supplierId: supplier._id },
            { $unset: { supplierId: '' } }
        );

        const snapshot = {
            id: supplier._id,
            email: supplier.email,
            supplierCode: supplier.supplierCode,
            companyName: supplier.companyName
        };

        await User.deleteOne({ _id: supplier._id });

        await logAudit({
            action: 'supplier_delete',
            req,
            entityType: 'User',
            entityId: snapshot.id,
            entityRef: snapshot.supplierCode || snapshot.email,
            before: snapshot,
            reason: req.body?.reason || `Supplier deleted by ${req.user.name}; ${unlinkResult.modifiedCount} chemicals unlinked`
        });

        res.json({
            message: `Supplier ${snapshot.supplierCode || snapshot.email} deleted`,
            chemicalsUnlinked: unlinkResult.modifiedCount
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Admin: Link chemical to supplier
app.put('/api/admin/chemicals/:chemicalId/link-supplier/:supplierId', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const chemical = await Chemical.findById(req.params.chemicalId);
        if (!chemical) {
            return res.status(404).json({ error: 'Chemical not found' });
        }

        const supplier = await User.findOne({ _id: req.params.supplierId, role: 'supplier' });
        if (!supplier) {
            return res.status(404).json({ error: 'Supplier not found' });
        }

        chemical.supplierId = supplier._id;
        chemical.sourceSupplier = supplier.supplierCode;
        await chemical.save();

        res.json({
            message: 'Chemical linked to supplier successfully',
            chemical: {
                id: chemical._id,
                productName: chemical.productName,
                supplierId: supplier._id,
                sourceSupplier: supplier.supplierCode
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ---- CHEMICAL ORDER ROUTES ----

// Create chemical order (customer)
app.post('/api/chemical-orders', authMiddleware, async (req, res) => {
    try {
        const {
            items,
            programId,
            programName,
            totalAcres,
            customerNotes,
            sprayParams,
            actingAsCustomerId,
            discountCode,
            marginAdjustment,
            marginAdjustmentReason
        } = req.body;

        // Only admin/distributor/superadmin can adjust margins - re-check on server
        const isAdminOrDistributor = ['admin', 'superadmin', 'distributor'].includes(req.user.role);
        const applyMarginAdjust = isAdminOrDistributor && marginAdjustment ? parseFloat(marginAdjustment) : 0;

        // Validate & normalize discount code
        let discountType = null;
        let discountDescription = '';
        if (isAdminOrDistributor && discountCode) {
            const code = String(discountCode).trim().toUpperCase();
            if (code === 'NODISTMARG') {
                discountType = 'no_dist_margin';
                discountDescription = 'No distributor margin (admin price)';
            } else if (code === 'ATCOSTAP') {
                discountType = 'at_cost';
                discountDescription = 'At cost (no markup)';
            }
        }

        // If distributor/admin is acting on behalf of a customer, the order userId = customer
        let orderUserId = req.user._id;
        let orderRepId = req.user.representative;
        if (actingAsCustomerId && isAdminOrDistributor) {
            const customer = await User.findById(actingAsCustomerId);
            if (customer && customer.role === 'customer') {
                orderUserId = customer._id;
                // Rep is the acting user (if they are a distributor) or the customer's assigned rep
                orderRepId = req.user.role === 'distributor' ? req.user._id : (customer.representative || customer.representativeId || req.user._id);
            }
        }

        // RUP COMPLIANCE CHECK
        const allChemicalsForCheck = await Chemical.find({ isActive: true }).select('productName isRestrictedUse requiredCertifications').lean();
        const customerForCompliance = (orderUserId.toString() !== req.user._id.toString())
            ? await User.findById(orderUserId).lean()
            : req.user;
        const itemsForCheck = items.map(i => {
            const chem = allChemicalsForCheck.find(c => c._id.toString() === (i.chemicalId || '').toString());
            return { productName: chem?.productName || i.productName };
        });
        const compliance = await validateRupCompliance({
            customer: customerForCompliance,
            items: itemsForCheck,
            allChemicals: allChemicalsForCheck
        });
        if (!compliance.ok) {
            await logAudit({
                action: 'rup_block',
                req,
                targetUser: customerForCompliance?._id,
                targetUserName: customerForCompliance?.name,
                entityType: 'ChemicalOrder',
                reason: 'RUP compliance check failed',
                after: { errors: compliance.errors }
            });
            return res.status(403).json({
                error: 'Restricted Use Pesticide compliance check failed',
                rupErrors: compliance.errors,
                userMessage: 'One or more products require a valid applicator license.'
            });
        }

        // Calculate totals with discount / margin adjustment applied server-side
        let subtotal = 0;
        let totalDiscount = 0;
        const orderItems = [];

        for (const item of items) {
            const chemical = await Chemical.findById(item.chemicalId);
            if (!chemical) continue;

            // Determine base unit price from discount code
            let unitPrice = chemical.sellPrice;
            if (discountType === 'at_cost') {
                unitPrice = chemical.costPrice || chemical.sellPrice;
            } else if (discountType === 'no_dist_margin') {
                unitPrice = chemical.adminPrice || chemical.sellPrice;
            }

            // Apply manual per-unit margin adjustment (floor at cost)
            if (applyMarginAdjust !== 0) {
                unitPrice = Math.max(chemical.costPrice || 0, unitPrice + applyMarginAdjust);
            }

            const totalPrice = Math.round(item.quantity * unitPrice * 100) / 100;
            subtotal += totalPrice;

            if (chemical.sellPrice) {
                totalDiscount += (chemical.sellPrice - unitPrice) * item.quantity;
            }

            orderItems.push({
                chemicalId: chemical._id,
                productName: chemical.productName,
                packSize: chemical.packSize,
                unit: chemical.unit,
                quantity: item.quantity,
                unitPrice,
                totalPrice,
                acres: item.acres,
                rate: item.rate,
                rateUnit: item.rateUnit,
                calculatedAmount: item.calculatedAmount
            });
        }

        subtotal = Math.round(subtotal * 100) / 100;
        totalDiscount = Math.round(totalDiscount * 100) / 100;

        // Build reason string for audit + order record
        let discountReasonFull = discountDescription;
        if (applyMarginAdjust !== 0) {
            const sign = applyMarginAdjust < 0 ? 'off' : 'added';
            discountReasonFull = (discountReasonFull ? discountReasonFull + ' | ' : '') +
                `$${Math.abs(applyMarginAdjust).toFixed(2)}/unit ${sign} by ${req.user.name}` +
                (marginAdjustmentReason ? ` - ${marginAdjustmentReason}` : '');
        }

        const order = new ChemicalOrder({
            userId: orderUserId,
            representativeId: orderRepId,
            createdBy: req.user._id,
            orderType: programId ? 'program' : 'direct',
            items: orderItems,
            programId,
            programName,
            totalAcres: totalAcres || sprayParams?.acres,
            subtotal,
            total: subtotal,
            customerNotes,
            sprayParams,
            discount: totalDiscount,
            discountReason: discountReasonFull || undefined,
            marginAdjustment: applyMarginAdjust || 0,
            marginAdjustmentReason: marginAdjustmentReason || '',
            marginAdjustedBy: applyMarginAdjust !== 0 ? req.user._id : undefined,
            status: 'draft'
        });

        await order.save();

        // Audit log margin adjustments or discount code usage
        if (applyMarginAdjust !== 0 || discountType) {
            await logAudit({
                action: 'margin_change',
                req,
                entityType: 'ChemicalOrder',
                entityId: order._id,
                entityRef: order.orderNumber,
                amount: totalDiscount,
                reason: discountReasonFull,
                before: { totalDiscount: 0 },
                after: {
                    totalDiscount,
                    marginAdjustment: applyMarginAdjust,
                    discountCode: discountCode || null,
                    programName: programName || null
                }
            });
        }

        // I5: create pending_verification RupSaleRecords for any RUP items on this order
        await autoCreateRupRecordsForOrder(order, { status: 'pending_verification' });

        res.status(201).json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ CHECKOUT ENDPOINT ============
// Full checkout flow - creates order with payment info
app.post('/api/chemical-orders/checkout', authMiddleware, async (req, res) => {
    try {
        const {
            year,
            location,
            items,
            subtotal,
            processingFee,
            total,
            paymentMethod,
            customerName,
            customerEmail,
            customerPhone,
            customerFarm,
            notes,
            discountCode,
            deliveryOption,
            deliveryAddress,
            marginAdjustment,
            marginAdjustmentReason
        } = req.body;

        // Only admin/distributor/superadmin can adjust margins
        const isAdminOrDistributor = ['admin', 'superadmin', 'distributor'].includes(req.user.role);
        const applyMarginAdjust = isAdminOrDistributor && marginAdjustment ? parseFloat(marginAdjustment) : 0;

        // Freight: minimum $400 for delivery
        const freightCharge = deliveryOption === 'delivery' ? 400 : 0;

        // Validate required fields
        if (!items || items.length === 0) {
            return res.status(400).json({ error: 'No items in order' });
        }

        if (!paymentMethod) {
            return res.status(400).json({ error: 'Payment method required' });
        }

        // Find or assign representative based on location
        let representativeId = req.user.representative;
        const repMap = {
            'mcconnell': 'kyle',
            'bamford': 'chad',
            'mollohan': 'ty'
        };
        const repCode = repMap[location] || 'kyle';

        // Try to find the rep user
        if (!representativeId) {
            const repUser = await User.findOne({
                $or: [
                    { representativeId: repCode },
                    { email: { $regex: new RegExp(repCode, 'i') } }
                ],
                role: { $in: ['admin', 'distributor', 'superadmin'] }
            });
            if (repUser) {
                representativeId = repUser._id;
            }
        }

        // Validate discount code
        let discountType = null;
        let discountDescription = '';
        if (discountCode) {
            const code = discountCode.trim().toUpperCase();
            if (code === 'NODISTMARG') {
                discountType = 'no_dist_margin';
                discountDescription = 'No distributor margin (admin price)';
            } else if (code === 'ATCOSTAP') {
                discountType = 'at_cost';
                discountDescription = 'At cost (no markup)';
            }
        }

        // Build order items with server-side price verification
        const allChemicals = await Chemical.find({ isActive: true }).lean();

        // RUP COMPLIANCE CHECK - block if customer doesn't have valid licenses
        // For acting-as orders, validate against the customer (orderUserId), not the distributor
        const customerForCompliance = (orderUserId && orderUserId.toString() !== req.user._id.toString())
            ? await User.findById(orderUserId).lean()
            : req.user;

        const compliance = await validateRupCompliance({
            customer: customerForCompliance,
            items,
            allChemicals
        });

        if (!compliance.ok) {
            await logAudit({
                action: 'rup_block',
                req,
                targetUser: customerForCompliance?._id,
                targetUserName: customerForCompliance?.name,
                entityType: 'ChemicalOrder',
                reason: 'RUP compliance check failed - order blocked',
                after: { errors: compliance.errors }
            });
            return res.status(403).json({
                error: 'Restricted Use Pesticide compliance check failed',
                rupErrors: compliance.errors,
                userMessage: 'One or more products in your order require a valid applicator license. See details below.'
            });
        }

        const orderItems = [];
        let verifiedSubtotal = 0;
        let totalDiscount = 0;

        for (const item of items) {
            const productName = item.productName || item.name;
            const qty = item.qty || item.quantity || 1;
            let unitPrice = item.price || 0;
            let isCustom = item.isCustom || false;

            // Server-side price verification: look up the chemical's pricing
            if (!isCustom && productName) {
                const chemical = allChemicals.find(c =>
                    c.productName === productName ||
                    c.productName.toLowerCase() === productName.toLowerCase()
                );
                if (chemical) {
                    if (discountType === 'at_cost') {
                        unitPrice = chemical.costPrice || chemical.sellPrice;
                    } else if (discountType === 'no_dist_margin') {
                        unitPrice = chemical.adminPrice || chemical.sellPrice;
                    } else {
                        unitPrice = chemical.sellPrice;
                    }

                    // Apply distributor margin adjustment (admin/distributor only)
                    // Negative value = discount, positive = surcharge
                    if (applyMarginAdjust !== 0) {
                        unitPrice = Math.max(chemical.costPrice || 0, unitPrice + applyMarginAdjust);
                    }

                    // Track the discount amount
                    if (chemical.sellPrice) {
                        totalDiscount += (chemical.sellPrice - unitPrice) * qty;
                    }
                }
            }

            const totalPrice = Math.round(qty * unitPrice * 100) / 100;
            verifiedSubtotal += totalPrice;

            orderItems.push({
                productName,
                packSize: item.packSize || '',
                unit: item.unit || 'pack',
                quantity: qty,
                unitPrice,
                totalPrice,
                timing: item.timing || '',
                isCustom
            });
        }

        // Server calculates the final totals (never trust frontend totals)
        const calculatedSubtotal = Math.round(verifiedSubtotal * 100) / 100;
        const calculatedWithFreight = calculatedSubtotal + freightCharge;
        const calculatedFee = paymentMethod === 'ach' ? Math.min(calculatedWithFreight * 0.008, 5) : 0;
        const calculatedTotal = Math.round((calculatedWithFreight + calculatedFee) * 100) / 100;

        // Build reason string
        let discountReasonFull = discountDescription;
        if (applyMarginAdjust !== 0) {
            const sign = applyMarginAdjust < 0 ? 'off' : 'added';
            discountReasonFull = (discountReasonFull ? discountReasonFull + ' | ' : '') +
                `$${Math.abs(applyMarginAdjust).toFixed(2)}/unit ${sign} by ${req.user.name}` +
                (marginAdjustmentReason ? ` - ${marginAdjustmentReason}` : '');
        }

        // Create the order
        const order = new ChemicalOrder({
            userId: req.user._id,
            representativeId: representativeId,
            orderType: 'direct',
            items: orderItems,
            totalAcres: 0,
            subtotal: calculatedSubtotal,
            freight: freightCharge,
            deliveryOption: deliveryOption || 'pickup',
            deliveryAddress: deliveryOption === 'delivery' ? deliveryAddress : '',
            processingFee: calculatedFee,
            total: calculatedTotal,
            customerNotes: notes,
            discount: Math.round(totalDiscount * 100) / 100,
            discountReason: discountReasonFull || undefined,
            marginAdjustment: applyMarginAdjust || 0,
            marginAdjustmentReason: marginAdjustmentReason || '',
            marginAdjustedBy: applyMarginAdjust !== 0 ? req.user._id : undefined,
            status: 'submitted',
            paymentMethod: paymentMethod,
            paymentStatus: paymentMethod === 'check' ? 'pending' : 'processing',
            submittedAt: new Date(),
            // Extra fields for contact info
            contactInfo: {
                name: customerName || req.user.name,
                email: customerEmail || req.user.email,
                phone: customerPhone || req.user.phone,
                farm: customerFarm
            },
            pickupLocation: location,
            year: year || new Date().getFullYear()
        });

        await order.save();

        // I5: create pending_verification RupSaleRecords for any RUP items on this order
        await autoCreateRupRecordsForOrder(order, { status: 'pending_verification' });

        // Audit log margin adjustments or discount code usage
        if (applyMarginAdjust !== 0 || discountType) {
            await logAudit({
                action: 'margin_change',
                req,
                entityType: 'ChemicalOrder',
                entityId: order._id,
                entityRef: order.orderNumber,
                amount: Math.round(totalDiscount * 100) / 100,
                reason: discountReasonFull,
                before: { totalDiscount: 0 },
                after: {
                    totalDiscount: Math.round(totalDiscount * 100) / 100,
                    marginAdjustment: applyMarginAdjust,
                    discountCode: discountCode || null
                }
            });
        }

        // Update user info if provided
        if (customerPhone && !req.user.phone) {
            req.user.phone = customerPhone;
            await req.user.save();
        }

        // Send order confirmation email
        try {
            const transporter = createEmailTransporter();
            const toEmail = customerEmail || req.user.email;
            if (transporter && toEmail) {
                const locationNames = {
                    mcconnell: 'McConnell Farm - Haxtun, CO',
                    bamford: 'Bamford Farm - Kirk, CO',
                    mollohan: 'Mollohan Farm - Otis, CO'
                };
                const paymentLabels = {
                    check: 'Check',
                    stripe: 'Credit Card',
                    stripe_ach: 'ACH Bank Transfer',
                    cash: 'Cash'
                };
                const pickupName = locationNames[location] || location || 'TBD';
                const paymentLabel = paymentLabels[paymentMethod] || paymentMethod;

                const itemsHtml = order.items.map(item =>
                    `<tr>
                        <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${item.productName}</td>
                        <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: center;">${item.quantity} ${item.unit || 'units'}</td>
                        <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right;">$${(item.totalPrice || 0).toFixed(2)}</td>
                    </tr>`
                ).join('');

                await transporter.sendMail({
                    from: process.env.EMAIL_FROM || process.env.SMTP_USER || process.env.GMAIL_USER,
                    to: toEmail,
                    subject: `Order Confirmation - ${order.orderNumber} - Acre Profit`,
                    html: `<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                        <h2 style="color: #2d5a27;">Order Confirmed!</h2>
                        <p>Hi ${customerName || req.user.name || 'Farmer'},</p>
                        <p>Thank you for your order. Here are your order details:</p>
                        <table style="width: 100%; margin: 16px 0; font-size: 0.95em;">
                            <tr><td style="padding: 4px 0; color: #666;">Order Number:</td><td style="font-weight: 600;">${order.orderNumber}</td></tr>
                            <tr><td style="padding: 4px 0; color: #666;">Payment Method:</td><td>${paymentLabel}</td></tr>
                            <tr><td style="padding: 4px 0; color: #666;">Pickup Location:</td><td>${pickupName}</td></tr>
                        </table>
                        <h3 style="color: #1a1a2e; margin-top: 24px;">Order Items</h3>
                        <table style="width: 100%; border-collapse: collapse;">
                            <thead>
                                <tr style="background: #f9fafb;">
                                    <th style="padding: 8px 12px; text-align: left; font-size: 0.85em; color: #666;">Product</th>
                                    <th style="padding: 8px 12px; text-align: center; font-size: 0.85em; color: #666;">Quantity</th>
                                    <th style="padding: 8px 12px; text-align: right; font-size: 0.85em; color: #666;">Total</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${itemsHtml}
                                <tr style="border-top: 2px solid #1a1a2e;">
                                    <td colspan="2" style="padding: 10px 12px; font-weight: 700;">Order Total</td>
                                    <td style="padding: 10px 12px; text-align: right; font-weight: 700; color: #2d5a27; font-size: 1.1em;">$${order.total.toFixed(2)}</td>
                                </tr>
                            </tbody>
                        </table>
                        <hr style="border: none; border-top: 1px solid #eee; margin: 24px 0;">
                        <p style="color: #888; font-size: 0.9em;">We'll notify you when your products are ready for pickup.<br>Questions? Contact your local representative.</p>
                        <p style="color: #888; font-size: 0.85em;">Thank you for choosing Acre Profit.</p>
                    </div>`
                });
                console.log(`Order confirmation email sent to ${toEmail} for order ${order.orderNumber}`);
            }
        } catch (emailError) {
            console.error('Failed to send order confirmation email:', emailError);
            // Don't fail the order if email fails
        }

        res.status(201).json({
            success: true,
            orderId: order._id,
            orderNumber: order.orderNumber,
            total: order.total,
            itemCount: order.items.length,
            paymentMethod: order.paymentMethod,
            paymentStatus: order.paymentStatus,
            location: location,
            items: order.items
        });

    } catch (error) {
        console.error('Checkout error:', error);
        res.status(400).json({ error: error.message });
    }
});

// Get Stripe publishable key
app.get('/api/stripe/config', (req, res) => {
    const publishableKey = process.env.STRIPE_PUBLISHABLE_KEY;
    if (!publishableKey) {
        return res.status(400).json({ error: 'Stripe not configured' });
    }
    res.json({ publishableKey });
});

// Get customer's chemical orders
app.get('/api/chemical-orders', authMiddleware, async (req, res) => {
    try {
        const orders = await ChemicalOrder.find({ userId: req.user._id })
            .sort({ createdAt: -1 });
        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single chemical order
app.get('/api/chemical-orders/:id', authMiddleware, async (req, res) => {
    try {
        const order = await ChemicalOrder.findById(req.params.id)
            .populate('items.chemicalId')
            .populate('userId', 'name email phone farm');

        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        // Check access: owner, admin, or assigned distributor
        const isOwner = order.userId?._id?.toString() === req.user._id.toString() ||
                        order.userId?.toString() === req.user._id.toString();
        const isAdmin = req.user.role === 'admin' || req.user.role === 'superadmin';
        const isAssignedDistributor = isDistributor(req.user) &&
                                      order.representativeId?.toString() === req.user._id.toString();

        if (!isOwner && !isAdmin && !isAssignedDistributor) {
            return res.status(403).json({ error: 'Not authorized to view this order' });
        }

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Submit chemical order
app.put('/api/chemical-orders/:id/submit', authMiddleware, async (req, res) => {
    try {
        const order = await ChemicalOrder.findOne({
            _id: req.params.id,
            userId: req.user._id
        });

        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        order.status = 'submitted';
        order.submittedAt = new Date();
        order.updatedAt = new Date();
        await order.save();

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update chemical order (admin/distributor)
app.put('/api/chemical-orders/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const order = await ChemicalOrder.findById(req.params.id);
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        // Check distributor access
        if (isDistributor(req.user) && order.representativeId?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Not authorized to edit this order' });
        }

        const { items, subtotal, discount, total, internalNotes, status } = req.body;

        // Track inventory changes if items are being modified
        if (items !== undefined && order.status !== 'cancelled' && order.status !== 'archived') {
            const oldItems = order.items || [];
            try {
                await updateOrderInventory({
                    oldItems,
                    newItems: items,
                    orderId: order._id,
                    orderNumber: order.orderNumber,
                    location: order.pickupLocation || 'main',
                    userId: req.user._id
                });
            } catch (invError) {
                console.error('Inventory update warning:', invError.message);
                // Continue with order update even if inventory tracking fails
            }
        }

        // Update fields if provided
        if (items !== undefined) order.items = items;
        if (subtotal !== undefined) order.subtotal = subtotal;
        if (discount !== undefined) order.discount = discount;
        if (total !== undefined) order.total = total;
        if (internalNotes !== undefined) order.internalNotes = internalNotes;
        if (status !== undefined) order.status = status;

        order.updatedAt = new Date();
        await order.save();

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Get all chemical orders
app.get('/api/admin/chemical-orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = {};
        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        }

        const orders = await ChemicalOrder.find(query)
            .populate('userId', 'name email phone farm')
            .sort({ createdAt: -1 });

        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Update chemical order status
app.put('/api/admin/chemical-orders/:id/status', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, internalNotes } = req.body;

        const order = await ChemicalOrder.findById(req.params.id);
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        const oldStatus = order.status;

        // Release inventory if order is being cancelled or archived
        if ((status === 'cancelled' || status === 'archived') &&
            oldStatus !== 'cancelled' && oldStatus !== 'archived') {
            try {
                // Release all reserved inventory for this order
                for (const item of order.items || []) {
                    if (item.chemicalId && item.quantity > 0) {
                        await releaseInventory({
                            chemicalId: item.chemicalId,
                            quantity: item.quantity,
                            location: order.pickupLocation || 'main',
                            orderId: order._id,
                            orderNumber: order.orderNumber,
                            userId: req.user._id,
                            notes: `Order ${status}: inventory released`
                        });
                    }
                }
            } catch (invError) {
                console.error('Inventory release warning:', invError.message);
            }
        }

        order.status = status;
        if (internalNotes) order.internalNotes = internalNotes;

        // Update timestamps based on status
        if (status === 'confirmed') order.confirmedAt = new Date();
        if (status === 'ordered_from_supplier') order.orderedFromSupplierAt = new Date();
        if (status === 'received') order.receivedAt = new Date();
        if (status === 'delivered') order.deliveredAt = new Date();

        order.updatedAt = new Date();
        await order.save();

        // I5: cancel any RupSaleRecords tied to this order when it's cancelled/archived
        if (status === 'cancelled' || status === 'archived') {
            await cancelRupRecordsForOrder(order);
        }

        // Auto-create ledger entry when order is delivered
        if (status === 'delivered' && order.representativeId) {
            const existingEntry = await LedgerEntry.findOne({
                referenceType: 'ChemicalOrder',
                referenceId: order._id,
                category: 'order'
            });

            if (!existingEntry) {
                let costTotal = 0;
                for (const item of order.items) {
                    if (item.chemicalId) {
                        const chemical = await Chemical.findById(item.chemicalId);
                        if (chemical) {
                            costTotal += item.quantity * chemical.costPrice * (chemical.unitsPerPack || 1);
                        }
                    } else {
                        costTotal += item.totalPrice || 0;
                    }
                }

                await createLedgerEntry({
                    representativeId: order.representativeId,
                    description: `Order ${order.orderNumber} delivered (${order.items.length} items)`,
                    amount: Math.round(costTotal * 100) / 100,
                    type: 'debit',
                    category: 'order',
                    referenceType: 'ChemicalOrder',
                    referenceId: order._id,
                    createdBy: req.user._id,
                    notes: `Auto-created on delivery. Order total: $${order.total}`
                });
            }
        }

        res.json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Archive chemical order (admin only) - soft delete for order issues
app.put('/api/chemical-orders/:id/archive', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const order = await ChemicalOrder.findById(req.params.id);
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        if (isDistributor(req.user) && order.representativeId?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Not authorized to archive this order' });
        }

        // Release inventory if not already cancelled/archived
        if (order.status !== 'cancelled' && order.status !== 'archived') {
            try {
                for (const item of order.items || []) {
                    if (item.chemicalId && item.quantity > 0) {
                        await releaseInventory({
                            chemicalId: item.chemicalId,
                            quantity: item.quantity,
                            location: order.pickupLocation || 'main',
                            orderId: order._id,
                            orderNumber: order.orderNumber,
                            userId: req.user._id,
                            notes: 'Order archived: inventory released'
                        });
                    }
                }
            } catch (invError) {
                console.error('Inventory release warning:', invError.message);
            }
        }

        order.status = 'archived';
        order.updatedAt = new Date();
        await order.save();

        // I5: cancel RupSaleRecords tied to this archived order
        await cancelRupRecordsForOrder(order);

        res.json({ message: 'Order archived successfully', order });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete chemical order permanently (admin only)
app.delete('/api/chemical-orders/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const order = await ChemicalOrder.findById(req.params.id);
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        if (isDistributor(req.user) && order.representativeId?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Not authorized to delete this order' });
        }

        // Only allow deletion if order is in draft, archived, or cancelled state
        if (!['draft', 'archived', 'cancelled'].includes(order.status)) {
            return res.status(400).json({ error: 'Can only delete orders that are draft, archived, or cancelled. Please archive the order first.' });
        }

        await ChemicalOrder.deleteOne({ _id: req.params.id });

        res.json({ message: 'Order deleted successfully' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- PRICE & CHEMICAL REQUEST ENDPOINTS ----

// Schema for price requests and chemical requests
const chemicalRequestSchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    requestType: {
        type: String,
        enum: ['price_match', 'unlisted_product'],
        required: true
    },
    // For price match requests
    productName: String,
    competitorSource: String,
    competitorPrice: Number,
    quantity: Number,
    unit: String,
    // For unlisted product requests
    requestedProduct: String,
    productDescription: String,
    estimatedQuantity: Number,
    // Common fields
    status: {
        type: String,
        enum: ['pending', 'reviewing', 'sourcing', 'quoted', 'fulfilled', 'declined'],
        default: 'pending'
    },
    adminNotes: String,
    supplierQuotes: [{
        supplierName: String,
        price: Number,
        notes: String,
        quotedAt: Date
    }],
    emailTemplate: String, // Generated email for forwarding to suppliers
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const ChemicalRequest = mongoose.model('ChemicalRequest', chemicalRequestSchema);

// Submit "Found it Cheaper" price match request
app.post('/api/chemical-orders/price-request', authMiddleware, async (req, res) => {
    try {
        const { productName, retailerSource, theirPrice, quantity, unit } = req.body;

        if (!productName || !retailerSource || !theirPrice) {
            return res.status(400).json({ error: 'Product name, source, and price are required' });
        }

        const request = new ChemicalRequest({
            userId: req.user._id,
            requestType: 'price_match',
            productName,
            competitorSource: retailerSource,
            competitorPrice: parseFloat(theirPrice),
            quantity: quantity || 1,
            unit: unit || 'unit',
            status: 'pending'
        });

        await request.save();

        res.status(201).json({
            message: 'Price match request submitted! We will pool orders and negotiate a better price.',
            requestId: request._id
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Request an unlisted chemical - generates email template
app.post('/api/chemical-orders/request-product', authMiddleware, async (req, res) => {
    try {
        const { productName, description, estimatedQuantity, unit } = req.body;

        if (!productName) {
            return res.status(400).json({ error: 'Product name is required' });
        }

        const user = await User.findById(req.user._id).populate('representative', 'name email');

        // Generate email template for suppliers
        const emailTemplate = `Subject: Request for Quote - ${productName}

Dear Supplier,

We are seeking competitive pricing on the following agricultural product:

PRODUCT DETAILS:
- Product: ${productName}
- Description: ${description || 'N/A'}
- Estimated Quantity: ${estimatedQuantity || 'TBD'} ${unit || 'units'}
- Delivery Location: Eastern Colorado

BUYER INFORMATION:
- Farm: ${user.farm?.name || user.name}
- Contact: ${user.name}
- Representative: ${user.representative?.name || 'Acre Profit'}

We are a group purchasing organization pooling orders from multiple farms to achieve volume pricing. Please provide:

1. Your best price per unit
2. Minimum order quantity
3. Availability/lead time
4. Freight terms (delivered vs. pickup)

Please reply to this email or contact us at:
- Email: ${user.representative?.email || 'orders@acreprofit.com'}
- Phone: ${user.phone || ''}

We look forward to your quote.

Best regards,
${user.name}
via Acre Profit Group Purchasing
`;

        const request = new ChemicalRequest({
            userId: req.user._id,
            requestType: 'unlisted_product',
            requestedProduct: productName,
            productDescription: description,
            estimatedQuantity: estimatedQuantity || 0,
            emailTemplate,
            status: 'pending'
        });

        await request.save();

        res.status(201).json({
            message: 'Product request created. Use the email template below to send to suppliers.',
            requestId: request._id,
            emailTemplate,
            instructions: 'Copy this email and send to your preferred suppliers. Forward any quotes to your rep or orders@acreprofit.com'
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get user's chemical requests
app.get('/api/chemical-orders/requests', authMiddleware, async (req, res) => {
    try {
        const requests = await ChemicalRequest.find({ userId: req.user._id })
            .sort({ createdAt: -1 });
        res.json(requests);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Get all chemical requests
app.get('/api/admin/chemical-requests', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const requests = await ChemicalRequest.find()
            .populate('userId', 'name email farm')
            .sort({ createdAt: -1 });
        res.json(requests);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Update chemical request status
app.put('/api/admin/chemical-requests/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, adminNotes, supplierQuote } = req.body;
        const request = await ChemicalRequest.findById(req.params.id);

        if (!request) {
            return res.status(404).json({ error: 'Request not found' });
        }

        if (status) request.status = status;
        if (adminNotes) request.adminNotes = adminNotes;
        if (supplierQuote) {
            request.supplierQuotes.push({
                ...supplierQuote,
                quotedAt: new Date()
            });
        }
        request.updatedAt = new Date();

        await request.save();
        res.json(request);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- USER SUBMITTED CHEMICALS ----
// Chemicals submitted by users to build the database

const userSubmittedChemicalSchema = new mongoose.Schema({
    // Submitted by
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    userName: String,

    // Chemical info
    brandName: { type: String, required: true },
    supplier: { type: String, required: true },
    uom: { type: String, required: true }, // Unit of measure (e.g., "gallon", "lb", "oz", "2.5 gal jug")
    packSize: String, // e.g., "Shuttle (250 gal)", "2x2.5 gal", "50 lb bag"
    costPerUnit: { type: Number, required: true }, // Cost per unit of measure

    // Additional info
    category: {
        type: String,
        enum: ['herbicide', 'fungicide', 'insecticide', 'adjuvant', 'fertilizer', 'seed_treatment', 'other'],
        default: 'herbicide'
    },
    crops: [String], // Which crops it can be used on
    notes: String, // Any additional info

    // Review status
    status: {
        type: String,
        enum: ['pending', 'approved', 'rejected'],
        default: 'pending'
    },
    reviewedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    reviewedAt: Date,
    reviewNotes: String,

    // If approved, link to the Chemical record created
    chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },

    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const UserSubmittedChemical = mongoose.model('UserSubmittedChemical', userSubmittedChemicalSchema);

// Password Reset Token Model
const passwordResetTokenSchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    token: { type: String, required: true },
    expiresAt: { type: Date, required: true },
    used: { type: Boolean, default: false },
    createdAt: { type: Date, default: Date.now }
});

const PasswordResetToken = mongoose.model('PasswordResetToken', passwordResetTokenSchema);

// Submit a new chemical to the database
app.post('/api/user-chemicals', authMiddleware, async (req, res) => {
    try {
        const { brandName, supplier, uom, packSize, costPerUnit, category, crops, notes } = req.body;

        if (!brandName || !supplier || !uom || !costPerUnit) {
            return res.status(400).json({
                error: 'Brand name, supplier, unit of measure, and cost are required'
            });
        }

        const user = await User.findById(req.user._id);

        const submission = new UserSubmittedChemical({
            userId: req.user._id,
            userName: user ? user.name : 'Unknown',
            brandName,
            supplier,
            uom,
            packSize: packSize || uom,
            costPerUnit: parseFloat(costPerUnit),
            category: category || 'herbicide',
            crops: crops || [],
            notes,
            status: 'pending'
        });

        await submission.save();

        res.status(201).json({
            message: 'Chemical submitted successfully! It is now available in the system.',
            chemical: submission
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get user's submitted chemicals
app.get('/api/user-chemicals', authMiddleware, async (req, res) => {
    try {
        const chemicals = await UserSubmittedChemical.find({ userId: req.user._id })
            .sort({ createdAt: -1 });
        res.json(chemicals);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all submitted chemicals (for use in ordering - shows all approved + user's own)
app.get('/api/user-chemicals/available', authMiddleware, async (req, res) => {
    try {
        // Get approved chemicals and user's own submissions
        const chemicals = await UserSubmittedChemical.find({
            $or: [
                { status: 'approved' },
                { userId: req.user._id }
            ]
        })
        .sort({ brandName: 1 });

        res.json(chemicals);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all user-submitted chemicals (admin view)
app.get('/api/admin/user-chemicals', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status } = req.query;
        let query = {};

        if (status) query.status = status;

        const chemicals = await UserSubmittedChemical.find(query)
            .populate('userId', 'name email farm')
            .populate('reviewedBy', 'name')
            .sort({ createdAt: -1 });

        res.json(chemicals);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Review/approve user-submitted chemical
app.put('/api/admin/user-chemicals/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, reviewNotes, addToDatabase } = req.body;
        const submission = await UserSubmittedChemical.findById(req.params.id);

        if (!submission) {
            return res.status(404).json({ error: 'Submission not found' });
        }

        submission.status = status;
        submission.reviewedBy = req.user._id;
        submission.reviewedAt = new Date();
        submission.reviewNotes = reviewNotes;
        submission.updatedAt = new Date();

        // If approved and addToDatabase is true, create a Chemical record
        if (status === 'approved' && addToDatabase) {
            const chemical = new Chemical({
                productName: submission.brandName,
                sourceSupplier: submission.supplier,
                packSize: submission.packSize,
                unit: submission.uom,
                costPrice: submission.costPerUnit,
                sellPrice: submission.costPerUnit, // Can be adjusted by admin
                category: submission.category,
                crops: submission.crops,
                notes: `User-submitted by ${submission.userName}. ${submission.notes || ''}`,
                isActive: true,
                availableForOrder: true,
                createdBy: req.user._id
            });

            await chemical.save();
            submission.chemicalId = chemical._id;
        }

        await submission.save();
        res.json(submission);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete user-submitted chemical (user can delete their own, admin can delete any)
app.delete('/api/user-chemicals/:id', authMiddleware, async (req, res) => {
    try {
        const submission = await UserSubmittedChemical.findById(req.params.id);

        if (!submission) {
            return res.status(404).json({ error: 'Submission not found' });
        }

        // Check if user owns this or is admin/distributor
        if (submission.userId.toString() !== req.user._id.toString() &&
            !isAdminLevel(req.user)) {
            return res.status(403).json({ error: 'Not authorized to delete this submission' });
        }

        await UserSubmittedChemical.findByIdAndDelete(req.params.id);
        res.json({ message: 'Chemical submission deleted' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- LEDGER ROUTES ----

// Get ledger entries
app.get('/api/admin/ledger', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { repId, startDate, endDate, category } = req.query;
        let query = {};

        if (req.user.role === 'superadmin') {
            if (repId) query.representativeId = repId;
        } else {
            query.representativeId = req.user._id;
        }

        if (category) query.category = category;
        if (startDate || endDate) {
            query.date = {};
            if (startDate) query.date.$gte = new Date(startDate);
            if (endDate) query.date.$lte = new Date(endDate);
        }

        const entries = await LedgerEntry.find(query)
            .populate('representativeId', 'name email')
            .populate('createdBy', 'name')
            .sort({ date: -1, createdAt: -1 });

        res.json(entries);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get balance summary per rep
app.get('/api/admin/ledger/summary', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let repFilter = {};

        if (isDistributor(req.user)) {
            repFilter = { _id: req.user._id };
        } else {
            repFilter = { role: { $in: ['admin', 'distributor', 'superadmin'] } };
        }

        const reps = await User.find(repFilter).select('name email role');

        const summaries = [];
        for (const rep of reps) {
            const lastEntry = await LedgerEntry.findOne({ representativeId: rep._id })
                .sort({ date: -1, createdAt: -1 });

            const totalDebits = await LedgerEntry.aggregate([
                { $match: { representativeId: rep._id, type: 'debit' } },
                { $group: { _id: null, total: { $sum: '$amount' } } }
            ]);

            const totalCredits = await LedgerEntry.aggregate([
                { $match: { representativeId: rep._id, type: 'credit' } },
                { $group: { _id: null, total: { $sum: '$amount' } } }
            ]);

            summaries.push({
                rep: { _id: rep._id, name: rep.name, email: rep.email },
                balance: lastEntry ? lastEntry.runningBalance : 0,
                totalDebits: totalDebits[0]?.total || 0,
                totalCredits: totalCredits[0]?.total || 0,
                entryCount: await LedgerEntry.countDocuments({ representativeId: rep._id })
            });
        }

        res.json(summaries);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create manual ledger entry (superadmin only for financial entries)
app.post('/api/admin/ledger', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { representativeId, description, amount, type, category, notes } = req.body;

        // Non-superadmin users cannot create ledger entries that change amounts
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can create ledger entries. Use the note endpoint to add notes.' });
        }

        if (!description || !amount || !type) {
            return res.status(400).json({ error: 'Description, amount, and type are required' });
        }

        let targetRepId = representativeId;
        if (isDistributor(req.user)) {
            targetRepId = req.user._id;
        }

        if (!targetRepId) {
            return res.status(400).json({ error: 'Representative ID is required' });
        }

        const entry = await createLedgerEntry({
            representativeId: targetRepId,
            description,
            amount: Math.abs(amount),
            type,
            category: category || 'adjustment',
            referenceType: 'Manual',
            createdBy: req.user._id,
            notes
        });

        const populated = await LedgerEntry.findById(entry._id)
            .populate('representativeId', 'name email')
            .populate('createdBy', 'name');

        res.status(201).json(populated);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update ledger entry note only (any admin/distributor)
app.put('/api/admin/ledger/:id/note', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { notes } = req.body;

        const entry = await LedgerEntry.findById(req.params.id);
        if (!entry) {
            return res.status(404).json({ error: 'Ledger entry not found' });
        }

        // Only update the notes field - no financial data changes allowed
        entry.notes = notes || '';
        await entry.save();

        const populated = await LedgerEntry.findById(entry._id)
            .populate('representativeId', 'name email')
            .populate('createdBy', 'name');

        res.json(populated);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update ledger entry fully (superadmin only) - can change amount, type, description, notes, category
app.put('/api/admin/ledger/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can edit ledger entries' });
        }

        const { description, amount, type, category, notes } = req.body;
        const entry = await LedgerEntry.findById(req.params.id);
        if (!entry) {
            return res.status(404).json({ error: 'Ledger entry not found' });
        }

        // Update fields
        if (description !== undefined) entry.description = description;
        if (amount !== undefined) entry.amount = Math.abs(amount);
        if (type !== undefined) entry.type = type;
        if (category !== undefined) entry.category = category;
        if (notes !== undefined) entry.notes = notes;

        await entry.save();

        // Recalculate running balances for all entries for this rep from this entry forward
        const allEntries = await LedgerEntry.find({ representativeId: entry.representativeId })
            .sort({ date: 1, createdAt: 1 });

        let runningBalance = 0;
        for (const e of allEntries) {
            const balanceChange = e.type === 'debit' ? e.amount : -e.amount;
            runningBalance = Math.round((runningBalance + balanceChange) * 100) / 100;
            if (e.runningBalance !== runningBalance) {
                e.runningBalance = runningBalance;
                await e.save();
            }
        }

        const populated = await LedgerEntry.findById(entry._id)
            .populate('representativeId', 'name email')
            .populate('createdBy', 'name');

        res.json(populated);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ PURCHASE ORDER ENDPOINTS ============

// Get all purchase orders
app.get('/api/admin/purchase-orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, supplier, paymentStatus } = req.query;
        let query = {};

        if (status) query.status = status;
        if (supplier) query['supplier.name'] = new RegExp(supplier, 'i');
        if (paymentStatus) query.paymentStatus = paymentStatus;

        const purchaseOrders = await PurchaseOrder.find(query)
            .populate('createdBy', 'name email')
            .populate('updatedBy', 'name email')
            .populate('paidBy', 'name email')
            .sort({ orderDate: -1 });

        // For each PO, calculate current distribution (who's holding what)
        const enriched = await Promise.all(purchaseOrders.map(async (po) => {
            const poObj = po.toObject();
            const itemsWithDistribution = await Promise.all((poObj.items || []).map(async (item) => {
                if (!item.chemicalId) return { ...item, distribution: [] };

                // Find all inventory locations for this chemical
                const invRecords = await Inventory.find({ chemicalId: item.chemicalId })
                    .populate('distributorId', 'name');

                const distribution = invRecords
                    .filter(inv => (inv.quantityOnHand || 0) > 0)
                    .map(inv => ({
                        distributor: inv.distributorId?.name || 'Unassigned',
                        location: inv.location,
                        quantity: inv.quantityOnHand,
                        value: Math.round((inv.quantityOnHand * item.pricePerUnit) * 100) / 100
                    }));

                return { ...item, distribution };
            }));
            poObj.items = itemsWithDistribution;
            return poObj;
        }));

        res.json(enriched);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single purchase order with splits
app.get('/api/admin/purchase-orders/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const po = await PurchaseOrder.findById(req.params.id)
            .populate('createdBy', 'name email')
            .populate('updatedBy', 'name email');

        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        // Get all splits for this PO
        const splits = await PurchaseOrderSplit.find({ purchaseOrderId: po._id })
            .populate('distributorId', 'name email')
            .sort({ splitCode: 1 });

        res.json({ purchaseOrder: po, splits });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create purchase order
app.post('/api/admin/purchase-orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { supplier, items, freight, otherFees, superAdminFee, expectedDeliveryDate, deliveryLocation, notes } = req.body;

        // Generate PO number
        const poNumber = await generatePONumber();

        // Calculate item totals and remaining quantities
        const processedItems = items.map(item => ({
            ...item,
            totalPrice: Math.round((item.quantityOrdered * item.pricePerUnit) * 100) / 100,
            quantityAllocated: 0,
            quantityRemaining: item.quantityOrdered
        }));

        // Calculate totals (superAdminFee only if user is superadmin)
        const subtotal = processedItems.reduce((sum, item) => sum + item.totalPrice, 0);
        const adminFee = req.user.role === 'superadmin' ? (superAdminFee || 0) : 0;
        const totalCost = subtotal + (freight || 0) + (otherFees || 0) + adminFee;

        const purchaseOrder = new PurchaseOrder({
            poNumber,
            supplier,
            items: processedItems,
            subtotal: Math.round(subtotal * 100) / 100,
            freight: freight || 0,
            otherFees: otherFees || 0,
            superAdminFee: adminFee,
            totalCost: Math.round(totalCost * 100) / 100,
            expectedDeliveryDate,
            deliveryLocation,
            notes,
            createdBy: req.user._id,
            updatedBy: req.user._id
        });

        await purchaseOrder.save();

        const populated = await PurchaseOrder.findById(purchaseOrder._id)
            .populate('createdBy', 'name email');

        res.status(201).json(populated);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update purchase order (general info and status)
app.put('/api/admin/purchase-orders/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { supplier, expectedDeliveryDate, deliveryLocation, bolNumber, trackingInfo, notes, internalNotes, status, superAdminFee } = req.body;

        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        // Update allowed fields
        if (supplier) po.supplier = supplier;
        if (expectedDeliveryDate !== undefined) po.expectedDeliveryDate = expectedDeliveryDate;
        if (deliveryLocation !== undefined) po.deliveryLocation = deliveryLocation;
        if (bolNumber !== undefined) po.bolNumber = bolNumber;
        if (trackingInfo !== undefined) po.trackingInfo = trackingInfo;
        if (notes !== undefined) po.notes = notes;
        if (internalNotes !== undefined) po.internalNotes = internalNotes;
        if (status) {
            po.status = status;
            if (status === 'received') po.receivedDate = new Date();
        }

        // Super admin fee (only superadmin can set)
        if (superAdminFee !== undefined && req.user.role === 'superadmin') {
            po.superAdminFee = superAdminFee;
            // Recalculate total
            po.totalCost = (po.subtotal || 0) + (po.freight || 0) + (po.otherFees || 0) + superAdminFee;
        }

        po.updatedBy = req.user._id;
        po.updatedAt = new Date();

        await po.save();

        const populated = await PurchaseOrder.findById(po._id)
            .populate('createdBy', 'name email')
            .populate('updatedBy', 'name email');

        res.json(populated);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update purchase order item price (key feature!)
app.put('/api/admin/purchase-orders/:id/items/:itemIndex/price', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { pricePerUnit } = req.body;
        const itemIndex = parseInt(req.params.itemIndex, 10);

        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        if (itemIndex < 0 || itemIndex >= po.items.length) {
            return res.status(400).json({ error: 'Invalid item index' });
        }

        // Update the price
        po.items[itemIndex].pricePerUnit = pricePerUnit;
        po.items[itemIndex].totalPrice = Math.round((po.items[itemIndex].quantityOrdered * pricePerUnit) * 100) / 100;

        // Recalculate totals
        const subtotal = po.items.reduce((sum, item) => sum + item.totalPrice, 0);
        po.subtotal = Math.round(subtotal * 100) / 100;
        po.totalCost = Math.round((subtotal + po.freight + po.otherFees) * 100) / 100;

        po.updatedBy = req.user._id;
        po.updatedAt = new Date();

        await po.save();

        res.json({ message: 'Price updated', purchaseOrder: po });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update purchase order item quantity
app.put('/api/admin/purchase-orders/:id/items/:itemIndex/quantity', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { quantityOrdered } = req.body;
        const itemIndex = parseInt(req.params.itemIndex, 10);

        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        if (itemIndex < 0 || itemIndex >= po.items.length) {
            return res.status(400).json({ error: 'Invalid item index' });
        }

        const item = po.items[itemIndex];

        // Ensure new quantity >= allocated quantity
        if (quantityOrdered < item.quantityAllocated) {
            return res.status(400).json({
                error: `Cannot reduce below allocated quantity (${item.quantityAllocated})`
            });
        }

        // Update the quantity
        item.quantityOrdered = quantityOrdered;
        item.totalPrice = Math.round((quantityOrdered * item.pricePerUnit) * 100) / 100;
        item.quantityRemaining = quantityOrdered - item.quantityAllocated;

        // Recalculate totals
        const subtotal = po.items.reduce((sum, i) => sum + i.totalPrice, 0);
        po.subtotal = Math.round(subtotal * 100) / 100;
        po.totalCost = Math.round((subtotal + po.freight + po.otherFees) * 100) / 100;

        po.updatedBy = req.user._id;
        po.updatedAt = new Date();

        await po.save();

        res.json({ message: 'Quantity updated', purchaseOrder: po });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Add item to purchase order
app.post('/api/admin/purchase-orders/:id/items', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { productName, chemicalId, description, packSize, unit, quantityOrdered, pricePerUnit } = req.body;

        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        const newItem = {
            productName,
            chemicalId,
            description,
            packSize,
            unit,
            quantityOrdered,
            pricePerUnit,
            totalPrice: Math.round((quantityOrdered * pricePerUnit) * 100) / 100,
            quantityAllocated: 0,
            quantityRemaining: quantityOrdered
        };

        po.items.push(newItem);

        // Recalculate totals
        const subtotal = po.items.reduce((sum, item) => sum + item.totalPrice, 0);
        po.subtotal = Math.round(subtotal * 100) / 100;
        po.totalCost = Math.round((subtotal + po.freight + po.otherFees) * 100) / 100;

        po.updatedBy = req.user._id;
        po.updatedAt = new Date();

        await po.save();

        res.json({ message: 'Item added', purchaseOrder: po });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete item from purchase order
app.delete('/api/admin/purchase-orders/:id/items/:itemIndex', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const itemIndex = parseInt(req.params.itemIndex, 10);

        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        if (itemIndex < 0 || itemIndex >= po.items.length) {
            return res.status(400).json({ error: 'Invalid item index' });
        }

        // Check if item has allocations
        if (po.items[itemIndex].quantityAllocated > 0) {
            return res.status(400).json({
                error: 'Cannot delete item with allocations. Remove splits first.'
            });
        }

        po.items.splice(itemIndex, 1);

        // Recalculate totals
        const subtotal = po.items.reduce((sum, item) => sum + item.totalPrice, 0);
        po.subtotal = Math.round(subtotal * 100) / 100;
        po.totalCost = Math.round((subtotal + po.freight + po.otherFees) * 100) / 100;

        po.updatedBy = req.user._id;
        po.updatedAt = new Date();

        await po.save();

        res.json({ message: 'Item deleted', purchaseOrder: po });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ PURCHASE ORDER SPLIT ENDPOINTS ============

// Get distributors for split dropdown
app.get('/api/admin/distributors', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const distributors = await User.find({
            role: { $in: ['admin', 'distributor'] }
        }).select('name email role').sort({ name: 1 });

        // Enrich with inventory value, order count, customer count
        const enriched = await Promise.all(distributors.map(async (d) => {
            const dObj = d.toObject();

            // Inventory value: sum of (quantityOnHand * averageCost) for this distributor
            const inventory = await Inventory.find({ distributorId: d._id });
            dObj.inventoryValue = inventory.reduce((sum, inv) => {
                return sum + ((inv.quantityOnHand || 0) * (inv.averageCost || inv.lastCost || 0));
            }, 0);
            dObj.inventoryItems = inventory.filter(i => (i.quantityOnHand || 0) > 0).length;

            // Orders: count orders where representativeId = this distributor
            const orderCount = await Order.countDocuments({ representativeId: d._id });
            const chemOrderCount = await ChemicalOrder.countDocuments({ representativeId: d._id });
            dObj.orderCount = orderCount + chemOrderCount;

            // Customers: count users with this distributor as representative
            dObj.customerCount = await User.countDocuments({
                role: 'customer',
                $or: [
                    { representative: d._id },
                    { representativeId: d._id }
                ]
            });

            return dObj;
        }));

        res.json(enriched);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create a split (allocate portion of PO to a distributor)
app.post('/api/admin/purchase-orders/:id/splits', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { distributorId, items, freightAllocation, notes } = req.body;

        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        const distributor = await User.findById(distributorId);
        if (!distributor) {
            return res.status(404).json({ error: 'Distributor not found' });
        }

        // Validate allocations don't exceed available quantities
        for (const splitItem of items) {
            const poItem = po.items[splitItem.originalItemIndex];
            if (!poItem) {
                return res.status(400).json({ error: `Invalid item index: ${splitItem.originalItemIndex}` });
            }

            const availableQty = poItem.quantityOrdered - poItem.quantityAllocated;
            if (splitItem.quantityAllocated > availableQty) {
                return res.status(400).json({
                    error: `Cannot allocate ${splitItem.quantityAllocated} of ${poItem.productName}. Only ${availableQty} available.`
                });
            }
        }

        // Generate split code (A, B, C, ...)
        const existingSplits = await PurchaseOrderSplit.countDocuments({ purchaseOrderId: po._id });
        const splitCode = String.fromCharCode(65 + existingSplits); // A=65, B=66, etc.

        // Process split items
        const processedItems = items.map(item => {
            const poItem = po.items[item.originalItemIndex];
            return {
                productName: poItem.productName,
                chemicalId: poItem.chemicalId,
                packSize: poItem.packSize,
                unit: poItem.unit,
                quantityAllocated: item.quantityAllocated,
                pricePerUnit: item.pricePerUnit || poItem.pricePerUnit,
                totalPrice: Math.round((item.quantityAllocated * (item.pricePerUnit || poItem.pricePerUnit)) * 100) / 100,
                originalItemIndex: item.originalItemIndex
            };
        });

        const subtotal = processedItems.reduce((sum, item) => sum + item.totalPrice, 0);

        const split = new PurchaseOrderSplit({
            purchaseOrderId: po._id,
            poNumber: po.poNumber,
            distributorId,
            distributorName: distributor.name,
            splitCode,
            items: processedItems,
            subtotal: Math.round(subtotal * 100) / 100,
            freightAllocation: freightAllocation || 0,
            totalCost: Math.round((subtotal + (freightAllocation || 0)) * 100) / 100,
            notes,
            createdBy: req.user._id
        });

        await split.save();

        // Update PO item allocated quantities
        for (const splitItem of items) {
            po.items[splitItem.originalItemIndex].quantityAllocated += splitItem.quantityAllocated;
            po.items[splitItem.originalItemIndex].quantityRemaining =
                po.items[splitItem.originalItemIndex].quantityOrdered -
                po.items[splitItem.originalItemIndex].quantityAllocated;
        }

        po.updatedBy = req.user._id;
        po.updatedAt = new Date();
        await po.save();

        const populated = await PurchaseOrderSplit.findById(split._id)
            .populate('distributorId', 'name email');

        res.status(201).json({ split: populated, purchaseOrder: po });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update split status
app.put('/api/admin/purchase-order-splits/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, deliveryDate, deliveryLocation, receivedBy, notes } = req.body;

        const split = await PurchaseOrderSplit.findById(req.params.id);
        if (!split) {
            return res.status(404).json({ error: 'Split not found' });
        }

        if (status) split.status = status;
        if (deliveryDate !== undefined) split.deliveryDate = deliveryDate;
        if (deliveryLocation !== undefined) split.deliveryLocation = deliveryLocation;
        if (receivedBy !== undefined) split.receivedBy = receivedBy;
        if (notes !== undefined) split.notes = notes;

        split.updatedAt = new Date();
        await split.save();

        res.json(split);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete split (unallocate items back to PO)
app.delete('/api/admin/purchase-order-splits/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const split = await PurchaseOrderSplit.findById(req.params.id);
        if (!split) {
            return res.status(404).json({ error: 'Split not found' });
        }

        // Don't allow deletion if already shipped/delivered
        if (['shipped', 'delivered', 'invoiced', 'paid'].includes(split.status)) {
            return res.status(400).json({
                error: 'Cannot delete split that has been shipped, delivered, invoiced, or paid'
            });
        }

        // Return quantities to PO
        const po = await PurchaseOrder.findById(split.purchaseOrderId);
        if (po) {
            for (const splitItem of split.items) {
                if (splitItem.originalItemIndex !== undefined && po.items[splitItem.originalItemIndex]) {
                    po.items[splitItem.originalItemIndex].quantityAllocated -= splitItem.quantityAllocated;
                    po.items[splitItem.originalItemIndex].quantityRemaining =
                        po.items[splitItem.originalItemIndex].quantityOrdered -
                        po.items[splitItem.originalItemIndex].quantityAllocated;
                }
            }
            po.updatedAt = new Date();
            await po.save();
        }

        await PurchaseOrderSplit.findByIdAndDelete(split._id);

        res.json({ message: 'Split deleted and quantities returned to purchase order' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all splits for a distributor
app.get('/api/admin/distributor-splits/:distributorId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const splits = await PurchaseOrderSplit.find({ distributorId: req.params.distributorId })
            .populate('purchaseOrderId', 'poNumber supplier status')
            .sort({ createdAt: -1 });

        res.json(splits);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- PRODUCTS ROUTE (with units sold) ----

app.get('/api/admin/products', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { search, category } = req.query;
        let query = {};

        if (search) {
            const searchRegex = new RegExp(search, 'i');
            query.$or = [
                { productName: searchRegex },
                { sourceSupplier: searchRegex }
            ];
        }
        if (category) query.category = category;

        const chemicals = await Chemical.find(query)
            .populate('createdBy', 'name email')
            .sort({ productName: 1, packSize: 1 });

        // Aggregate units sold per chemical from non-cancelled orders
        const salesData = await ChemicalOrder.aggregate([
            { $match: { status: { $nin: ['cancelled', 'draft'] } } },
            { $unwind: '$items' },
            { $group: {
                _id: '$items.chemicalId',
                unitsSold: { $sum: '$items.quantity' },
                totalRevenue: { $sum: '$items.totalPrice' }
            }}
        ]);

        const salesMap = {};
        for (const s of salesData) {
            if (s._id) salesMap[s._id.toString()] = { unitsSold: s.unitsSold, totalRevenue: s.totalRevenue };
        }

        const products = chemicals.map(c => ({
            _id: c._id,
            productName: c.productName,
            sourceSupplier: c.sourceSupplier,
            category: c.category,
            packSize: c.packSize,
            unit: c.unit,
            unitsPerPack: c.unitsPerPack,
            costPrice: c.costPrice,
            adminPrice: c.adminPrice,
            adminMargin: c.adminMargin,
            sellPrice: c.sellPrice,
            margin: c.margin,
            activeIngredients: c.activeIngredients,
            isActive: c.isActive,
            availableForOrder: c.availableForOrder,
            notes: c.notes,
            createdBy: c.createdBy,
            createdAt: c.createdAt,
            unitsSold: salesMap[c._id.toString()]?.unitsSold || 0,
            totalRevenue: salesMap[c._id.toString()]?.totalRevenue || 0
        }));

        res.json(products);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- SPRAY PROGRAM CALCULATOR ROUTES ----

// Helper: Convert hardcoded sprayPrograms to program format
const convertHardcodedToPrograms = () => {
    const programs = [];
    for (const [cropKey, cropPrograms] of Object.entries(sprayPrograms)) {
        for (const [programKey, programData] of Object.entries(cropPrograms)) {
            // Build chemicals list for restriction generation
            const chemList = programData.chemicals.map(chem => ({
                productName: chem.name,
                suggestedRate: chem.defaultRate,
                rateUnit: chem.rateUnit,
                packSize: chem.packageSize,
                unit: chem.packageUnit,
                isAdjuvant: chem.isAdjuvant || false
            }));

            // Auto-generate restrictions from the chemicals in this template
            const generated = generateRecipeRestrictions(chemList);

            programs.push({
                _id: `template-${cropKey}-${programKey}`,
                name: programData.name,
                description: programData.description,
                crop: cropKey,
                type: 'template',
                isPublic: true,
                isTemplate: true,
                groundType: generated.groundType,
                rotationRestrictions: generated.rotationRestrictions,
                grazingRestrictions: generated.grazingRestrictions,
                applications: [{
                    name: programData.name,
                    chemicals: chemList
                }]
            });
        }
    }
    return programs;
};

// GET all programs — merge hardcoded templates + MongoDB custom/suggestion programs
// Auth: all logged-in roles
app.get('/api/spray-programs', authMiddleware, async (req, res) => {
    try {
        const { crop, type } = req.query;
        const userId = req.user._id;
        const userRole = req.user.role;

        // 1. Get hardcoded templates
        let templatePrograms = convertHardcodedToPrograms();
        if (crop) {
            templatePrograms = templatePrograms.filter(p => p.crop === crop);
        }

        // 2. Build MongoDB query
        let dbQuery = {};

        if (isAdminLevel(req.user)) {
            // Admins see all MongoDB programs
            if (crop) dbQuery.crop = crop;
            if (type && type !== 'template') dbQuery.type = type;
        } else {
            // Regular users see: public programs OR their own programs
            dbQuery.$or = [
                { isPublic: true },
                { createdBy: userId }
            ];
            if (crop) dbQuery.crop = crop;
            if (type && type !== 'template') dbQuery.type = type;
        }

        const dbPrograms = await SprayProgram.find(dbQuery)
            .populate('createdBy', 'name email')
            .sort({ crop: 1, name: 1 });

        // 3. Merge and return
        // Filter templates if specific type requested
        let allPrograms = [];
        if (!type || type === 'template') {
            allPrograms = [...templatePrograms];
        }
        if (!type || type !== 'template') {
            allPrograms = [...allPrograms, ...dbPrograms.map(p => p.toObject())];
        }

        res.json({
            programs: allPrograms,
            counts: {
                templates: templatePrograms.length,
                database: dbPrograms.length,
                total: allPrograms.length
            }
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get user's custom programs only
app.get('/api/spray-programs/my', authMiddleware, async (req, res) => {
    try {
        const programs = await SprayProgram.find({
            createdBy: req.user._id
        }).sort({ createdAt: -1 });

        res.json(programs);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single program with chemical details
app.get('/api/spray-programs/:id', authMiddleware, async (req, res) => {
    try {
        const programId = req.params.id;

        // Check if it's a hardcoded template
        if (programId.startsWith('template-')) {
            const templates = convertHardcodedToPrograms();
            const template = templates.find(t => t._id === programId);
            if (template) {
                return res.json(template);
            }
            return res.status(404).json({ error: 'Template not found' });
        }

        const program = await SprayProgram.findById(programId)
            .populate('createdBy', 'name email');

        if (!program) {
            return res.status(404).json({ error: 'Program not found' });
        }

        // Check access: public, owner, or admin
        if (!program.isPublic &&
            program.createdBy._id.toString() !== req.user._id.toString() &&
            !isAdminLevel(req.user)) {
            return res.status(403).json({ error: 'Access denied' });
        }

        res.json(program);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// POST save custom program
// Auth: all logged-in roles
app.post('/api/spray-programs', authMiddleware, async (req, res) => {
    try {
        const { name, description, crop, applications, isPublic, groundType, rotationRestrictions, grazingRestrictions } = req.body;

        // Validate required fields
        if (!name || !crop) {
            return res.status(400).json({ error: 'Name and crop are required' });
        }

        // Determine program type and visibility
        const isAdmin = isAdminLevel(req.user);
        const programType = isAdmin ? 'suggestion' : 'custom';
        // Only admins can create public programs
        const publicFlag = isAdmin ? (isPublic || false) : false;

        // Auto-generate restrictions if not provided
        let autoGroundType = groundType || null;
        let autoRotation = rotationRestrictions || null;
        let autoGrazing = grazingRestrictions || null;

        if ((!autoGroundType || !autoRotation || !autoGrazing) && applications && applications.length > 0) {
            // Collect all chemicals from all applications
            const allChemicals = [];
            for (const app of applications) {
                for (const chem of app.chemicals || []) {
                    allChemicals.push(chem);
                }
            }
            const generated = generateRecipeRestrictions(allChemicals);
            if (!autoGroundType) autoGroundType = generated.groundType;
            if (!autoRotation) autoRotation = generated.rotationRestrictions;
            if (!autoGrazing) autoGrazing = generated.grazingRestrictions;
        }

        const program = new SprayProgram({
            name,
            description,
            crop,
            applications: applications || [],
            type: programType,
            isPublic: publicFlag,
            createdBy: req.user._id,
            groundType: autoGroundType,
            rotationRestrictions: autoRotation,
            grazingRestrictions: autoGrazing
        });

        // Calculate estimated cost per acre if applications provided
        if (applications && applications.length > 0) {
            let totalCost = 0;
            for (const app of applications) {
                for (const chem of app.chemicals || []) {
                    if (chem.chemicalId) {
                        const chemical = await Chemical.findById(chem.chemicalId);
                        if (chemical && chemical.sellPrice && chem.suggestedRate) {
                            // Convert rate to gallons and multiply by price
                            totalCost += (chem.suggestedRate / 128) * chemical.sellPrice;
                        }
                    }
                }
            }
            program.estimatedCostPerAcre = Math.round(totalCost * 100) / 100;
        }

        await program.save();
        res.status(201).json(program);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// POST calculate order from program or custom product list
// Auth: all logged-in roles
app.post('/api/spray-programs/calculate', authMiddleware, async (req, res) => {
    try {
        const { acres, gpa = 10, products } = req.body;

        // Validate inputs
        if (!acres || acres <= 0) {
            return res.status(400).json({ error: 'Acres is required and must be greater than 0' });
        }
        if (!products || !Array.isArray(products) || products.length === 0) {
            return res.status(400).json({ error: 'Products array is required' });
        }

        const orderLines = [];
        let totalConfirmedPrice = 0;
        let hasNeedsQuote = false;
        let valorWarning = false;
        const totalWaterVolume = acres * gpa;

        // STEP 1 & 2: Process each product
        for (const prod of products) {
            const { chemicalId, rate, unit } = prod;

            if (!chemicalId || !rate) {
                continue; // Skip invalid entries
            }

            // Fetch chemical from database
            const chemical = await Chemical.findById(chemicalId);
            if (!chemical) {
                orderLines.push({
                    chemicalId,
                    productName: 'Unknown Product',
                    error: 'Product not found in database',
                    status: 'error'
                });
                continue;
            }

            // Check for Valor warning
            if (chemical.productName && chemical.productName.toLowerCase().includes('valor')) {
                valorWarning = true;
            }

            // Calculate total needed based on rate unit
            let totalNeeded = 0;
            const rateUnit = unit || chemical.rateUnit || 'oz/acre';

            switch (rateUnit) {
                case 'oz/acre':
                    totalNeeded = (rate * acres) / 128; // oz to gallons
                    break;
                case 'pt/acre':
                    totalNeeded = (rate * acres) / 8; // pints to gallons
                    break;
                case 'qt/acre':
                    totalNeeded = (rate * acres) / 4; // quarts to gallons
                    break;
                case 'gal/acre':
                    totalNeeded = rate * acres;
                    break;
                case 'lb/acre':
                    totalNeeded = rate * acres; // lbs
                    break;
                case '% v/v':
                    // Percentage of total water volume
                    totalNeeded = totalWaterVolume * (rate / 100);
                    break;
                default:
                    totalNeeded = rate * acres;
            }

            // Get package size (unitsPerPack is gallons/units per package)
            const packageSize = chemical.unitsPerPack || 1;

            // ALWAYS round UP to nearest package
            const packagesNeeded = Math.ceil(totalNeeded / packageSize);

            // Check inventory
            const inventory = await Inventory.findOne({
                chemicalId: chemical._id,
                location: 'main' // Default location
            });
            const onHandQuantity = inventory ? inventory.quantityAvailable : 0;

            // Determine status
            let status, pricePerPackage, lineTotal;
            if (onHandQuantity >= packagesNeeded) {
                status = 'confirmed';
                pricePerPackage = chemical.sellPrice * packageSize;
                lineTotal = packagesNeeded * pricePerPackage;
                totalConfirmedPrice += lineTotal;
            } else {
                status = 'needs_quote';
                pricePerPackage = null;
                lineTotal = null;
                hasNeedsQuote = true;
            }

            orderLines.push({
                chemicalId: chemical._id,
                productName: chemical.productName,
                category: chemical.category,
                rate,
                unit: rateUnit,
                totalNeeded: Math.round(totalNeeded * 1000) / 1000,
                packageSize,
                packSize: chemical.packSize,
                packageUnit: chemical.unit,
                packagesNeeded,
                onHandQuantity,
                pricePerPackage: pricePerPackage ? Math.round(pricePerPackage * 100) / 100 : null,
                lineTotal: lineTotal ? Math.round(lineTotal * 100) / 100 : null,
                status,
                supplier: chemical.sourceSupplier
            });
        }

        // STEP 3: Hydrovant auto-calculation (LOCKED — cannot be passed in by client)
        let hydrovantLine = null;
        const hydrovantGallonsNeeded = totalWaterVolume * 0.001; // 0.1% of total water volume

        if (hydrovantGallonsNeeded > 0) {
            const hydrovant = await Chemical.findOne({
                productName: { $regex: /hydrovant/i }
            });

            if (hydrovant) {
                const hvPackageSize = hydrovant.unitsPerPack || 2.5; // Default 2.5 gal
                const hvPackagesNeeded = Math.ceil(hydrovantGallonsNeeded / hvPackageSize);

                const hvInventory = await Inventory.findOne({
                    chemicalId: hydrovant._id,
                    location: 'main'
                });
                const hvOnHand = hvInventory ? hvInventory.quantityAvailable : 0;

                const hvStatus = hvOnHand >= hvPackagesNeeded ? 'confirmed' : 'needs_quote';
                const hvPricePerPack = hvStatus === 'confirmed' ? (hydrovant.sellPrice * hvPackageSize) : null;
                const hvLineTotal = hvStatus === 'confirmed' ? (hvPackagesNeeded * hvPricePerPack) : null;

                if (hvStatus === 'confirmed' && hvLineTotal) {
                    totalConfirmedPrice += hvLineTotal;
                } else {
                    hasNeedsQuote = true;
                }

                hydrovantLine = {
                    chemicalId: hydrovant._id,
                    productName: 'Hydrovant (Auto-Added)',
                    category: 'adjuvant',
                    gallonsNeeded: Math.round(hydrovantGallonsNeeded * 1000) / 1000,
                    packageSize: hvPackageSize,
                    packSize: hydrovant.packSize,
                    packageUnit: hydrovant.unit,
                    packagesNeeded: hvPackagesNeeded,
                    onHandQuantity: hvOnHand,
                    pricePerPackage: hvPricePerPack ? Math.round(hvPricePerPack * 100) / 100 : null,
                    lineTotal: hvLineTotal ? Math.round(hvLineTotal * 100) / 100 : null,
                    status: hvStatus,
                    isAutoAdded: true,
                    note: '0.1% of total spray volume'
                };
            }
        }

        // STEP 5: Build response
        const orderStatus = hasNeedsQuote ? 'pending_quote' : 'ready_for_checkout';

        res.json({
            acres,
            gpa,
            totalWaterVolume,
            orderLines,
            hydrovant: hydrovantLine,
            valorWarning: valorWarning ? 'Valor requires application 7–30 days preplant. Minimum 1/4 inch rainfall required before planting.' : null,
            orderStatus,
            totalConfirmedPrice: Math.round(totalConfirmedPrice * 100) / 100,
            costPerAcre: acres > 0 ? Math.round((totalConfirmedPrice / acres) * 100) / 100 : 0
        });
    } catch (error) {
        console.error('Calculate error:', error);
        res.status(400).json({ error: error.message });
    }
});

// POST submit order or quote request
// Auth: all logged-in roles
app.post('/api/spray-programs/submit-order', authMiddleware, async (req, res) => {
    try {
        const {
            acres, gpa, orderLines, hydrovant, orderStatus,
            totalConfirmedPrice, crop, programName
        } = req.body;

        if (!acres || !orderLines || orderLines.length === 0) {
            return res.status(400).json({ error: 'Invalid order data' });
        }

        // Build chemicals array for order
        const chemicals = orderLines.map(line => ({
            name: line.productName,
            chemicalId: line.chemicalId,
            rate: line.rate,
            rateUnit: line.unit,
            totalAmount: line.totalNeeded,
            totalUnit: 'gal',
            packageSize: line.packageSize,
            packageUnit: line.packageUnit,
            packagesNeeded: line.packagesNeeded,
            pricePerPackage: line.pricePerPackage,
            totalPrice: line.lineTotal,
            status: line.status
        }));

        // Add Hydrovant if present
        if (hydrovant) {
            chemicals.push({
                name: hydrovant.productName,
                chemicalId: hydrovant.chemicalId,
                rate: 0.1,
                rateUnit: '% v/v',
                totalAmount: hydrovant.gallonsNeeded,
                totalUnit: 'gal',
                packageSize: hydrovant.packageSize,
                packageUnit: hydrovant.packageUnit,
                packagesNeeded: hydrovant.packagesNeeded,
                pricePerPackage: hydrovant.pricePerPackage,
                totalPrice: hydrovant.lineTotal,
                status: hydrovant.status,
                isAutoAdded: true
            });
        }

        if (orderStatus === 'ready_for_checkout') {
            // All items confirmed — proceed to Stripe checkout
            if (!stripe) {
                return res.status(500).json({ error: 'Payment processing not configured' });
            }

            // Create or get Stripe customer
            let customerId = req.user.stripeCustomerId;
            if (!customerId) {
                const customer = await stripe.customers.create({
                    email: req.user.email,
                    name: req.user.name,
                    metadata: { userId: req.user._id.toString() }
                });
                customerId = customer.id;
                await User.findByIdAndUpdate(req.user._id, { stripeCustomerId: customerId });
            }

            // Create line items for Stripe
            const lineItems = chemicals
                .filter(c => c.status === 'confirmed' && c.totalPrice)
                .map(c => ({
                    price_data: {
                        currency: 'usd',
                        product_data: {
                            name: c.name,
                            description: `${c.packagesNeeded} x ${c.packageSize} ${c.packageUnit}`
                        },
                        unit_amount: Math.round(c.totalPrice * 100) // Stripe uses cents
                    },
                    quantity: 1
                }));

            // Create Stripe checkout session
            const session = await stripe.checkout.sessions.create({
                customer: customerId,
                payment_method_types: ['card'],
                line_items: lineItems,
                mode: 'payment',
                success_url: `${process.env.FRONTEND_URL || 'https://acreprofit.com'}/order-success?session_id={CHECKOUT_SESSION_ID}`,
                cancel_url: `${process.env.FRONTEND_URL || 'https://acreprofit.com'}/calculator`,
                metadata: {
                    userId: req.user._id.toString(),
                    acres: acres.toString(),
                    crop: crop || 'unknown'
                }
            });

            // Save order as draft with Stripe session
            const order = new Order({
                userId: req.user._id,
                representativeId: req.user.representative || req.user._id,
                crop: crop || 'unknown',
                program: programName || 'Custom Order',
                acres,
                gpa,
                chemicals,
                totalCost: totalConfirmedPrice,
                costPerAcre: acres > 0 ? totalConfirmedPrice / acres : 0,
                status: 'draft',
                paymentStatus: 'pending',
                stripePaymentIntentId: session.payment_intent
            });
            await order.save();

            return res.json({
                success: true,
                orderStatus: 'checkout',
                checkoutUrl: session.url,
                orderId: order._id
            });

        } else {
            // Has items needing quotes — save and notify admin
            const order = new Order({
                userId: req.user._id,
                representativeId: req.user.representative || req.user._id,
                crop: crop || 'unknown',
                program: programName || 'Custom Order',
                acres,
                gpa,
                chemicals,
                totalCost: totalConfirmedPrice,
                costPerAcre: acres > 0 ? totalConfirmedPrice / acres : 0,
                status: 'draft',
                paymentStatus: 'pending',
                notes: 'Quote requested - some items need pricing'
            });
            await order.save();

            // Send email notification to admin
            const transporter = createEmailTransporter();
            if (transporter) {
                const needsQuoteItems = chemicals.filter(c => c.status === 'needs_quote');
                const confirmedItems = chemicals.filter(c => c.status === 'confirmed');

                await transporter.sendMail({
                    from: process.env.EMAIL_FROM || '"Acre Profit" <noreply@acreprofit.com>',
                    to: 'contact@acreprofit.com',
                    subject: `New Quote Request — ${req.user.name} — ${acres} acres`,
                    html: `
                        <div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">
                            <div style="background-color: #2d5a27; padding: 20px; text-align: center;">
                                <h1 style="color: white; margin: 0;">Quote Request</h1>
                            </div>
                            <div style="padding: 30px; background-color: #f9f9f9;">
                                <h2 style="color: #333;">Customer Information</h2>
                                <p><strong>Name:</strong> ${req.user.name}</p>
                                <p><strong>Email:</strong> ${req.user.email}</p>
                                <p><strong>Phone:</strong> ${req.user.phone || 'Not provided'}</p>

                                <h2 style="color: #333; margin-top: 20px;">Order Details</h2>
                                <p><strong>Crop:</strong> ${crop || 'Not specified'}</p>
                                <p><strong>Acres:</strong> ${acres}</p>
                                <p><strong>GPA:</strong> ${gpa}</p>

                                <h3 style="color: #c00; margin-top: 20px;">Items Needing Quote (${needsQuoteItems.length})</h3>
                                <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
                                    <tr style="background-color: #fdd;">
                                        <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Product</th>
                                        <th style="padding: 8px; border: 1px solid #ddd;">Packages Needed</th>
                                        <th style="padding: 8px; border: 1px solid #ddd;">Size</th>
                                    </tr>
                                    ${needsQuoteItems.map(item => `
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #ddd;">${item.name}</td>
                                            <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${item.packagesNeeded}</td>
                                            <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${item.packageSize} ${item.packageUnit}</td>
                                        </tr>
                                    `).join('')}
                                </table>

                                ${confirmedItems.length > 0 ? `
                                <h3 style="color: #2d5a27; margin-top: 20px;">Confirmed Items (${confirmedItems.length})</h3>
                                <table style="width: 100%; border-collapse: collapse;">
                                    <tr style="background-color: #dfd;">
                                        <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Product</th>
                                        <th style="padding: 8px; border: 1px solid #ddd;">Qty</th>
                                        <th style="padding: 8px; border: 1px solid #ddd;">Price</th>
                                    </tr>
                                    ${confirmedItems.map(item => `
                                        <tr>
                                            <td style="padding: 8px; border: 1px solid #ddd;">${item.name}</td>
                                            <td style="padding: 8px; border: 1px solid #ddd; text-align: center;">${item.packagesNeeded}</td>
                                            <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">$${item.totalPrice?.toFixed(2) || 'N/A'}</td>
                                        </tr>
                                    `).join('')}
                                </table>
                                <p style="text-align: right; font-weight: bold; margin-top: 10px;">
                                    Confirmed Total: $${totalConfirmedPrice.toFixed(2)}
                                </p>
                                ` : ''}

                                <p style="margin-top: 30px; padding: 15px; background-color: #fff3cd; border-radius: 4px;">
                                    <strong>Action Required:</strong> Please provide quotes for the items listed above and contact the customer.
                                </p>
                            </div>
                        </div>
                    `
                });
            }

            return res.json({
                success: true,
                orderStatus: 'quote_pending',
                orderId: order._id,
                message: 'Quote request submitted. You will be contacted within 24 hours.'
            });
        }
    } catch (error) {
        console.error('Submit order error:', error);
        res.status(400).json({ error: error.message });
    }
});

// Calculate order from existing program by ID (legacy support)
app.post('/api/spray-programs/:id/calculate', authMiddleware, async (req, res) => {
    try {
        const { acres, gpa = 10 } = req.body;
        const programId = req.params.id;

        if (!acres || acres <= 0) {
            return res.status(400).json({ error: 'Acres is required' });
        }

        let program;
        let chemicals = [];

        // Check if it's a hardcoded template
        if (programId.startsWith('template-')) {
            const templates = convertHardcodedToPrograms();
            program = templates.find(t => t._id === programId);
            if (!program) {
                return res.status(404).json({ error: 'Template not found' });
            }
            // For templates, we need to match chemicals by name
            for (const app of program.applications) {
                for (const chem of app.chemicals) {
                    const dbChem = await Chemical.findOne({
                        productName: { $regex: new RegExp(chem.productName, 'i') }
                    });
                    if (dbChem) {
                        chemicals.push({
                            chemicalId: dbChem._id,
                            rate: chem.suggestedRate,
                            unit: chem.rateUnit
                        });
                    }
                }
            }
        } else {
            program = await SprayProgram.findById(programId);
            if (!program) {
                return res.status(404).json({ error: 'Program not found' });
            }
            // Extract chemicals from program applications
            for (const app of program.applications) {
                for (const chem of app.chemicals) {
                    chemicals.push({
                        chemicalId: chem.chemicalId,
                        rate: chem.suggestedRate,
                        unit: chem.rateUnit
                    });
                }
            }
        }

        // Use the main calculate endpoint logic
        // Forward to calculate endpoint
        req.body = { acres, gpa, products: chemicals };

        // Call calculate logic directly (could also use next() pattern)
        const result = await calculateOrder(acres, gpa, chemicals);

        res.json({
            program: program.name,
            crop: program.crop,
            ...result
        });
    } catch (error) {
        console.error('Program calculate error:', error);
        res.status(400).json({ error: error.message });
    }
});

// Helper function for order calculation (reusable)
async function calculateOrder(acres, gpa, products) {
    const orderLines = [];
    let totalConfirmedPrice = 0;
    let hasNeedsQuote = false;
    let valorWarning = false;
    const totalWaterVolume = acres * gpa;

    for (const prod of products) {
        const { chemicalId, rate, unit } = prod;
        if (!chemicalId || !rate) continue;

        const chemical = await Chemical.findById(chemicalId);
        if (!chemical) continue;

        if (chemical.productName?.toLowerCase().includes('valor')) {
            valorWarning = true;
        }

        let totalNeeded = 0;
        const rateUnit = unit || chemical.rateUnit || 'oz/acre';

        switch (rateUnit) {
            case 'oz/acre': totalNeeded = (rate * acres) / 128; break;
            case 'pt/acre': totalNeeded = (rate * acres) / 8; break;
            case 'qt/acre': totalNeeded = (rate * acres) / 4; break;
            case 'gal/acre': totalNeeded = rate * acres; break;
            case 'lb/acre': totalNeeded = rate * acres; break;
            case '% v/v': totalNeeded = totalWaterVolume * (rate / 100); break;
            default: totalNeeded = rate * acres;
        }

        const packageSize = chemical.unitsPerPack || 1;
        const packagesNeeded = Math.ceil(totalNeeded / packageSize);

        const inventory = await Inventory.findOne({ chemicalId: chemical._id, location: 'main' });
        const onHandQuantity = inventory ? inventory.quantityAvailable : 0;

        let status, pricePerPackage, lineTotal;
        if (onHandQuantity >= packagesNeeded) {
            status = 'confirmed';
            pricePerPackage = chemical.sellPrice * packageSize;
            lineTotal = packagesNeeded * pricePerPackage;
            totalConfirmedPrice += lineTotal;
        } else {
            status = 'needs_quote';
            pricePerPackage = null;
            lineTotal = null;
            hasNeedsQuote = true;
        }

        orderLines.push({
            chemicalId: chemical._id,
            productName: chemical.productName,
            category: chemical.category,
            rate,
            unit: rateUnit,
            totalNeeded: Math.round(totalNeeded * 1000) / 1000,
            packageSize,
            packSize: chemical.packSize,
            packageUnit: chemical.unit,
            packagesNeeded,
            onHandQuantity,
            pricePerPackage: pricePerPackage ? Math.round(pricePerPackage * 100) / 100 : null,
            lineTotal: lineTotal ? Math.round(lineTotal * 100) / 100 : null,
            status,
            supplier: chemical.sourceSupplier
        });
    }

    // Hydrovant auto-calculation
    let hydrovant = null;
    const hvGallons = totalWaterVolume * 0.001;
    if (hvGallons > 0) {
        const hvChem = await Chemical.findOne({ productName: { $regex: /hydrovant/i } });
        if (hvChem) {
            const hvPkgSize = hvChem.unitsPerPack || 2.5;
            const hvPkgs = Math.ceil(hvGallons / hvPkgSize);
            const hvInv = await Inventory.findOne({ chemicalId: hvChem._id, location: 'main' });
            const hvOnHand = hvInv ? hvInv.quantityAvailable : 0;
            const hvStatus = hvOnHand >= hvPkgs ? 'confirmed' : 'needs_quote';
            const hvPrice = hvStatus === 'confirmed' ? hvChem.sellPrice * hvPkgSize : null;
            const hvTotal = hvStatus === 'confirmed' ? hvPkgs * hvPrice : null;

            if (hvStatus === 'confirmed' && hvTotal) totalConfirmedPrice += hvTotal;
            else hasNeedsQuote = true;

            hydrovant = {
                chemicalId: hvChem._id,
                productName: 'Hydrovant (Auto-Added)',
                gallonsNeeded: Math.round(hvGallons * 1000) / 1000,
                packageSize: hvPkgSize,
                packagesNeeded: hvPkgs,
                status: hvStatus,
                pricePerPackage: hvPrice ? Math.round(hvPrice * 100) / 100 : null,
                lineTotal: hvTotal ? Math.round(hvTotal * 100) / 100 : null
            };
        }
    }

    return {
        acres,
        gpa,
        totalWaterVolume,
        orderLines,
        hydrovant,
        valorWarning: valorWarning ? 'Valor requires application 7–30 days preplant. Minimum 1/4 inch rainfall required before planting.' : null,
        orderStatus: hasNeedsQuote ? 'pending_quote' : 'ready_for_checkout',
        totalConfirmedPrice: Math.round(totalConfirmedPrice * 100) / 100,
        costPerAcre: acres > 0 ? Math.round((totalConfirmedPrice / acres) * 100) / 100 : 0
    };
}

// ============ INVENTORY API ENDPOINTS ============

// Get all inventory (with option to include products without inventory records)
app.get('/api/admin/inventory', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { location, lowStock, includeAll } = req.query;
        let query = {};

        if (location) query.location = location;
        if (lowStock === 'true') {
            query.$expr = { $lte: ['$quantityOnHand', '$reorderPoint'] };
        }

        const inventory = await Inventory.find(query)
            .populate('chemicalId', 'productName packSize unit sellPrice costPrice adminPrice category')
            .populate('distributorId', 'name email')
            .sort({ productName: 1 });

        // Get "On Order" quantities from pending POs (not yet received)
        const pendingPOs = await PurchaseOrder.find({
            status: { $in: ['draft', 'submitted', 'confirmed', 'partial_received'] }
        }).select('items');

        // Build map of chemicalId -> quantity on order
        const onOrderMap = {};
        for (const po of pendingPOs) {
            for (const item of po.items || []) {
                if (item.chemicalId) {
                    const key = item.chemicalId.toString();
                    onOrderMap[key] = (onOrderMap[key] || 0) + (item.quantityOrdered || 0);
                }
            }
        }

        // Add onOrder quantity to each inventory item
        const enrichedInventory = inventory.map(inv => {
            const invObj = inv.toObject();
            const chemId = inv.chemicalId?._id?.toString() || inv.chemicalId?.toString();
            invObj.quantityOnOrder = chemId ? (onOrderMap[chemId] || 0) : 0;
            return invObj;
        });

        // If includeAll is true, also include chemicals without inventory records
        if (includeAll === 'true') {
            const existingChemicalIds = inventory.map(i => i.chemicalId?._id?.toString()).filter(Boolean);
            const missingChemicals = await Chemical.find({
                _id: { $nin: existingChemicalIds }
            }).select('productName packSize unit sellPrice costPrice adminPrice category');

            // Add placeholder records for missing chemicals
            const placeholders = missingChemicals.map(chem => ({
                _id: null,
                chemicalId: chem,
                productName: chem.productName,
                packSize: chem.packSize,
                unit: chem.unit,
                quantityOnHand: 0,
                quantityReserved: 0,
                quantityAvailable: 0,
                quantityOnOrder: onOrderMap[chem._id.toString()] || 0,
                averageCost: chem.costPrice || 0,
                location: 'main',
                needsInventoryRecord: true
            }));

            res.json([...enrichedInventory, ...placeholders]);
            return;
        }

        res.json(enrichedInventory);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Sync all products to inventory - creates inventory records for products that don't have one
app.post('/api/admin/inventory/sync', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const result = await syncCatalogToInventory();
        res.json(result);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Shared sync function - keeps product catalog and inventory in sync
async function syncCatalogToInventory(location = 'main') {
    const chemicals = await Chemical.find({ isActive: { $ne: false } });
    const existingInventory = await Inventory.find({ location });

    // Build map of existing inventory by chemicalId
    const invByChemId = {};
    existingInventory.forEach(inv => {
        const key = inv.chemicalId?.toString();
        if (key) invByChemId[key] = inv;
    });

    const created = [];
    const updated = [];
    const errors = [];

    for (const chem of chemicals) {
        const chemIdStr = chem._id.toString();
        const existing = invByChemId[chemIdStr];

        if (existing) {
            // Update existing inventory record if product info changed
            let needsSave = false;

            if (existing.productName !== chem.productName) {
                existing.productName = chem.productName;
                needsSave = true;
            }
            if (existing.packSize !== chem.packSize) {
                existing.packSize = chem.packSize;
                needsSave = true;
            }
            if (existing.unit !== chem.unit) {
                existing.unit = chem.unit;
                needsSave = true;
            }

            if (needsSave) {
                existing.updatedAt = new Date();
                await existing.save();
                updated.push(`${chem.productName} (${chem.packSize})`);
            }
        } else {
            // Create new inventory record
            try {
                await new Inventory({
                    chemicalId: chem._id,
                    productName: chem.productName,
                    packSize: chem.packSize,
                    unit: chem.unit,
                    location,
                    quantityOnHand: 0,
                    quantityReserved: 0,
                    quantityAvailable: 0,
                    averageCost: chem.costPrice || 0,
                    lastCost: chem.costPrice || 0,
                    reorderPoint: 0,
                    reorderQuantity: 0
                }).save();
                created.push(`${chem.productName} (${chem.packSize})`);
            } catch (err) {
                // Might be a duplicate - skip
                if (err.code !== 11000) {
                    errors.push({ name: chem.productName, error: err.message });
                }
            }
        }
    }

    if (created.length > 0 || updated.length > 0) {
        console.log(`Catalog sync: ${created.length} created, ${updated.length} updated`);
    }

    return {
        success: true,
        message: `Synced ${created.length} new, updated ${updated.length} existing`,
        created,
        updated,
        errors: errors.length > 0 ? errors : undefined,
        totalProducts: chemicals.length
    };
}

// Add inventory for a specific chemical
app.post('/api/admin/inventory', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { chemicalId, quantity, location = 'main', costPerUnit, notes } = req.body;

        if (!chemicalId) {
            return res.status(400).json({ error: 'Chemical ID is required' });
        }

        const chemical = await Chemical.findById(chemicalId);
        if (!chemical) {
            return res.status(404).json({ error: 'Chemical not found' });
        }

        // Check if inventory record exists
        let inventory = await Inventory.findOne({ chemicalId, location });

        if (inventory) {
            // Update existing record
            const previousQty = inventory.quantityOnHand;
            inventory.quantityOnHand += (quantity || 0);
            inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
            if (costPerUnit) {
                inventory.lastCost = costPerUnit;
                // Update average cost
                if (previousQty > 0) {
                    inventory.averageCost = ((previousQty * inventory.averageCost) + (quantity * costPerUnit)) / inventory.quantityOnHand;
                } else {
                    inventory.averageCost = costPerUnit;
                }
            }
            inventory.updatedAt = new Date();
            if (quantity > 0) inventory.lastReceivedDate = new Date();
            await inventory.save();
        } else {
            // Create new record
            inventory = new Inventory({
                chemicalId,
                productName: chemical.productName,
                packSize: chemical.packSize,
                unit: chemical.unit,
                location,
                quantityOnHand: quantity || 0,
                quantityReserved: 0,
                quantityAvailable: quantity || 0,
                averageCost: costPerUnit || chemical.costPrice || 0,
                lastCost: costPerUnit || chemical.costPrice || 0,
                lastReceivedDate: quantity > 0 ? new Date() : null
            });
            await inventory.save();
        }

        // Create transaction record if quantity was added
        if (quantity > 0) {
            const transaction = new InventoryTransaction({
                inventoryId: inventory._id,
                chemicalId,
                productName: chemical.productName,
                type: 'receive',
                quantityChange: quantity,
                previousQuantity: inventory.quantityOnHand - quantity,
                newQuantity: inventory.quantityOnHand,
                unitCost: costPerUnit || chemical.costPrice || 0,
                totalCost: quantity * (costPerUnit || chemical.costPrice || 0),
                location,
                notes: notes || 'Manual inventory addition',
                createdBy: req.user._id
            });
            await transaction.save();
        }

        res.json(inventory);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get inventory for a specific product
app.get('/api/admin/inventory/product/:chemicalId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const inventory = await Inventory.find({ chemicalId: req.params.chemicalId })
            .populate('distributorId', 'name email');

        res.json(inventory);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get inventory transactions (audit trail)
app.get('/api/admin/inventory/:id/transactions', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const transactions = await InventoryTransaction.find({ inventoryId: req.params.id })
            .populate('createdBy', 'name email')
            .sort({ createdAt: -1 })
            .limit(50);

        res.json(transactions);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Manual inventory adjustment
app.post('/api/admin/inventory/adjust', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { chemicalId, location, quantityChange, reason, notes } = req.body;

        const inventory = await Inventory.findOne({ chemicalId, location: location || 'main' });
        if (!inventory) {
            return res.status(404).json({ error: 'Inventory record not found' });
        }

        const previousQuantity = inventory.quantityOnHand;
        const newQuantity = previousQuantity + quantityChange;

        if (newQuantity < 0) {
            return res.status(400).json({ error: 'Adjustment would result in negative inventory' });
        }

        inventory.quantityOnHand = newQuantity;
        inventory.quantityAvailable = newQuantity - inventory.quantityReserved;
        inventory.updatedAt = new Date();

        await inventory.save();

        // Create transaction record
        const transaction = new InventoryTransaction({
            inventoryId: inventory._id,
            chemicalId,
            productName: inventory.productName,
            type: 'adjustment',
            quantityChange,
            previousQuantity,
            newQuantity,
            location: location || 'main',
            reason,
            notes,
            createdBy: req.user._id
        });

        await transaction.save();

        res.json({ inventory, transaction });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ INVENTORY BATCH ENDPOINTS ============

// Get all active batches (with optional filters)
app.get('/api/admin/inventory/batches', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { chemicalId, poNumber, status = 'active', location } = req.query;
        let query = {};

        if (chemicalId) query.chemicalId = chemicalId;
        if (poNumber) query.poNumber = { $regex: poNumber, $options: 'i' };
        if (status && status !== 'all') query.status = status;
        if (location) query.location = location;

        const batches = await InventoryBatch.find(query)
            .populate('chemicalId', 'productName packSize unit category')
            .sort({ receivedDate: -1 });

        res.json(batches);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get batches for a specific product
app.get('/api/admin/inventory/batches/product/:chemicalId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const batches = await InventoryBatch.find({
            chemicalId: req.params.chemicalId,
            status: { $in: ['active', 'depleted'] }
        }).sort({ receivedDate: -1 });

        res.json(batches);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update batch info (lot number, notes, etc.)
app.patch('/api/admin/inventory/batches/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { lotNumber, status, notes } = req.body;
        const batch = await InventoryBatch.findById(req.params.id);

        if (!batch) {
            return res.status(404).json({ error: 'Batch not found' });
        }

        if (lotNumber !== undefined) batch.lotNumber = lotNumber;
        if (status) batch.status = status;
        if (notes !== undefined) batch.notes = notes;
        batch.updatedAt = new Date();

        await batch.save();
        res.json(batch);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Backfill inventory from received POs (manual trigger)
app.post('/api/admin/inventory/backfill', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        await backfillInventoryFromPOs();
        res.json({ message: 'Inventory backfill completed. Check server logs for details.' });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// One-time: receive March 2026 PO inventory into stock
// Safe to run multiple times - skips products already received
app.post('/api/admin/backfill-march-2026-inventory', authMiddleware, superAdminMiddleware, async (req, res) => {
    const TARGET_POS = ['JABCO-SO2129', 'SIMS-103850', 'SIMS-103849'];
    const results = { received: [], skipped: [], errors: [] };

    const session = await mongoose.startSession();
    try {
        await session.withTransaction(async () => {
            for (const poNumber of TARGET_POS) {
                const po = await PurchaseOrder.findOne({ poNumber }).session(session);
                if (!po) {
                    results.errors.push({ poNumber, error: 'PO not found' });
                    continue;
                }

                for (const item of po.items) {
                    // Idempotency check - skip if inventory already exists for this chemical
                    const existing = await Inventory.findOne({ chemicalId: item.chemicalId, location: 'main' }).session(session);
                    if (existing && existing.quantityOnHand > 0) {
                        results.skipped.push({ poNumber, product: item.productName, qty: existing.quantityOnHand });
                        continue;
                    }

                    await receiveInventory({
                        chemicalId: item.chemicalId,
                        productName: item.productName,
                        packSize: item.packSize,
                        unit: item.unit,
                        quantity: item.quantityOrdered,
                        unitCost: item.pricePerUnit,
                        location: 'main',
                        purchaseOrderId: po._id,
                        poNumber: po.poNumber,
                        supplierName: po.supplier?.name || '',
                        userId: req.user._id,
                        session
                    });

                    results.received.push({ poNumber, product: item.productName, qty: item.quantityOrdered, cost: item.pricePerUnit });
                }

                // Mark PO as received
                po.status = 'received';
                po.receivedDate = new Date();
                await po.save({ session });
            }
        });

        res.json({
            message: 'March 2026 inventory backfill complete',
            summary: {
                received: results.received.length,
                skipped: results.skipped.length,
                errors: results.errors.length
            },
            results
        });
    } catch (error) {
        res.status(500).json({ error: error.message, partial: results });
    } finally {
        session.endSession();
    }
});

// Seed or update the milo 2-pass program (KSU-based, NE Colorado dryland).
// Idempotent: finds the program by name, updates if exists, creates if not.
// Uses the real schema field names (suggestedRate, deliveryWindow, isAdjuvant,
// precautions) added in commit 14cd798.
app.post('/api/admin/seed-milo-program', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const miloProgram = {
            name: 'Milo 2-Pass Program - Hardy Economical',
            crop: 'milo',
            type: 'template',
            isPublic: true,
            description: 'KSU-based 2-pass program for NE Colorado dryland milo. Pass 1 burndown 30 days before planting, Pass 2 at-planting pre-emerge. Requires Concep-safened seed.',
            applications: [
                {
                    name: 'Pass 1 — Pre-Plant Burndown',
                    timing: '30 days before planting (Early-Mid April)',
                    deliveryWindow: 'Late March',
                    chemicals: [
                        { productName: 'XSATE Glyphosate 53.8%', suggestedRate: 22, rateUnit: 'oz/acre', notes: 'Kills emerged weeds' },
                        { productName: 'Atrazine 4L', suggestedRate: 1, rateUnit: 'qt/acre', notes: 'Residual broadleaf - pigweed, ragweed, mustards' },
                        { productName: 'Flumioxazin 51% WDG', suggestedRate: 2, rateUnit: 'oz/acre', notes: 'Palmer amaranth, kochia, marestail residual. MUST be 30 days before planting.' },
                        { productName: 'Dicamba 49.8% SL', suggestedRate: 4, rateUnit: 'oz/acre', notes: 'Broadleaf burndown boost' },
                        { productName: 'Hydrovant fA', suggestedRate: 1.28, rateUnit: 'fl oz/acre', isAdjuvant: true, notes: 'Drift reduction adjuvant' }
                    ]
                },
                {
                    name: 'Pass 2 — At-Planting / Pre-Emerge',
                    timing: 'Mid May. Apply after seed germination but before crop emergence. Dicamba must be applied in this window — after germination, before emergence — to avoid crop injury.',
                    deliveryWindow: 'Early May',
                    chemicals: [
                        { productName: 'S-Metolachlor (Dual II Magnum)', suggestedRate: 1.33, rateUnit: 'pt/acre' },
                        { productName: 'Meso 4SC', suggestedRate: 6, rateUnit: 'fl oz/acre' },
                        { productName: 'Atrazine 4L', suggestedRate: 1, rateUnit: 'qt/acre' },
                        { productName: 'XSATE Glyphosate 53.8%', suggestedRate: 28, rateUnit: 'fl oz/acre' },
                        { productName: 'Dicamba 49.8% SL', suggestedRate: 6, rateUnit: 'fl oz/acre' },
                        { productName: 'Hydrovant fA', suggestedRate: 1.28, rateUnit: 'fl oz/acre', isAdjuvant: true }
                    ]
                }
            ],
            precautions: [
                'CONCEP-SAFENED SEED REQUIRED - S-Metolachlor (Dual II Magnum) will injure milo without Concep III safener on seed. Corn safener does NOT work.',
                'FLUMIOXAZIN 30-DAY RULE - Must be applied minimum 30 days before planting. At least 1 inch rainfall required between application and planting.',
                'NO POST-EMERGE GRASS CONTROL - There are no herbicides labeled for post-emergence grass control in conventional grain sorghum. Pass 1 and Pass 2 residuals are your only grass protection.',
                'DICAMBA WAIT - 15-day waiting period between dicamba application and sorghum planting when using 8 fl oz Clarity. At 4 oz rate, 7-day wait recommended.',
                'ATRAZINE RUNOFF - In sensitive watersheds, do not exceed 1 lb ai/acre at planting due to runoff risk.',
                'GRAZING - Do not graze sorghum forage for 60 days after atrazine application.',
                'SCOUT 14-21 days after each application. Adjust rates based on soil type, organic matter, and weed pressure. Label is the law.'
            ]
        };

        // Link chemicals to their catalog IDs where possible (for cost calculations)
        for (const app of miloProgram.applications) {
            for (const chem of app.chemicals) {
                const catalogMatch = await Chemical.findOne({ productName: chem.productName }).select('_id packSize unit unitsPerPack').lean();
                if (catalogMatch) {
                    chem.chemicalId = catalogMatch._id;
                    chem.packSize = catalogMatch.packSize;
                    chem.unit = catalogMatch.unit;
                    chem.unitsPerPack = catalogMatch.unitsPerPack;
                }
            }
        }

        const existing = await SprayProgram.findOne({ name: miloProgram.name });
        if (existing) {
            Object.assign(existing, miloProgram);
            existing.updatedAt = new Date();
            await existing.save();
            return res.json({ message: 'Milo program updated', id: existing._id, program: existing });
        }

        const program = new SprayProgram({
            ...miloProgram,
            createdBy: req.user._id
        });
        await program.save();
        res.json({ message: 'Milo program created', id: program._id, program });
    } catch (error) {
        console.error('seed-milo-program error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Seed Field Pea Pre-Emerge program. Single-pass pre-emerge weed control for
// field peas. Applied at planting before pea emergence. Idempotent: updates
// the existing doc in-place via Object.assign (same pattern as seed-milo-program)
// so re-runs pick up spec changes — including unitsPerPack backfill.
app.post('/api/admin/seed-field-pea-program', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const programSpec = {
            name: 'Field Pea Pre-Emerge Program',
            crop: 'fieldpeas',
            description: 'Pre-emerge weed control for field peas. Apply at planting before pea emergence. Proven on this crop in NE Colorado.',
            type: 'template',
            isPublic: true,
            isActive: true,
            applications: [
                {
                    name: 'Pass 1: Pre-Emerge',
                    timing: 'At planting — apply before pea emergence. Works in the soil ahead of germinating weeds.',
                    deliveryWindow: 'Early May',
                    chemicals: [
                        { productName: 'Sulfentrazone 39.6% SC', suggestedRate: 6, rateUnit: 'fl oz/acre' },
                        { productName: 'S-Metolachlor (Dual II Magnum)', suggestedRate: 21, rateUnit: 'fl oz/acre' },
                        { productName: 'XSATE Glyphosate 53.8%', suggestedRate: 21, rateUnit: 'fl oz/acre' },
                        { productName: 'Hydrovant fA', suggestedRate: 1.28, rateUnit: 'fl oz/acre', isAdjuvant: true }
                    ]
                }
            ],
            precautions: [
                'SULFENTRAZONE — 18-month rotation restriction to corn and sorghum. Safe to wheat the following fall at labeled rates.',
                'S-METOLACHLOR (Dual II Magnum) — labeled for field peas. No seed safener required for peas (safener required for sorghum only).',
                'Confirm your field pea variety tolerance with your seed rep before application.',
                'Activation requires rainfall or irrigation within 7 days of application for best residual control.',
                'Avoid sandy soils or soils with pH above 7.5 for sulfentrazone — binding is reduced and crop injury risk increases.'
            ],
            rotationRestrictions: 'Corn/Sorghum: 18 months after sulfentrazone. Wheat: 4 months. Soybeans: 12 months. Always check the full sulfentrazone label for your planned rotation.',
            grazingRestrictions: 'Do not graze treated areas or cut for hay for 28 days after sulfentrazone application.',
            groundType: 'Field pea ground with annual grass and broadleaf pressure. Medium to heavy textured soils preferred for sulfentrazone. Avoid coarse sands or high-pH soils.'
        };

        // Link chemicals to catalog (same pattern as seed-milo-program).
        // Mutates programSpec.applications in place before save/assign.
        for (const app of programSpec.applications) {
            for (const chem of app.chemicals) {
                const catalogMatch = await Chemical.findOne({ productName: chem.productName }).select('_id packSize unit unitsPerPack').lean();
                if (catalogMatch) {
                    chem.chemicalId = catalogMatch._id;
                    chem.packSize = catalogMatch.packSize;
                    chem.unit = catalogMatch.unit;
                    chem.unitsPerPack = catalogMatch.unitsPerPack;
                }
            }
        }

        const existing = await SprayProgram.findOne({ name: programSpec.name });
        if (existing) {
            Object.assign(existing, programSpec);
            existing.updatedAt = new Date();
            await existing.save();
            return res.json({ message: 'Field Pea program updated', id: existing._id, program: existing });
        }

        const program = new SprayProgram({
            ...programSpec,
            createdBy: req.user._id
        });
        await program.save();
        res.json({ message: 'Field Pea Pre-Emerge Program seeded successfully', id: program._id, program });
    } catch (error) {
        console.error('seed-field-pea-program error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Transfer inventory between locations
app.post('/api/admin/inventory/transfer', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { chemicalId, quantity, fromLocation, toLocation, notes, fromRepId, toRepId } = req.body;

        if (!chemicalId || !quantity || !fromLocation || !toLocation) {
            return res.status(400).json({ error: 'chemicalId, quantity, fromLocation, and toLocation are required' });
        }
        if (fromLocation === toLocation) {
            return res.status(400).json({ error: 'From and to locations must be different' });
        }
        if (quantity <= 0) {
            return res.status(400).json({ error: 'Quantity must be positive' });
        }

        // Deduct from source
        const sourceInv = await Inventory.findOne({ chemicalId, location: fromLocation });
        if (!sourceInv) {
            return res.status(400).json({ error: `No inventory at ${fromLocation}` });
        }
        if (sourceInv.quantityAvailable < quantity) {
            return res.status(400).json({ error: `Insufficient stock at ${fromLocation}. Available: ${sourceInv.quantityAvailable}` });
        }

        sourceInv.quantityOnHand -= quantity;
        sourceInv.quantityAvailable = sourceInv.quantityOnHand - sourceInv.quantityReserved;
        sourceInv.updatedAt = new Date();
        await sourceInv.save();

        // Add to destination
        let destInv = await Inventory.findOne({ chemicalId, location: toLocation });
        if (!destInv) {
            destInv = new Inventory({
                chemicalId,
                productName: sourceInv.productName,
                packSize: sourceInv.packSize,
                unit: sourceInv.unit,
                location: toLocation,
                quantityOnHand: 0,
                quantityReserved: 0,
                quantityAvailable: 0,
                averageCost: sourceInv.averageCost,
                lastCost: sourceInv.lastCost
            });
        }
        destInv.quantityOnHand += quantity;
        destInv.quantityAvailable = destInv.quantityOnHand - destInv.quantityReserved;
        destInv.averageCost = destInv.averageCost || sourceInv.averageCost;
        if (toRepId) destInv.distributorId = toRepId;
        destInv.updatedAt = new Date();
        await destInv.save();

        // Update source distributorId if transferring to a different rep
        if (fromRepId) {
            sourceInv.distributorId = fromRepId;
            await sourceInv.save();
        }

        // Audit trail - source
        await new InventoryTransaction({
            inventoryId: sourceInv._id,
            chemicalId,
            productName: sourceInv.productName,
            type: 'transfer',
            quantityChange: -quantity,
            previousQuantity: sourceInv.quantityOnHand + quantity,
            newQuantity: sourceInv.quantityOnHand,
            unitCost: sourceInv.averageCost,
            totalCost: quantity * sourceInv.averageCost,
            referenceType: 'Transfer',
            location: fromLocation,
            fromLocation,
            toLocation,
            notes: notes || `Transfer to ${toLocation}`,
            createdBy: req.user._id
        }).save();

        // Audit trail - destination
        await new InventoryTransaction({
            inventoryId: destInv._id,
            chemicalId,
            productName: destInv.productName,
            type: 'transfer',
            quantityChange: quantity,
            previousQuantity: destInv.quantityOnHand - quantity,
            newQuantity: destInv.quantityOnHand,
            unitCost: sourceInv.averageCost,
            totalCost: quantity * sourceInv.averageCost,
            referenceType: 'Transfer',
            location: toLocation,
            fromLocation,
            toLocation,
            notes: notes || `Transfer from ${fromLocation}`,
            createdBy: req.user._id
        }).save();

        // Create ledger entries for transfer between reps
        const transferValue = Math.round(quantity * (sourceInv.averageCost || 0) * 100) / 100;
        const productLabel = `${sourceInv.productName} (${quantity} ${sourceInv.unit})`;

        if (fromRepId && transferValue > 0) {
            // Credit the sender - they gave product, reducing what they owe
            await createLedgerEntry({
                representativeId: fromRepId,
                description: `Transfer OUT: ${productLabel} to ${toLocation}`,
                amount: transferValue,
                type: 'credit',
                category: 'adjustment',
                referenceType: 'Manual',
                createdBy: req.user._id,
                notes: notes || `Inventory transferred from ${fromLocation} to ${toLocation}`
            });
        }

        if (toRepId && transferValue > 0) {
            // Debit the receiver - they got product, increasing what they owe
            await createLedgerEntry({
                representativeId: toRepId,
                description: `Transfer IN: ${productLabel} from ${fromLocation}`,
                amount: transferValue,
                type: 'debit',
                category: 'adjustment',
                referenceType: 'Manual',
                createdBy: req.user._id,
                notes: notes || `Inventory received from ${fromLocation} to ${toLocation}`
            });
        }

        res.json({
            message: `Transferred ${quantity} ${sourceInv.unit} of ${sourceInv.productName} from ${fromLocation} to ${toLocation}`,
            transferValue,
            source: sourceInv,
            destination: destInv
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Receive Purchase Order - Add items to inventory
app.post('/api/admin/purchase-orders/:id/receive', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { items, location, notes } = req.body;
        // items: [{ itemIndex: 0, quantityReceived: 10, lotNumber: 'ABC123' }, ...]

        const po = await PurchaseOrder.findById(req.params.id).populate('supplierId', 'name');
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        const results = [];

        for (const receiveItem of items) {
            const poItem = po.items[receiveItem.itemIndex];
            if (!poItem) {
                return res.status(400).json({ error: `Invalid item index: ${receiveItem.itemIndex}` });
            }

            // Get chemicalId - use from PO item, or look up by product name
            let chemicalId = poItem.chemicalId;
            if (!chemicalId && poItem.productName) {
                // Try to find chemical by product name (fallback for old POs without chemicalId)
                const chemical = await Chemical.findOne({
                    productName: { $regex: new RegExp(`^${poItem.productName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i') }
                });
                if (chemical) {
                    chemicalId = chemical._id;
                    // Update the PO item with the found chemicalId for future reference
                    poItem.chemicalId = chemicalId;
                } else {
                    console.warn(`Could not find chemical for product: ${poItem.productName}`);
                    return res.status(400).json({
                        error: `Product "${poItem.productName}" not found in catalog. Please add it to Products first.`
                    });
                }
            }

            if (!chemicalId) {
                return res.status(400).json({
                    error: `Missing product reference for item: ${poItem.productName || 'Unknown'}`
                });
            }

            const quantityReceived = receiveItem.quantityReceived || poItem.quantityOrdered;

            // Add to inventory with batch tracking
            const result = await receiveInventory({
                chemicalId: chemicalId,
                productName: poItem.productName,
                packSize: poItem.packSize,
                unit: poItem.unit,
                quantity: quantityReceived,
                unitCost: poItem.pricePerUnit,
                location: location || 'main',
                purchaseOrderId: po._id,
                poNumber: po.poNumber,
                lotNumber: receiveItem.lotNumber || '',
                supplierName: po.supplierId?.name || po.supplierName || '',
                userId: req.user._id
            });

            results.push({
                productName: poItem.productName,
                quantityReceived,
                inventory: result.inventory,
                batch: result.batch
            });
        }

        // Track received quantities on each PO item
        for (const receiveItem of items) {
            const poItem = po.items[receiveItem.itemIndex];
            if (poItem) {
                const qtyReceived = receiveItem.quantityReceived || poItem.quantityOrdered;
                poItem.quantityReceived = (poItem.quantityReceived || 0) + qtyReceived;
            }
        }

        // Determine PO status based on total received vs ordered
        const allFullyReceived = po.items.every(item =>
            (item.quantityReceived || 0) >= item.quantityOrdered
        );
        const anyReceived = po.items.some(item => (item.quantityReceived || 0) > 0);

        if (allFullyReceived) {
            po.status = 'received';
            po.receivedDate = new Date();
        } else if (anyReceived) {
            po.status = 'partial_received';
        }

        po.updatedBy = req.user._id;
        po.updatedAt = new Date();
        if (notes) po.internalNotes = (po.internalNotes || '') + '\n' + notes;

        await po.save();

        res.json({
            message: allFullyReceived ? 'All items received successfully' : 'Partial shipment received',
            purchaseOrder: po,
            inventoryUpdates: results
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get products with inventory levels (for Products tab)
app.get('/api/admin/products-with-inventory', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Filter: ?status=active (default) | archived | all
        const statusFilter = (req.query.status || 'active').toLowerCase();
        const query = {};
        if (statusFilter === 'active') query.isActive = true;
        else if (statusFilter === 'archived') query.isActive = false;
        // 'all' leaves query empty

        const chemicals = await Chemical.find(query)
            .sort({ productName: 1, packSize: 1 });

        // Get inventory for all chemicals
        const inventoryData = await Inventory.find({}).lean();
        const inventoryMap = {};
        for (const inv of inventoryData) {
            const key = inv.chemicalId.toString();
            if (!inventoryMap[key]) inventoryMap[key] = [];
            inventoryMap[key].push(inv);
        }

        // Get sales data
        const salesData = await ChemicalOrder.aggregate([
            { $match: { status: { $nin: ['cancelled', 'draft'] } } },
            { $unwind: '$items' },
            { $group: {
                _id: '$items.chemicalId',
                unitsSold: { $sum: '$items.quantity' },
                totalRevenue: { $sum: '$items.totalPrice' }
            }}
        ]);

        const salesMap = {};
        for (const s of salesData) {
            if (s._id) salesMap[s._id.toString()] = { unitsSold: s.unitsSold, totalRevenue: s.totalRevenue };
        }

        const products = chemicals.map(c => {
            const inventoryRecords = inventoryMap[c._id.toString()] || [];
            const totalOnHand = inventoryRecords.reduce((sum, inv) => sum + (inv.quantityOnHand || 0), 0);
            const totalAvailable = inventoryRecords.reduce((sum, inv) => sum + (inv.quantityAvailable || 0), 0);

            return {
                _id: c._id,
                productName: c.productName,
                sourceSupplier: c.sourceSupplier,
                category: c.category,
                packSize: c.packSize,
                unit: c.unit,
                costPrice: c.costPrice,
                adminPrice: c.adminPrice,
                adminMargin: c.adminMargin,
                sellPrice: c.sellPrice,
                margin: c.margin,
                activeIngredients: c.activeIngredients,
                isActive: c.isActive,
                // Inventory data
                quantityOnHand: totalOnHand,
                quantityAvailable: totalAvailable,
                inventoryByLocation: inventoryRecords.map(inv => ({
                    location: inv.location,
                    onHand: inv.quantityOnHand,
                    available: inv.quantityAvailable
                })),
                // Sales data
                unitsSold: salesMap[c._id.toString()]?.unitsSold || 0,
                totalRevenue: salesMap[c._id.toString()]?.totalRevenue || 0
            };
        });

        res.json(products);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ INVOICE API ENDPOINTS ============

// Get all invoices
app.get('/api/admin/invoices', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, customerId } = req.query;
        let query = {};

        if (status) query.status = status;
        if (customerId) query.customerId = customerId;

        // For non-superadmin, only show their customers' invoices
        if (req.user.role !== 'superadmin') {
            const customers = await User.find({ representative: req.user._id }).select('_id');
            const customerIds = customers.map(c => c._id);
            query.customerId = { $in: customerIds };
        }

        const invoices = await Invoice.find(query)
            .populate('customerId', 'name email phone')
            .populate('representativeId', 'name email')
            .sort({ invoiceDate: -1 });

        res.json(invoices);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Customer-facing: get the invoice for one of their orders (or admin/rep lookup)
// Returns 404 if no invoice exists yet - frontend falls back to client-side synthesis
// Returns 403 if the user doesn't own the order
app.get('/api/invoices/by-order/:orderId', authMiddleware, async (req, res) => {
    try {
        const invoice = await Invoice.findOne({ orderId: req.params.orderId })
            .populate('customerId', 'name email phone farm address')
            .populate('representativeId', 'name email phone');

        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }

        // Ownership: customer on the invoice, OR admin/superadmin/distributor
        const isOwner = invoice.customerId?._id?.toString() === req.user._id.toString();
        const isAdmin = ['admin', 'superadmin', 'distributor'].includes(req.user.role);
        if (!isOwner && !isAdmin) {
            return res.status(403).json({ error: 'Not authorized to view this invoice' });
        }

        res.json(invoice);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single invoice
app.get('/api/admin/invoices/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const invoice = await Invoice.findById(req.params.id)
            .populate('customerId', 'name email phone farm address')
            .populate('representativeId', 'name email phone')
            .populate('orderId');

        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }

        res.json(invoice);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create invoice from order
app.post('/api/admin/invoices/from-order/:orderId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Try ChemicalOrder first, then Order
        let order = await ChemicalOrder.findById(req.params.orderId)
            .populate('userId', 'name email phone address farm');

        let isChemicalOrder = !!order;

        if (!order) {
            // Try regular Order model
            order = await Order.findById(req.params.orderId)
                .populate('userId', 'name email phone farm');
        }

        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        // Check if invoice already exists for this order
        const existing = await Invoice.findOne({ orderId: order._id });
        if (existing) {
            return res.status(400).json({ error: 'Invoice already exists for this order', invoiceId: existing._id });
        }

        const invoiceNumber = await generateInvoiceNumber();
        const customer = order.userId;

        let invoiceItems = [];
        let subtotal = 0;
        let total = 0;
        let orderNumber = '';

        if (isChemicalOrder) {
            // ChemicalOrder has items array - enrich with cost/margin data from catalog
            const chemicalIds = order.items.map(i => i.chemicalId).filter(Boolean);
            const chemicalsMap = {};
            if (chemicalIds.length > 0) {
                const chems = await Chemical.find({ _id: { $in: chemicalIds } }).lean();
                chems.forEach(c => { chemicalsMap[c._id.toString()] = c; });
            }

            invoiceItems = order.items.map(item => {
                const chem = item.chemicalId ? chemicalsMap[item.chemicalId.toString()] : null;
                const costPrice = chem?.costPrice || 0;
                const adminPrice = chem?.adminPrice || 0;
                const unitPrice = item.unitPrice || item.pricePerUnit || chem?.sellPrice || 0;
                const qty = item.quantity || 0;
                return {
                    productName: item.productName,
                    description: `${item.packSize || ''} ${item.unit || ''}`.trim(),
                    packSize: item.packSize,
                    unit: item.unit,
                    unitsPerPack: chem?.unitsPerPack || 1,
                    quantity: qty,
                    packQuantity: item.packQuantity || (chem?.unitsPerPack ? Math.ceil(qty / chem.unitsPerPack) : qty),
                    unitPrice,
                    costPrice,
                    adminPrice,
                    totalPrice: qty * unitPrice,
                    margin: (unitPrice - costPrice) * qty
                };
            });
            subtotal = order.subtotal;
            total = order.total;
            orderNumber = order.orderNumber;
        } else {
            // Regular Order has chemicals, seeds, pivotBio arrays
            // Convert chemicals to invoice items
            if (order.chemicals && order.chemicals.length > 0) {
                order.chemicals.forEach(chem => {
                    invoiceItems.push({
                        productName: chem.name,
                        description: `${chem.packageSize} ${chem.packageUnit}`,
                        packSize: `${chem.packageSize}`,
                        unit: chem.packageUnit,
                        quantity: chem.packagesNeeded,
                        unitPrice: chem.pricePerPackage,
                        totalPrice: chem.totalPrice
                    });
                });
            }
            // Convert seeds to invoice items
            if (order.seeds && order.seeds.length > 0) {
                order.seeds.forEach(seed => {
                    invoiceItems.push({
                        productName: seed.name,
                        description: `${seed.crop} seed`,
                        packSize: 'bag',
                        unit: 'bags',
                        quantity: seed.bagsNeeded,
                        unitPrice: seed.pricePerBag,
                        totalPrice: seed.totalPrice
                    });
                });
            }
            // Convert pivotBio to invoice items
            if (order.pivotBio && order.pivotBio.length > 0) {
                order.pivotBio.forEach(pb => {
                    invoiceItems.push({
                        productName: pb.product,
                        description: 'PivotBio',
                        packSize: 'unit',
                        unit: 'units',
                        quantity: Math.ceil(pb.totalAmount),
                        unitPrice: pb.pricePerUnit,
                        totalPrice: pb.totalPrice
                    });
                });
            }
            subtotal = order.totalCost || invoiceItems.reduce((sum, item) => sum + (item.totalPrice || 0), 0);
            total = subtotal;
            orderNumber = `ORD-${order._id.toString().slice(-8).toUpperCase()}`;
        }

        const invoice = new Invoice({
            invoiceNumber,
            customerId: customer._id,
            customerName: customer.name,
            customerEmail: customer.email,
            customerPhone: customer.phone,
            customerAddress: customer.address,
            orderId: order._id,
            orderNumber: orderNumber,
            representativeId: order.representativeId,
            items: invoiceItems,
            subtotal: subtotal,
            discount: order.discount || 0,
            total: total,
            amountDue: total - (order.discount || 0),
            dueDate: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000), // 30 days
            createdBy: req.user._id
        });

        await invoice.save();

        res.status(201).json(invoice);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create manual invoice
app.post('/api/admin/invoices', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { customerId, items, discount, discountReason, notes, dueDate, representativeId } = req.body;

        const customer = await User.findById(customerId);
        if (!customer) {
            return res.status(404).json({ error: 'Customer not found' });
        }

        const invoiceNumber = await generateInvoiceNumber();

        // Calculate totals
        const subtotal = items.reduce((sum, item) => sum + (item.quantity * item.unitPrice), 0);
        const total = subtotal - (discount || 0);

        // Determine representative: allow superadmin to specify, otherwise use customer's rep or current user
        let repId = customer.representative || req.user._id;
        if (req.user.role === 'superadmin' && representativeId) {
            repId = representativeId;
        }

        const invoice = new Invoice({
            invoiceNumber,
            customerId: customer._id,
            customerName: customer.name,
            customerEmail: customer.email,
            customerPhone: customer.phone,
            customerAddress: customer.address,
            representativeId: repId,
            items: items.map(item => ({
                ...item,
                totalPrice: item.quantity * item.unitPrice
            })),
            subtotal,
            discount: discount || 0,
            discountReason,
            total,
            amountDue: total,
            dueDate: dueDate ? new Date(dueDate) : new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
            notes,
            createdBy: req.user._id
        });

        await invoice.save();

        res.status(201).json(invoice);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update invoice status
app.put('/api/admin/invoices/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, paymentStatus, amountPaid, paymentMethod, paymentDate, notes } = req.body;

        const invoice = await Invoice.findById(req.params.id);
        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }

        if (status) invoice.status = status;
        if (paymentStatus) invoice.paymentStatus = paymentStatus;
        if (amountPaid !== undefined) {
            invoice.amountPaid = amountPaid;
            invoice.amountDue = invoice.total - amountPaid;
            if (amountPaid >= invoice.total) {
                invoice.paymentStatus = 'paid';
            } else if (amountPaid > 0) {
                invoice.paymentStatus = 'partial';
            }
        }
        if (paymentMethod) invoice.paymentMethod = paymentMethod;
        if (paymentDate) invoice.paymentDate = new Date(paymentDate);
        if (notes !== undefined) invoice.notes = notes;

        invoice.updatedAt = new Date();
        await invoice.save();

        res.json(invoice);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete a DRAFT invoice. Sent/paid invoices stay - deleting a record that was
// already emailed to a customer would create accounting/audit gaps. Use the
// 'cancelled' status path for those.
app.delete('/api/admin/invoices/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const invoice = await Invoice.findById(req.params.id);
        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }
        if (invoice.status && invoice.status !== 'draft') {
            return res.status(400).json({
                error: `Cannot delete invoice in status '${invoice.status}'. Only draft invoices can be hard-deleted.`
            });
        }

        const snapshot = {
            invoiceNumber: invoice.invoiceNumber,
            customerName: invoice.customerName,
            total: invoice.total,
            status: invoice.status
        };

        await Invoice.deleteOne({ _id: invoice._id });

        await logAudit({
            action: 'invoice_delete',
            req,
            entityType: 'Invoice',
            entityId: invoice._id,
            entityRef: invoice.invoiceNumber,
            before: snapshot,
            reason: req.body?.reason || `Draft invoice deleted by ${req.user.name}`
        });

        res.json({ message: `Invoice ${snapshot.invoiceNumber} deleted` });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Send invoice to customer
app.post('/api/admin/invoices/:id/send', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const invoice = await Invoice.findById(req.params.id)
            .populate('customerId', 'name email phone')
            .populate('orderId');

        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }

        // Get customer email
        const customerEmail = invoice.customerEmail || invoice.customerId?.email;
        if (!customerEmail) {
            return res.status(400).json({ error: 'No customer email found' });
        }

        // Format dates
        const invoiceDate = new Date(invoice.invoiceDate || invoice.createdAt).toLocaleDateString('en-US', {
            year: 'numeric', month: 'long', day: 'numeric'
        });
        const dueDate = new Date(invoice.dueDate || new Date(Date.now() + 30 * 24 * 60 * 60 * 1000)).toLocaleDateString('en-US', {
            year: 'numeric', month: 'long', day: 'numeric'
        });

        // Build items table HTML
        const itemsHtml = (invoice.items || []).map(item => `
            <tr>
                <td style="padding: 14px 16px; border-bottom: 1px solid #e0e0e0;">${item.productName || item.description || 'Product'}</td>
                <td style="padding: 14px 16px; border-bottom: 1px solid #e0e0e0; text-align: center;">${item.quantity || 0}</td>
                <td style="padding: 14px 16px; border-bottom: 1px solid #e0e0e0; text-align: right;">$${(item.unitPrice || 0).toFixed(2)}</td>
                <td style="padding: 14px 16px; border-bottom: 1px solid #e0e0e0; text-align: right; font-weight: 600;">$${(item.total || (item.quantity * item.unitPrice) || 0).toFixed(2)}</td>
            </tr>
        `).join('');

        // Calculate totals
        const subtotal = invoice.subtotal || invoice.items?.reduce((sum, i) => sum + (i.total || i.quantity * i.unitPrice || 0), 0) || 0;
        const discount = invoice.discount || 0;
        const total = invoice.total || (subtotal - discount);

        // Get order details if available
        const order = invoice.orderId;
        const crop = order?.programName || 'General';
        const acres = order?.totalAcres || 0;
        const year = order?.year || new Date().getFullYear();

        // Build professional invoice email HTML
        const emailHtml = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f5f5f0;">
    <div style="max-width: 700px; margin: 0 auto; background-color: #ffffff;">
        <!-- Header -->
        <div style="background: linear-gradient(135deg, #2d5a27 0%, #1e3d1a 100%); padding: 32px; text-align: center;">
            <h1 style="color: #ffffff; margin: 0; font-size: 28px; letter-spacing: 2px;">ACRE PROFIT</h1>
            <p style="color: rgba(255,255,255,0.9); margin: 8px 0 0 0; font-size: 14px;">Agricultural Chemical Solutions</p>
        </div>

        <!-- Invoice Title Bar -->
        <div style="background-color: #e8b923; padding: 16px 32px; display: flex; justify-content: space-between;">
            <div>
                <h2 style="margin: 0; color: #1a1a1a; font-size: 24px;">INVOICE</h2>
            </div>
            <div style="text-align: right;">
                <p style="margin: 0; font-size: 18px; font-weight: 700; color: #1a1a1a;">${invoice.invoiceNumber}</p>
            </div>
        </div>

        <!-- Content -->
        <div style="padding: 32px;">
            <!-- Info Bar -->
            <div style="background-color: #f5f5f0; border-radius: 12px; padding: 20px; margin-bottom: 28px;">
                <table width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                        <td style="text-align: center; padding: 8px;">
                            <p style="margin: 0; font-size: 12px; color: #666; text-transform: uppercase;">Invoice Date</p>
                            <p style="margin: 4px 0 0 0; font-weight: 600;">${invoiceDate}</p>
                        </td>
                        <td style="text-align: center; padding: 8px;">
                            <p style="margin: 0; font-size: 12px; color: #666; text-transform: uppercase;">Due Date</p>
                            <p style="margin: 4px 0 0 0; font-weight: 600;">${dueDate}</p>
                        </td>
                        <td style="text-align: center; padding: 8px;">
                            <p style="margin: 0; font-size: 12px; color: #666; text-transform: uppercase;">Order #</p>
                            <p style="margin: 4px 0 0 0; font-weight: 600;">${invoice.orderNumber || 'N/A'}</p>
                        </td>
                        <td style="text-align: center; padding: 8px;">
                            <p style="margin: 0; font-size: 12px; color: #666; text-transform: uppercase;">Status</p>
                            <p style="margin: 4px 0 0 0;">
                                <span style="background: ${invoice.paymentStatus === 'paid' ? '#d1fae5' : '#fee2e2'}; color: ${invoice.paymentStatus === 'paid' ? '#065f46' : '#991b1b'}; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 700;">
                                    ${invoice.paymentStatus === 'paid' ? 'PAID' : 'UNPAID'}
                                </span>
                            </p>
                        </td>
                    </tr>
                </table>
            </div>

            <!-- Addresses -->
            <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom: 28px;">
                <tr>
                    <td style="width: 48%; vertical-align: top;">
                        <p style="margin: 0 0 12px 0; font-size: 12px; text-transform: uppercase; color: #2d5a27; font-weight: 700; border-bottom: 2px solid #e8b923; padding-bottom: 8px;">Bill To</p>
                        <p style="margin: 0; font-weight: 700; font-size: 18px;">${invoice.customerName || invoice.customerId?.name || 'Customer'}</p>
                        <p style="margin: 4px 0; color: #666;">${invoice.customerFarm || ''}</p>
                        <p style="margin: 4px 0; color: #666;">${customerEmail}</p>
                        <p style="margin: 4px 0; color: #666;">${invoice.customerPhone || invoice.customerId?.phone || ''}</p>
                    </td>
                    <td style="width: 4%;"></td>
                    <td style="width: 48%; vertical-align: top;">
                        <p style="margin: 0 0 12px 0; font-size: 12px; text-transform: uppercase; color: #2d5a27; font-weight: 700; border-bottom: 2px solid #e8b923; padding-bottom: 8px;">From</p>
                        <p style="margin: 0; font-weight: 700; font-size: 18px;">Acre Profit LLC</p>
                        <p style="margin: 4px 0; color: #666;">Agricultural Chemical Distribution</p>
                        <p style="margin: 4px 0; color: #666;">Haxtun, CO</p>
                        <p style="margin: 4px 0; color: #666;">info@acreprofit.com</p>
                    </td>
                </tr>
            </table>

            <!-- Order Details -->
            <div style="margin-bottom: 28px;">
                <span style="display: inline-block; background: #e8f5e9; color: #2d5a27; padding: 8px 16px; border-radius: 20px; font-weight: 600; margin-right: 8px;">${crop}</span>
                <span style="display: inline-block; background: #e8f5e9; color: #2d5a27; padding: 8px 16px; border-radius: 20px; font-weight: 600; margin-right: 8px;">${acres.toLocaleString()} Acres</span>
                <span style="display: inline-block; background: #e8f5e9; color: #2d5a27; padding: 8px 16px; border-radius: 20px; font-weight: 600;">${year}</span>
            </div>

            <!-- Items Table -->
            <table width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; margin-bottom: 28px;">
                <thead>
                    <tr style="background-color: #2d5a27;">
                        <th style="padding: 14px 16px; text-align: left; color: white; font-size: 13px; text-transform: uppercase;">Product</th>
                        <th style="padding: 14px 16px; text-align: center; color: white; font-size: 13px; text-transform: uppercase;">Qty</th>
                        <th style="padding: 14px 16px; text-align: right; color: white; font-size: 13px; text-transform: uppercase;">Unit Price</th>
                        <th style="padding: 14px 16px; text-align: right; color: white; font-size: 13px; text-transform: uppercase;">Total</th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsHtml || '<tr><td colspan="4" style="padding: 20px; text-align: center; color: #666;">No items</td></tr>'}
                </tbody>
            </table>

            <!-- Totals -->
            <div style="display: flex; justify-content: flex-end;">
                <div style="width: 280px; background: #f5f5f0; border-radius: 12px; padding: 20px;">
                    <div style="display: flex; justify-content: space-between; padding: 8px 0;">
                        <span>Subtotal</span>
                        <span>$${subtotal.toFixed(2)}</span>
                    </div>
                    ${discount > 0 ? `
                    <div style="display: flex; justify-content: space-between; padding: 8px 0; color: #059669;">
                        <span>Discount</span>
                        <span>-$${discount.toFixed(2)}</span>
                    </div>
                    ` : ''}
                    <div style="display: flex; justify-content: space-between; padding: 12px 0; margin-top: 8px; border-top: 2px solid #2d5a27; font-size: 20px; font-weight: 700; color: #2d5a27;">
                        <span>Total Due</span>
                        <span>$${total.toFixed(2)}</span>
                    </div>
                </div>
            </div>

            <!-- Payment Info -->
            <div style="background: #fffbeb; border: 1px solid #fcd34d; border-radius: 12px; padding: 20px; margin: 28px 0;">
                <h4 style="margin: 0 0 8px 0; color: #92400e;">Payment Information</h4>
                <p style="margin: 4px 0; color: #78350f; font-size: 14px;">Please make payment via check or ACH transfer:</p>
                <div style="margin-top: 12px; padding: 12px; background: rgba(255,255,255,0.7); border-radius: 8px;">
                    <p style="margin: 0; font-weight: 700;">Acre Profit LLC</p>
                    <p style="margin: 4px 0; font-size: 14px; color: #666;">Contact your representative for ACH details or mail check to your pickup location.</p>
                </div>
            </div>

            <!-- Footer -->
            <div style="text-align: center; padding-top: 20px; border-top: 2px solid #e0e0e0;">
                <p style="margin: 4px 0; color: #2d5a27; font-weight: 600;">Thank you for your business!</p>
                <p style="margin: 4px 0; color: #666; font-size: 14px;">Questions? Contact us at info@acreprofit.com</p>
            </div>
        </div>

        <!-- Bottom Bar -->
        <div style="background-color: #2d5a27; padding: 16px; text-align: center;">
            <p style="margin: 0; color: rgba(255,255,255,0.8); font-size: 12px;">&copy; ${new Date().getFullYear()} Acre Profit LLC. All rights reserved.</p>
        </div>
    </div>
</body>
</html>`;

        // Send email
        const transporter = createEmailTransporter();
        if (transporter) {
            await transporter.sendMail({
                from: process.env.EMAIL_FROM || '"Acre Profit" <noreply@acreprofit.com>',
                to: customerEmail,
                subject: `Invoice ${invoice.invoiceNumber} from Acre Profit - $${total.toFixed(2)} Due`,
                html: emailHtml
            });

            invoice.status = 'sent';
            invoice.sentAt = new Date();
            invoice.sentBy = req.user._id;
            await invoice.save();

            res.json({ message: 'Invoice sent successfully', invoice });
        } else {
            // No email transporter configured, just update status
            invoice.status = 'sent';
            invoice.sentAt = new Date();
            invoice.sentBy = req.user._id;
            await invoice.save();

            res.json({
                message: 'Invoice marked as sent (email not configured)',
                invoice,
                warning: 'Email transporter not configured. Please set up SMTP or Gmail credentials.'
            });
        }
    } catch (error) {
        console.error('Error sending invoice:', error);
        res.status(400).json({ error: error.message });
    }
});

// ============ PROGRAM QUOTE ROUTES ============
// Create a program quote snapshot for a customer and send it via email.
// Distributors can only quote their own customers. Admins/superadmins can
// quote anyone. Quote is stored with status='sent' once the email leaves —
// draft is a transient internal state used only if transporter is null.
app.post('/api/admin/quotes', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const {
            customerId, programId, programName, crop, totalAcres,
            items, subtotal, total, costPerAcre, sprayParams, notes
        } = req.body;

        if (!customerId || !Array.isArray(items) || items.length === 0) {
            return res.status(400).json({ error: 'customerId and non-empty items[] required' });
        }

        const customer = await User.findById(customerId);
        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Distributors scoped to their own customers
        if (isDistributor(req.user) && String(customer.representative) !== String(req.user._id)) {
            return res.status(403).json({ error: 'Customer is not assigned to you' });
        }

        const quoteNumber = await generateQuoteNumber();
        const expiresAt = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000); // 30 days

        const quote = new ProgramQuote({
            quoteNumber,
            customerId: customer._id,
            createdBy: req.user._id,
            representativeId: customer.representative || req.user._id,
            programId: programId || null,
            programName: programName || '',
            crop: crop || '',
            totalAcres: totalAcres || 0,
            items: items.map(i => ({
                chemicalId: i.chemicalId || null,
                productName: i.productName || '',
                packSize: i.packSize || '',
                unit: i.unit || '',
                quantity: i.quantity || 0,
                unitPrice: i.unitPrice || 0,
                totalPrice: i.totalPrice || (i.quantity || 0) * (i.unitPrice || 0),
                acres: i.acres || 0,
                rate: i.rate || 0,
                rateUnit: i.rateUnit || '',
                calculatedAmount: i.calculatedAmount || 0
            })),
            subtotal: subtotal || 0,
            total: total || 0,
            costPerAcre: costPerAcre || 0,
            sprayParams: sprayParams || {},
            notes: notes || '',
            expiresAt,
            status: 'draft'
        });

        await quote.save();

        // Send the quote email. Same transporter pattern as password-reset
        // (server/index.js:4259) — null transporter fails silently and the
        // quote stays at status='draft' for manual resend.
        const transporter = createEmailTransporter();
        const loginUrl = `${process.env.FRONTEND_URL || 'https://acreprofit.com'}/login.html`;
        const expiresStr = expiresAt.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

        const itemsRows = quote.items.map(it => `
            <tr>
                <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb;">${it.productName}</td>
                <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb; text-align:center;">${it.rate || '—'} ${it.rateUnit || ''}</td>
                <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb; text-align:center;">${it.quantity} × ${it.packSize || ''}</td>
                <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb; text-align:right; font-weight:600;">$${(it.totalPrice || 0).toFixed(2)}</td>
            </tr>
        `).join('');

        if (transporter && customer.email) {
            try {
                await transporter.sendMail({
                    from: process.env.EMAIL_FROM || '"Acre Profit" <noreply@acreprofit.com>',
                    to: customer.email,
                    subject: `Program Quote ${quote.quoteNumber} — ${quote.programName || 'AcreProfit'}`,
                    html: `
                        <div style="font-family:Arial,sans-serif; max-width:640px; margin:0 auto;">
                            <div style="background:#2d5a27; padding:24px; text-align:center;">
                                <h1 style="color:white; margin:0; font-size:22px;">Acre Profit</h1>
                                <div style="color:rgba(255,255,255,0.85); font-size:13px; margin-top:4px;">Program Estimate</div>
                            </div>
                            <div style="padding:28px; background:#f9f9f9; color:#333; line-height:1.55;">
                                <p>Hi ${customer.name.split(' ')[0]},</p>
                                <p>Here's a program estimate for your <strong>${quote.crop || 'field'}</strong> acres based on standard rates. Prices are estimated and subject to change until your order is placed.</p>

                                <div style="background:white; border:1px solid #e5e7eb; border-radius:8px; padding:16px; margin:16px 0;">
                                    <div style="font-weight:700; color:#2d5a27; margin-bottom:4px;">${quote.programName || 'Custom program'}</div>
                                    <div style="color:#666; font-size:13px;">${quote.totalAcres} acres · Quote #${quote.quoteNumber}</div>
                                </div>

                                <table width="100%" style="border-collapse:collapse; background:white; border:1px solid #e5e7eb; border-radius:8px; overflow:hidden;">
                                    <thead>
                                        <tr style="background:#2d5a27; color:white;">
                                            <th style="padding:10px 12px; text-align:left;">Product</th>
                                            <th style="padding:10px 12px; text-align:center;">Rate / Acre</th>
                                            <th style="padding:10px 12px; text-align:center;">Qty</th>
                                            <th style="padding:10px 12px; text-align:right;">Estimated Cost</th>
                                        </tr>
                                    </thead>
                                    <tbody>${itemsRows}</tbody>
                                    <tfoot>
                                        <tr style="background:#f0f7ef; font-weight:700;">
                                            <td colspan="3" style="padding:12px; text-align:right;">Estimated Total</td>
                                            <td style="padding:12px; text-align:right; color:#2d5a27;">$${(quote.total || 0).toFixed(2)}</td>
                                        </tr>
                                        <tr style="background:#f0f7ef;">
                                            <td colspan="3" style="padding:8px 12px; text-align:right; color:#666; font-size:13px;">Estimated Cost Per Acre</td>
                                            <td style="padding:8px 12px; text-align:right; color:#2d5a27; font-weight:600;">$${(quote.costPerAcre || 0).toFixed(2)}</td>
                                        </tr>
                                    </tfoot>
                                </table>

                                <p style="margin-top:24px;">Want to adjust rates or quantities? <strong>Log in to AcreProfit to build your own order</strong> — you control the rates, we handle the sourcing.</p>

                                <div style="text-align:center; margin:28px 0;">
                                    <a href="${loginUrl}" style="background:#d4a017; color:white; padding:14px 28px; text-decoration:none; border-radius:6px; font-weight:bold; display:inline-block;">
                                        Log In & Build Your Order
                                    </a>
                                </div>

                                <p style="color:#666; font-size:13px;">Ready to order at these rates? Reply to this email or contact your rep.</p>
                                <hr style="border:none; border-top:1px solid #ddd; margin:24px 0;">
                                <p style="color:#999; font-size:12px;">Quote expires: ${expiresStr}. Prices subject to change.</p>
                            </div>
                        </div>
                    `
                });
                quote.status = 'sent';
                quote.sentAt = new Date();
                await quote.save();
            } catch (mailErr) {
                console.error('Quote email send failed:', mailErr.message);
                // Leave status='draft' so admin can resend from the customer card.
            }
        }

        res.json({ quote });
    } catch (error) {
        console.error('POST /api/admin/quotes error:', error);
        res.status(400).json({ error: error.message });
    }
});

// List program quotes for a customer. Distributors see only their own
// customers' quotes; admins/superadmins see all.
app.get('/api/admin/quotes/customer/:customerId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const customer = await User.findById(req.params.customerId);
        if (!customer) return res.status(404).json({ error: 'Customer not found' });

        if (isDistributor(req.user) && String(customer.representative) !== String(req.user._id)) {
            return res.status(403).json({ error: 'Customer is not assigned to you' });
        }

        const quotes = await ProgramQuote.find({ customerId: customer._id })
            .sort({ createdAt: -1 })
            .lean();
        res.json(quotes);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Convert a sent quote into an Invoice and trigger the existing invoice
// flow. Quote items snapshot is copied verbatim — invoice re-computes
// subtotal/total inline. Quote transitions to 'converted' with a pointer
// to the new invoice for audit.
app.post('/api/admin/quotes/:id/convert', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const quote = await ProgramQuote.findById(req.params.id);
        if (!quote) return res.status(404).json({ error: 'Quote not found' });
        if (quote.status === 'converted') {
            return res.status(400).json({ error: 'Quote already converted', convertedToInvoiceId: quote.convertedToInvoiceId });
        }

        const customer = await User.findById(quote.customerId);
        if (!customer) return res.status(404).json({ error: 'Customer not found' });

        if (isDistributor(req.user) && String(customer.representative) !== String(req.user._id)) {
            return res.status(403).json({ error: 'Customer is not assigned to you' });
        }

        const invoiceNumber = await generateInvoiceNumber();
        const items = quote.items.map(it => ({
            productName: it.productName,
            description: it.productName,
            quantity: it.quantity,
            unitPrice: it.unitPrice,
            totalPrice: (it.quantity || 0) * (it.unitPrice || 0)
        }));
        const subtotal = items.reduce((sum, it) => sum + (it.totalPrice || 0), 0);

        // Representative: superadmin may override via body, else quote rep,
        // else customer's rep, else the converter.
        let repId = quote.representativeId || customer.representative || req.user._id;
        if (req.user.role === 'superadmin' && req.body.representativeId) {
            repId = req.body.representativeId;
        }

        const invoice = new Invoice({
            invoiceNumber,
            customerId: customer._id,
            customerName: customer.name,
            customerEmail: customer.email,
            customerPhone: customer.phone,
            customerAddress: customer.address,
            representativeId: repId,
            items,
            subtotal,
            discount: 0,
            total: subtotal,
            notes: `Converted from quote ${quote.quoteNumber}`,
            dueDate: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000)
        });
        await invoice.save();

        quote.status = 'converted';
        quote.convertedAt = new Date();
        quote.convertedToInvoiceId = invoice._id;
        await quote.save();

        res.json({ invoice, quote });
    } catch (error) {
        console.error('POST /api/admin/quotes/:id/convert error:', error);
        res.status(400).json({ error: error.message });
    }
});

// Record payment on invoice
app.post('/api/admin/invoices/:id/payment', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { amount, method, notes } = req.body;

        const invoice = await Invoice.findById(req.params.id);
        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }

        const newAmountPaid = (invoice.amountPaid || 0) + amount;
        invoice.amountPaid = newAmountPaid;
        invoice.amountDue = invoice.total - newAmountPaid;
        invoice.paymentMethod = method;
        invoice.paymentDate = new Date();

        if (newAmountPaid >= invoice.total) {
            invoice.paymentStatus = 'paid';
            invoice.status = 'paid';
        } else {
            invoice.paymentStatus = 'partial';
        }

        if (notes) {
            invoice.notes = (invoice.notes || '') + '\n' + `Payment of $${amount} received via ${method}. ${notes}`;
        }

        invoice.updatedAt = new Date();
        await invoice.save();

        // Deduct inventory if order exists and status allows
        if (invoice.orderId) {
            const order = await ChemicalOrder.findById(invoice.orderId);
            if (order && order.status !== 'delivered') {
                // Optionally deduct inventory here when payment is received
                // This depends on business logic - might want to deduct on delivery instead
            }
        }

        res.json({ message: 'Payment recorded', invoice });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ DELIVERY CONFIRMATION ENDPOINTS ============

// Update delivery status with signature
app.post('/api/admin/invoices/:id/delivery', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { deliveryStatus, deliveryDate, deliveryLocation, deliverySignature, deliverySignedBy, deliveryNotes } = req.body;

        const invoice = await Invoice.findById(req.params.id);
        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }

        invoice.deliveryStatus = deliveryStatus;
        if (deliveryDate) invoice.deliveryDate = new Date(deliveryDate);
        if (deliveryLocation) invoice.deliveryLocation = deliveryLocation;
        if (deliverySignature) invoice.deliverySignature = deliverySignature;
        if (deliverySignedBy) invoice.deliverySignedBy = deliverySignedBy;
        if (deliveryStatus === 'signed') {
            invoice.deliverySignedAt = new Date();
        }
        if (deliveryNotes) invoice.deliveryNotes = deliveryNotes;

        invoice.updatedAt = new Date();
        await invoice.save();

        // If delivered and signed, deduct inventory
        if (deliveryStatus === 'signed' && invoice.orderId) {
            try {
                const order = await ChemicalOrder.findById(invoice.orderId);
                if (order) {
                    for (const item of order.items) {
                        if (item.chemicalId) {
                            await deductInventory({
                                chemicalId: item.chemicalId,
                                quantity: item.quantity,
                                location: 'main',
                                orderId: order._id,
                                orderNumber: order.orderNumber,
                                userId: req.user._id,
                                notes: `Delivered - Invoice ${invoice.invoiceNumber}`
                            });
                        }
                    }

                    // Update order status
                    order.status = 'delivered';
                    order.deliveredAt = new Date();
                    await order.save();
                }
            } catch (invError) {
                // Log but don't fail the delivery confirmation
                console.error('Inventory deduction error:', invError.message);
            }
        }

        res.json({ message: 'Delivery confirmed', invoice });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Customer-facing: Confirm delivery with signature
app.post('/api/orders/:id/confirm-delivery', authMiddleware, async (req, res) => {
    try {
        const { signature, signedBy } = req.body;

        const invoice = await Invoice.findOne({ orderId: req.params.id, customerId: req.user._id });
        if (!invoice) {
            return res.status(404).json({ error: 'Invoice not found' });
        }

        invoice.deliveryStatus = 'signed';
        invoice.deliverySignature = signature;
        invoice.deliverySignedBy = signedBy || req.user.name;
        invoice.deliverySignedAt = new Date();
        invoice.updatedAt = new Date();

        await invoice.save();

        res.json({ message: 'Delivery confirmed', invoice });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ SHAREPOINT/EXCEL PRICE SYNC ============

// Price Sync Log Model - Track sync history
const priceSyncLogSchema = new mongoose.Schema({
    syncDate: { type: Date, default: Date.now },
    source: { type: String, default: 'sharepoint' },
    fileName: String,
    productsUpdated: { type: Number, default: 0 },
    productsAdded: { type: Number, default: 0 },
    productsSkipped: { type: Number, default: 0 },
    errors: [String],
    status: { type: String, enum: ['success', 'partial', 'failed'], default: 'success' },
    syncedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    details: mongoose.Schema.Types.Mixed
});

const PriceSyncLog = mongoose.model('PriceSyncLog', priceSyncLogSchema);

// Initialize Microsoft Graph client
let graphClient = null;

function initGraphClient() {
    if (!Client || !ClientSecretCredential) {
        console.log('Microsoft Graph SDK not available');
        return null;
    }

    const clientId = process.env.MICROSOFT_CLIENT_ID;
    const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
    const tenantId = process.env.MICROSOFT_TENANT_ID;

    if (!clientId || !clientSecret || !tenantId) {
        console.log('Microsoft Graph credentials not configured');
        return null;
    }

    try {
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

        graphClient = Client.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => {
                    const token = await credential.getToken(['https://graph.microsoft.com/.default']);
                    return token.token;
                }
            }
        });

        console.log('Microsoft Graph client initialized');
        return graphClient;
    } catch (error) {
        console.error('Failed to initialize Graph client:', error.message);
        return null;
    }
}

// Parse Excel data and update prices
async function syncPricesFromExcel(fileBuffer, userId) {
    if (!XLSX) {
        throw new Error('XLSX library not available');
    }

    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    const log = {
        productsUpdated: 0,
        productsAdded: 0,
        productsSkipped: 0,
        errors: [],
        details: []
    };

    for (const row of data) {
        try {
            // Expected columns: Product Name, Pack Size, Cost Price, Sell Price
            // Adjust column names based on actual spreadsheet
            const productName = row['Product Name'] || row['Product'] || row['Name'];
            const packSize = row['Pack Size'] || row['Size'] || row['Package'];
            const costPrice = parseFloat(row['Cost Price'] || row['Cost'] || row['Wholesale'] || 0);
            const sellPrice = parseFloat(row['Sell Price'] || row['Price'] || row['Retail'] || 0);
            const supplier = row['Supplier'] || row['Source'] || 'JABCO';

            if (!productName) {
                log.productsSkipped++;
                continue;
            }

            // Find existing product
            let chemical = await Chemical.findOne({
                productName: { $regex: new RegExp(`^${productName.trim()}$`, 'i') },
                packSize: packSize ? { $regex: new RegExp(packSize.trim(), 'i') } : undefined
            });

            if (chemical) {
                // Update existing
                const oldCost = chemical.costPrice;
                const oldSell = chemical.sellPrice;

                if (costPrice > 0) chemical.costPrice = costPrice;
                if (sellPrice > 0) chemical.sellPrice = sellPrice;
                chemical.updatedAt = new Date();

                // Record price history
                if (costPrice !== oldCost || sellPrice !== oldSell) {
                    const history = new ChemicalPriceHistory({
                        chemicalId: chemical._id,
                        productName: chemical.productName,
                        sourceSupplier: chemical.sourceSupplier,
                        packSize: chemical.packSize,
                        costPrice,
                        sellPrice,
                        priceVersion: 'sync-' + new Date().toISOString().slice(0, 10),
                        changedBy: userId
                    });
                    await history.save();
                }

                await chemical.save();
                log.productsUpdated++;
                log.details.push({ product: productName, action: 'updated', oldCost, newCost: costPrice, oldSell, newSell: sellPrice });
            } else if (costPrice > 0 || sellPrice > 0) {
                // Create new product
                chemical = new Chemical({
                    productName: productName.trim(),
                    sourceSupplier: supplier,
                    packSize: packSize || 'Unknown',
                    unit: 'gl',
                    costPrice: costPrice || 0,
                    sellPrice: sellPrice || costPrice * 1.15, // 15% default markup
                    category: 'herbicide',
                    isActive: true,
                    availableForOrder: true,
                    createdBy: userId
                });
                await chemical.save();
                log.productsAdded++;
                log.details.push({ product: productName, action: 'added', cost: costPrice, sell: sellPrice });
            } else {
                log.productsSkipped++;
            }
        } catch (error) {
            log.errors.push(`Row error: ${error.message}`);
            log.productsSkipped++;
        }
    }

    return log;
}

// Fetch file from SharePoint and sync prices
async function syncFromSharePoint(userId) {
    if (!graphClient) {
        graphClient = initGraphClient();
    }

    if (!graphClient) {
        throw new Error('Microsoft Graph client not configured. Check MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, and MICROSOFT_TENANT_ID.');
    }

    const siteId = process.env.SHAREPOINT_SITE_ID;
    const filePath = process.env.SHAREPOINT_FILE_PATH;

    if (!siteId || !filePath) {
        throw new Error('SharePoint site ID or file path not configured');
    }

    try {
        // Get the file content
        // For personal OneDrive: /users/{user-id}/drive/root:/{path}:/content
        // For SharePoint: /sites/{site-id}/drive/root:/{path}:/content

        let fileBuffer;

        // Try OneDrive personal path first
        try {
            const response = await graphClient
                .api(`/users/jeff@cropprotectdirect.com/drive/root:${filePath}:/content`)
                .get();
            fileBuffer = Buffer.from(response);
        } catch (e) {
            // Try SharePoint site path
            const response = await graphClient
                .api(`/sites/${siteId}/drive/root:${filePath}:/content`)
                .get();
            fileBuffer = Buffer.from(response);
        }

        const result = await syncPricesFromExcel(fileBuffer, userId);

        // Save sync log
        const syncLog = new PriceSyncLog({
            source: 'sharepoint',
            fileName: filePath.split('/').pop(),
            ...result,
            status: result.errors.length === 0 ? 'success' : (result.productsUpdated > 0 ? 'partial' : 'failed'),
            syncedBy: userId
        });
        await syncLog.save();

        return { success: true, ...result };
    } catch (error) {
        console.error('SharePoint sync error:', error);

        // Save failed sync log
        const syncLog = new PriceSyncLog({
            source: 'sharepoint',
            status: 'failed',
            errors: [error.message],
            syncedBy: userId
        });
        await syncLog.save();

        throw error;
    }
}

// API Endpoints for price sync

// Manual sync trigger (admin only)
app.post('/api/admin/sync-prices/sharepoint', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can trigger price sync' });
        }

        const result = await syncFromSharePoint(req.user._id);
        res.json({
            message: 'Price sync completed',
            ...result
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Upload Excel file manually for sync
app.post('/api/admin/sync-prices/upload', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can upload price files' });
        }

        // Expect base64 encoded file in request body
        const { fileData, fileName } = req.body;

        if (!fileData) {
            return res.status(400).json({ error: 'No file data provided' });
        }

        const fileBuffer = Buffer.from(fileData, 'base64');
        const result = await syncPricesFromExcel(fileBuffer, req.user._id);

        // Save sync log
        const syncLog = new PriceSyncLog({
            source: 'upload',
            fileName: fileName || 'uploaded-file.xlsx',
            ...result,
            status: result.errors.length === 0 ? 'success' : 'partial',
            syncedBy: req.user._id
        });
        await syncLog.save();

        res.json({
            message: 'Price sync from upload completed',
            ...result
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get sync history
app.get('/api/admin/sync-prices/history', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const logs = await PriceSyncLog.find()
            .populate('syncedBy', 'name email')
            .sort({ syncDate: -1 })
            .limit(50);

        res.json(logs);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get sync status/config
app.get('/api/admin/sync-prices/status', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const lastSync = await PriceSyncLog.findOne().sort({ syncDate: -1 });

        const configured = !!(
            process.env.MICROSOFT_CLIENT_ID &&
            process.env.MICROSOFT_CLIENT_SECRET &&
            process.env.MICROSOFT_TENANT_ID
        );

        res.json({
            configured,
            lastSync: lastSync ? {
                date: lastSync.syncDate,
                status: lastSync.status,
                productsUpdated: lastSync.productsUpdated,
                productsAdded: lastSync.productsAdded
            } : null,
            cronSchedule: process.env.PRICE_SYNC_CRON || '0 6 */3 * *'
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ PURCHASE ORDER DOCUMENT UPLOADS ============

// Configure multer for file uploads
// S7: UPLOAD_DIR points at a Render persistent disk mount in production
// (e.g. /var/data/purchase-orders). Falls back to a repo-root relative path
// for local dev. License uploads go to a sibling dir - route not wired yet,
// but the directory is pre-created so the future upload endpoint just drops in.
const poDocumentsPath = process.env.UPLOAD_DIR || path.join(__dirname, '..', 'purchase-orders');
const licensesPath = process.env.LICENSES_UPLOAD_DIR || path.join(path.dirname(poDocumentsPath), 'licenses');

// Ensure upload directories exist
for (const dir of [poDocumentsPath, licensesPath]) {
    try {
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    } catch (e) {
        console.error(`Failed to create upload dir ${dir}: ${e.message}`);
    }
}

let uploadPO;
if (multer) {
    const poStorage = multer.diskStorage({
        destination: (req, file, cb) => {
            cb(null, poDocumentsPath);
        },
        filename: (req, file, cb) => {
            // Generate filename: PO-2026-00001-invoice-timestamp.pdf
            const ext = path.extname(file.originalname);
            const poNumber = req.params.poNumber || 'unknown';
            const docType = req.body.documentType || 'document';
            const timestamp = Date.now();
            cb(null, `${poNumber}-${docType}-${timestamp}${ext}`);
        }
    });

    uploadPO = multer({
        storage: poStorage,
        limits: { fileSize: 10 * 1024 * 1024 }, // 10MB max
        fileFilter: (req, file, cb) => {
            const allowedTypes = [
                'application/pdf',
                'image/jpeg',
                'image/png',
                'image/gif',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // xlsx
                'application/vnd.ms-excel' // xls
            ];
            if (allowedTypes.includes(file.mimetype)) {
                cb(null, true);
            } else {
                cb(new Error('Invalid file type. Allowed: PDF, JPEG, PNG, GIF, Excel'));
            }
        }
    });
}

// Upload document to a purchase order
app.post('/api/admin/purchase-orders/:id/documents', authMiddleware, adminMiddleware, (req, res, next) => {
    if (!uploadPO) {
        return res.status(500).json({ error: 'File upload not configured' });
    }
    next();
}, async (req, res) => {
    try {
        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        // Set poNumber for filename generation
        req.params.poNumber = po.poNumber;

        // Handle file upload
        uploadPO.single('document')(req, res, async (err) => {
            if (err) {
                return res.status(400).json({ error: err.message });
            }

            if (!req.file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }

            // Add document to PO
            const doc = {
                fileName: req.file.originalname,
                storedName: req.file.filename,
                fileType: req.file.mimetype,
                fileSize: req.file.size,
                documentType: req.body.documentType || 'invoice',
                uploadedBy: req.user._id,
                uploadedAt: new Date(),
                notes: req.body.notes || ''
            };

            po.documents.push(doc);
            po.updatedAt = new Date();
            po.updatedBy = req.user._id;
            await po.save();

            res.json({
                message: 'Document uploaded successfully',
                document: doc,
                purchaseOrder: po
            });
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get/download document from purchase order
app.get('/api/admin/purchase-orders/:id/documents/:docId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        const doc = po.documents.id(req.params.docId);
        if (!doc) {
            return res.status(404).json({ error: 'Document not found' });
        }

        const filePath = path.join(poDocumentsPath, doc.storedName);

        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ error: 'File not found on server' });
        }

        res.setHeader('Content-Type', doc.fileType);
        res.setHeader('Content-Disposition', `inline; filename="${doc.fileName}"`);
        res.sendFile(filePath);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Download document (force download)
app.get('/api/admin/purchase-orders/:id/documents/:docId/download', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        const doc = po.documents.id(req.params.docId);
        if (!doc) {
            return res.status(404).json({ error: 'Document not found' });
        }

        const filePath = path.join(poDocumentsPath, doc.storedName);

        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ error: 'File not found on server' });
        }

        res.download(filePath, doc.fileName);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete document from purchase order
app.delete('/api/admin/purchase-orders/:id/documents/:docId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Only superadmin can delete documents
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can delete documents' });
        }

        const po = await PurchaseOrder.findById(req.params.id);
        if (!po) {
            return res.status(404).json({ error: 'Purchase order not found' });
        }

        const doc = po.documents.id(req.params.docId);
        if (!doc) {
            return res.status(404).json({ error: 'Document not found' });
        }

        // Delete file from disk
        const filePath = path.join(poDocumentsPath, doc.storedName);
        if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
        }

        // Remove from PO
        po.documents.pull(req.params.docId);
        po.updatedAt = new Date();
        po.updatedBy = req.user._id;
        await po.save();

        res.json({ message: 'Document deleted', purchaseOrder: po });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// List all PO documents (superadmin only - for browsing the folder)
app.get('/api/admin/purchase-order-documents', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can view all documents' });
        }

        // Get all POs with documents
        const pos = await PurchaseOrder.find({ 'documents.0': { $exists: true } })
            .select('poNumber supplier.name documents status orderDate')
            .sort({ orderDate: -1 });

        const result = pos.map(po => ({
            poId: po._id,
            poNumber: po.poNumber,
            supplier: po.supplier.name,
            status: po.status,
            orderDate: po.orderDate,
            documents: po.documents.map(d => ({
                _id: d._id,
                fileName: d.fileName,
                documentType: d.documentType,
                fileSize: d.fileSize,
                uploadedAt: d.uploadedAt
            }))
        }));

        res.json(result);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Serve static files from purchase-orders folder (authenticated)
app.use('/purchase-orders', authMiddleware, adminMiddleware, express.static(poDocumentsPath));

// ============ DISTRIBUTOR PRICING ENDPOINTS ============

// Get products for distributor pricing (hides wholesale prices)
app.get('/api/distributor/products', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        // Only distributors and admins can access
        if (!['distributor', 'admin', 'superadmin'].includes(req.user.role)) {
            return res.status(403).json({ error: 'Access denied' });
        }

        const chemicals = await Chemical.find({ isActive: true })
            .select('productName sourceSupplier category packSize unit costPrice adminPrice sellPrice adminMarginDollars marginDollars')
            .sort({ productName: 1 });

        // Get this distributor's custom pricing
        const distributorPricing = await DistributorPricing.find({ distributorId: req.user._id });
        const pricingMap = {};
        distributorPricing.forEach(p => {
            pricingMap[p.chemicalId.toString()] = p;
        });

        // Build response
        const products = chemicals.map(c => {
            const customPrice = pricingMap[c._id.toString()];
            return {
                _id: c._id,
                productName: c.productName,
                sourceSupplier: c.sourceSupplier,
                category: c.category,
                packSize: c.packSize,
                unit: c.unit,
                costPrice: c.costPrice,
                adminPrice: c.adminPrice,
                sellPrice: c.sellPrice,
                price: customPrice?.retailPrice || c.sellPrice,
                isAvailable: customPrice?.isAvailable !== false,
                notes: customPrice?.notes || ''
            };
        });

        res.json(products);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Set distributor's retail price for a product
app.put('/api/distributor/products/:chemicalId/price', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (!['distributor', 'admin', 'superadmin'].includes(req.user.role)) {
            return res.status(403).json({ error: 'Access denied' });
        }

        const { retailPrice, isAvailable, notes } = req.body;

        if (retailPrice === undefined || retailPrice < 0) {
            return res.status(400).json({ error: 'Valid retail price required' });
        }

        // Find or create distributor pricing
        let pricing = await DistributorPricing.findOne({
            distributorId: req.user._id,
            chemicalId: req.params.chemicalId
        });

        if (pricing) {
            pricing.retailPrice = retailPrice;
            if (isAvailable !== undefined) pricing.isAvailable = isAvailable;
            if (notes !== undefined) pricing.notes = notes;
            pricing.updatedAt = new Date();
            pricing.updatedBy = req.user._id;
        } else {
            pricing = new DistributorPricing({
                distributorId: req.user._id,
                chemicalId: req.params.chemicalId,
                retailPrice,
                isAvailable: isAvailable !== false,
                notes: notes || '',
                updatedBy: req.user._id
            });
        }

        await pricing.save();

        res.json({ message: 'Price updated', pricing });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Bulk update distributor prices
app.put('/api/distributor/products/bulk-price', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (!['distributor', 'admin', 'superadmin'].includes(req.user.role)) {
            return res.status(403).json({ error: 'Access denied' });
        }

        const { prices } = req.body;
        // prices: [{ chemicalId, retailPrice }, ...]

        if (!Array.isArray(prices)) {
            return res.status(400).json({ error: 'prices array required' });
        }

        let updated = 0;
        for (const item of prices) {
            if (!item.chemicalId || item.retailPrice === undefined) continue;

            await DistributorPricing.findOneAndUpdate(
                { distributorId: req.user._id, chemicalId: item.chemicalId },
                {
                    retailPrice: item.retailPrice,
                    updatedAt: new Date(),
                    updatedBy: req.user._id
                },
                { upsert: true }
            );
            updated++;
        }

        res.json({ message: `${updated} prices updated` });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all distributor pricing (superadmin only - to view all distributors' prices)
app.get('/api/admin/distributor-pricing', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        if (req.user.role !== 'superadmin') {
            return res.status(403).json({ error: 'Only superadmin can view all distributor pricing' });
        }

        const pricing = await DistributorPricing.find()
            .populate('distributorId', 'name email')
            .populate('chemicalId', 'productName packSize sellPrice costPrice')
            .sort({ 'distributorId': 1 });

        res.json(pricing);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ QUOTE REQUEST / SUBMIT QUANTITY FOR BID ============

// Submit a quote request (customer or admin on behalf of customer)
app.post('/api/quote-requests', authMiddleware, async (req, res) => {
    try {
        const { items, customerNotes, deliveryLocation, preferredDeliveryDate, customerId, crop, timing, acres, gallonsPerAcre } = req.body;

        if (!items || items.length === 0) {
            return res.status(400).json({ error: 'At least one item is required' });
        }

        // Determine customer
        let targetCustomerId = req.user._id;
        let customerName = req.user.name;
        let customerEmail = req.user.email;
        let customerPhone = req.user.phone;

        // Admin can submit on behalf of a customer
        if (customerId && (req.user.role === 'admin' || req.user.role === 'superadmin' || req.user.role === 'distributor')) {
            const customer = await User.findById(customerId);
            if (customer) {
                targetCustomerId = customer._id;
                customerName = customer.name;
                customerEmail = customer.email;
                customerPhone = customer.phone;
            }
        }

        const quoteRequest = new QuoteRequest({
            customerId: targetCustomerId,
            customerName,
            customerEmail,
            customerPhone,
            representativeId: req.user.role === 'customer' ? req.user.representativeId : req.user._id,
            crop,
            timing,
            acres,
            gallonsPerAcre: gallonsPerAcre || 15,
            items: items.map(item => ({
                productName: item.productName,
                chemicalId: item.chemicalId,
                category: item.category,
                packSize: item.packSize,
                unit: item.unit,
                quantityNeeded: item.quantityNeeded,
                isPriced: false
            })),
            customerNotes,
            deliveryLocation,
            preferredDeliveryDate,
            status: 'submitted',
            createdBy: req.user._id
        });

        await quoteRequest.save();

        res.status(201).json(quoteRequest);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get quote requests (customer sees their own, admin sees all)
app.get('/api/quote-requests', authMiddleware, async (req, res) => {
    try {
        let query = {};

        if (req.user.role === 'customer') {
            query.customerId = req.user._id;
        } else if (isDistributor(req.user)) {
            // Distributors see quotes from their customers
            query.representativeId = req.user._id;
        }
        // Superadmin sees all

        const quotes = await QuoteRequest.find(query)
            .populate('customerId', 'name email')
            .populate('representativeId', 'name')
            .sort({ submittedAt: -1 });

        res.json(quotes);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single quote request
app.get('/api/quote-requests/:id', authMiddleware, async (req, res) => {
    try {
        const quote = await QuoteRequest.findById(req.params.id)
            .populate('customerId', 'name email phone')
            .populate('representativeId', 'name')
            .populate('items.chemicalId', 'productName packSize unit costPrice sellPrice');

        if (!quote) {
            return res.status(404).json({ error: 'Quote request not found' });
        }

        // Check access
        if (req.user.role === 'customer' && quote.customerId._id.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Access denied' });
        }

        res.json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Add pricing to quote request items
app.put('/api/admin/quote-requests/:id/price', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { items, adminNotes, expiresAt } = req.body;

        const quote = await QuoteRequest.findById(req.params.id);
        if (!quote) {
            return res.status(404).json({ error: 'Quote request not found' });
        }

        // Update pricing for each item
        let estimatedTotal = 0;
        for (const updatedItem of items) {
            const item = quote.items.id(updatedItem._id);
            if (item) {
                item.costPrice = updatedItem.costPrice;
                item.adminPrice = updatedItem.adminPrice;
                item.sellPrice = updatedItem.sellPrice;
                item.totalPrice = (updatedItem.sellPrice || 0) * item.quantityNeeded;
                item.priceNotes = updatedItem.priceNotes;
                item.isPriced = updatedItem.sellPrice > 0;
                estimatedTotal += item.totalPrice;
            }
        }

        quote.estimatedTotal = estimatedTotal;
        quote.status = 'pricing';
        quote.pricedAt = new Date();
        quote.adminNotes = adminNotes;
        quote.updatedBy = req.user._id;
        quote.updatedAt = new Date();

        if (expiresAt) {
            quote.expiresAt = new Date(expiresAt);
        }

        await quote.save();

        res.json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Send quote to customer
app.put('/api/admin/quote-requests/:id/send', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const quote = await QuoteRequest.findById(req.params.id);
        if (!quote) {
            return res.status(404).json({ error: 'Quote request not found' });
        }

        // Ensure all items are priced
        const unpricedItems = quote.items.filter(i => !i.isPriced);
        if (unpricedItems.length > 0) {
            return res.status(400).json({ error: `${unpricedItems.length} item(s) still need pricing` });
        }

        quote.status = 'quoted';
        quote.quotedAt = new Date();
        quote.expiresAt = quote.expiresAt || new Date(Date.now() + 7 * 24 * 60 * 60 * 1000); // 7 days default
        quote.updatedBy = req.user._id;
        quote.updatedAt = new Date();

        await quote.save();

        // TODO: Send email notification to customer

        res.json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Customer: Accept or decline quote
app.put('/api/quote-requests/:id/respond', authMiddleware, async (req, res) => {
    try {
        const { response } = req.body; // 'accept' or 'decline'

        const quote = await QuoteRequest.findById(req.params.id);
        if (!quote) {
            return res.status(404).json({ error: 'Quote request not found' });
        }

        // Verify customer owns this quote
        if (req.user.role === 'customer' && quote.customerId.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Access denied' });
        }

        if (quote.status !== 'quoted') {
            return res.status(400).json({ error: 'Quote is not in a state that can be responded to' });
        }

        // Check expiration
        if (quote.expiresAt && new Date() > quote.expiresAt) {
            quote.status = 'expired';
            await quote.save();
            return res.status(400).json({ error: 'Quote has expired' });
        }

        if (response === 'accept') {
            quote.status = 'accepted';
            quote.respondedAt = new Date();
            // Quote can now be converted to an order
        } else if (response === 'decline') {
            quote.status = 'declined';
            quote.respondedAt = new Date();
        } else {
            return res.status(400).json({ error: 'Invalid response. Use "accept" or "decline"' });
        }

        quote.updatedAt = new Date();
        await quote.save();

        res.json(quote);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Convert accepted quote to order
app.post('/api/admin/quote-requests/:id/convert-to-order', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const quote = await QuoteRequest.findById(req.params.id)
            .populate('customerId');

        if (!quote) {
            return res.status(404).json({ error: 'Quote request not found' });
        }

        if (quote.status !== 'accepted') {
            return res.status(400).json({ error: 'Quote must be accepted before converting to order' });
        }

        // Create order from quote
        const order = new Order({
            userId: quote.customerId._id,
            representativeId: quote.representativeId,
            crop: 'Quote Order',
            acres: 0,
            chemicals: quote.items.map(item => ({
                name: item.productName,
                chemicalId: item.chemicalId,
                qty: item.quantityNeeded,
                packSize: item.packSize,
                unit: item.unit,
                pricePerUnit: item.sellPrice,
                totalPrice: item.totalPrice,
                sourceSupplier: 'Quote'
            })),
            totalCost: quote.estimatedTotal,
            status: 'payment_pending',
            paymentStatus: 'pending',
            deliveryAddress: { notes: quote.deliveryLocation },
            notes: `Converted from Quote ${quote.quoteNumber}`
        });

        await order.save();

        // Update quote
        quote.status = 'converted';
        quote.convertedToOrderId = order._id;
        quote.convertedAt = new Date();
        quote.updatedBy = req.user._id;
        quote.updatedAt = new Date();
        await quote.save();

        res.json({ quote, order });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Get all quote requests with filters
app.get('/api/admin/quote-requests', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, repId } = req.query;

        let query = {};
        if (status) query.status = status;

        if (isDistributor(req.user)) {
            query.representativeId = req.user._id;
        } else if (repId) {
            query.representativeId = repId;
        }

        const quotes = await QuoteRequest.find(query)
            .populate('customerId', 'name email')
            .populate('representativeId', 'name')
            .sort({ submittedAt: -1 });

        res.json(quotes);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ CHEMICAL VOLUME AGGREGATOR ============

// Get aggregated volume needs across all pending orders and quotes
app.get('/api/admin/chemical-volume-needs', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { includeOrders = true, includeQuotes = true } = req.query;

        const volumeNeeds = {};

        // Aggregate from pending/confirmed orders
        if (includeOrders === 'true' || includeOrders === true) {
            const orderStatuses = ['pending', 'payment_pending', 'payment_secured', 'manufacturer_ordered'];
            let orderQuery = { status: { $in: orderStatuses } };

            if (isDistributor(req.user)) {
                orderQuery.representativeId = req.user._id;
            }

            const orders = await Order.find(orderQuery).select('chemicals');

            for (const order of orders) {
                for (const chem of order.chemicals || []) {
                    const key = `${chem.name || chem.productName}|${chem.packSize || 'N/A'}|${chem.unit || 'unit'}`;
                    if (!volumeNeeds[key]) {
                        volumeNeeds[key] = {
                            productName: chem.name || chem.productName,
                            packSize: chem.packSize || 'N/A',
                            unit: chem.unit || 'unit',
                            chemicalId: chem.chemicalId,
                            totalQuantityOrders: 0,
                            totalQuantityQuotes: 0,
                            orderCount: 0,
                            quoteCount: 0
                        };
                    }
                    volumeNeeds[key].totalQuantityOrders += (chem.qty || chem.quantity || chem.packagesNeeded || 0);
                    volumeNeeds[key].orderCount++;
                }
            }
        }

        // Aggregate from submitted/pricing quotes
        if (includeQuotes === 'true' || includeQuotes === true) {
            const quoteStatuses = ['submitted', 'pricing', 'quoted', 'accepted'];
            let quoteQuery = { status: { $in: quoteStatuses } };

            if (isDistributor(req.user)) {
                quoteQuery.representativeId = req.user._id;
            }

            const quotes = await QuoteRequest.find(quoteQuery).select('items');

            for (const quote of quotes) {
                for (const item of quote.items || []) {
                    const key = `${item.productName}|${item.packSize || 'N/A'}|${item.unit || 'unit'}`;
                    if (!volumeNeeds[key]) {
                        volumeNeeds[key] = {
                            productName: item.productName,
                            packSize: item.packSize || 'N/A',
                            unit: item.unit || 'unit',
                            chemicalId: item.chemicalId,
                            totalQuantityOrders: 0,
                            totalQuantityQuotes: 0,
                            orderCount: 0,
                            quoteCount: 0
                        };
                    }
                    volumeNeeds[key].totalQuantityQuotes += (item.quantityNeeded || 0);
                    volumeNeeds[key].quoteCount++;
                }
            }
        }

        // Convert to array and add totals
        const result = Object.values(volumeNeeds).map(item => ({
            ...item,
            totalQuantity: item.totalQuantityOrders + item.totalQuantityQuotes,
            sources: `${item.orderCount} orders, ${item.quoteCount} quotes`
        }));

        // Sort by total quantity descending
        result.sort((a, b) => b.totalQuantity - a.totalQuantity);

        res.json(result);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ SUPPLIER BID SHEETS ============

// Create a bid sheet from volume needs
app.post('/api/admin/bid-sheets', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { title, description, items, invitedSuppliers, responseDueDate } = req.body;

        if (!items || items.length === 0) {
            return res.status(400).json({ error: 'At least one item is required' });
        }

        const bidSheet = new SupplierBidSheet({
            title: title || `Bid Request - ${new Date().toLocaleDateString()}`,
            description,
            items: items.map(item => ({
                productName: item.productName,
                chemicalId: item.chemicalId,
                category: item.category,
                packSize: item.packSize,
                unit: item.unit,
                quantityNeeded: item.quantityNeeded,
                notes: item.notes
            })),
            invitedSuppliers: (invitedSuppliers || []).map(s => ({
                supplierId: s.supplierId,
                supplierName: s.supplierName,
                contactEmail: s.contactEmail,
                contactPhone: s.contactPhone,
                status: 'invited'
            })),
            responseDueDate: responseDueDate ? new Date(responseDueDate) : new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
            status: 'draft',
            createdBy: req.user._id
        });

        await bidSheet.save();
        res.status(201).json(bidSheet);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create bid sheet from aggregated volume needs
app.post('/api/admin/bid-sheets/from-volume-needs', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { title, description, invitedSuppliers, responseDueDate } = req.body;

        // Default: aggregate only requests needed within the next 14 days.
        // Override with ?withinDays=N (0 or 'all' disables the date filter).
        const rawWithin = req.query.withinDays;
        const withinDays = (rawWithin === 'all' || rawWithin === '0')
            ? null
            : (rawWithin !== undefined ? parseInt(rawWithin, 10) : 14);
        const dateHorizon = withinDays && withinDays > 0
            ? new Date(Date.now() + withinDays * 24 * 60 * 60 * 1000)
            : null;

        // Product-name normalization: fold free-text customer requests onto
        // canonical catalog entries. Case-insensitive match; fallback to raw.
        const catalog = await Chemical.find({ isActive: true })
            .select('productName packSize unit')
            .lean();
        const catalogByName = {};
        catalog.forEach(c => {
            if (c.productName) catalogByName[c.productName.toLowerCase().trim()] = c;
        });
        const normalize = (rawName) => {
            if (!rawName) return null;
            const match = catalogByName[String(rawName).toLowerCase().trim()];
            if (match) return { productName: match.productName, packSize: match.packSize, unit: match.unit };
            return { productName: String(rawName).trim(), packSize: null, unit: null };
        };

        const volumeNeeds = {};
        const addToBucket = (rawName, packSize, unit, chemicalId, quantity) => {
            if (!rawName || !(quantity > 0)) return;
            const norm = normalize(rawName);
            if (!norm) return;
            const finalPackSize = norm.packSize || packSize || 'N/A';
            const finalUnit = norm.unit || unit || 'unit';
            const key = `${norm.productName}|${finalPackSize}|${finalUnit}`;
            if (!volumeNeeds[key]) {
                volumeNeeds[key] = {
                    productName: norm.productName,
                    packSize: finalPackSize,
                    unit: finalUnit,
                    chemicalId: chemicalId || null,
                    quantityNeeded: 0
                };
            }
            volumeNeeds[key].quantityNeeded += quantity;
            if (!volumeNeeds[key].chemicalId && chemicalId) {
                volumeNeeds[key].chemicalId = chemicalId;
            }
        };

        // Source 1: ChemicalOrder (modern customer-placed orders)
        const chemicalOrderStatuses = ['submitted', 'confirmed', 'ordered_from_supplier', 'received', 'ready_for_pickup', 'payment_pending', 'payment_secured'];
        const chemOrders = await ChemicalOrder.find({ status: { $in: chemicalOrderStatuses } })
            .select('items')
            .lean();
        for (const order of chemOrders) {
            for (const item of order.items || []) {
                addToBucket(item.productName, item.packSize, item.unit, item.chemicalId, item.quantity || 0);
            }
        }

        // Source 2: Legacy Order (admin-created via for-customer route).
        // Keep-alive until the pricing refactor retires this write path.
        const legacyOrderStatuses = ['pending', 'payment_pending', 'payment_secured', 'manufacturer_ordered'];
        const legacyOrders = await Order.find({ status: { $in: legacyOrderStatuses } })
            .select('chemicals')
            .lean();
        for (const order of legacyOrders) {
            for (const chem of order.chemicals || []) {
                const qty = chem.qty || chem.quantity || chem.packagesNeeded || 0;
                addToBucket(chem.name || chem.productName, chem.packSize, chem.unit, chem.chemicalId, qty);
            }
        }

        // Source 3: PriceMiningQuote (customer 2-week-out requests, not yet orders).
        // Filtered by neededBy when present; records without neededBy fall through
        // (can't filter them without losing legacy data).
        const pmqQuery = { status: 'open' };
        if (dateHorizon) {
            pmqQuery.$or = [
                { neededBy: { $lte: dateHorizon } },
                { neededBy: null },
                { neededBy: { $exists: false } }
            ];
        }
        const pmQuotes = await PriceMiningQuote.find(pmqQuery)
            .select('lines product supplier')
            .lean();
        for (const quote of pmQuotes) {
            if (Array.isArray(quote.lines) && quote.lines.length > 0) {
                for (const line of quote.lines) {
                    const qty = parseFloat(line.quantity) || 0;
                    addToBucket(line.product, null, null, null, qty);
                }
            } else if (quote.product) {
                // Legacy single-line shape: quantity was stuffed into 'supplier'
                addToBucket(quote.product, null, null, null, parseFloat(quote.supplier) || 0);
            }
        }

        // Source 4: QuoteRequest (kept from original aggregation)
        const quoteStatuses = ['submitted', 'pricing', 'quoted', 'accepted'];
        const quotes = await QuoteRequest.find({ status: { $in: quoteStatuses } })
            .select('items')
            .lean();
        for (const q of quotes) {
            for (const item of q.items || []) {
                addToBucket(item.productName, item.packSize, item.unit, item.chemicalId, item.quantityNeeded || 0);
            }
        }

        const items = Object.values(volumeNeeds).filter(i => i.quantityNeeded > 0);

        if (items.length === 0) {
            return res.status(400).json({ error: 'No pending volume needs found' });
        }

        const horizonNote = withinDays
            ? ` (within ${withinDays} days)`
            : '';
        const bidSheet = new SupplierBidSheet({
            title: title || `Volume Needs Bid - ${new Date().toLocaleDateString()}`,
            description: description || `Auto-generated from pending orders, quote requests, and price mining submissions${horizonNote}`,
            items,
            invitedSuppliers: (invitedSuppliers || []).map(s => ({
                supplierId: s.supplierId,
                supplierName: s.supplierName,
                contactEmail: s.contactEmail,
                contactPhone: s.contactPhone,
                status: 'invited'
            })),
            responseDueDate: responseDueDate ? new Date(responseDueDate) : new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
            status: 'draft',
            createdBy: req.user._id
        });

        await bidSheet.save();
        res.status(201).json(bidSheet);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get all bid sheets
app.get('/api/admin/bid-sheets', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status } = req.query;
        let query = {};
        if (status) query.status = status;

        const bidSheets = await SupplierBidSheet.find(query)
            .sort({ createdAt: -1 })
            .populate('createdBy', 'name');

        res.json(bidSheets);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get single bid sheet
app.get('/api/admin/bid-sheets/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const bidSheet = await SupplierBidSheet.findById(req.params.id)
            .populate('createdBy', 'name')
            .populate('items.chemicalId', 'productName costPrice sellPrice');

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        res.json(bidSheet);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Add suppliers to bid sheet
app.put('/api/admin/bid-sheets/:id/suppliers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { suppliers } = req.body;
        const bidSheet = await SupplierBidSheet.findById(req.params.id);

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        // Add new suppliers
        for (const s of suppliers) {
            const exists = bidSheet.invitedSuppliers.find(
                inv => inv.supplierName === s.supplierName || inv.supplierId?.toString() === s.supplierId
            );
            if (!exists) {
                bidSheet.invitedSuppliers.push({
                    supplierId: s.supplierId,
                    supplierName: s.supplierName,
                    contactEmail: s.contactEmail,
                    contactPhone: s.contactPhone,
                    status: 'invited'
                });
            }
        }

        bidSheet.updatedAt = new Date();
        bidSheet.updatedBy = req.user._id;
        await bidSheet.save();

        res.json(bidSheet);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Send bid sheet to suppliers - actually emails each invited supplier with the
// bid request. Skips placeholder @acreprofit.com emails (Sims until Kyle sources
// a real contact). Partial success: one failed email doesn't abort the batch.
app.put('/api/admin/bid-sheets/:id/send', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const bidSheet = await SupplierBidSheet.findById(req.params.id);

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        if (!bidSheet.invitedSuppliers || bidSheet.invitedSuppliers.length === 0) {
            return res.status(400).json({ error: 'No suppliers invited to bid' });
        }

        // Freshen contact info from the supplier User doc in case email/name
        // drifted since the bid sheet was created
        const supplierIds = bidSheet.invitedSuppliers
            .map(s => s.supplierId)
            .filter(Boolean);
        const freshSuppliers = supplierIds.length > 0
            ? await User.find({ _id: { $in: supplierIds }, role: 'supplier' })
                .select('email name companyName bidEligible')
                .lean()
            : [];
        const freshMap = {};
        freshSuppliers.forEach(s => { freshMap[s._id.toString()] = s; });

        // Build the bid items block once - same HTML used in every supplier email
        const itemsHtml = (bidSheet.items || []).map(i => `
            <tr>
                <td style="padding:8px; border-bottom:1px solid #eee;"><strong>${i.productName}</strong></td>
                <td style="padding:8px; border-bottom:1px solid #eee;">${i.packSize || '-'}</td>
                <td style="padding:8px; border-bottom:1px solid #eee; text-align:right;">${i.quantityNeeded || 0} ${i.unit || ''}</td>
                <td style="padding:8px; border-bottom:1px solid #eee; color:#666; font-size:0.9em;">${i.notes || ''}</td>
            </tr>
        `).join('');

        const responseDueStr = bidSheet.responseDueDate
            ? new Date(bidSheet.responseDueDate).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })
            : '7 days from now';
        const frontendUrl = process.env.FRONTEND_URL || 'https://acreprofit.com';

        const transporter = createEmailTransporter();
        const sendResults = { sent: 0, skipped: 0, failed: 0, details: [] };

        for (const invited of bidSheet.invitedSuppliers) {
            // Use freshest email from supplier User doc when available
            const fresh = invited.supplierId ? freshMap[invited.supplierId.toString()] : null;
            const email = (fresh?.email || invited.contactEmail || '').toLowerCase().trim();
            const displayName = fresh?.companyName || invited.supplierName || 'Supplier';

            // Placeholder-email skip: any @acreprofit.com address means mail
            // forwarding isn't set up for this supplier yet. Mark as skipped
            // rather than sending mail that will never route anywhere.
            if (!email || email.endsWith('@acreprofit.com')) {
                invited.lastEmailError = 'skipped: placeholder or missing email';
                sendResults.skipped++;
                sendResults.details.push({ supplier: displayName, result: 'skipped', reason: 'placeholder email' });
                console.warn(`Bid ${bidSheet.bidNumber}: skipped ${displayName} - ${invited.lastEmailError}`);
                continue;
            }

            if (!transporter) {
                invited.lastEmailError = 'skipped: email transporter not configured (SMTP env vars missing)';
                sendResults.skipped++;
                sendResults.details.push({ supplier: displayName, result: 'skipped', reason: 'SMTP not configured' });
                continue;
            }

            const mailBody = `
                <div style="font-family: Arial, sans-serif; max-width: 680px; margin: 0 auto;">
                    <div style="background:#2d5a27; color:white; padding:18px 24px;">
                        <h1 style="margin:0; font-size:1.4rem;">Acre Profit — Bid Request</h1>
                        <div style="font-size:0.95rem; opacity:0.9; margin-top:4px;">${bidSheet.bidNumber} · ${bidSheet.title}</div>
                    </div>
                    <div style="padding:20px 24px; background:#f9f9f9;">
                        <p>Hi ${displayName},</p>
                        <p>We're pooling orders across our farmers and requesting quotes on the following products.
                        ${bidSheet.description ? `<br><em>${bidSheet.description}</em>` : ''}</p>

                        <div style="background:#fff; padding:14px; border-radius:6px; margin:16px 0;">
                            <div style="font-weight:700; color:#2d5a27; margin-bottom:8px;">Please respond by: ${responseDueStr}</div>
                            <div style="color:#666; font-size:0.9em;">Reply to this email with your per-unit pricing, availability, lead time, and any freight/payment terms.</div>
                        </div>

                        <table style="width:100%; border-collapse:collapse; background:#fff; margin-bottom:16px;">
                            <thead>
                                <tr style="background:#f1f5f0;">
                                    <th style="padding:10px; text-align:left; border-bottom:2px solid #2d5a27;">Product</th>
                                    <th style="padding:10px; text-align:left; border-bottom:2px solid #2d5a27;">Pack Size</th>
                                    <th style="padding:10px; text-align:right; border-bottom:2px solid #2d5a27;">Quantity Needed</th>
                                    <th style="padding:10px; text-align:left; border-bottom:2px solid #2d5a27;">Notes</th>
                                </tr>
                            </thead>
                            <tbody>${itemsHtml}</tbody>
                        </table>

                        <p style="color:#666; font-size:0.9em;">
                            Questions: reply to this email or contact ${req.user.name || 'our team'} directly.<br>
                            Acre Profit · ${frontendUrl}
                        </p>
                    </div>
                </div>
            `;

            try {
                await transporter.sendMail({
                    from: process.env.EMAIL_FROM || '"Acre Profit" <contact@acreprofit.com>',
                    to: email,
                    subject: `Bid Request ${bidSheet.bidNumber} — ${bidSheet.title}`,
                    html: mailBody
                });
                invited.emailedAt = new Date();
                invited.lastEmailError = undefined;
                sendResults.sent++;
                sendResults.details.push({ supplier: displayName, result: 'sent', email });
                console.log(`Bid ${bidSheet.bidNumber} sent to ${displayName} at ${email}`);
            } catch (err) {
                invited.lastEmailError = err.message;
                sendResults.failed++;
                sendResults.details.push({ supplier: displayName, result: 'failed', reason: err.message });
                console.error(`Bid ${bidSheet.bidNumber}: send failed for ${displayName}: ${err.message}`);
            }

            invited.invitedAt = invited.invitedAt || new Date();
        }

        bidSheet.status = 'sent';
        bidSheet.sentAt = new Date();
        bidSheet.updatedAt = new Date();
        bidSheet.updatedBy = req.user._id;
        // Mongoose's nested array change detection is unreliable when modifying
        // sub-document fields in a loop. Force the dirty flag to guarantee writes.
        bidSheet.markModified('invitedSuppliers');

        await bidSheet.save();

        await logAudit({
            action: 'bid_sheet_sent',
            req,
            entityType: 'SupplierBidSheet',
            entityId: bidSheet._id,
            entityRef: bidSheet.bidNumber,
            reason: `Sent bid ${bidSheet.bidNumber} to suppliers`,
            after: sendResults
        });

        res.json({
            message: `Bid sheet ${bidSheet.bidNumber} processed: ${sendResults.sent} sent, ${sendResults.skipped} skipped, ${sendResults.failed} failed`,
            results: sendResults,
            bidSheet
        });
    } catch (error) {
        console.error('Bid sheet send error:', error);
        res.status(400).json({ error: error.message });
    }
});

// Record a supplier's bid response
app.post('/api/admin/bid-sheets/:id/responses', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { supplierName, supplierId, itemPricing, freight, validUntil, paymentTerms, deliveryTerms, bidNotes } = req.body;

        const bidSheet = await SupplierBidSheet.findById(req.params.id);

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        // Calculate totals
        let subtotal = 0;
        const pricedItems = itemPricing.map((ip, idx) => {
            const item = bidSheet.items[ip.itemIndex] || bidSheet.items[idx];
            const totalPrice = (ip.pricePerUnit || 0) * (item?.quantityNeeded || ip.quantity || 0);
            subtotal += totalPrice;
            return {
                productName: item?.productName || ip.productName,
                itemIndex: ip.itemIndex ?? idx,
                pricePerUnit: ip.pricePerUnit,
                totalPrice,
                availableQuantity: ip.availableQuantity,
                leadTimeDays: ip.leadTimeDays,
                notes: ip.notes
            };
        });

        const totalBid = subtotal + (freight || 0);

        // Add or update supplier response
        const existingBidIdx = bidSheet.supplierBids.findIndex(
            b => b.supplierName === supplierName || b.supplierId?.toString() === supplierId
        );

        const bidResponse = {
            supplierId,
            supplierName,
            receivedAt: new Date(),
            itemPricing: pricedItems,
            subtotal,
            freight: freight || 0,
            totalBid,
            validUntil: validUntil ? new Date(validUntil) : new Date(Date.now() + 14 * 24 * 60 * 60 * 1000),
            paymentTerms,
            deliveryTerms,
            bidNotes,
            isSelected: false
        };

        if (existingBidIdx >= 0) {
            bidSheet.supplierBids[existingBidIdx] = bidResponse;
        } else {
            bidSheet.supplierBids.push(bidResponse);
        }

        // Update invited supplier status
        const invitedIdx = bidSheet.invitedSuppliers.findIndex(
            s => s.supplierName === supplierName || s.supplierId?.toString() === supplierId
        );
        if (invitedIdx >= 0) {
            bidSheet.invitedSuppliers[invitedIdx].status = 'responded';
        }

        // Update bid sheet status
        if (bidSheet.status === 'sent') {
            bidSheet.status = 'responses_received';
        }

        bidSheet.updatedAt = new Date();
        bidSheet.updatedBy = req.user._id;
        await bidSheet.save();

        res.json(bidSheet);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Award bid to a supplier
app.put('/api/admin/bid-sheets/:id/award', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { supplierBidId, supplierName } = req.body;

        const bidSheet = await SupplierBidSheet.findById(req.params.id);

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        // Find the winning bid
        const winningBid = bidSheet.supplierBids.find(
            b => b._id.toString() === supplierBidId || b.supplierName === supplierName
        );

        if (!winningBid) {
            return res.status(400).json({ error: 'Supplier bid not found' });
        }

        // Mark as selected
        bidSheet.supplierBids.forEach(b => {
            b.isSelected = false;
        });
        winningBid.isSelected = true;
        winningBid.selectedAt = new Date();

        bidSheet.status = 'awarded';
        bidSheet.awardedAt = new Date();
        bidSheet.awardedSupplierId = winningBid.supplierId;
        bidSheet.awardedSupplierName = winningBid.supplierName;
        bidSheet.updatedAt = new Date();
        bidSheet.updatedBy = req.user._id;

        await bidSheet.save();

        res.json({ message: `Bid awarded to ${winningBid.supplierName}`, bidSheet });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create PO from awarded bid
app.post('/api/admin/bid-sheets/:id/create-po', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const bidSheet = await SupplierBidSheet.findById(req.params.id);

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        if (bidSheet.status !== 'awarded') {
            return res.status(400).json({ error: 'Bid must be awarded before creating PO' });
        }

        const winningBid = bidSheet.supplierBids.find(b => b.isSelected);
        if (!winningBid) {
            return res.status(400).json({ error: 'No winning bid found' });
        }

        // Create PO from winning bid
        const poItems = winningBid.itemPricing.map((ip, idx) => {
            const bidItem = bidSheet.items[ip.itemIndex] || bidSheet.items[idx];
            return {
                productName: bidItem?.productName || ip.productName,
                chemicalId: bidItem?.chemicalId,
                packSize: bidItem?.packSize,
                unit: bidItem?.unit,
                quantityOrdered: bidItem?.quantityNeeded || ip.availableQuantity,
                pricePerUnit: ip.pricePerUnit,
                totalPrice: ip.totalPrice
            };
        });

        const po = new PurchaseOrder({
            supplier: {
                name: winningBid.supplierName,
                contact: '',
                phone: '',
                email: ''
            },
            items: poItems,
            subtotal: winningBid.subtotal,
            freight: winningBid.freight,
            totalCost: winningBid.totalBid,
            status: 'draft',
            orderDate: new Date(),
            notes: `Created from Bid Sheet ${bidSheet.bidNumber}`,
            createdBy: req.user._id
        });

        await po.save();

        // Update bid sheet
        bidSheet.status = 'po_created';
        bidSheet.purchaseOrderId = po._id;
        bidSheet.updatedAt = new Date();
        bidSheet.updatedBy = req.user._id;
        await bidSheet.save();

        res.json({ message: 'Purchase Order created', purchaseOrder: po, bidSheet });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Compare supplier bids (returns ranked comparison)
app.get('/api/admin/bid-sheets/:id/compare', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const bidSheet = await SupplierBidSheet.findById(req.params.id);

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        if (bidSheet.supplierBids.length === 0) {
            return res.json({ message: 'No supplier bids to compare', comparison: [] });
        }

        // Rank by total bid (lowest first)
        const ranked = bidSheet.supplierBids
            .map(bid => ({
                supplierName: bid.supplierName,
                totalBid: bid.totalBid,
                subtotal: bid.subtotal,
                freight: bid.freight,
                itemCount: bid.itemPricing.length,
                avgPricePerItem: bid.subtotal / bid.itemPricing.length,
                validUntil: bid.validUntil,
                paymentTerms: bid.paymentTerms,
                isSelected: bid.isSelected,
                receivedAt: bid.receivedAt
            }))
            .sort((a, b) => a.totalBid - b.totalBid);

        // Add rank and savings info
        const lowestBid = ranked[0]?.totalBid || 0;
        ranked.forEach((bid, idx) => {
            bid.rank = idx + 1;
            bid.savingsVsHighest = ranked[ranked.length - 1].totalBid - bid.totalBid;
            bid.percentAboveLowest = lowestBid > 0 ? ((bid.totalBid - lowestBid) / lowestBid * 100).toFixed(1) : 0;
        });

        res.json({
            bidNumber: bidSheet.bidNumber,
            itemCount: bidSheet.items.length,
            supplierCount: bidSheet.supplierBids.length,
            comparison: ranked
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ---- CHEMICAL MIX RECIPE BOOK ROUTES ----

// Get all public mixes (recipe book browse)
app.get('/api/mixes', async (req, res) => {
    try {
        const {
            crop,
            timing,
            search,
            tags,
            sort = 'popular',
            page = 1,
            limit = 20
        } = req.query;

        const query = { status: 'published', isPublic: true };

        if (crop) query.crop = crop;
        if (timing) query.timing = timing;
        if (tags) query.tags = { $in: tags.split(',') };

        // Text search
        if (search) {
            query.$text = { $search: search };
        }

        // Sorting options
        let sortOption = {};
        switch (sort) {
            case 'popular':
                sortOption = { totalOrders: -1, averageRating: -1 };
                break;
            case 'rating':
                sortOption = { averageRating: -1, totalRatings: -1 };
                break;
            case 'newest':
                sortOption = { publishedAt: -1 };
                break;
            case 'views':
                sortOption = { totalViews: -1 };
                break;
            default:
                sortOption = { totalOrders: -1 };
        }

        const skip = (parseInt(page) - 1) * parseInt(limit);

        const [mixes, total] = await Promise.all([
            ChemicalMix.find(query)
                .sort(sortOption)
                .skip(skip)
                .limit(parseInt(limit))
                .populate('ingredients.chemicalId', 'productName category sellPrice packSize unit')
                .lean(),
            ChemicalMix.countDocuments(query)
        ]);

        // Hide creator info for anonymous mixes
        const safeMixes = mixes.map(mix => {
            if (mix.isAnonymous) {
                delete mix.createdBy;
            }
            return mix;
        });

        res.json({
            mixes: safeMixes,
            pagination: {
                page: parseInt(page),
                limit: parseInt(limit),
                total,
                pages: Math.ceil(total / parseInt(limit))
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get single mix by slug or ID
app.get('/api/mixes/:identifier', async (req, res) => {
    try {
        const { identifier } = req.params;

        // Try to find by slug first, then by ID
        let mix = await ChemicalMix.findOne({ slug: identifier })
            .populate('ingredients.chemicalId', 'productName category sellPrice packSize unit epaRegistrationNumber isRestrictedUse signalWord activeIngredients')
            .populate('createdBy', 'name farm.name');

        if (!mix && mongoose.Types.ObjectId.isValid(identifier)) {
            mix = await ChemicalMix.findById(identifier)
                .populate('ingredients.chemicalId', 'productName category sellPrice packSize unit epaRegistrationNumber isRestrictedUse signalWord activeIngredients')
                .populate('createdBy', 'name farm.name');
        }

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        // Increment view count
        await ChemicalMix.findByIdAndUpdate(mix._id, { $inc: { totalViews: 1 } });

        // Convert to object and handle anonymity
        const mixObj = mix.toObject();
        if (mixObj.isAnonymous) {
            mixObj.createdBy = mixObj.creatorDisplayName ? { name: mixObj.creatorDisplayName } : { name: 'Anonymous Farmer' };
        }

        res.json(mixObj);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get my mixes (authenticated)
app.get('/api/my-mixes', authMiddleware, async (req, res) => {
    try {
        const mixes = await ChemicalMix.find({ createdBy: req.user._id })
            .sort({ updatedAt: -1 })
            .populate('ingredients.chemicalId', 'productName category sellPrice packSize unit')
            .lean();

        res.json(mixes);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Create a new mix
app.post('/api/mixes', authMiddleware, async (req, res) => {
    try {
        let {
            name,
            crop,
            timing,
            timingNotes,
            ingredients,
            gallonsPerAcre,
            description,
            story,
            tips,
            tags,
            isAnonymous,
            creatorDisplayName,
            status,
            groundType,
            rotationRestrictions,
            grazingRestrictions,
            applications,
            visibility,
            sprayParams,
            basedOn
        } = req.body;

        // If frontend sent an "applications" array (from Build a Recipe / Save Custom Recipe),
        // flatten it into ingredients + timing
        if ((!ingredients || ingredients.length === 0) && applications && applications.length > 0) {
            timing = applications[0].timing || applications[0].name || 'Custom';
            ingredients = [];
            applications.forEach(app => {
                (app.chemicals || []).forEach(c => {
                    ingredients.push({
                        productName: c.productName,
                        rate: c.rate,
                        rateUnit: c.rateUnit,
                        timing: app.name,
                        packSize: c.packSize,
                        unit: c.unit,
                        unitsPerPack: c.unitsPerPack
                    });
                });
            });
            if (sprayParams?.gallonsPerAcre) gallonsPerAcre = sprayParams.gallonsPerAcre;
            if (basedOn && !description) description = `Based on ${basedOn}`;
        }

        // Map visibility to isPublic
        const isPublic = visibility === 'public';

        if (!name || !crop || !timing || !ingredients || ingredients.length === 0) {
            return res.status(400).json({ error: 'Name, crop, timing, and at least one ingredient are required' });
        }

        // Validate and populate ingredient names
        const populatedIngredients = await Promise.all(
            ingredients.map(async (ing) => {
                if (ing.chemicalId) {
                    const chemical = await Chemical.findById(ing.chemicalId);
                    if (chemical) {
                        return {
                            ...ing,
                            productName: chemical.productName,
                            category: chemical.category
                        };
                    }
                }
                return ing;
            })
        );

        // Auto-generate restrictions if not provided by the user
        let autoGroundType = groundType || null;
        let autoRotation = rotationRestrictions || null;
        let autoGrazing = grazingRestrictions || null;

        if (!autoGroundType || !autoRotation || !autoGrazing) {
            const generated = generateRecipeRestrictions(populatedIngredients);
            if (!autoGroundType) autoGroundType = generated.groundType;
            if (!autoRotation) autoRotation = generated.rotationRestrictions;
            if (!autoGrazing) autoGrazing = generated.grazingRestrictions;
        }

        const mix = new ChemicalMix({
            name,
            crop,
            timing,
            timingNotes,
            ingredients: populatedIngredients,
            gallonsPerAcre: gallonsPerAcre || 15,
            description,
            story,
            tips,
            tags: tags || [],
            isAnonymous: isAnonymous !== false, // Default to anonymous
            isPublic,
            creatorDisplayName,
            createdBy: req.user._id,
            status: status || (isPublic ? 'published' : 'draft'),
            publishedAt: (status === 'published' || isPublic) ? new Date() : null,
            groundType: autoGroundType,
            rotationRestrictions: autoRotation,
            grazingRestrictions: autoGrazing
        });

        await mix.save();
        res.status(201).json(mix);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update a mix
app.put('/api/mixes/:id', authMiddleware, async (req, res) => {
    try {
        const mix = await ChemicalMix.findById(req.params.id);

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        // Only creator can edit (or admin)
        if (mix.createdBy.toString() !== req.user._id.toString() && !isAdminLevel(req.user)) {
            return res.status(403).json({ error: 'Not authorized to edit this mix' });
        }

        const updates = req.body;

        // If changing to published, set publishedAt
        if (updates.status === 'published' && mix.status !== 'published') {
            updates.publishedAt = new Date();
        }

        // Re-populate ingredient names if ingredients changed
        if (updates.ingredients) {
            updates.ingredients = await Promise.all(
                updates.ingredients.map(async (ing) => {
                    if (ing.chemicalId) {
                        const chemical = await Chemical.findById(ing.chemicalId);
                        if (chemical) {
                            return {
                                ...ing,
                                productName: chemical.productName,
                                category: chemical.category
                            };
                        }
                    }
                    return ing;
                })
            );
        }

        Object.assign(mix, updates);
        await mix.save();

        res.json(mix);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete a mix
app.delete('/api/mixes/:id', authMiddleware, async (req, res) => {
    try {
        const mix = await ChemicalMix.findById(req.params.id);

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        // Only creator can delete (or admin)
        if (mix.createdBy.toString() !== req.user._id.toString() && !isAdminLevel(req.user)) {
            return res.status(403).json({ error: 'Not authorized to delete this mix' });
        }

        await ChemicalMix.findByIdAndDelete(req.params.id);
        await ChemicalMixRating.deleteMany({ mixId: req.params.id });

        res.json({ message: 'Mix deleted successfully' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Add yield data to a mix
app.post('/api/mixes/:id/yield-data', authMiddleware, async (req, res) => {
    try {
        const mix = await ChemicalMix.findById(req.params.id);

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        if (mix.createdBy.toString() !== req.user._id.toString() && !isAdminLevel(req.user)) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        const yieldEntry = {
            year: req.body.year || new Date().getFullYear(),
            crop: req.body.crop || mix.crop,
            acres: req.body.acres,
            yieldPerAcre: req.body.yieldPerAcre,
            yieldUnit: req.body.yieldUnit || 'bu/acre',
            location: req.body.location,
            notes: req.body.notes,
            imageUrl: req.body.imageUrl
        };

        mix.yieldData.push(yieldEntry);
        await mix.save();

        res.json(mix);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Add images to a mix
app.post('/api/mixes/:id/images', authMiddleware, async (req, res) => {
    try {
        const mix = await ChemicalMix.findById(req.params.id);

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        if (mix.createdBy.toString() !== req.user._id.toString() && !isAdminLevel(req.user)) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        const { url, caption } = req.body;

        if (!url) {
            return res.status(400).json({ error: 'Image URL is required' });
        }

        mix.images.push({ url, caption, uploadedAt: new Date() });
        await mix.save();

        res.json(mix);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Rate a mix
app.post('/api/mixes/:id/ratings', authMiddleware, async (req, res) => {
    try {
        const { rating, review, usedOnCrop, usedOnAcres, yieldResult, yieldUnit, wouldRecommend, madeModifications, modifications, images } = req.body;

        if (!rating || rating < 1 || rating > 5) {
            return res.status(400).json({ error: 'Rating must be between 1 and 5' });
        }

        const mix = await ChemicalMix.findById(req.params.id);
        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        // Check for existing rating
        let existingRating = await ChemicalMixRating.findOne({
            mixId: req.params.id,
            userId: req.user._id
        });

        if (existingRating) {
            // Update existing rating
            existingRating.rating = rating;
            existingRating.review = review;
            existingRating.usedOnCrop = usedOnCrop;
            existingRating.usedOnAcres = usedOnAcres;
            existingRating.yieldResult = yieldResult;
            existingRating.yieldUnit = yieldUnit;
            existingRating.wouldRecommend = wouldRecommend;
            existingRating.madeModifications = madeModifications;
            existingRating.modifications = modifications;
            existingRating.images = images || [];
            await existingRating.save();
        } else {
            // Create new rating
            existingRating = new ChemicalMixRating({
                mixId: req.params.id,
                userId: req.user._id,
                rating,
                review,
                usedOnCrop,
                usedOnAcres,
                yieldResult,
                yieldUnit,
                wouldRecommend,
                madeModifications,
                modifications,
                images: images || []
            });
            await existingRating.save();
        }

        // Update mix stats
        await updateMixRatingStats(req.params.id);

        res.json(existingRating);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get ratings for a mix
app.get('/api/mixes/:id/ratings', async (req, res) => {
    try {
        const { page = 1, limit = 10 } = req.query;
        const skip = (parseInt(page) - 1) * parseInt(limit);

        const [ratings, total] = await Promise.all([
            ChemicalMixRating.find({ mixId: req.params.id })
                .sort({ helpfulVotes: -1, createdAt: -1 })
                .skip(skip)
                .limit(parseInt(limit))
                .populate('userId', 'name farm.state')
                .lean(),
            ChemicalMixRating.countDocuments({ mixId: req.params.id })
        ]);

        res.json({
            ratings,
            pagination: {
                page: parseInt(page),
                limit: parseInt(limit),
                total,
                pages: Math.ceil(total / parseInt(limit))
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Vote a rating as helpful
app.post('/api/ratings/:id/helpful', authMiddleware, async (req, res) => {
    try {
        const rating = await ChemicalMixRating.findByIdAndUpdate(
            req.params.id,
            { $inc: { helpfulVotes: 1 } },
            { new: true }
        );

        if (!rating) {
            return res.status(404).json({ error: 'Rating not found' });
        }

        res.json(rating);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Order ingredients from a mix (create order from mix)
app.post('/api/mixes/:id/order', authMiddleware, async (req, res) => {
    try {
        const { acres, gallonsPerAcre } = req.body;

        if (!acres || acres <= 0) {
            return res.status(400).json({ error: 'Acres is required' });
        }

        const mix = await ChemicalMix.findById(req.params.id)
            .populate('ingredients.chemicalId');

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        // Calculate quantities needed for each ingredient
        const orderItems = mix.ingredients.map(ing => {
            const chemical = ing.chemicalId;
            if (!chemical) {
                return null; // Chemical no longer exists
            }

            // Calculate total needed based on rate and acres
            const ratePerAcre = ing.rate;
            const totalNeeded = ratePerAcre * acres;

            // Convert to packages based on pack size
            const unitsPerPack = chemical.unitsPerPack || 1;
            const packagesNeeded = Math.ceil(totalNeeded / unitsPerPack);

            return {
                chemicalId: chemical._id,
                productName: chemical.productName,
                category: chemical.category,
                rate: ing.rate,
                rateUnit: ing.rateUnit,
                totalNeeded,
                packagesNeeded,
                packSize: chemical.packSize,
                unit: chemical.unit,
                pricePerPackage: chemical.sellPrice,
                lineTotal: packagesNeeded * chemical.sellPrice
            };
        }).filter(item => item !== null);

        // Increment mix order count
        await ChemicalMix.findByIdAndUpdate(mix._id, { $inc: { totalOrders: 1 } });

        // Return order summary (user can proceed to checkout)
        res.json({
            mix: {
                id: mix._id,
                name: mix.name,
                crop: mix.crop,
                timing: mix.timing
            },
            acres,
            gallonsPerAcre: gallonsPerAcre || mix.gallonsPerAcre,
            totalWaterVolume: (gallonsPerAcre || mix.gallonsPerAcre) * acres,
            items: orderItems,
            subtotal: orderItems.reduce((sum, item) => sum + item.lineTotal, 0)
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// AI Analysis - Analyze mix for restrictions, compatibility, etc.
app.post('/api/mixes/:id/analyze', authMiddleware, async (req, res) => {
    try {
        const mix = await ChemicalMix.findById(req.params.id)
            .populate('ingredients.chemicalId', 'productName category activeIngredients isRestrictedUse signalWord hazardClassifications requiredCertifications sdsUrl labelUrl');

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        // Build analysis data from chemicals
        const ingredientData = mix.ingredients.map(ing => {
            const chem = ing.chemicalId;
            if (!chem) return { productName: ing.productName, rate: ing.rate, rateUnit: ing.rateUnit };

            return {
                productName: chem.productName,
                category: chem.category,
                rate: ing.rate,
                rateUnit: ing.rateUnit,
                activeIngredients: chem.activeIngredients,
                isRestrictedUse: chem.isRestrictedUse,
                signalWord: chem.signalWord,
                hazardClassifications: chem.hazardClassifications,
                requiredCertifications: chem.requiredCertifications
            };
        });

        // Extract modes of action from active ingredients
        const modesOfAction = [];
        const activeIngredientsList = [];

        ingredientData.forEach(ing => {
            if (ing.activeIngredients) {
                ing.activeIngredients.forEach(ai => {
                    if (ai.name && !activeIngredientsList.includes(ai.name)) {
                        activeIngredientsList.push(ai.name);
                    }
                });
            }
        });

        // Check for restricted use products
        const hasRestrictedUse = ingredientData.some(ing => ing.isRestrictedUse);

        // Check required certifications
        const requiredCerts = new Set();
        ingredientData.forEach(ing => {
            if (ing.requiredCertifications) {
                ing.requiredCertifications.forEach(cert => requiredCerts.add(cert));
            }
        });

        // Build analysis object
        const analysis = {
            lastAnalyzedAt: new Date(),
            modesOfAction: modesOfAction,
            warnings: [],
            compatibility: 'Analysis based on product labels. Always refer to product labels for tank mix compatibility.',
            restrictions: hasRestrictedUse ? 'This mix contains Restricted Use Pesticides (RUP). Applicator certification required.' : 'No restricted use products in this mix.',
            ppeRequired: 'Refer to individual product labels for PPE requirements. Use the most restrictive PPE when tank mixing.',
            fullAnalysis: JSON.stringify({
                ingredients: ingredientData,
                activeIngredients: activeIngredientsList,
                requiredCertifications: Array.from(requiredCerts),
                hasRestrictedUse,
                analyzedAt: new Date().toISOString()
            }, null, 2)
        };

        if (hasRestrictedUse) {
            analysis.warnings.push('Contains Restricted Use Pesticide(s) - Applicator license required');
        }

        if (requiredCerts.has('dicamba_training')) {
            analysis.warnings.push('Contains Dicamba product - Annual training certification required');
        }

        if (requiredCerts.has('paraquat_training')) {
            analysis.warnings.push('Contains Paraquat - EPA Paraquat training certification required');
        }

        // Save analysis to mix
        mix.aiAnalysis = analysis;
        await mix.save();

        res.json({
            message: 'Analysis complete',
            analysis,
            ingredientDetails: ingredientData
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get featured mixes
app.get('/api/mixes-featured', async (req, res) => {
    try {
        const featured = await ChemicalMix.find({
            status: 'published',
            isPublic: true,
            isFeatured: true
        })
            .sort({ averageRating: -1 })
            .limit(6)
            .populate('ingredients.chemicalId', 'productName category')
            .lean();

        res.json(featured);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get top mixes by crop
app.get('/api/mixes-by-crop/:crop', async (req, res) => {
    try {
        const { timing } = req.query;
        const query = {
            status: 'published',
            isPublic: true,
            crop: req.params.crop.toLowerCase()
        };

        if (timing) query.timing = timing;

        const mixes = await ChemicalMix.find(query)
            .sort({ averageRating: -1, totalOrders: -1 })
            .limit(20)
            .populate('ingredients.chemicalId', 'productName category')
            .lean();

        res.json(mixes);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Admin: Feature/unfeature a mix
app.put('/api/admin/mixes/:id/feature', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { isFeatured } = req.body;

        const mix = await ChemicalMix.findByIdAndUpdate(
            req.params.id,
            { isFeatured: !!isFeatured },
            { new: true }
        );

        if (!mix) {
            return res.status(404).json({ error: 'Mix not found' });
        }

        res.json(mix);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get available crops and timings for filtering
app.get('/api/mixes-filters', async (req, res) => {
    try {
        const [crops, tags] = await Promise.all([
            ChemicalMix.distinct('crop', { status: 'published', isPublic: true }),
            ChemicalMix.distinct('tags', { status: 'published', isPublic: true })
        ]);

        res.json({
            crops: crops.sort(),
            timings: [
                { value: 'burndown', label: 'Burndown', description: 'Pre-plant weed control' },
                { value: 'pre-emerge', label: 'Pre-Emerge', description: 'After planting, before crop emerges' },
                { value: 'early-post', label: 'Early Post', description: 'V1-V3 corn / VC-V2 beans' },
                { value: 'post-emerge', label: 'Post-Emerge', description: 'General post-emergence' },
                { value: 'v4-v6', label: 'V4-V6 Window', description: 'Mid-season corn window' },
                { value: 'v6-plus', label: 'V6+ / Late Post', description: 'Late post applications' },
                { value: 'r1-r3', label: 'R1-R3 (Beans)', description: 'Reproductive stage soybeans' },
                { value: 'layby', label: 'Layby', description: 'Last app before canopy closes' },
                { value: 'tassel', label: 'Tassel', description: 'At or around tasseling' },
                { value: 'harvest-aid', label: 'Harvest Aid', description: 'Pre-harvest desiccant' },
                { value: 'fall-application', label: 'Fall Application', description: 'Post-harvest' },
                { value: 'cover-crop', label: 'Cover Crop', description: 'Cover crop termination' },
                { value: 'other', label: 'Other', description: 'Other timing' }
            ],
            tags: tags.sort(),
            defaultCrops: ['corn', 'soybeans', 'wheat', 'sorghum', 'cotton', 'sunflowers', 'dry-beans', 'sugar-beets']
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ============ START SERVER ============

// Migration: Fix customers without representative ObjectId
async function migrateCustomerRepresentatives() {
    const repEmails = {
        kyle: 'office@togoag.com',
        ty: 'tymollohan77@gmail.com',
        chad: 'ckbamford@yahoo.com',
        seth: 'seth@acreprofit.com'
    };

    // Find all customers without representative but with representativeId
    const customersToUpdate = await User.find({
        role: 'customer',
        representative: { $exists: false },
        representativeId: { $exists: true, $ne: null }
    });

    if (customersToUpdate.length === 0) {
        console.log('Customer representative migration: No customers need updating');
        return;
    }

    let updated = 0;
    for (const customer of customersToUpdate) {
        const repEmail = repEmails[customer.representativeId];
        if (repEmail) {
            const distributor = await User.findOne({ email: repEmail });
            if (distributor) {
                customer.representative = distributor._id;
                await customer.save();
                updated++;
            }
        }
    }
    console.log(`Customer representative migration: Updated ${updated} of ${customersToUpdate.length} customers`);
}

connectDB().then(async () => {
    // Run migrations
    await migrateCustomerRepresentatives();

    // Fix Shuttle/Tote unitsPerPack: 250 → 265
    try {
        const fixed = await Chemical.updateMany(
            { packSize: { $in: ['Shuttle', 'Tote'] }, unit: 'gal', unitsPerPack: 250 },
            { $set: { unitsPerPack: 265, updatedAt: new Date() } }
        );
        if (fixed.modifiedCount > 0) {
            console.log(`Fixed ${fixed.modifiedCount} Shuttle/Tote products: unitsPerPack 250 → 265`);
        }
    } catch (e) { console.error('Shuttle fix error:', e.message); }

    // Set label and SDS links for all products
    try {
        const labelData = [
            { name: /meso 4sc/i, epa: '101458-3',
              labelUrl: 'https://www3.epa.gov/pesticides/chem_search/ppls/101458-00003-20230608.pdf',
              sdsUrl: 'https://cropprotectdirect.com/products/meso-4sc/' },
            { name: /xsate.*glyphosate/i, epa: '88343-5',
              labelUrl: 'https://s3-us-west-1.amazonaws.com/agrian-cg-fs1-production/pdfs/XSATE_GLYPHOSATE_53.8_Label.pdf',
              sdsUrl: 'https://s3-us-west-1.amazonaws.com/agrian-cg-fs1-production/pdfs/XSATE_GLYPHOSATE_53.8_MSDS.pdf' },
            { name: /dicamba 49/i, epa: '85678-46',
              labelUrl: 'https://www.redeagleinternational.com/wp-content/uploads/2018/10/Dicamba-48.6-SL-EM-30Oct18.pdf',
              sdsUrl: 'https://www.redeagleinternational.com/wp-content/uploads/2018/12/Dicamba_49.8_SL-SDS-20Nov18.pdf' },
            { name: /flumioxazin/i, epa: '85678-34',
              labelUrl: 'https://s3-us-west-1.amazonaws.com/agrian-cg-fs1-production/pdfs/Flumioxazin_51_WDG_Label.pdf',
              sdsUrl: 'https://assets.greenbook.net/18-40-26-28-05-2019-Ms_384_RedEagle_Flumioxazin_51__WDG.pdf' },
            { name: /defy lv/i, epa: '66222-220',
              labelUrl: 'https://s3-us-west-1.amazonaws.com/agrian-cg-fs1-production/pdfs/Defy_LV-6_Label1g.pdf',
              sdsUrl: 'https://www.adama.com/us/en/products/herbicide/defy-lv-6' },
            { name: /sulfentrazone/i, epa: '85678-59',
              labelUrl: 'https://www.redeagleinternational.com/wp-content/uploads/2018/10/Sulfentrazone-39.6-SC-EM15Aug18.pdf',
              sdsUrl: 'https://www.redeagleinternational.com/wp-content/uploads/2021/04/SDS_Sulfentrazone_39.6_SC.pdf' },
            { name: /atrazine 4l/i, epa: '35915-4',
              labelUrl: 'https://www.sipcam.com/downloads/10890/2648/ATRIZINE_4L_bilingual.pdf',
              sdsUrl: 'https://www.sipcam.com/us/en/agriculture/agrochemicals/atrazine-4l' },
            { name: /rancor/i, epa: '91234-73',
              labelUrl: 'https://atticusllc.com/wp-content/uploads/2020/08/Rancor-4-F-Specimen.pdf',
              sdsUrl: 'https://atticusllc.com/wp-content/uploads/2020/08/Rancor-4-F-SDS-v1.1.pdf' },
            { name: /hydrovant/i, epa: '',
              labelUrl: 'https://hydrovant.com/wp-content/uploads/sites/2/2021/08/Hydrovant-fA-English-Specimen-Label-082021.pdf',
              sdsUrl: 'https://hydrovant.com/wp-content/uploads/sites/2/2021/08/HYDROVANT-fA-SDS-ENGLISH-082021.pdf' },
            { name: /dicamba.*tigris/i, epa: '92647-11',
              labelUrl: 'https://assets.greenbook.net/23-09-54-31-05-2024-Tigris_Dicamba_DMA_-_label.pdf',
              sdsUrl: 'https://tigrisag.com/wp-content/uploads/2022/08/TIGRIS0519_SDS_DicambaDMA_v0r1.pdf' },
            { name: /lv6 de-ester/i, epa: '19713-655',
              labelUrl: 'https://s3-us-west-1.amazonaws.com/agrian-cg-fs1-production/pdfs/Drexel_De-EsterTM_LV6_Label.pdf',
              sdsUrl: 'https://www.drexchem.com/products/de-ester-lv6/' },
            { name: /anthem nxt/i, epa: '279-9674',
              labelUrl: 'https://ag.fmc.com/us/en/herbicides/anthem-nxt-herbicide',
              sdsUrl: 'https://ag.fmc.com/us/en/herbicides/anthem-nxt-herbicide' },
            { name: /mivum/i, epa: '83529-89',
              labelUrl: 'https://www.shardausa.com/product/mivum',
              sdsUrl: 'https://www.shardausa.com/product/mivum' },
        ];

        let labelCount = 0;
        for (const ld of labelData) {
            if (!ld.labelUrl && !ld.sdsUrl && !ld.epa) continue;
            const chems = await Chemical.find({ productName: ld.name });
            for (const chem of chems) {
                let changed = false;
                if (ld.labelUrl && !chem.labelUrl) { chem.labelUrl = ld.labelUrl; changed = true; }
                if (ld.sdsUrl && !chem.sdsUrl) { chem.sdsUrl = ld.sdsUrl; changed = true; }
                if (ld.epa && !chem.epaRegistrationNumber) { chem.epaRegistrationNumber = ld.epa; changed = true; }
                if (changed) { await chem.save(); labelCount++; }
            }
        }
        if (labelCount > 0) console.log(`Set label/SDS links on ${labelCount} products`);
    } catch (e) { console.error('Label link error:', e.message); }

    // Rename products to match actual label names
    try {
        const renames = [
            { old: 'Dicamba DMA', new: 'Dicamba 49.8% SL', manufacturer: 'Red Eagle Agricultural Chemicals' },
            { old: 'CPD Mesotrione', new: 'Meso 4SC', manufacturer: 'JABCO LLC (Crop Protect Direct)' },
            { old: '2,4-D LV6', new: 'Defy LV-6', manufacturer: 'ADAMA Essentials' },
            { old: 'Flumioxazin 51%', new: 'Flumioxazin 51% WDG', manufacturer: 'Red Eagle Agricultural Chemicals' },
            { old: 'Sulfentrazone 4SC', new: 'Sulfentrazone 39.6% SC', manufacturer: '' },
            { old: 'Sulfentrazone', new: 'Sulfentrazone 39.6% SC', manufacturer: '' },
            { old: 'Hydrovant', new: 'Hydrovant fA', manufacturer: 'Corbet Scientific, LLC' },
        ];

        for (const r of renames) {
            const updated = await Chemical.updateMany(
                { productName: r.old },
                { $set: { productName: r.new, ...(r.manufacturer ? { manufacturer: r.manufacturer } : {}), updatedAt: new Date() } }
            );
            if (updated.modifiedCount > 0) {
                console.log(`Renamed "${r.old}" → "${r.new}" (${updated.modifiedCount} records)`);
                // Also update inventory records
                await Inventory.updateMany({ productName: r.old }, { $set: { productName: r.new, updatedAt: new Date() } });
                await InventoryBatch.updateMany({ productName: r.old }, { $set: { productName: r.new } });
            }
        }
    } catch (e) { console.error('Product rename error:', e.message); }

    // Ensure Flumioxazin 51% WDG has inventory (from JABCO-SO2129: 1440 lb)
    try {
        const flumi = await Chemical.findOne({ productName: /flumioxazin/i });
        if (flumi) {
            const inv = await Inventory.findOne({ chemicalId: flumi._id, location: 'main' });
            if (!inv) {
                await receiveInventory({
                    chemicalId: flumi._id,
                    productName: flumi.productName,
                    packSize: flumi.packSize,
                    unit: flumi.unit || 'lb',
                    quantity: 1440,
                    unitCost: 14.00,
                    location: 'main',
                    poNumber: 'JABCO-SO2129',
                    lotNumber: 'jabco-so2129-flumi',
                    supplierName: 'JABCO',
                    userId: null
                });
                console.log('Created Flumioxazin 51% inventory: 1,440 lb @ $14.00/lb');
            }
            // Ensure it's available for order
            if (!flumi.availableForOrder) {
                flumi.availableForOrder = true;
                flumi.isActive = true;
                await flumi.save();
            }
        }
    } catch (e) { console.error('Flumioxazin inventory fix error:', e.message); }

    // Paraquat RUP flag migration — original seed used isRUP (not a schema field)
    // so Mongoose dropped it silently. Production Paraquat docs had no
    // isRestrictedUse flag and no requiredCertifications, meaning the RUP
    // compliance check never fired on them. Backfill any existing Paraquat
    // records and verify the flag is set so I5 auto-creation catches them.
    try {
        const paraquatResult = await Chemical.updateMany(
            {
                productName: /paraquat/i,
                $or: [
                    { isRestrictedUse: { $ne: true } },
                    { requiredCertifications: { $size: 0 } },
                    { requiredCertifications: { $exists: false } }
                ]
            },
            {
                $set: {
                    isRestrictedUse: true,
                    requiredCertifications: ['private_applicator', 'paraquat_training']
                }
            }
        );
        if (paraquatResult.modifiedCount > 0) {
            console.log(`Paraquat RUP flag migration: fixed ${paraquatResult.modifiedCount} record(s)`);
        } else {
            console.log(`Paraquat RUP flag migration: matched ${paraquatResult.matchedCount}, modified 0 (nothing to fix)`);
        }
    } catch (e) { console.error('Paraquat RUP flag migration error:', e.message); }

    // Add Rancor 4F (Metribuzin 4F) - JABCO Invoice 1622, SO# 2131: 180 gal @ $45.50
    try {
        let rancor = await Chemical.findOne({ productName: /rancor/i });
        if (!rancor) {
            rancor = await Chemical.create({
                productName: 'Rancor 4F',
                packSize: '2x2.5 gal',
                unit: 'gal',
                unitsPerPack: 5,
                costPrice: 45.50,
                adminMarginDollars: 0,
                adminPrice: 45.50,
                marginDollars: 0,
                sellPrice: 45.50,
                category: 'herbicide',
                sourceSupplier: 'JABCO',
                manufacturer: 'Crop Protect Direct',
                signalWord: 'CAUTION',
                notes: 'Metribuzin 4F herbicide',
                activeIngredients: [{ name: 'Metribuzin', percentage: 39.6 }],
                isActive: true,
                availableForOrder: true
            });
            console.log('Created Rancor 4F product');
        }

        const rancorInv = await Inventory.findOne({ chemicalId: rancor._id, location: 'main' });
        if (!rancorInv) {
            await receiveInventory({
                chemicalId: rancor._id,
                productName: 'Rancor 4F',
                packSize: '2x2.5 gal',
                unit: 'gal',
                quantity: 180,
                unitCost: 45.50,
                location: 'main',
                poNumber: 'JABCO-SO2131',
                lotNumber: 'jabco-so2131-rancor',
                supplierName: 'JABCO',
                userId: null
            });
            console.log('Added Rancor 4F inventory: 180 gal @ $45.50/gal ($8,190)');
        }
    } catch (e) { console.error('Rancor 4F setup error:', e.message); }

    // Rename CPD → JABCO in all existing records
    try {
        const chemFixed = await Chemical.updateMany(
            { sourceSupplier: { $in: ['CPD', 'Crop Protect Direct'] } },
            { $set: { sourceSupplier: 'JABCO' } }
        );
        if (chemFixed.modifiedCount > 0) console.log(`Renamed ${chemFixed.modifiedCount} products: CPD → JABCO`);
    } catch (e) { console.error('CPD rename error:', e.message); }

    // CLEAN LEDGER: Wipe old confusing entries and create simple loan entries
    // Kyle loaned AP $167,987 to buy JABCO + Hydrovant inventory
    // Ty loaned AP $146,445 to buy Sims inventory
    // AP owns all inventory. AP pays them back as customers pay.
    try {
        const kyle = await User.findOne({ email: 'office@togoag.com' });
        const ty = await User.findOne({ email: 'tymollohan77@gmail.com' });

        if (kyle && ty) {
            // Check if clean ledger already set up (v7 - loans only, no transfer entries)
            // Wipe if v5 transfer entries exist (force clean rebuild)
            const hasOldTransferEntries = await LedgerEntry.findOne({ description: { $regex: /^\(v5\) Transfer/ } });
            const cleanV7Marker = await LedgerEntry.findOne({ description: '(v7) Ledger initialized - loans only' });
            if (!cleanV7Marker || hasOldTransferEntries) {
                // Wipe all old seed ledger entries
                await LedgerEntry.deleteMany({});
                console.log('Cleared old ledger entries for v5 product-line rebuild');

                // Kyle's loans - ONE LINE PER PRODUCT - total $167,987
                const kyleProducts = [
                    { product: 'XSATE Glyphosate 53.8%', qty: 4240, unit: 'gal', cost: 13.25, amount: 56180.00, po: 'JABCO-2119' },
                    { product: 'Meso 4SC', qty: 720, unit: 'gal', cost: 45.75, amount: 32940.00, po: 'JABCO SO# 2129' },
                    { product: 'Flumioxazin 51% WDG', qty: 1440, unit: 'lb', cost: 14.00, amount: 20160.00, po: 'JABCO SO# 2129' },
                    { product: 'Sulfentrazone 39.6% SC', qty: 180, unit: 'gal', cost: 71.50, amount: 12870.00, po: 'JABCO SO# 2129' },
                    { product: 'Dicamba 49.8% SL', qty: 180, unit: 'gal', cost: 30.25, amount: 5445.00, po: 'JABCO SO# 2129' },
                    { product: 'Defy LV-6', qty: 180, unit: 'gal', cost: 28.90, amount: 5202.00, po: 'JABCO SO# 2129' },
                    { product: 'Rancor 4F', qty: 180, unit: 'gal', cost: 45.50, amount: 8190.00, po: 'JABCO SO# 2131' },
                    { product: 'Hydrovant fA', qty: 360, unit: 'gal', cost: 75.00, amount: 27000.00, po: 'Direct purchase' }
                ];

                for (const p of kyleProducts) {
                    await createLedgerEntry({
                        representativeId: kyle._id,
                        description: `(v5) ${p.product} - ${p.qty.toLocaleString()} ${p.unit}`,
                        amount: p.amount,
                        type: 'credit',
                        category: 'supplier_payment',
                        referenceType: 'Manual',
                        notes: `${p.po} | ${p.qty.toLocaleString()} ${p.unit} @ $${p.cost.toFixed(2)}/${p.unit} | Kyle funded`
                    });
                }

                // Ty's loans - ONE LINE PER PRODUCT - total $146,445
                const tyProducts = [
                    { product: 'Atrazine 4L', qty: 4240, unit: 'gal', cost: 12.50, amount: 53000.00, po: 'Sims #103849' },
                    { product: 'Dicamba DMA (Tigris)', qty: 1060, unit: 'gal', cost: 25.50, amount: 27030.00, po: 'Sims #103850' },
                    { product: 'LV6 De-Ester (Drexel)', qty: 1060, unit: 'gal', cost: 22.75, amount: 24115.00, po: 'Sims #103850' },
                    { product: 'Anthem NXT', qty: 90, unit: 'gal', cost: 430.00, amount: 38700.00, po: 'Sims #103850' },
                    { product: 'Mivum', qty: 1600, unit: 'oz', cost: 2.25, amount: 3600.00, po: 'Sims #103850' }
                ];

                for (const p of tyProducts) {
                    await createLedgerEntry({
                        representativeId: ty._id,
                        description: `(v5) ${p.product} - ${p.qty.toLocaleString()} ${p.unit}`,
                        amount: p.amount,
                        type: 'credit',
                        category: 'supplier_payment',
                        referenceType: 'Manual',
                        notes: `${p.po} | ${p.qty.toLocaleString()} ${p.unit} @ $${p.cost.toFixed(2)}/${p.unit} | Ty funded`
                    });
                }

                // Marker entry to prevent re-seed
                await createLedgerEntry({
                    representativeId: kyle._id,
                    description: '(v7) Ledger initialized - loans only',
                    amount: 0,
                    type: 'debit',
                    category: 'adjustment',
                    referenceType: 'Manual',
                    notes: 'System marker - safe to ignore. Created when ledger was cleaned and initialized with loan entries only.'
                });

                // Mark PO payment status - who paid which supplier invoice
                const jabcoKyle = ['JABCO-2119', 'JABCO-SO2129', 'JABCO-SO2131'];
                const simsTy = ['SIMS-103849', 'SIMS-103850'];

                for (const poNum of jabcoKyle) {
                    await PurchaseOrder.updateMany(
                        { poNumber: poNum },
                        { $set: { paidBy: kyle._id, paymentStatus: 'paid', paidDate: new Date('2026-03-30'), paymentMethod: 'check' } }
                    );
                }
                for (const poNum of simsTy) {
                    await PurchaseOrder.updateMany(
                        { poNumber: poNum },
                        { $set: { paidBy: ty._id, paymentStatus: 'paid', paidDate: new Date('2026-03-30'), paymentMethod: 'check' } }
                    );
                }
                console.log('PO payment status updated: JABCO=Kyle, Sims=Ty');

                console.log(`Clean v5 ledger: Kyle ${kyleProducts.length} product lines ($167,987), Ty ${tyProducts.length} product lines ($146,445) - Total AP debt: $314,432`);

                // NOTE: Inventory transfers between Kyle and Ty (Rancor/Flumi/Meso/Sulf/Hydrovant
                // to Ty, Atrazine to Kyle) are tracked in Inventory records, NOT the ledger.
                // AP's loan debt stays at $314,432 until AP writes checks to pay them back.
                // Transfers create debts BETWEEN Kyle and Ty (not with AP) - handled separately.
            }
        }

        // Clean up any stale JABCO vendor entries from previous version
        const jabcoVendor = await User.findOne({ email: 'vendor-jabco@acreprofit.com' });
        if (jabcoVendor) {
            await LedgerEntry.deleteMany({ representativeId: jabcoVendor._id });
        }
    } catch (e) { console.error('Clean ledger setup error:', e.message); }

    // Correct inventory quantities to match physical count
    // Kyle has: Meso 180, XSATE 4240, Dicamba 180, Defy LV-6 180, Hydrovant 180, Flumi 720
    // Ty has: Rancor 180, Meso 360, Flumi 720, Sulfentrazone 180, Hydrovant 180
    // Plus Ty's Sims: Atrazine, Dicamba Tigris, LV6 Drexel, Anthem NXT, Mivum
    try {
        const kyle = await User.findOne({ email: 'office@togoag.com' });
        const ty = await User.findOne({ email: 'tymollohan77@gmail.com' });
        if (kyle && ty) {
            // Kyle's JABCO inventory - correct quantities
            // Kyle's inventory (JABCO products + 6 totes Atrazine from Sims)
            const kyleProducts = [
                { name: /meso 4sc/i, qty: 360, owner: kyle._id },
                { name: /xsate/i, qty: 4240, owner: kyle._id },
                { name: /dicamba 49/i, qty: 180, owner: kyle._id },
                { name: /defy lv/i, qty: 180, owner: kyle._id },
                { name: /hydrovant/i, qty: 180, owner: kyle._id },
                { name: /flumioxazin/i, qty: 720, owner: kyle._id },
                { name: /atrazine 4l/i, qty: 1590, owner: kyle._id },
            ];
            // Ty's JABCO products (transferred from Kyle)
            const tyJabcoProducts = [
                { name: /rancor/i, qty: 180, owner: ty._id },
                { name: /meso 4sc/i, qty: 360, owner: ty._id },
                { name: /flumioxazin/i, qty: 720, owner: ty._id },
                { name: /sulfentrazone/i, qty: 180, owner: ty._id },
                { name: /hydrovant/i, qty: 180, owner: ty._id },
            ];
            // Ty's Sims products (Ty paid for all of these)
            const tySimsProducts = [
                { name: /atrazine 4l/i, qty: 2650, owner: ty._id },
                { name: /dicamba.*tigris/i, qty: 1060, owner: ty._id },
                { name: /lv6 de-ester/i, qty: 1060, owner: ty._id },
                { name: /anthem nxt/i, qty: 90, owner: ty._id },
                { name: /mivum/i, qty: 1600, owner: ty._id },
            ];

            // Set Kyle's quantities
            for (const p of kyleProducts) {
                const chem = await Chemical.findOne({ productName: p.name });
                if (chem) {
                    const inv = await Inventory.findOne({ chemicalId: chem._id, location: 'main' });
                    if (inv) {
                        inv.quantityOnHand = p.qty;
                        inv.quantityAvailable = p.qty - inv.quantityReserved;
                        inv.distributorId = p.owner;
                        inv.updatedAt = new Date();
                        await inv.save();
                    }
                }
            }

            // Create separate inventory records for Ty's JABCO products
            for (const p of tyJabcoProducts) {
                const chem = await Chemical.findOne({ productName: p.name });
                if (chem) {
                    let inv = await Inventory.findOne({ chemicalId: chem._id, location: 'ty' });
                    if (!inv) {
                        inv = new Inventory({
                            chemicalId: chem._id,
                            productName: chem.productName,
                            packSize: chem.packSize,
                            unit: chem.unit,
                            location: 'ty',
                            quantityOnHand: p.qty,
                            quantityReserved: 0,
                            quantityAvailable: p.qty,
                            distributorId: p.owner,
                            averageCost: chem.costPrice,
                            lastCost: chem.costPrice
                        });
                        await inv.save();
                    } else {
                        inv.quantityOnHand = p.qty;
                        inv.quantityAvailable = p.qty - inv.quantityReserved;
                        inv.distributorId = p.owner;
                        inv.updatedAt = new Date();
                        await inv.save();
                    }
                }
            }

            // Set Ty's Sims inventory
            for (const p of tySimsProducts) {
                const chem = await Chemical.findOne({ productName: p.name });
                if (chem) {
                    let inv = await Inventory.findOne({ chemicalId: chem._id, location: 'ty' });
                    if (!inv) {
                        // Check if there's a 'main' record and reassign it
                        inv = await Inventory.findOne({ chemicalId: chem._id, location: 'main' });
                        if (inv) {
                            inv.location = 'ty';
                            inv.quantityOnHand = p.qty;
                            inv.quantityAvailable = p.qty - inv.quantityReserved;
                            inv.distributorId = p.owner;
                            inv.updatedAt = new Date();
                            await inv.save();
                        } else {
                            inv = new Inventory({
                                chemicalId: chem._id,
                                productName: chem.productName,
                                packSize: chem.packSize,
                                unit: chem.unit,
                                location: 'ty',
                                quantityOnHand: p.qty,
                                quantityReserved: 0,
                                quantityAvailable: p.qty,
                                distributorId: p.owner,
                                averageCost: chem.costPrice,
                                lastCost: chem.costPrice
                            });
                            await inv.save();
                        }
                    } else {
                        inv.quantityOnHand = p.qty;
                        inv.quantityAvailable = p.qty - inv.quantityReserved;
                        inv.distributorId = p.owner;
                        inv.updatedAt = new Date();
                        await inv.save();
                    }
                }
            }
            console.log('Inventory quantities corrected to match physical count');
        }
    } catch (e) { console.error('Inventory correction error:', e.message); }

    // Assign inventory ownership: JABCO products → Kyle, Sims products → Ty
    try {
        const kyle = await User.findOne({ email: 'office@togoag.com' });
        const ty = await User.findOne({ email: 'tymollohan77@gmail.com' });

        if (kyle && ty) {
            // JABCO products belong to Kyle
            const jabcoChemicals = await Chemical.find({
                sourceSupplier: { $in: ['JABCO', 'Jabco', 'CPD', 'Crop Protect Direct'] }
            });
            for (const chem of jabcoChemicals) {
                await Inventory.updateMany(
                    { chemicalId: chem._id, distributorId: { $exists: false } },
                    { $set: { distributorId: kyle._id, updatedAt: new Date() } }
                );
                await Inventory.updateMany(
                    { chemicalId: chem._id, distributorId: null },
                    { $set: { distributorId: kyle._id, updatedAt: new Date() } }
                );
            }

            // Sims products belong to Ty
            const simsChemicals = await Chemical.find({
                sourceSupplier: { $in: ['Sims Fertilizer & Chemical', 'Sims', 'sims'] }
            });
            for (const chem of simsChemicals) {
                await Inventory.updateMany(
                    { chemicalId: chem._id, distributorId: { $exists: false } },
                    { $set: { distributorId: ty._id, updatedAt: new Date() } }
                );
                await Inventory.updateMany(
                    { chemicalId: chem._id, distributorId: null },
                    { $set: { distributorId: ty._id, updatedAt: new Date() } }
                );
            }

            // Hydrovant fA → Kyle (direct purchase)
            const hydrovant = await Chemical.findOne({ productName: /hydrovant/i });
            if (hydrovant) {
                await Inventory.updateMany(
                    { chemicalId: hydrovant._id, distributorId: null },
                    { $set: { distributorId: kyle._id, updatedAt: new Date() } }
                );
            }

            console.log('Inventory ownership assigned: JABCO/CPD → Kyle, Sims → Ty');
        }
    } catch (e) { console.error('Inventory ownership assignment error:', e.message); }

    app.listen(PORT, () => {
        console.log(`Acre Profit API running on port ${PORT}`);

        // Initialize Graph client if configured
        if (process.env.MICROSOFT_CLIENT_ID) {
            initGraphClient();
        }

        // Set up scheduled price sync if cron is available
        if (cron && process.env.MICROSOFT_CLIENT_ID) {
            const cronSchedule = process.env.PRICE_SYNC_CRON || '0 6 */3 * *';
            cron.schedule(cronSchedule, async () => {
                console.log('Running scheduled price sync...');
                try {
                    const result = await syncFromSharePoint(null);
                    console.log(`Scheduled sync completed: ${result.productsUpdated} updated, ${result.productsAdded} added`);
                } catch (error) {
                    console.error('Scheduled sync failed:', error.message);
                }
            });
            console.log(`Price sync scheduled: ${cronSchedule}`);
        }
    });
});

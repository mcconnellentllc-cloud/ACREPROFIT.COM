require('dotenv').config();
const express = require('express');
const cors = require('cors');
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
const JWT_SECRET = process.env.JWT_SECRET || 'acreprofit-secret-key-change-in-production';

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));

// MongoDB Connection
const connectDB = async () => {
    try {
        if (process.env.MONGODB_URI) {
            await mongoose.connect(process.env.MONGODB_URI);
            console.log('MongoDB connected successfully');
            // Initialize admin users
            await initializeAdmins();
            // Seed initial inventory
            await seedJabcoInventory();
            // Seed March 2026 purchase orders
            await seedMarch2026PurchaseOrders();
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
    // Supplier-specific fields
    companyName: String, // For suppliers - company/business name
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
        enum: ['draft', 'submitted', 'confirmed', 'ordered', 'shipped', 'delivered', 'archived', 'cancelled', 'quote_pending', 'quote_sent'],
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
    supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'Supplier' },

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
        enum: ['draft', 'submitted', 'confirmed', 'ordered_from_supplier', 'received', 'ready_for_pickup', 'delivered', 'cancelled', 'archived'],
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

    // Notes
    customerNotes: String,
    internalNotes: String,

    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

// Auto-generate order number
chemicalOrderSchema.pre('save', async function(next) {
    if (!this.orderNumber) {
        const count = await mongoose.model('ChemicalOrder').countDocuments();
        this.orderNumber = `CO-${new Date().getFullYear()}-${String(count + 1).padStart(5, '0')}`;
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
        enum: ['order', 'payment', 'supplier_payment', 'commission', 'adjustment', 'refund'],
        default: 'adjustment'
    },
    referenceType: { type: String, enum: ['ChemicalOrder', 'Order', 'Manual'], default: 'Manual' },
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
async function createLedgerEntry({ representativeId, description, amount, type, category, referenceType, referenceId, createdBy, notes }) {
    const lastEntry = await LedgerEntry.findOne({ representativeId })
        .sort({ date: -1, createdAt: -1 })
        .lean();

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

    await entry.save();
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
        chemicals: [{
            chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
            productName: String,
            suggestedRate: Number, // Use "suggested" not "recommended"
            rateUnit: String, // oz/acre, pt/acre, qt/acre, gal/acre, lb/acre
            packSize: String,
            unit: String,
            notes: String // e.g., "Adjust based on weed pressure"
        }]
    }],

    // Cost estimate per acre (calculated)
    estimatedCostPerAcre: Number,

    // Status
    isActive: { type: Boolean, default: true },

    // Owner
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },

    // Metadata
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const SprayProgram = mongoose.model('SprayProgram', sprayProgramSchema);

// Merch Order Model
const merchOrderSchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    items: [{
        productId: String,
        name: String,
        size: String,
        quantity: Number,
        price: Number
    }],
    subtotal: Number,
    creditApplied: { type: Number, default: 0 },
    totalDue: Number,
    shippingAddress: {
        name: String,
        address: String,
        city: String,
        state: String,
        zip: String
    },
    printifyOrderId: String, // Printify order ID once submitted
    status: {
        type: String,
        enum: ['pending', 'submitted', 'production', 'shipped', 'delivered'],
        default: 'pending'
    },
    trackingNumber: String,
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const MerchOrder = mongoose.model('MerchOrder', merchOrderSchema);

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

        // For tracking splits
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
const supplierSchema = new mongoose.Schema({
    name: { type: String, required: true },
    contact: String,
    phone: String,
    email: String,
    address: {
        street: String,
        city: String,
        state: String,
        zip: String
    },
    notes: String,
    isActive: { type: Boolean, default: true },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

supplierSchema.index({ name: 1 });
const Supplier = mongoose.model('Supplier', supplierSchema);

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
        const count = await QuoteRequest.countDocuments();
        this.quoteNumber = `QR-${year}-${String(count + 1).padStart(5, '0')}`;
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
        supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'Supplier' },
        supplierName: String,
        contactEmail: String,
        contactPhone: String,
        invitedAt: Date,
        status: {
            type: String,
            enum: ['invited', 'viewed', 'responded', 'declined', 'no_response'],
            default: 'invited'
        }
    }],

    // Supplier responses/bids
    supplierBids: [{
        supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'Supplier' },
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
    awardedSupplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'Supplier' },
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
        const count = await SupplierBidSheet.countDocuments();
        this.bidNumber = `BID-${year}-${String(count + 1).padStart(5, '0')}`;
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
    supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'Supplier' },
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
        quantity: Number,
        unitPrice: Number,
        totalPrice: Number
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
    const prefix = `INV-${year}-`;

    const lastInvoice = await Invoice.findOne({ invoiceNumber: { $regex: `^${prefix}` } })
        .sort({ invoiceNumber: -1 })
        .lean();

    let nextNum = 1;
    if (lastInvoice && lastInvoice.invoiceNumber) {
        const lastNum = parseInt(lastInvoice.invoiceNumber.split('-')[2], 10);
        if (!isNaN(lastNum)) {
            nextNum = lastNum + 1;
        }
    }

    return `${prefix}${String(nextNum).padStart(5, '0')}`;
}

// Helper: Update inventory when receiving a PO
async function receiveInventory({ chemicalId, productName, packSize, unit, quantity, unitCost, location, purchaseOrderId, poNumber, lotNumber, supplierName, userId }) {
    // Find or create inventory record (aggregate tracking)
    let inventory = await Inventory.findOne({ chemicalId, location: location || 'main' });

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

    // Calculate new weighted average cost
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

    await inventory.save();

    // Create batch record for PO-based tracking
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

    await batch.save();

    // Create transaction record
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

    await transaction.save();

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
async function reserveInventory({ chemicalId, quantity, location, orderId, orderNumber, userId, notes }) {
    let inventory = await Inventory.findOne({ chemicalId, location: location || 'main' });

    if (!inventory) {
        // No inventory record yet - create one with zero quantities
        // This tracks the reservation even before stock arrives
        const chemical = await Chemical.findById(chemicalId);
        inventory = new Inventory({
            chemicalId,
            productName: chemical?.name || 'Unknown Product',
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

    await inventory.save();

    // Create transaction record for audit trail
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

    await transaction.save();

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
        {
            name: 'Acre Profit Admin',
            email: 'contact@acreprofit.com',
            password: 'Farm2026!',
            phone: '970-571-1015',
            role: 'superadmin'
        },
        {
            name: 'Kyle McConnell',
            email: 'office@togoag.com',
            password: 'Farm2026!',
            phone: '970-571-1015',
            role: 'distributor'
        },
        {
            name: 'Ty Mollohan',
            email: 'tymollohan77@gmail.com',
            password: 'Farm2026!',
            phone: '970-520-2340',
            role: 'distributor'
        },
        {
            name: 'Chad Bamford',
            email: 'ckbamford@yahoo.com',
            password: 'Farm2026!',
            phone: '970-520-3716',
            role: 'distributor'
        },
        {
            name: 'Seth Rolfs',
            email: 'seth@acreprofit.com',
            password: 'Farm2026!',
            phone: '785-531-0680',
            role: 'distributor'
        }
    ];

    for (const admin of admins) {
        try {
            const existing = await User.findOne({ email: admin.email });
            if (!existing) {
                await User.create(admin);
                console.log(`Created admin: ${admin.name}`);
            }
        } catch (error) {
            console.log(`Admin ${admin.email} may already exist`);
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
                adminPrice: 14.50,
                sellPrice: 16.00,
                category: 'herbicide',
                sourceSupplier: 'Jabco',
                epaRegistrationNumber: '89343-5',
                signalWord: 'CAUTION',
                notes: '5.4 lb/gal glyphosate - Xingfa USA',
                activeIngredients: [{ name: 'Glyphosate', percentage: 53.8, poundsPerGallon: 5.4 }]
            });
            console.log('Created XSATE Glyphosate 53.8% product');
        } else if (product.sellPrice === 0) {
            // Update prices if not set
            product.adminPrice = 14.50;
            product.sellPrice = 16.00;
            await product.save();
            console.log('Updated XSATE Glyphosate prices');
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
                    productName: 'CPD Mesotrione',
                    packSize: '2x2.5 GL Case',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 720,
                    unitPrice: 45.75,
                    lineTotal: 32940.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Flumioxazin 51%',
                    packSize: '4x5 Lb Case',
                    unit: 'lb',
                    unitsPerPack: 20,
                    qty: 1440,
                    unitPrice: 14.00,
                    lineTotal: 20160.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Sulfentrazone 4SC',
                    packSize: '2x2.5 Gl Case',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 180,
                    unitPrice: 71.50,
                    lineTotal: 12870.00,
                    category: 'herbicide'
                },
                {
                    productName: 'Dicamba DMA',
                    packSize: '2x2.5 Gl Case',
                    unit: 'gal',
                    unitsPerPack: 5,
                    qty: 180,
                    unitPrice: 30.25,
                    lineTotal: 5445.00,
                    category: 'herbicide'
                },
                {
                    productName: '2,4-D LV6',
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
                        adminPrice: prod.unitPrice * 1.05, // 5% admin margin
                        sellPrice: prod.unitPrice * 1.15, // 15% retail margin
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
                status: 'confirmed',
                orderDate: new Date('2026-03-30'),
                notes: 'Billed to: Acre Profit LLC'
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
                        adminPrice: prod.unitPrice * 1.05,
                        sellPrice: prod.unitPrice * 1.15,
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
                status: 'confirmed',
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
                    adminPrice: 13.00,
                    sellPrice: 14.50,
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
                status: 'confirmed',
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

// ============ SPRAY PROGRAMS DATA ============

const sprayPrograms = {
    corn: {
        '2-pass': {
            name: '2-Pass Corn Program',
            description: 'Pre-emergent + Post-emergent application',
            chemicals: [
                { name: 'Glyphosate', defaultRate: 32, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Atrazine 4L', defaultRate: 1.5, rateUnit: 'qt/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Metolachlor', defaultRate: 1.3, rateUnit: 'pt/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 }
            ]
        },
        '3-pass': {
            name: '3-Pass Corn Program',
            description: 'Burndown + Pre-emergent + Post-emergent application',
            chemicals: [
                { name: 'Glyphosate', defaultRate: 32, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
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
                { name: 'Glyphosate', defaultRate: 32, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Metribuzin', defaultRate: 0.5, rateUnit: 'lb/acre', packageSize: 50, packageUnit: 'lb', pricePerPackage: 0 }
            ]
        }
    },
    'dryland-corn': {
        'preplant': {
            name: 'Option 1 - Corn Preplant',
            description: '30 days pre-plant - wheat stubble with atrazine & valor',
            chemicals: [
                { name: 'Glyphosate', defaultRate: 22, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Valor SX', defaultRate: 3, rateUnit: 'oz/acre', packageSize: 5, packageUnit: 'lb', pricePerPackage: 0 },
                { name: 'Atrazine 4L', defaultRate: 1, rateUnit: 'lb/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Hydrovant', defaultRate: 0.1, rateUnit: '% v/v', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0, isAdjuvant: true }
            ]
        },
        'post-plant-pre-emerge': {
            name: 'Option 2 - Corn Post Plant Pre-Emerge',
            description: 'Post plant pre-emerge - fall atrazine already applied',
            chemicals: [
                { name: 'Glyphosate', defaultRate: 22, rateUnit: 'oz/acre', packageSize: 250, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Mesotrione', defaultRate: 6, rateUnit: 'oz/acre', packageSize: 1, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Dicamba DMA', defaultRate: 4, rateUnit: 'oz/acre', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Anthem Max', defaultRate: 3, rateUnit: 'oz/acre', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0 },
                { name: 'Hydrovant', defaultRate: 0.1, rateUnit: '% v/v', packageSize: 2.5, packageUnit: 'gal', pricePerPackage: 0, isAdjuvant: true }
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
app.post('/api/admin/reset-admins', async (req, res) => {
    try {
        const { secretKey } = req.body;

        // Simple secret key protection
        if (secretKey !== 'acreprofit2026reset') {
            return res.status(403).json({ error: 'Invalid secret key' });
        }

        const admins = [
            { name: 'Acre Profit Admin', email: 'contact@acreprofit.com', password: 'Farm2026!', phone: '970-571-1015', role: 'superadmin' },
            { name: 'Kyle McConnell', email: 'office@togoag.com', password: 'Farm2026!', phone: '970-571-1015', role: 'distributor' },
            { name: 'Ty Mollohan', email: 'tymollohan77@gmail.com', password: 'Farm2026!', phone: '970-520-2340', role: 'distributor' },
            { name: 'Chad Bamford', email: 'ckbamford@yahoo.com', password: 'Farm2026!', phone: '970-520-3716', role: 'distributor' },
            { name: 'Seth Rolfs', email: 'seth@acreprofit.com', password: 'Farm2026!', phone: '785-531-0680', role: 'distributor' }
        ];

        const results = [];
        for (const admin of admins) {
            // Delete existing user if exists
            await User.deleteOne({ email: admin.email.toLowerCase() });
            // Create fresh
            const user = new User(admin);
            await user.save();
            results.push(`Created/reset: ${admin.email}`);
        }

        res.json({ message: 'Admin users reset successfully', results });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// ---- AUTH ROUTES ----

app.post('/api/auth/signup', async (req, res) => {
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

app.post('/api/auth/login', async (req, res) => {
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
            token
        });
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
app.post('/api/auth/forgot-password', async (req, res) => {
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
app.post('/api/auth/reset-password', async (req, res) => {
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

        await req.user.save();
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

        // Get order count and recent orders
        const orders = await Order.find({ userId: customer._id })
            .sort({ createdAt: -1 })
            .limit(10);

        res.json({
            ...customer.toObject(),
            orders,
            orderCount: await Order.countDocuments({ userId: customer._id })
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

        const orders = await Order.find({ userId: customer._id })
            .populate('representativeId', 'name email')
            .sort({ createdAt: -1 });

        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create order on behalf of a customer (admin only)
app.post('/api/admin/orders/for-customer', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { customerId, crop, programId, acres, gpa, year, chemicals, seeds, pivotBio, totalPrice, status, notes } = req.body;

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

        const costPerAcre = acres > 0 ? Math.round((totalPrice / acres) * 100) / 100 : 0;

        // Normalize chemicals array to use consistent field names
        const normalizedChemicals = (chemicals || []).map(c => ({
            name: c.name || c.productName,
            qty: c.qty || c.quantity || 0,
            unit: c.unit || 'gal',
            pricePerUnit: c.pricePerUnit || c.price || 0,
            totalPrice: c.totalPrice || c.total || (c.qty || c.quantity || 0) * (c.pricePerUnit || c.price || 0),
            chemicalId: c.chemicalId,
            packSize: c.packSize,
            sourceSupplier: c.sourceSupplier
        }));

        const order = new Order({
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
            totalCost: totalPrice,
            costPerAcre,
            status: status || 'draft',
            notes,
            createdBy: req.user._id // Track who created this order
        });

        await order.save();

        // Reserve inventory and calculate commissions for each chemical
        const chemicalIds = (chemicals || []).map(c => c.chemicalId).filter(Boolean);
        if (chemicalIds.length > 0) {
            try {
                // Fetch chemical pricing data for commission calculation
                const chemicalPricing = await Chemical.find({ _id: { $in: chemicalIds } })
                    .select('productName costPrice adminPrice sellPrice marginDollars adminMarginDollars');

                const pricingMap = {};
                chemicalPricing.forEach(c => {
                    pricingMap[c._id.toString()] = c;
                });

                let totalRepCommission = 0;
                let totalAdminRevenue = 0;

                // Reserve inventory and calculate commissions for each item
                for (const chem of normalizedChemicals) {
                    const qty = chem.qty || chem.quantity || 0;
                    if (chem.chemicalId && qty > 0) {
                        // Reserve inventory
                        try {
                            await reserveInventory({
                                chemicalId: chem.chemicalId,
                                quantity: qty,
                                location: 'main',
                                orderId: order._id,
                                orderNumber: `ORD-${order._id.toString().slice(-8).toUpperCase()}`,
                                userId: req.user._id,
                                notes: `Reserved for order - ${crop}`
                            });
                        } catch (invErr) {
                            console.warn('Inventory reservation warning:', invErr.message);
                            // Continue even if reservation fails (might not have inventory records yet)
                        }

                        // Calculate commission from pricing data
                        const pricing = pricingMap[chem.chemicalId.toString()];
                        if (pricing) {
                            // Rep commission = marginDollars * quantity
                            totalRepCommission += (pricing.marginDollars || 0) * qty;
                            // Admin revenue = adminMarginDollars * quantity
                            totalAdminRevenue += (pricing.adminMarginDollars || 0) * qty;
                        }
                    }
                }

                // Update order with commission data
                if (totalRepCommission > 0 || totalAdminRevenue > 0) {
                    order.repCommission = totalRepCommission;
                    order.adminRevenue = totalAdminRevenue;
                    await order.save();
                }
            } catch (err) {
                console.error('Error processing inventory/commissions:', err.message);
                // Don't fail order creation if this fails
            }
        }

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

        const order = await Order.findOne(query)
            .populate('userId', 'name email phone farm')
            .populate('representativeId', 'name email');

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

        const orders = await Order.find(query)
            .populate('userId', 'name email phone farm')
            .populate('representativeId', 'name email')
            .sort({ createdAt: -1 });
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

// Get order stats (admin only)
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

// ---- MERCH ORDER ROUTES ----

// Printify product mapping (maps our product IDs to Printify blueprint IDs)
// This would be configured once you set up your Printify store
const PRINTIFY_PRODUCTS = {
    'chore-coat': { blueprintId: '6', printProviderId: '99' },
    'quarter-zip': { blueprintId: '578', printProviderId: '99' },
    'full-zip-hoodie': { blueprintId: '77', printProviderId: '99' },
    'pullover-hoodie': { blueprintId: '77', printProviderId: '99' },
    'trucker-green': { blueprintId: '380', printProviderId: '99' },
    'trucker-black': { blueprintId: '380', printProviderId: '99' },
    'trucker-camo': { blueprintId: '380', printProviderId: '99' },
    'fitted-black': { blueprintId: '381', printProviderId: '99' },
    'beanie': { blueprintId: '432', printProviderId: '99' },
    'classic-tee-green': { blueprintId: '5', printProviderId: '99' },
    'classic-tee-black': { blueprintId: '5', printProviderId: '99' },
    'farmers-tee': { blueprintId: '5', printProviderId: '99' }
};

// Submit merch order
app.post('/api/merch-orders', authMiddleware, async (req, res) => {
    try {
        const { items, subtotal, creditApplied, totalDue, shippingAddress } = req.body;

        // Create order in database
        const merchOrder = new MerchOrder({
            userId: req.user._id,
            items,
            subtotal,
            creditApplied,
            totalDue,
            shippingAddress,
            status: 'pending'
        });

        await merchOrder.save();

        // TODO: Integrate with Printify API
        // When PRINTIFY_API_KEY is configured, this will auto-submit to Printify
        if (process.env.PRINTIFY_API_KEY && process.env.PRINTIFY_SHOP_ID) {
            try {
                // Submit to Printify (example structure)
                const printifyOrder = await submitToPrintify(merchOrder, req.user);
                merchOrder.printifyOrderId = printifyOrder.id;
                merchOrder.status = 'submitted';
                await merchOrder.save();
            } catch (printifyError) {
                console.error('Printify submission failed:', printifyError);
                // Order saved but not submitted to Printify - manual intervention needed
            }
        }

        res.status(201).json({
            message: 'Order placed successfully',
            orderId: merchOrder._id,
            status: merchOrder.status
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get user's merch orders
app.get('/api/merch-orders', authMiddleware, async (req, res) => {
    try {
        const orders = await MerchOrder.find({ userId: req.user._id })
            .sort({ createdAt: -1 });
        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Admin: Get all merch orders
app.get('/api/admin/merch-orders', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const orders = await MerchOrder.find()
            .populate('userId', 'name email')
            .sort({ createdAt: -1 });
        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Printify submission helper (implement when API key available)
async function submitToPrintify(merchOrder, user) {
    const PRINTIFY_API_KEY = process.env.PRINTIFY_API_KEY;
    const PRINTIFY_SHOP_ID = process.env.PRINTIFY_SHOP_ID;

    // Build Printify order payload
    const lineItems = merchOrder.items.map(item => ({
        product_id: PRINTIFY_PRODUCTS[item.productId]?.blueprintId,
        variant_id: 1, // Would need to map sizes to variant IDs
        quantity: item.quantity
    }));

    const response = await fetch(`https://api.printify.com/v1/shops/${PRINTIFY_SHOP_ID}/orders.json`, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${PRINTIFY_API_KEY}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            external_id: merchOrder._id.toString(),
            line_items: lineItems,
            shipping_method: 1,
            address_to: {
                first_name: user.name.split(' ')[0],
                last_name: user.name.split(' ').slice(1).join(' ') || '',
                email: user.email,
                phone: user.phone || '',
                country: 'US',
                region: merchOrder.shippingAddress?.state || '',
                address1: merchOrder.shippingAddress?.address || '',
                city: merchOrder.shippingAddress?.city || '',
                zip: merchOrder.shippingAddress?.zip || ''
            }
        })
    });

    if (!response.ok) {
        throw new Error('Printify API error');
    }

    return response.json();
}

// Printify webhook for order updates
app.post('/api/webhooks/printify', async (req, res) => {
    try {
        const { type, resource } = req.body;

        if (type === 'order:shipment:created') {
            const merchOrder = await MerchOrder.findOne({ printifyOrderId: resource.id });
            if (merchOrder) {
                merchOrder.status = 'shipped';
                merchOrder.trackingNumber = resource.shipments?.[0]?.tracking_number;
                merchOrder.updatedAt = new Date();
                await merchOrder.save();
            }
        }

        res.json({ received: true });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

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

        const order = await Order.findById(orderId);
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        // Verify the rep owns this order (unless superadmin)
        if (req.user.role !== 'superadmin' &&
            order.representativeId.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        order.paymentMethod = 'check';
        order.paymentStatus = 'paid';
        order.checkNumber = checkNumber;
        order.checkReceivedDate = new Date();
        order.paidAt = new Date();
        order.status = 'confirmed';
        order.updatedAt = new Date();
        await order.save();

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
            { productName: 'Dicamba DMA', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 30.25, sellPrice: 0, category: 'herbicide' },
            { productName: 'Dicamba DMA', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 28.25, sellPrice: 0, category: 'herbicide' },
            { productName: 'Dicamba HD', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 30.57, sellPrice: 0, category: 'herbicide' },

            // LV 6 (2,4-D)
            { productName: 'LV 6', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 29.90, sellPrice: 0, category: 'herbicide' },
            { productName: 'LV 6', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 27.90, sellPrice: 0, category: 'herbicide' },

            // Glyphosate - AgSaver is CPD equiv for RT3/Glystar Supreme
            { productName: 'AgSaver', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 13.32, sellPrice: 0, category: 'herbicide', equivalentProduct: 'RT3, Glystar Supreme', notes: 'Glyphosate' },

            // Glyphosate 5.4 lb - from JABCO/CPD
            { productName: 'Glyphosate 5.4', packSize: 'Tote', unit: 'gal', unitsPerPack: 250, costPrice: 13.25, adminPrice: 13.32, sellPrice: 14.09, category: 'herbicide', notes: '5.4 lb/gal glyphosate' },

            // XSATE Glyphosate 53.8% - Xingfa USA via Jabco (EPA 89343-5) - $13.25/gal cost, $16/gal retail
            { productName: 'XSATE Glyphosate 53.8%', packSize: '265 gal', unit: 'gal', unitsPerPack: 265, costPrice: 13.25, adminPrice: 14.50, sellPrice: 16.00, category: 'herbicide', sourceSupplier: 'Jabco', epaRegistrationNumber: '89343-5', signalWord: 'CAUTION', notes: '5.4 lb/gal glyphosate - Xingfa USA', activeIngredients: [{ name: 'Glyphosate', percentage: 53.8, poundsPerGallon: 5.4 }] },
            { productName: 'XSATE Glyphosate 53.8%', packSize: '30 gal', unit: 'gal', unitsPerPack: 30, costPrice: 13.25, adminPrice: 14.50, sellPrice: 16.00, category: 'herbicide', sourceSupplier: 'Jabco', epaRegistrationNumber: '89343-5', signalWord: 'CAUTION', notes: '5.4 lb/gal glyphosate - Xingfa USA', activeIngredients: [{ name: 'Glyphosate', percentage: 53.8, poundsPerGallon: 5.4 }] },
            { productName: 'XSATE Glyphosate 53.8%', packSize: '2.5 gal', unit: 'gal', unitsPerPack: 2.5, costPrice: 13.25, adminPrice: 14.50, sellPrice: 16.00, category: 'herbicide', sourceSupplier: 'Jabco', epaRegistrationNumber: '89343-5', signalWord: 'CAUTION', notes: '5.4 lb/gal glyphosate - Xingfa USA', activeIngredients: [{ name: 'Glyphosate', percentage: 53.8, poundsPerGallon: 5.4 }] },

            // Atrazine
            { productName: 'Aatrex', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 13.35, sellPrice: 0, category: 'herbicide' },

            // Agri-Star is CPD equiv for Level Best Pro
            { productName: 'Agri-Star', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 39.25, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Level Best Pro' },
            { productName: 'Agri-Star', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 38.13, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Level Best Pro' },

            // Agri-Star Tapran is CPD equiv for Tapran
            { productName: 'Agri-Star Tapran', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 19.00, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Tapran' },
            { productName: 'Agri-Star Tapran', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 18.00, sellPrice: 0, category: 'herbicide', equivalentProduct: 'Tapran' },

            // Aggrestrol
            { productName: 'Aggrestrol', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 21.00, sellPrice: 0, category: 'herbicide' },
            { productName: 'Aggrestrol', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 20.00, sellPrice: 0, category: 'herbicide' },

            // Sulfentrazone
            { productName: 'Sulfentrazone', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 70.50, sellPrice: 0, category: 'herbicide' },

            // Valor SX
            { productName: 'Valor SX', packSize: '4x5', unit: 'lb', unitsPerPack: 20, costPrice: 14.25, sellPrice: 0, category: 'herbicide' },

            // CPD-only products
            { productName: 'Glufosinate', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 16.00, sellPrice: 0, category: 'herbicide' },
            { productName: 'Paraquat', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 15.25, sellPrice: 0, category: 'herbicide', isRUP: true, notes: 'Restricted Use Pesticide - requires certification' },
            { productName: 'Mesotrione', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 48.25, sellPrice: 0, category: 'herbicide' },
            { productName: 'Clethodim', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 33.50, sellPrice: 0, category: 'herbicide' },
            { productName: 'Clethodim', packSize: '135', unit: 'gal', unitsPerPack: 135, costPrice: 33.00, sellPrice: 0, category: 'herbicide' },

            // AMS (Ammonium Sulfate)
            { productName: 'AMS', packSize: '24 lb', unit: 'lb', unitsPerPack: 24, costPrice: 1.45, sellPrice: 0, category: 'adjuvant', notes: 'Ammonium Sulfate - water conditioner/adjuvant' },
            { productName: 'AMS x5', packSize: '24', unit: 'lb', unitsPerPack: 120, costPrice: 1.45, sellPrice: 0, category: 'adjuvant', notes: 'Ammonium Sulfate - 5 bag bundle' },

            // Hydrovant
            { productName: 'Hydrovant', packSize: '2x2.5', unit: 'gal', unitsPerPack: 5, costPrice: 95.00, sellPrice: 145.00, category: 'adjuvant', notes: 'Drift reduction/deposition aid adjuvant' },
            { productName: 'Hydrovant', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 95.00, sellPrice: 145.00, category: 'adjuvant', notes: 'Drift reduction/deposition aid adjuvant' }
        ];

        const results = { created: [], existing: [] };

        // Insert CPD products
        for (const chem of cpdProducts) {
            const existing = await Chemical.findOne({
                productName: chem.productName,
                sourceSupplier: 'CPD',
                packSize: chem.packSize
            });

            if (!existing) {
                // Calculate prices with margins if not set
                const cost = chem.costPrice || 0;
                const admin = chem.adminPrice || cost * 1.10; // 10% Acre Profit margin
                const sell = chem.sellPrice || admin * 1.15; // 15% rep margin on top

                const newChem = await Chemical.create({
                    ...chem,
                    costPrice: cost,
                    adminPrice: admin,
                    sellPrice: sell,
                    margin: sell > 0 ? Math.round(((sell - cost) / sell) * 100) : 0,
                    sourceSupplier: 'CPD',
                    priceVersion,
                    isActive: true,
                    availableForOrder: true,
                    createdBy: req.user._id
                });
                await ChemicalPriceHistory.create({
                    chemicalId: newChem._id,
                    productName: chem.productName,
                    sourceSupplier: 'CPD',
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
                    const admin = chem.adminPrice || cost * 1.10;
                    const sell = chem.sellPrice || admin * 1.15;

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

        // For public view - ONLY show products with valid retail pricing
        // NEVER expose wholesale/cost pricing to public
        const publicChemicals = chemicals
            .filter(c => c.sellPrice > 0) // Only show products with retail price set
            .map(c => ({
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
                availableForOrder: c.availableForOrder
            }));

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

// Get all chemicals with FULL pricing (admin only)
app.get('/api/chemicals/admin', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { sourceSupplier, productName, category } = req.query;
        let query = {};

        if (sourceSupplier) query.sourceSupplier = new RegExp(sourceSupplier, 'i');
        if (productName) query.productName = new RegExp(productName, 'i');
        if (category) query.category = category;

        const chemicals = await Chemical.find(query)
            .sort({ productName: 1, packSize: 1 });

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

        res.status(201).json(chemical);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update chemical price (admin only)
app.put('/api/chemicals/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { costPrice, adminPrice, sellPrice, adminMarginDollars, marginDollars, priceVersion, notes, equivalentProduct, isActive, availableForOrder,
                category, crops, defaultRate, rateUnit, unitsPerPack } = req.body;

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
        if (category !== undefined) chemical.category = category;
        if (crops !== undefined) chemical.crops = crops;
        if (defaultRate !== undefined) chemical.defaultRate = defaultRate;
        if (rateUnit !== undefined) chemical.rateUnit = rateUnit;
        if (unitsPerPack !== undefined) chemical.unitsPerPack = unitsPerPack;

        chemical.updatedAt = new Date();
        await chemical.save();

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
        const chemical = await Chemical.findByIdAndDelete(req.params.id);
        if (!chemical) {
            return res.status(404).json({ error: 'Chemical not found' });
        }
        res.json({ message: 'Chemical deleted', chemical });
    } catch (error) {
        res.status(400).json({ error: error.message });
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

// Admin: Get all suppliers
app.get('/api/admin/suppliers', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const suppliers = await User.find({ role: 'supplier' })
            .select('name email companyName supplierCode phone createdAt')
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
        const { name, companyName, phone, email } = req.body;

        const supplier = await User.findOne({ _id: req.params.id, role: 'supplier' });
        if (!supplier) {
            return res.status(404).json({ error: 'Supplier not found' });
        }

        if (name) supplier.name = name;
        if (companyName) supplier.companyName = companyName;
        if (phone) supplier.phone = phone;
        if (email) supplier.email = email.toLowerCase();

        await supplier.save();

        res.json({
            message: 'Supplier updated successfully',
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
        const { items, programId, programName, totalAcres, customerNotes } = req.body;

        // Calculate totals
        let subtotal = 0;
        const orderItems = [];

        for (const item of items) {
            const chemical = await Chemical.findById(item.chemicalId);
            if (!chemical) continue;

            const totalPrice = item.quantity * chemical.sellPrice;
            subtotal += totalPrice;

            orderItems.push({
                chemicalId: chemical._id,
                productName: chemical.productName,
                packSize: chemical.packSize,
                unit: chemical.unit,
                quantity: item.quantity,
                unitPrice: chemical.sellPrice,
                totalPrice,
                acres: item.acres,
                rate: item.rate,
                rateUnit: item.rateUnit,
                calculatedAmount: item.calculatedAmount
            });
        }

        const order = new ChemicalOrder({
            userId: req.user._id,
            representativeId: req.user.representative,
            orderType: programId ? 'program' : 'direct',
            items: orderItems,
            programId,
            programName,
            totalAcres,
            subtotal,
            total: subtotal,
            customerNotes,
            status: 'draft'
        });

        await order.save();
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
            notes
        } = req.body;

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

        // Build order items
        const orderItems = items.map(item => ({
            productName: item.productName || item.name,
            packSize: item.packSize || '',
            unit: item.unit || 'pack',
            quantity: item.qty || item.quantity || 1,
            unitPrice: item.price || 0,
            totalPrice: item.total || 0,
            timing: item.timing || '',
            isCustom: item.isCustom || false
        }));

        // Calculate totals if not provided
        const calculatedSubtotal = subtotal || items.reduce((sum, item) => sum + (item.total || 0), 0);
        const calculatedTotal = total || calculatedSubtotal + (processingFee || 0);

        // Create the order
        const order = new ChemicalOrder({
            userId: req.user._id,
            representativeId: representativeId,
            orderType: 'direct',
            items: orderItems,
            totalAcres: 0,
            subtotal: calculatedSubtotal,
            processingFee: processingFee || 0,
            total: calculatedTotal,
            customerNotes: notes,
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

        // Update user info if provided
        if (customerPhone && !req.user.phone) {
            req.user.phone = customerPhone;
            await req.user.save();
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

// Create manual ledger entry
app.post('/api/admin/ledger', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { representativeId, description, amount, type, category, notes } = req.body;

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

// ============ SUPPLIER ENDPOINTS ============

// Get all suppliers
app.get('/api/admin/suppliers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const suppliers = await Supplier.find({ isActive: true }).sort({ name: 1 });
        res.json(suppliers);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create a new supplier
app.post('/api/admin/suppliers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { name, contact, phone, email, address, notes } = req.body;

        if (!name) {
            return res.status(400).json({ error: 'Supplier name is required' });
        }

        const supplier = new Supplier({
            name,
            contact,
            phone,
            email,
            address,
            notes,
            createdBy: req.user._id
        });

        await supplier.save();
        res.status(201).json(supplier);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Update a supplier
app.put('/api/admin/suppliers/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { name, contact, phone, email, address, notes } = req.body;

        const supplier = await Supplier.findByIdAndUpdate(
            req.params.id,
            { name, contact, phone, email, address, notes, updatedAt: new Date() },
            { new: true }
        );

        if (!supplier) {
            return res.status(404).json({ error: 'Supplier not found' });
        }

        res.json(supplier);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Delete (deactivate) a supplier
app.delete('/api/admin/suppliers/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const supplier = await Supplier.findByIdAndUpdate(
            req.params.id,
            { isActive: false, updatedAt: new Date() },
            { new: true }
        );

        if (!supplier) {
            return res.status(404).json({ error: 'Supplier not found' });
        }

        res.json({ message: 'Supplier deleted' });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ PURCHASE ORDER ENDPOINTS ============

// Get all purchase orders
app.get('/api/admin/purchase-orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, supplier } = req.query;
        let query = {};

        if (status) query.status = status;
        if (supplier) query['supplier.name'] = new RegExp(supplier, 'i');

        const purchaseOrders = await PurchaseOrder.find(query)
            .populate('createdBy', 'name email')
            .populate('updatedBy', 'name email')
            .sort({ orderDate: -1 });

        res.json(purchaseOrders);
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
            role: { $in: ['admin', 'superadmin', 'distributor'] }
        }).select('name email role').sort({ name: 1 });

        res.json(distributors);
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
            programs.push({
                _id: `template-${cropKey}-${programKey}`,
                name: programData.name,
                description: programData.description,
                crop: cropKey,
                type: 'template',
                isPublic: true,
                isTemplate: true,
                applications: [{
                    name: programData.name,
                    chemicals: programData.chemicals.map(chem => ({
                        productName: chem.name,
                        suggestedRate: chem.defaultRate,
                        rateUnit: chem.rateUnit,
                        packSize: chem.packageSize,
                        unit: chem.packageUnit,
                        isAdjuvant: chem.isAdjuvant || false
                    }))
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
        const { name, description, crop, applications, isPublic } = req.body;

        // Validate required fields
        if (!name || !crop) {
            return res.status(400).json({ error: 'Name and crop are required' });
        }

        // Determine program type and visibility
        const isAdmin = isAdminLevel(req.user);
        const programType = isAdmin ? 'suggestion' : 'custom';
        // Only admins can create public programs
        const publicFlag = isAdmin ? (isPublic || false) : false;

        const program = new SprayProgram({
            name,
            description,
            crop,
            applications: applications || [],
            type: programType,
            isPublic: publicFlag,
            createdBy: req.user._id
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
            status: { $in: ['pending', 'submitted', 'approved', 'ordered'] }
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
        const { location = 'main' } = req.body;

        // Get all chemicals (active products)
        const chemicals = await Chemical.find({ isActive: { $ne: false } });

        // Get existing inventory records for this location
        const existingInventory = await Inventory.find({ location });

        // Build sets for matching - by chemicalId AND by productName+packSize
        const existingChemicalIds = new Set(existingInventory.map(i => i.chemicalId?.toString()).filter(Boolean));
        const existingProductKeys = new Set(existingInventory.map(i => `${i.productName?.toLowerCase()}_${i.packSize?.toLowerCase()}`));

        const created = [];
        const skipped = [];
        const errors = [];

        for (const chem of chemicals) {
            const chemIdStr = chem._id.toString();
            const productKey = `${chem.productName?.toLowerCase()}_${chem.packSize?.toLowerCase()}`;

            // Skip if already exists by chemicalId OR by product name + pack size
            if (existingChemicalIds.has(chemIdStr)) {
                skipped.push({ name: chem.productName, packSize: chem.packSize, reason: 'chemicalId exists' });
                continue;
            }
            if (existingProductKeys.has(productKey)) {
                skipped.push({ name: chem.productName, packSize: chem.packSize, reason: 'name+packSize exists' });
                continue;
            }

            try {
                const inventory = new Inventory({
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
                });

                await inventory.save();
                created.push(`${chem.productName} (${chem.packSize})`);
            } catch (err) {
                errors.push({ name: chem.productName, packSize: chem.packSize, error: err.message });
            }
        }

        res.json({
            success: true,
            message: `Synced ${created.length} products to inventory`,
            created,
            skipped: skipped.length,
            skippedDetails: skipped,
            errors: errors.length > 0 ? errors : undefined,
            totalProducts: chemicals.length
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

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

        // Update PO status
        po.status = 'received';
        po.receivedDate = new Date();
        po.updatedBy = req.user._id;
        po.updatedAt = new Date();
        if (notes) po.internalNotes = (po.internalNotes || '') + '\n' + notes;

        await po.save();

        res.json({
            message: 'Inventory received successfully',
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
        const chemicals = await Chemical.find({ isActive: true })
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
            // ChemicalOrder has items array
            invoiceItems = order.items.map(item => ({
                productName: item.productName,
                description: `${item.packSize} ${item.unit}`,
                packSize: item.packSize,
                unit: item.unit,
                quantity: item.quantity,
                unitPrice: item.unitPrice,
                totalPrice: item.totalPrice
            }));
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
            const supplier = row['Supplier'] || row['Source'] || 'CPD';

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
const poDocumentsPath = path.join(__dirname, '..', 'purchase-orders');

// Ensure upload directory exists
if (!fs.existsSync(poDocumentsPath)) {
    fs.mkdirSync(poDocumentsPath, { recursive: true });
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
            .select('productName sourceSupplier category packSize unit sellPrice') // NO costPrice for distributors
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
                // Base retail price (set by super admin)
                baseRetailPrice: c.sellPrice,
                // Distributor's custom retail price (if set)
                myRetailPrice: customPrice?.retailPrice || null,
                // Effective price (custom or base)
                effectivePrice: customPrice?.retailPrice || c.sellPrice,
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
        const { items, customerNotes, deliveryLocation, preferredDeliveryDate, customerId } = req.body;

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

        // Get current volume needs
        const volumeNeeds = {};
        const orderStatuses = ['pending', 'payment_pending', 'payment_secured', 'manufacturer_ordered'];
        const quoteStatuses = ['submitted', 'pricing', 'quoted', 'accepted'];

        // From orders
        const orders = await Order.find({ status: { $in: orderStatuses } }).select('chemicals');
        for (const order of orders) {
            for (const chem of order.chemicals || []) {
                const key = `${chem.name || chem.productName}|${chem.packSize || 'N/A'}|${chem.unit || 'unit'}`;
                if (!volumeNeeds[key]) {
                    volumeNeeds[key] = {
                        productName: chem.name || chem.productName,
                        packSize: chem.packSize || 'N/A',
                        unit: chem.unit || 'unit',
                        chemicalId: chem.chemicalId,
                        quantityNeeded: 0
                    };
                }
                volumeNeeds[key].quantityNeeded += (chem.qty || chem.quantity || chem.packagesNeeded || 0);
            }
        }

        // From quotes
        const quotes = await QuoteRequest.find({ status: { $in: quoteStatuses } }).select('items');
        for (const quote of quotes) {
            for (const item of quote.items || []) {
                const key = `${item.productName}|${item.packSize || 'N/A'}|${item.unit || 'unit'}`;
                if (!volumeNeeds[key]) {
                    volumeNeeds[key] = {
                        productName: item.productName,
                        packSize: item.packSize || 'N/A',
                        unit: item.unit || 'unit',
                        chemicalId: item.chemicalId,
                        quantityNeeded: 0
                    };
                }
                volumeNeeds[key].quantityNeeded += (item.quantityNeeded || 0);
            }
        }

        const items = Object.values(volumeNeeds).filter(i => i.quantityNeeded > 0);

        if (items.length === 0) {
            return res.status(400).json({ error: 'No pending volume needs found' });
        }

        const bidSheet = new SupplierBidSheet({
            title: title || `Volume Needs Bid - ${new Date().toLocaleDateString()}`,
            description: description || 'Auto-generated from pending orders and quote requests',
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

// Send bid sheet to suppliers (mark as sent)
app.put('/api/admin/bid-sheets/:id/send', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const bidSheet = await SupplierBidSheet.findById(req.params.id);

        if (!bidSheet) {
            return res.status(404).json({ error: 'Bid sheet not found' });
        }

        if (bidSheet.invitedSuppliers.length === 0) {
            return res.status(400).json({ error: 'No suppliers invited to bid' });
        }

        bidSheet.status = 'sent';
        bidSheet.sentAt = new Date();
        bidSheet.invitedSuppliers.forEach(s => {
            s.invitedAt = new Date();
        });
        bidSheet.updatedAt = new Date();
        bidSheet.updatedBy = req.user._id;

        await bidSheet.save();

        // TODO: Send emails to suppliers with bid request details

        res.json({ message: 'Bid sheet sent to suppliers', bidSheet });
    } catch (error) {
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
        const {
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
            status
        } = req.body;

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
            creatorDisplayName,
            createdBy: req.user._id,
            status: status || 'draft',
            publishedAt: status === 'published' ? new Date() : null
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

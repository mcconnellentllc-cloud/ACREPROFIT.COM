require('dotenv').config();
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const Stripe = require('stripe');

const app = express();

// Initialize Stripe (will be configured per-request for Connect)
const stripe = process.env.STRIPE_SECRET_KEY ? new Stripe(process.env.STRIPE_SECRET_KEY) : null;
const PORT = process.env.PORT || 3001;
const JWT_SECRET = process.env.JWT_SECRET || 'acreprofit-secret-key-change-in-production';

// Middleware
app.use(cors());
app.use(express.json());

// MongoDB Connection
const connectDB = async () => {
    try {
        if (process.env.MONGODB_URI) {
            await mongoose.connect(process.env.MONGODB_URI);
            console.log('MongoDB connected successfully');
            // Initialize admin users
            await initializeAdmins();
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
        enum: ['customer', 'admin', 'superadmin'],
        default: 'customer'
    },
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
    representativeId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    crop: { type: String, required: true },
    program: { type: String, required: true },
    acres: { type: Number, required: true },
    chemicals: [{
        name: String,
        rate: Number,
        rateUnit: String,
        totalAmount: Number,
        totalUnit: String,
        packageSize: Number,
        packageUnit: String,
        packagesNeeded: Number,
        pricePerPackage: Number,
        totalPrice: Number
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
    status: {
        type: String,
        enum: ['draft', 'submitted', 'confirmed', 'bundled', 'ordered', 'shipped', 'delivered'],
        default: 'draft'
    },
    bundleId: { type: mongoose.Schema.Types.ObjectId, ref: 'Bundle' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const Order = mongoose.model('Order', orderSchema);

// Bundle Model (for truckload bundling)
const bundleSchema = new mongoose.Schema({
    product: String,
    targetQuantity: Number, // e.g., 250 gal for a shuttle
    currentQuantity: { type: Number, default: 0 },
    orders: [{ type: mongoose.Schema.Types.ObjectId, ref: 'Order' }],
    status: {
        type: String,
        enum: ['collecting', 'ready', 'ordered', 'shipped', 'delivered'],
        default: 'collecting'
    },
    createdAt: { type: Date, default: Date.now }
});

const Bundle = mongoose.model('Bundle', bundleSchema);

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

    // Category and crop info
    category: { type: String, enum: ['herbicide', 'fungicide', 'insecticide', 'adjuvant', 'fertilizer', 'other'], default: 'herbicide' },
    crops: [String], // Which crops this can be used on: ['corn', 'soybeans', 'wheat']

    // Packaging
    packSize: { type: String, required: true }, // e.g., "2x2.5", "Shuttle", "4x5", "20"
    unit: { type: String, required: true }, // e.g., "gl" (gallon), "oz", "lb"
    unitsPerPack: { type: Number }, // e.g., 250 for a Shuttle (250 gal)

    // Pricing - 3-tier pricing model
    costPrice: { type: Number, required: true }, // Tier 1: What we pay the supplier (per unit)
    adminPrice: { type: Number }, // Tier 2: Cost + Acre Profit admin margin (per unit)
    sellPrice: { type: Number, required: true }, // Tier 3: Retail price - what customer pays (per unit)
    // Rep commission = sellPrice - adminPrice (goes to the customer's assigned rep)
    margin: { type: Number }, // Calculated: (sellPrice - costPrice) / sellPrice * 100

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

// Auto-calculate margin before save
chemicalSchema.pre('save', function(next) {
    if (this.sellPrice && this.costPrice) {
        this.margin = Math.round(((this.sellPrice - this.costPrice) / this.sellPrice) * 100 * 100) / 100;
    }
    next();
});

// Index for quick lookups
chemicalSchema.index({ productName: 1, sourceSupplier: 1, packSize: 1 });
chemicalSchema.index({ sourceSupplier: 1 });
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
    total: Number,

    // Status tracking
    status: {
        type: String,
        enum: ['draft', 'submitted', 'confirmed', 'ordered_from_supplier', 'received', 'ready_for_pickup', 'delivered', 'cancelled'],
        default: 'draft'
    },

    // Payment
    paymentStatus: { type: String, enum: ['pending', 'paid', 'partial'], default: 'pending' },
    paymentMethod: String,

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
        description: 'Water conditioning agent',
        unit: 'gal',
        pricing: [
            { minQty: 1, maxQty: 9, pricePerUnit: 165 },
            { minQty: 10, maxQty: 60, pricePerUnit: 135 },
            { minQty: 61, maxQty: 180, pricePerUnit: 125 }
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
    if (req.user.role !== 'admin' && req.user.role !== 'superadmin') {
        return res.status(403).json({ error: 'Admin access required' });
    }
    next();
};

const superAdminMiddleware = async (req, res, next) => {
    if (req.user.role !== 'superadmin') {
        return res.status(403).json({ error: 'Super admin access required' });
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

        const user = new User({
            name,
            email,
            password,
            phone,
            address,
            farm,
            crops,
            representativeId,
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

        res.json({
            user: {
                id: user._id,
                name: user.name,
                email: user.email,
                role: user.role,
                farm: user.farm
            },
            token
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
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
        const reps = await User.find({ role: { $in: ['admin', 'superadmin'] } })
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
            totalPrice,
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
        if (req.user.role === 'admin') {
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

        // Check access for non-superadmin
        if (req.user.role === 'admin' &&
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

        // Check access for non-superadmin
        if (req.user.role === 'admin' &&
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
                ...customer.privateApplicatorLicense,
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

// Get all orders for a specific customer (admin only)
app.get('/api/admin/customers/:customerId/orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const customer = await User.findById(req.params.customerId);

        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Check access for non-superadmin
        if (req.user.role === 'admin' &&
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
        const { customerId, crop, programId, acres, chemicals, seeds, pivotBio, totalPrice, status } = req.body;

        // Validate customer exists and admin has access
        const customer = await User.findById(customerId);
        if (!customer || customer.role !== 'customer') {
            return res.status(404).json({ error: 'Customer not found' });
        }

        // Check access for non-superadmin
        if (req.user.role === 'admin' &&
            customer.representative?.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Access denied - not your customer' });
        }

        const costPerAcre = acres > 0 ? Math.round((totalPrice / acres) * 100) / 100 : 0;

        const order = new Order({
            userId: customerId,
            representativeId: req.user._id,
            crop,
            program: programId,
            acres,
            chemicals,
            seeds,
            pivotBio,
            totalPrice,
            costPerAcre,
            status: status || 'draft',
            createdBy: req.user._id // Track who created this order
        });

        await order.save();
        res.status(201).json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Edit an order (admin only)
app.put('/api/admin/orders/:orderId', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { crop, acres, chemicals, seeds, pivotBio, totalPrice, status } = req.body;

        let query = { _id: req.params.orderId };

        // If not superadmin, can only edit their own customers' orders
        if (req.user.role === 'admin') {
            query.representativeId = req.user._id;
        }

        const order = await Order.findOne(query);
        if (!order) {
            return res.status(404).json({ error: 'Order not found or access denied' });
        }

        // Update fields
        if (crop !== undefined) order.crop = crop;
        if (acres !== undefined) order.acres = acres;
        if (chemicals !== undefined) order.chemicals = chemicals;
        if (seeds !== undefined) order.seeds = seeds;
        if (pivotBio !== undefined) order.pivotBio = pivotBio;
        if (totalPrice !== undefined) {
            order.totalPrice = totalPrice;
            order.costPerAcre = order.acres > 0 ? Math.round((totalPrice / order.acres) * 100) / 100 : 0;
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

// Get all orders (admin only)
app.get('/api/admin/orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = {};

        // If not superadmin, only show orders for their customers
        if (req.user.role === 'admin') {
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
        if (req.user.role === 'admin') {
            query.representativeId = req.user._id;
        }

        const order = await Order.findOneAndUpdate(
            query,
            { status, updatedAt: new Date() },
            { new: true }
        );

        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

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

// Get order stats (admin only)
app.get('/api/admin/stats', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = {};
        if (req.user.role === 'admin') {
            query.representativeId = req.user._id;
        }

        const totalOrders = await Order.countDocuments(query);
        const submittedOrders = await Order.countDocuments({ ...query, status: 'submitted' });
        const totalAcres = await Order.aggregate([
            { $match: query },
            { $group: { _id: null, total: { $sum: '$acres' } } }
        ]);

        let customerQuery = { role: 'customer' };
        if (req.user.role === 'admin') {
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

// Bundle orders for truckload (superadmin only)
app.post('/api/admin/bundles', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const { product, orderIds, targetQuantity } = req.body;

        const bundle = new Bundle({
            product,
            targetQuantity,
            orders: orderIds
        });

        // Calculate current quantity from orders
        const orders = await Order.find({ _id: { $in: orderIds } });
        let currentQuantity = 0;
        orders.forEach(order => {
            order.chemicals.forEach(chem => {
                if (chem.name === product) {
                    currentQuantity += chem.totalAmount;
                }
            });
        });

        bundle.currentQuantity = currentQuantity;
        if (currentQuantity >= targetQuantity) {
            bundle.status = 'ready';
        }

        await bundle.save();

        // Update orders with bundle ID
        await Order.updateMany(
            { _id: { $in: orderIds } },
            { bundleId: bundle._id, status: 'bundled' }
        );

        res.status(201).json(bundle);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get bundles (superadmin only)
app.get('/api/admin/bundles', authMiddleware, superAdminMiddleware, async (req, res) => {
    try {
        const bundles = await Bundle.find()
            .populate({
                path: 'orders',
                populate: { path: 'userId', select: 'name email farm' }
            })
            .sort({ createdAt: -1 });
        res.json(bundles);
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

        // If approved, create admin user
        if (status === 'approved') {
            const existingUser = await User.findOne({ email: application.email.toLowerCase() });
            if (!existingUser) {
                const tempPassword = 'Farm2026!'; // They should change this
                await User.create({
                    name: `${application.firstName} ${application.lastName}`,
                    email: application.email,
                    password: tempPassword,
                    phone: application.phone,
                    role: 'admin'
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
        const { orderId, paymentMethod } = req.body;

        if (!stripe) {
            return res.status(400).json({ error: 'Stripe not configured' });
        }

        const order = await Order.findById(orderId).populate('representativeId');
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        if (order.userId.toString() !== req.user._id.toString()) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        const rep = order.representativeId;
        if (!rep.stripeAccountId || rep.stripeAccountStatus !== 'active') {
            return res.status(400).json({
                error: 'Representative has not set up payment processing. Please pay by check.',
                checkPayableTo: rep.checkPayableTo || rep.name,
                checkMailingAddress: rep.checkMailingAddress
            });
        }

        // Calculate platform fee (optional - 0% for now, can add later)
        const platformFeePercent = 0;
        const applicationFee = Math.round(order.totalCost * 100 * platformFeePercent);

        // Create payment intent with Stripe Connect
        const paymentIntentParams = {
            amount: Math.round(order.totalCost * 100), // Convert to cents
            currency: 'usd',
            payment_method_types: paymentMethod === 'ach' ? ['us_bank_account'] : ['card'],
            transfer_data: {
                destination: rep.stripeAccountId
            },
            metadata: {
                orderId: order._id.toString(),
                customerId: req.user._id.toString(),
                customerName: req.user.name,
                repName: rep.name
            }
        };

        if (applicationFee > 0) {
            paymentIntentParams.application_fee_amount = applicationFee;
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
        order.paymentMethod = 'stripe_ach'; // ACH only - no credit cards
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

// Stripe webhook for payment confirmations
app.post('/api/webhooks/stripe', express.raw({ type: 'application/json' }), async (req, res) => {
    const sig = req.headers['stripe-signature'];
    const webhookSecret = process.env.STRIPE_WEBHOOK_SECRET;

    let event;
    try {
        event = stripe.webhooks.constructEvent(req.body, sig, webhookSecret);
    } catch (err) {
        return res.status(400).send(`Webhook Error: ${err.message}`);
    }

    if (event.type === 'payment_intent.succeeded') {
        const paymentIntent = event.data.object;
        const orderId = paymentIntent.metadata.orderId;

        if (orderId) {
            const order = await Order.findById(orderId);
            if (order) {
                order.paymentStatus = 'paid';
                order.paidAt = new Date();
                order.status = 'confirmed';
                order.updatedAt = new Date();
                await order.save();
            }
        }
    }

    if (event.type === 'payment_intent.payment_failed') {
        const paymentIntent = event.data.object;
        const orderId = paymentIntent.metadata.orderId;

        if (orderId) {
            const order = await Order.findById(orderId);
            if (order) {
                order.paymentStatus = 'failed';
                order.updatedAt = new Date();
                await order.save();
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
            { productName: 'AgSaver', packSize: 'Shuttle', unit: 'gal', unitsPerPack: 250, costPrice: 12.90, sellPrice: 0, category: 'herbicide', equivalentProduct: 'RT3, Glystar Supreme', notes: 'Glyphosate' },

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
            { productName: 'AMS', packSize: '24 lb', unit: 'lb', unitsPerPack: 24, costPrice: 1.45, sellPrice: 0, category: 'adjuvant', notes: 'Ammonium Sulfate - water conditioner/adjuvant' }
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
                const newChem = await Chemical.create({
                    ...chem,
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
                results.existing.push(`${chem.productName} (${chem.packSize})`);
            }
        }

        res.json({
            message: 'CPD products loaded successfully',
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

        // For public view, show sellPrice not costPrice
        const publicChemicals = chemicals.map(c => ({
            _id: c._id,
            productName: c.productName,
            category: c.category,
            crops: c.crops,
            packSize: c.packSize,
            unit: c.unit,
            unitsPerPack: c.unitsPerPack,
            price: c.sellPrice, // Customer sees sell price
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
        const { costPrice, sellPrice, priceVersion, notes, equivalentProduct, isActive, availableForOrder,
                category, crops, defaultRate, rateUnit, unitsPerPack } = req.body;

        const chemical = await Chemical.findById(req.params.id);
        if (!chemical) {
            return res.status(404).json({ error: 'Chemical not found' });
        }

        // If price changed, save to history
        if ((costPrice !== undefined && costPrice !== chemical.costPrice) ||
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
            if (sellPrice !== undefined) chemical.sellPrice = sellPrice;
            chemical.priceDate = new Date();
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
        const { chemicals, sourceSupplier, priceVersion } = req.body;

        if (!Array.isArray(chemicals)) {
            return res.status(400).json({ error: 'chemicals must be an array' });
        }

        const results = [];
        const version = priceVersion || new Date().toISOString().slice(0, 10);

        for (const chem of chemicals) {
            let existing = await Chemical.findOne({
                productName: chem.productName,
                sourceSupplier: sourceSupplier || chem.sourceSupplier,
                packSize: chem.packSize
            });

            if (existing) {
                if (existing.costPrice !== chem.costPrice || existing.sellPrice !== chem.sellPrice) {
                    await ChemicalPriceHistory.create({
                        chemicalId: existing._id,
                        productName: existing.productName,
                        sourceSupplier: existing.sourceSupplier,
                        packSize: existing.packSize,
                        unit: existing.unit,
                        costPrice: chem.costPrice,
                        sellPrice: chem.sellPrice,
                        priceVersion: version,
                        changedBy: req.user._id
                    });

                    existing.costPrice = chem.costPrice;
                    existing.sellPrice = chem.sellPrice;
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
                    costPrice: chem.costPrice,
                    sellPrice: chem.sellPrice,
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

        // Only admin or the user themselves can update
        if (req.user.role !== 'admin' && req.user.role !== 'superadmin' && req.user._id.toString() !== userId) {
            return res.status(403).json({ error: 'Not authorized' });
        }

        const user = await User.findById(userId);
        if (!user) {
            return res.status(404).json({ error: 'User not found' });
        }

        // Update license fields
        if (updateData.privateApplicatorLicense) {
            user.privateApplicatorLicense = { ...user.privateApplicatorLicense?.toObject(), ...updateData.privateApplicatorLicense };
            if (req.user.role === 'admin' || req.user.role === 'superadmin') {
                user.privateApplicatorLicense.verifiedBy = req.user._id;
                user.privateApplicatorLicense.verifiedAt = new Date();
            }
        }

        if (updateData.commercialApplicatorLicense) {
            user.commercialApplicatorLicense = { ...user.commercialApplicatorLicense?.toObject(), ...updateData.commercialApplicatorLicense };
            if (req.user.role === 'admin' || req.user.role === 'superadmin') {
                user.commercialApplicatorLicense.verifiedBy = req.user._id;
                user.commercialApplicatorLicense.verifiedAt = new Date();
            }
        }

        if (updateData.paraquatCertification) {
            user.paraquatCertification = { ...user.paraquatCertification?.toObject(), ...updateData.paraquatCertification };
            if (req.user.role === 'admin' || req.user.role === 'superadmin') {
                user.paraquatCertification.verifiedBy = req.user._id;
                user.paraquatCertification.verifiedAt = new Date();
            }
        }

        if (updateData.dicambaCertification) {
            user.dicambaCertification = { ...user.dicambaCertification?.toObject(), ...updateData.dicambaCertification };
            if (req.user.role === 'admin' || req.user.role === 'superadmin') {
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
        const order = await ChemicalOrder.findOne({
            _id: req.params.id,
            userId: req.user._id
        }).populate('items.chemicalId');

        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
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

// Admin: Get all chemical orders
app.get('/api/admin/chemical-orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = {};
        if (req.user.role === 'admin') {
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

        // Check if user owns this or is admin
        if (submission.userId.toString() !== req.user._id.toString() &&
            req.user.role !== 'admin' && req.user.role !== 'superadmin') {
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

        if (req.user.role === 'admin') {
            repFilter = { _id: req.user._id };
        } else {
            repFilter = { role: { $in: ['admin', 'superadmin'] } };
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
        if (req.user.role === 'admin') {
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
        const { supplier, items, freight, otherFees, expectedDeliveryDate, deliveryLocation, notes } = req.body;

        // Generate PO number
        const poNumber = await generatePONumber();

        // Calculate item totals and remaining quantities
        const processedItems = items.map(item => ({
            ...item,
            totalPrice: Math.round((item.quantityOrdered * item.pricePerUnit) * 100) / 100,
            quantityAllocated: 0,
            quantityRemaining: item.quantityOrdered
        }));

        // Calculate totals
        const subtotal = processedItems.reduce((sum, item) => sum + item.totalPrice, 0);
        const totalCost = subtotal + (freight || 0) + (otherFees || 0);

        const purchaseOrder = new PurchaseOrder({
            poNumber,
            supplier,
            items: processedItems,
            subtotal: Math.round(subtotal * 100) / 100,
            freight: freight || 0,
            otherFees: otherFees || 0,
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
        const { supplier, expectedDeliveryDate, deliveryLocation, bolNumber, trackingInfo, notes, internalNotes, status } = req.body;

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
            crops: c.crops,
            packSize: c.packSize,
            unit: c.unit,
            unitsPerPack: c.unitsPerPack,
            costPrice: c.costPrice,
            sellPrice: c.sellPrice,
            margin: c.margin,
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

// ---- SPRAY PROGRAM ROUTES ----

// Get all public/recommended programs
app.get('/api/spray-programs', async (req, res) => {
    try {
        const { crop, type } = req.query;
        let query = { isPublic: true };

        if (crop) query.crop = crop;
        if (type) query.type = type;

        const programs = await SprayProgram.find(query)
            .sort({ crop: 1, name: 1 });

        res.json(programs);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get user's custom programs
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
app.get('/api/spray-programs/:id', async (req, res) => {
    try {
        const program = await SprayProgram.findById(req.params.id);
        if (!program) {
            return res.status(404).json({ error: 'Program not found' });
        }
        res.json(program);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Create spray program (admin for recommended, customer for custom)
app.post('/api/spray-programs', authMiddleware, async (req, res) => {
    try {
        const { name, description, crop, applications, type } = req.body;

        // Only admins can create recommended programs
        const programType = (req.user.role === 'admin' || req.user.role === 'superadmin') ? (type || 'template') : 'custom';
        const isPublic = programType === 'recommended';

        const program = new SprayProgram({
            name,
            description,
            crop,
            applications,
            type: programType,
            isPublic,
            createdBy: req.user._id
        });

        // Calculate estimated cost per acre
        let totalCost = 0;
        for (const app of applications) {
            for (const chem of app.chemicals) {
                const chemical = await Chemical.findById(chem.chemicalId);
                if (chemical && chemical.sellPrice && chem.rate) {
                    // Convert rate to gallons and multiply by price
                    totalCost += (chem.rate / 128) * chemical.sellPrice; // Assuming oz to gal
                }
            }
        }
        program.estimatedCostPerAcre = Math.round(totalCost * 100) / 100;

        await program.save();
        res.status(201).json(program);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Calculate order from program (preview before ordering)
app.post('/api/spray-programs/:id/calculate', authMiddleware, async (req, res) => {
    try {
        const { acres } = req.body;
        const program = await SprayProgram.findById(req.params.id);

        if (!program) {
            return res.status(404).json({ error: 'Program not found' });
        }

        const orderItems = [];
        let totalCost = 0;

        for (const app of program.applications) {
            for (const chemItem of app.chemicals) {
                const chemical = await Chemical.findById(chemItem.chemicalId);
                if (!chemical) continue;

                // Calculate amount needed
                let amountNeeded = 0;
                const rate = chemItem.rate;

                // Convert based on rate unit
                switch (chemItem.rateUnit) {
                    case 'oz/acre':
                        amountNeeded = (rate * acres) / 128; // oz to gallons
                        break;
                    case 'pt/acre':
                        amountNeeded = (rate * acres) / 8; // pints to gallons
                        break;
                    case 'qt/acre':
                        amountNeeded = (rate * acres) / 4; // quarts to gallons
                        break;
                    case 'gal/acre':
                        amountNeeded = rate * acres;
                        break;
                    case 'lb/acre':
                        amountNeeded = rate * acres;
                        break;
                    default:
                        amountNeeded = rate * acres;
                }

                // Round up to nearest package
                const unitsPerPack = chemical.unitsPerPack || 1;
                const packsNeeded = Math.ceil(amountNeeded / unitsPerPack);
                const totalPrice = packsNeeded * chemical.sellPrice * unitsPerPack;

                orderItems.push({
                    chemicalId: chemical._id,
                    productName: chemical.productName,
                    packSize: chemical.packSize,
                    unit: chemical.unit,
                    rate,
                    rateUnit: chemItem.rateUnit,
                    acres,
                    calculatedAmount: Math.round(amountNeeded * 100) / 100,
                    quantity: packsNeeded,
                    unitPrice: chemical.sellPrice,
                    totalPrice: Math.round(totalPrice * 100) / 100,
                    applicationName: app.name
                });

                totalCost += totalPrice;
            }
        }

        res.json({
            program: program.name,
            crop: program.crop,
            acres,
            items: orderItems,
            totalCost: Math.round(totalCost * 100) / 100,
            costPerAcre: Math.round((totalCost / acres) * 100) / 100
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// ============ START SERVER ============

connectDB().then(() => {
    app.listen(PORT, () => {
        console.log(`Acre Profit API running on port ${PORT}`);
    });
});

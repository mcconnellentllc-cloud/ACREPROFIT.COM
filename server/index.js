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
    // Payment information
    paymentMethod: {
        type: String,
        enum: ['stripe_ach', 'stripe_card', 'check', 'pending'],
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

// Merch Credit Model
const merchCreditSchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    totalSpent: { type: Number, default: 0 },
    creditEarned: { type: Number, default: 0 },
    creditUsed: { type: Number, default: 0 },
    creditAvailable: { type: Number, default: 0 },
    history: [{
        type: { type: String, enum: ['earned', 'used'] },
        amount: Number,
        orderId: mongoose.Schema.Types.ObjectId,
        description: String,
        date: { type: Date, default: Date.now }
    }],
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const MerchCredit = mongoose.model('MerchCredit', merchCreditSchema);

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

    // Pricing - COST is what Acre Profit pays, SELL is customer price
    costPrice: { type: Number, required: true }, // What we pay the supplier (per unit)
    sellPrice: { type: Number, required: true }, // What we charge customers (per unit)
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

// Spray Program Model (saved custom programs)
const sprayProgramSchema = new mongoose.Schema({
    name: { type: String, required: true },
    description: String,

    // Program type
    type: { type: String, enum: ['recommended', 'custom', 'template'], default: 'custom' },
    isPublic: { type: Boolean, default: false }, // Recommended programs are public

    // Target crop
    crop: { type: String, required: true }, // corn, soybeans, wheat, etc.

    // Program passes/applications
    applications: [{
        name: String, // e.g., "Burndown", "Pre-emergent", "Post-emergent"
        timing: String, // e.g., "14 days before planting", "At planting", "V4-V6"
        chemicals: [{
            chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical' },
            productName: String,
            rate: Number,
            rateUnit: String,
            packSize: String,
            unit: String
        }]
    }],

    // Cost estimate per acre
    estimatedCostPerAcre: Number,

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

// ============ INITIALIZE ADMIN USERS ============

async function initializeAdmins() {
    const admins = [
        {
            name: 'Kyle McConnell',
            email: 'kyle@togoag.com',
            password: 'Farm2026!',
            phone: '970-571-1015',
            role: 'superadmin'
        },
        {
            name: 'Ty Mollohan',
            email: 'tymollohan77@gmail.com',
            password: 'Farm2026!',
            phone: '970-520-2340',
            role: 'admin'
        },
        {
            name: 'Chad Bamford',
            email: 'ckbamford@yahoo.com',
            password: 'Farm2026!',
            phone: '970-520-3716',
            role: 'admin'
        },
        {
            name: 'Seth Rolfs',
            email: 'seth@acreprofit.com',
            password: 'Farm2026!',
            phone: '785-531-0680',
            role: 'admin'
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
            { name: 'Kyle McConnell', email: 'kyle@togoag.com', password: 'Farm2026!', phone: '970-571-1015', role: 'superadmin' },
            { name: 'Ty Mollohan', email: 'tymollohan77@gmail.com', password: 'Farm2026!', phone: '970-520-2340', role: 'admin' },
            { name: 'Chad Bamford', email: 'ckbamford@yahoo.com', password: 'Farm2026!', phone: '970-520-3716', role: 'admin' },
            { name: 'Seth Rolfs', email: 'seth@acreprofit.com', password: 'Farm2026!', phone: '785-531-0680', role: 'admin' }
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
        const { name, email, password, phone, farm, crops, representativeId } = req.body;

        const existingUser = await User.findOne({ email: email.toLowerCase() });
        if (existingUser) {
            return res.status(400).json({ error: 'Email already registered' });
        }

        const user = new User({
            name,
            email,
            password,
            phone,
            farm,
            crops,
            representative: representativeId,
            role: 'customer'
        });
        await user.save();

        const token = jwt.sign({ userId: user._id }, JWT_SECRET, { expiresIn: '30d' });

        res.status(201).json({
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
            representative: user.representative
        }
    });
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

// Get all customers (admin only)
app.get('/api/admin/customers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        let query = { role: 'customer' };

        // If not superadmin, only show their own customers
        if (req.user.role === 'admin') {
            query.representative = req.user._id;
        }

        const customers = await User.find(query)
            .select('-password')
            .populate('representative', 'name email');
        res.json(customers);
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

// ---- MERCH CREDIT ROUTES ----

// Get user's merch credit
app.get('/api/merch-credit', authMiddleware, async (req, res) => {
    try {
        let credit = await MerchCredit.findOne({ userId: req.user._id });

        if (!credit) {
            credit = await MerchCredit.create({ userId: req.user._id });
        }

        res.json({
            totalSpent: credit.totalSpent,
            creditEarned: credit.creditEarned,
            creditUsed: credit.creditUsed,
            creditAvailable: credit.creditAvailable,
            history: credit.history
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Calculate and add merch credit after order completion
app.post('/api/merch-credit/add', authMiddleware, async (req, res) => {
    try {
        const { orderId, orderAmount } = req.body;

        let credit = await MerchCredit.findOne({ userId: req.user._id });
        if (!credit) {
            credit = await MerchCredit.create({ userId: req.user._id });
        }

        // Calculate new credit: $20 for every $2000 spent
        const previousTotal = credit.totalSpent;
        const newTotal = previousTotal + orderAmount;

        const previousCredits = Math.floor(previousTotal / 2000) * 20;
        const newCredits = Math.floor(newTotal / 2000) * 20;
        const creditToAdd = newCredits - previousCredits;

        if (creditToAdd > 0) {
            credit.history.push({
                type: 'earned',
                amount: creditToAdd,
                orderId,
                description: `Earned $${creditToAdd} merch credit from order`
            });
        }

        credit.totalSpent = newTotal;
        credit.creditEarned += creditToAdd;
        credit.creditAvailable += creditToAdd;
        credit.updatedAt = new Date();

        await credit.save();

        res.json({
            totalSpent: credit.totalSpent,
            creditEarned: credit.creditEarned,
            creditAvailable: credit.creditAvailable,
            newCreditAdded: creditToAdd
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Use merch credit
app.post('/api/merch-credit/use', authMiddleware, async (req, res) => {
    try {
        const { amount, description } = req.body;

        const credit = await MerchCredit.findOne({ userId: req.user._id });
        if (!credit || credit.creditAvailable < amount) {
            return res.status(400).json({ error: 'Insufficient merch credit' });
        }

        credit.creditUsed += amount;
        credit.creditAvailable -= amount;
        credit.history.push({
            type: 'used',
            amount,
            description: description || 'Merch purchase'
        });
        credit.updatedAt = new Date();

        await credit.save();

        res.json({
            creditUsed: amount,
            creditAvailable: credit.creditAvailable
        });
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
            kyle: 'kyle@togoag.com',
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

        // Add merch credit for the customer
        const orderAmount = order.totalCost || 0;
        if (orderAmount > 0) {
            let credit = await MerchCredit.findOne({ userId: order.userId });
            if (!credit) {
                credit = await MerchCredit.create({ userId: order.userId });
            }

            const previousTotal = credit.totalSpent;
            const newTotal = previousTotal + orderAmount;
            const previousCredits = Math.floor(previousTotal / 2000) * 20;
            const newCredits = Math.floor(newTotal / 2000) * 20;
            const creditToAdd = newCredits - previousCredits;

            if (creditToAdd > 0) {
                credit.history.push({
                    type: 'earned',
                    amount: creditToAdd,
                    orderId: order._id,
                    description: `Earned $${creditToAdd} merch credit from order`
                });
            }

            credit.totalSpent = newTotal;
            credit.creditEarned += creditToAdd;
            credit.creditAvailable += creditToAdd;
            credit.updatedAt = new Date();
            await credit.save();
        }

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

                // Add merch credit
                const orderAmount = order.totalCost || 0;
                if (orderAmount > 0) {
                    let credit = await MerchCredit.findOne({ userId: order.userId });
                    if (!credit) {
                        credit = await MerchCredit.create({ userId: order.userId });
                    }

                    const previousTotal = credit.totalSpent;
                    const newTotal = previousTotal + orderAmount;
                    const previousCredits = Math.floor(previousTotal / 2000) * 20;
                    const newCredits = Math.floor(newTotal / 2000) * 20;
                    const creditToAdd = newCredits - previousCredits;

                    if (creditToAdd > 0) {
                        credit.history.push({
                            type: 'earned',
                            amount: creditToAdd,
                            orderId: order._id,
                            description: `Earned $${creditToAdd} merch credit from order`
                        });
                    }

                    credit.totalSpent = newTotal;
                    credit.creditEarned += creditToAdd;
                    credit.creditAvailable += creditToAdd;
                    credit.updatedAt = new Date();
                    await credit.save();
                }
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

        const priceVersion = '2026-01-18';

        // Products that Acre Profit will sell (sourced from CPD)
        // costPrice = what we pay CPD, sellPrice = what we charge customers (matching CHS for now)
        const acreProfit_CPD_Products = [
            // Product, Pack, Unit, CPD Cost, CHS Price (our sell price to beat), Notes
            { productName: 'Dicamba DMA', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 30.25, sellPrice: 31.24, category: 'herbicide' },
            { productName: 'Dicamba DMA', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 28.25, sellPrice: 30.89, category: 'herbicide' },
            { productName: 'Dicamba HD', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 30.57, sellPrice: 35.31, category: 'herbicide' },
            { productName: 'LV 6', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 29.90, sellPrice: 32.43, category: 'herbicide' },
            { productName: 'LV 6', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 27.90, sellPrice: 30.14, category: 'herbicide' },
            { productName: 'AgSaver', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 13.25, sellPrice: 15.00, category: 'herbicide', equivalentProduct: 'RT3/Glystar Supreme', notes: 'Glyphosate - Formulation equiv' },
            { productName: 'Aatrex', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 13.35, sellPrice: 12.78, category: 'herbicide', notes: 'CPD higher than CHS' },
            { productName: 'Agri-Star', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 39.25, sellPrice: 43.83, category: 'herbicide', equivalentProduct: 'Level Best Pro', notes: 'Need to get equivalents' },
            { productName: 'Agri-Star', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 38.13, sellPrice: 42.42, category: 'herbicide', equivalentProduct: 'Level Best Pro', notes: 'Need to get equivalents' },
            { productName: 'Agri-Star Tapran', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 19.00, sellPrice: 28.71, category: 'herbicide', equivalentProduct: 'Tapran', notes: 'Need to get equivalents' },
            { productName: 'Agri-Star Tapran', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 18.00, sellPrice: 28.16, category: 'herbicide', equivalentProduct: 'Tapran', notes: 'Need to get equivalents' },
            { productName: 'Aggrestrol', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 21.00, sellPrice: 33.81, category: 'herbicide', notes: 'Need to get equivalents' },
            { productName: 'Aggrestrol', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 20.00, sellPrice: 32.37, category: 'herbicide', notes: 'Need to get equivalents' },
            { productName: 'Sulfentrazone', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 70.50, sellPrice: 75.20, category: 'herbicide' },
            { productName: 'Valor SX', packSize: '4x5', unit: 'lb', unitsPerPack: 20, costPrice: 14.25, sellPrice: 15.06, category: 'herbicide' },
            // CPD-only products (no CHS equivalent)
            { productName: 'Glufosinate', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 16.50, sellPrice: 18.00, category: 'herbicide' },
            { productName: 'Paraquat', packSize: 'Shuttle', unit: 'gl', unitsPerPack: 250, costPrice: 16.00, sellPrice: 18.00, category: 'herbicide' },
            { productName: 'Mesotrione', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 48.25, sellPrice: 52.00, category: 'herbicide' },
            { productName: 'Clethodim', packSize: '2x2.5', unit: 'gl', unitsPerPack: 5, costPrice: 33.50, sellPrice: 36.00, category: 'herbicide' },
            { productName: 'Clethodim', packSize: '135', unit: 'gl', unitsPerPack: 135, costPrice: 33.00, sellPrice: 35.50, category: 'herbicide' }
        ];

        // CHS prices for reference/comparison
        const chsPrices = [
            { productName: 'Dicamba DMA', packSize: '2x2.5', unit: 'gl', price: 31.24 },
            { productName: 'Dicamba DMA', packSize: 'Shuttle', unit: 'gl', price: 30.89 },
            { productName: 'Dicamba HD', packSize: 'Shuttle', unit: 'gl', price: 35.31 },
            { productName: 'LV 6', packSize: '2x2.5', unit: 'gl', price: 32.43 },
            { productName: 'LV 6', packSize: 'Shuttle', unit: 'gl', price: 30.14 },
            { productName: 'RT3', packSize: 'Shuttle', unit: 'gl', price: 17.50 },
            { productName: 'Glystar Supreme', packSize: 'Shuttle', unit: 'gl', price: 15.53 },
            { productName: 'Aatrex', packSize: 'Shuttle', unit: 'gl', price: 12.78 },
            { productName: 'Level Best Pro', packSize: '2x2.5', unit: 'gl', price: 43.83 },
            { productName: 'Level Best Pro', packSize: 'Shuttle', unit: 'gl', price: 42.42 },
            { productName: 'Tapran', packSize: '2x2.5', unit: 'gl', price: 28.71 },
            { productName: 'Tapran', packSize: 'Shuttle', unit: 'gl', price: 28.16 },
            { productName: 'Aggrestrol', packSize: '2x2.5', unit: 'gl', price: 33.81 },
            { productName: 'Aggrestrol', packSize: 'Shuttle', unit: 'gl', price: 32.37 },
            { productName: 'Artect FI', packSize: '2x2.5', unit: 'gl', price: 84.71 },
            { productName: 'Artect FI', packSize: 'Shuttle', unit: 'gl', price: 84.71 },
            { productName: 'Sulfentrazone', packSize: '2x2.5', unit: 'gl', price: 75.20 },
            { productName: 'Autumn Super', packSize: '20', unit: 'oz', price: 22.58 },
            { productName: 'Valor SX', packSize: '4x5', unit: 'lb', price: 15.06 }
        ];

        const results = { acreProfit: [], chs: [] };

        // Insert Acre Profit products (sourced from CPD)
        for (const chem of acreProfit_CPD_Products) {
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
                results.acreProfit.push({ action: 'created', product: chem.productName });
            } else {
                results.acreProfit.push({ action: 'exists', product: chem.productName });
            }
        }

        res.json({
            message: 'Seed data loaded',
            summary: {
                acreProfit: {
                    created: results.acreProfit.filter(r => r.action === 'created').length,
                    existed: results.acreProfit.filter(r => r.action === 'exists').length
                },
                totalProducts: acreProfit_CPD_Products.length,
                chsReferenceProducts: chsPrices.length
            }
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

        res.json(order);
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

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

const app = express();
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
    farm: {
        name: String,
        acres: Number,
        state: String,
        county: String
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
    totalPrice: Number,
    costPerAcre: Number,
    status: {
        type: String,
        enum: ['draft', 'submitted', 'confirmed', 'bundled', 'ordered', 'delivered'],
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

// ============ START SERVER ============

connectDB().then(() => {
    app.listen(PORT, () => {
        console.log(`Acre Profit API running on port ${PORT}`);
    });
});

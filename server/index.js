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
    crop: { type: String, required: true },
    program: { type: String, required: true }, // '2-pass', '3-pass', etc.
    acres: { type: Number, required: true },
    chemicals: [{
        name: String,
        rate: Number, // rate per acre
        rateUnit: String, // 'oz/acre', 'pt/acre', 'gal/acre'
        totalAmount: Number,
        totalUnit: String,
        packageSize: Number,
        packageUnit: String,
        packagesNeeded: Number,
        pricePerPackage: Number,
        totalPrice: Number
    }],
    totalPrice: Number,
    status: {
        type: String,
        enum: ['draft', 'submitted', 'confirmed', 'ordered', 'delivered'],
        default: 'draft'
    },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const Order = mongoose.model('Order', orderSchema);

// ============ SPRAY PROGRAMS DATA ============
// Based on Lueking Crop Planner - UPDATE WITH ACTUAL DATA

const sprayPrograms = {
    corn: {
        '2-pass': {
            name: '2-Pass Corn Program',
            description: 'Pre-emergent + Post-emergent application',
            chemicals: [
                {
                    name: 'Glyphosate',
                    defaultRate: 32,
                    rateUnit: 'oz/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0 // Price TBD
                },
                {
                    name: 'Atrazine 4L',
                    defaultRate: 1.5,
                    rateUnit: 'qt/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0
                },
                {
                    name: 'Metolachlor',
                    defaultRate: 1.3,
                    rateUnit: 'pt/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0
                }
            ]
        },
        '3-pass': {
            name: '3-Pass Corn Program',
            description: 'Burndown + Pre-emergent + Post-emergent application',
            chemicals: [
                {
                    name: 'Glyphosate',
                    defaultRate: 32,
                    rateUnit: 'oz/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0
                },
                {
                    name: 'Atrazine 4L',
                    defaultRate: 2,
                    rateUnit: 'qt/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0
                },
                {
                    name: 'Metolachlor',
                    defaultRate: 1.5,
                    rateUnit: 'pt/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0
                },
                {
                    name: '2,4-D Amine',
                    defaultRate: 1,
                    rateUnit: 'pt/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0
                }
            ]
        }
    },
    soybeans: {
        '2-pass': {
            name: '2-Pass Soybean Program',
            description: 'Pre-emergent + Post-emergent application',
            chemicals: [
                {
                    name: 'Glyphosate',
                    defaultRate: 32,
                    rateUnit: 'oz/acre',
                    packageSize: 250,
                    packageUnit: 'gal',
                    pricePerPackage: 0
                },
                {
                    name: 'Metribuzin',
                    defaultRate: 0.5,
                    rateUnit: 'lb/acre',
                    packageSize: 50,
                    packageUnit: 'lb',
                    pricePerPackage: 0
                }
            ]
        }
    }
};

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

// ============ ROUTES ============

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', message: 'Acre Profit API is running' });
});

// ---- AUTH ROUTES ----

// Sign up
app.post('/api/auth/signup', async (req, res) => {
    try {
        const { name, email, password, phone, farm, crops } = req.body;

        // Check if user exists
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
            crops
        });
        await user.save();

        const token = jwt.sign({ userId: user._id }, JWT_SECRET, { expiresIn: '30d' });

        res.status(201).json({
            user: {
                id: user._id,
                name: user.name,
                email: user.email,
                farm: user.farm
            },
            token
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Login
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
                farm: user.farm
            },
            token
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get current user
app.get('/api/auth/me', authMiddleware, async (req, res) => {
    res.json({
        user: {
            id: req.user._id,
            name: req.user.name,
            email: req.user.email,
            phone: req.user.phone,
            farm: req.user.farm,
            crops: req.user.crops
        }
    });
});

// Update user profile
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

// ---- SPRAY PROGRAM ROUTES ----

// Get all crops
app.get('/api/crops', (req, res) => {
    const crops = Object.keys(sprayPrograms).map(crop => ({
        id: crop,
        name: crop.charAt(0).toUpperCase() + crop.slice(1),
        programs: Object.keys(sprayPrograms[crop])
    }));
    res.json(crops);
});

// Get programs for a crop
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

// Get program details with chemicals
app.get('/api/crops/:crop/programs/:programId', authMiddleware, (req, res) => {
    const { crop, programId } = req.params;
    const program = sprayPrograms[crop.toLowerCase()]?.[programId];

    if (!program) {
        return res.status(404).json({ error: 'Program not found' });
    }

    res.json(program);
});

// ---- ORDER/CALCULATION ROUTES ----

// Calculate order based on acres and rates
app.post('/api/calculate', authMiddleware, (req, res) => {
    try {
        const { crop, programId, acres, customRates } = req.body;

        const program = sprayPrograms[crop.toLowerCase()]?.[programId];
        if (!program) {
            return res.status(404).json({ error: 'Program not found' });
        }

        const calculations = program.chemicals.map(chemical => {
            // Use custom rate if provided, otherwise use default
            const rate = customRates?.[chemical.name] ?? chemical.defaultRate;

            // Convert rate to gallons for calculation
            let totalGallons;
            switch (chemical.rateUnit) {
                case 'oz/acre':
                    totalGallons = (rate * acres) / 128; // 128 oz per gallon
                    break;
                case 'pt/acre':
                    totalGallons = (rate * acres) / 8; // 8 pints per gallon
                    break;
                case 'qt/acre':
                    totalGallons = (rate * acres) / 4; // 4 quarts per gallon
                    break;
                case 'gal/acre':
                    totalGallons = rate * acres;
                    break;
                case 'lb/acre':
                    // For dry products, keep in lbs
                    totalGallons = rate * acres;
                    break;
                default:
                    totalGallons = rate * acres;
            }

            // Calculate packages needed
            const packageSize = chemical.packageSize;
            const packagesNeeded = Math.ceil(totalGallons / packageSize);
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

        res.json({
            crop,
            program: program.name,
            acres,
            chemicals: calculations,
            totalPrice
        });
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Save order
app.post('/api/orders', authMiddleware, async (req, res) => {
    try {
        const { crop, programId, acres, chemicals, totalPrice } = req.body;

        const order = new Order({
            userId: req.user._id,
            crop,
            program: programId,
            acres,
            chemicals,
            totalPrice,
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
        const orders = await Order.find({ userId: req.user._id }).sort({ createdAt: -1 });
        res.json(orders);
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

// ============ START SERVER ============

connectDB().then(() => {
    app.listen(PORT, () => {
        console.log(`Acre Profit API running on port ${PORT}`);
    });
});

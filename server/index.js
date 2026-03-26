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

const app = express();

// ============ CONFIGURATION ============
const PORT = process.env.PORT || 3001;
const JWT_SECRET = process.env.JWT_SECRET || 'acreprofit-dev-secret-change-in-production';
const stripe = process.env.STRIPE_SECRET_KEY ? new Stripe(process.env.STRIPE_SECRET_KEY) : null;

// Email transporter
let emailTransporter = null;
if (process.env.SMTP_HOST) {
    emailTransporter = nodemailer.createTransport({
        host: process.env.SMTP_HOST,
        port: parseInt(process.env.SMTP_PORT) || 587,
        secure: process.env.SMTP_SECURE === 'true',
        auth: {
            user: process.env.SMTP_USER,
            pass: process.env.SMTP_PASS
        }
    });
}

// ============ MIDDLEWARE ============
app.use(cors());

// Stripe webhook needs raw body
app.use('/api/payments/webhook', express.raw({ type: 'application/json' }));

// JSON parsing for all other routes
app.use(express.json());

// Serve static files from root directory
app.use(express.static(path.join(__dirname, '..')));

// Request logging
app.use((req, res, next) => {
    console.log(`${new Date().toISOString()} ${req.method} ${req.path}`);
    next();
});

// ============ MONGOOSE MODELS ============

// User Schema
const userSchema = new mongoose.Schema({
    email: { type: String, required: true, unique: true, lowercase: true, trim: true },
    password: { type: String, required: true },
    role: { type: String, enum: ['customer', 'admin'], default: 'customer' },
    name: { type: String, required: true, trim: true },
    phone: { type: String, trim: true },
    address: {
        street: String,
        city: String,
        state: String,
        zip: String
    },
    stripeCustomerId: String,
    resetPasswordToken: String,
    resetPasswordExpires: Date,
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

userSchema.pre('save', async function(next) {
    if (!this.isModified('password')) return next();
    this.password = await bcrypt.hash(this.password, 12);
    this.updatedAt = new Date();
    next();
});

userSchema.methods.comparePassword = async function(candidatePassword) {
    return bcrypt.compare(candidatePassword, this.password);
};

userSchema.methods.toJSON = function() {
    const obj = this.toObject();
    delete obj.password;
    delete obj.resetPasswordToken;
    delete obj.resetPasswordExpires;
    return obj;
};

const User = mongoose.model('User', userSchema);

// Customer Profile Schema
const customerProfileSchema = new mongoose.Schema({
    userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, unique: true },
    acres: { type: Number, default: 0 },
    crops: [{
        name: { type: String, required: true },
        acres: { type: Number, default: 0 }
    }],
    sprayHistory: [{
        date: Date,
        product: { type: mongoose.Schema.Types.ObjectId, ref: 'Product' },
        quantity: Number,
        acres: Number,
        notes: String
    }],
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const CustomerProfile = mongoose.model('CustomerProfile', customerProfileSchema);

// Product Schema
const productSchema = new mongoose.Schema({
    name: { type: String, required: true, trim: true },
    description: { type: String, trim: true },
    category: {
        type: String,
        enum: ['herbicide', 'insecticide', 'fungicide', 'adjuvant', 'fertilizer', 'seed-treatment', 'other'],
        required: true
    },
    activeIngredient: { type: String, trim: true },
    manufacturer: { type: String, trim: true },
    applicationRate: {
        min: Number,
        max: Number,
        unit: { type: String, default: 'oz/acre' }
    },
    packageSizes: [{
        volume: { type: Number, required: true },
        unit: { type: String, required: true, enum: ['oz', 'pt', 'qt', 'gal', '2.5gal', '30gal', 'lb', 'case'] },
        price: { type: Number, required: true },
        inventory: { type: Number, default: 0 },
        sku: String
    }],
    supplierCost: { type: Number },
    imageUrl: String,
    safetyDataSheet: String,
    label: String,
    active: { type: Boolean, default: true },
    featured: { type: Boolean, default: false },
    crops: [String],
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

productSchema.index({ name: 'text', description: 'text', activeIngredient: 'text' });

const Product = mongoose.model('Product', productSchema);

// Order Schema
const orderSchema = new mongoose.Schema({
    customerId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true },
    orderNumber: { type: String, unique: true },
    items: [{
        product: { type: mongoose.Schema.Types.ObjectId, ref: 'Product', required: true },
        productName: String,
        packageSize: {
            volume: Number,
            unit: String,
            price: Number
        },
        quantity: { type: Number, required: true },
        unitPrice: { type: Number, required: true },
        totalPrice: { type: Number, required: true }
    }],
    subtotal: { type: Number, required: true },
    tax: { type: Number, default: 0 },
    shipping: { type: Number, default: 0 },
    total: { type: Number, required: true },
    status: {
        type: String,
        enum: ['pending', 'paid', 'processing', 'shipped', 'delivered', 'cancelled', 'refunded'],
        default: 'pending'
    },
    paymentStatus: {
        type: String,
        enum: ['pending', 'paid', 'failed', 'refunded'],
        default: 'pending'
    },
    stripePaymentIntentId: String,
    stripePaymentMethodId: String,
    shippingAddress: {
        name: String,
        street: String,
        city: String,
        state: String,
        zip: String
    },
    trackingNumber: String,
    notes: String,
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

orderSchema.pre('save', async function(next) {
    if (!this.orderNumber) {
        const count = await mongoose.model('Order').countDocuments();
        this.orderNumber = `AP-${String(count + 1001).padStart(6, '0')}`;
    }
    this.updatedAt = new Date();
    next();
});

const Order = mongoose.model('Order', orderSchema);

// Supplier Purchase Schema (for admin inventory management)
const supplierPurchaseSchema = new mongoose.Schema({
    product: { type: mongoose.Schema.Types.ObjectId, ref: 'Product', required: true },
    productName: String,
    packageSizeIndex: { type: Number, required: true },
    quantity: { type: Number, required: true },
    costPerUnit: { type: Number, required: true },
    totalCost: { type: Number, required: true },
    supplier: { type: String, required: true },
    invoiceNumber: String,
    purchaseDate: { type: Date, default: Date.now },
    notes: String,
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now }
});

const SupplierPurchase = mongoose.model('SupplierPurchase', supplierPurchaseSchema);

// ============ AUTH MIDDLEWARE ============

const authMiddleware = async (req, res, next) => {
    try {
        const authHeader = req.headers.authorization;
        if (!authHeader || !authHeader.startsWith('Bearer ')) {
            return res.status(401).json({ success: false, message: 'No token provided' });
        }

        const token = authHeader.split(' ')[1];
        const decoded = jwt.verify(token, JWT_SECRET);

        const user = await User.findById(decoded.userId);
        if (!user) {
            return res.status(401).json({ success: false, message: 'User not found' });
        }

        req.user = user;
        next();
    } catch (error) {
        if (error.name === 'TokenExpiredError') {
            return res.status(401).json({ success: false, message: 'Token expired' });
        }
        return res.status(401).json({ success: false, message: 'Invalid token' });
    }
};

const adminMiddleware = (req, res, next) => {
    if (req.user.role !== 'admin') {
        return res.status(403).json({ success: false, message: 'Admin access required' });
    }
    next();
};

// ============ HELPER FUNCTIONS ============

const generateToken = (userId) => {
    return jwt.sign({ userId }, JWT_SECRET, { expiresIn: '7d' });
};

const sendEmail = async (to, subject, html) => {
    if (!emailTransporter) {
        console.log('Email not configured. Would send:', { to, subject });
        return false;
    }

    try {
        await emailTransporter.sendMail({
            from: process.env.EMAIL_FROM || 'noreply@acreprofit.com',
            to,
            subject,
            html
        });
        return true;
    } catch (error) {
        console.error('Email send error:', error);
        return false;
    }
};

const apiResponse = (res, status, success, data = null, message = '') => {
    return res.status(status).json({ success, data, message });
};

// ============ AUTH ROUTES ============

// POST /api/auth/signup
app.post('/api/auth/signup', async (req, res) => {
    try {
        const { email, password, name, phone } = req.body;

        if (!email || !password || !name) {
            return apiResponse(res, 400, false, null, 'Email, password, and name are required');
        }

        if (password.length < 8) {
            return apiResponse(res, 400, false, null, 'Password must be at least 8 characters');
        }

        const existingUser = await User.findOne({ email: email.toLowerCase() });
        if (existingUser) {
            return apiResponse(res, 400, false, null, 'Email already registered');
        }

        const user = new User({
            email: email.toLowerCase(),
            password,
            name,
            phone,
            role: 'customer'
        });

        await user.save();

        // Create empty customer profile
        const profile = new CustomerProfile({
            userId: user._id,
            acres: 0,
            crops: []
        });
        await profile.save();

        const token = generateToken(user._id);

        return apiResponse(res, 201, true, {
            user: user.toJSON(),
            token
        }, 'Account created successfully');
    } catch (error) {
        console.error('Signup error:', error);
        return apiResponse(res, 500, false, null, 'Error creating account');
    }
});

// POST /api/auth/login
app.post('/api/auth/login', async (req, res) => {
    try {
        const { email, password } = req.body;

        if (!email || !password) {
            return apiResponse(res, 400, false, null, 'Email and password are required');
        }

        const user = await User.findOne({ email: email.toLowerCase() });
        if (!user) {
            return apiResponse(res, 401, false, null, 'Invalid email or password');
        }

        const isMatch = await user.comparePassword(password);
        if (!isMatch) {
            return apiResponse(res, 401, false, null, 'Invalid email or password');
        }

        const token = generateToken(user._id);

        return apiResponse(res, 200, true, {
            user: user.toJSON(),
            token
        }, 'Login successful');
    } catch (error) {
        console.error('Login error:', error);
        return apiResponse(res, 500, false, null, 'Error logging in');
    }
});

// POST /api/auth/forgot-password
app.post('/api/auth/forgot-password', async (req, res) => {
    try {
        const { email } = req.body;

        if (!email) {
            return apiResponse(res, 400, false, null, 'Email is required');
        }

        const user = await User.findOne({ email: email.toLowerCase() });
        if (!user) {
            // Don't reveal if email exists
            return apiResponse(res, 200, true, null, 'If an account exists, a reset link has been sent');
        }

        const resetToken = crypto.randomBytes(32).toString('hex');
        const hashedToken = crypto.createHash('sha256').update(resetToken).digest('hex');

        user.resetPasswordToken = hashedToken;
        user.resetPasswordExpires = Date.now() + 3600000; // 1 hour
        await user.save();

        const resetUrl = `${req.headers.origin || 'https://acreprofit.com'}/reset-password.html?token=${resetToken}`;

        await sendEmail(
            user.email,
            'Reset Your AcreProfit Password',
            `
            <h2>Password Reset Request</h2>
            <p>Hello ${user.name},</p>
            <p>You requested a password reset. Click the link below to reset your password:</p>
            <p><a href="${resetUrl}" style="background-color: #22c55e; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; display: inline-block;">Reset Password</a></p>
            <p>This link expires in 1 hour.</p>
            <p>If you didn't request this, please ignore this email.</p>
            <p>- The AcreProfit Team</p>
            `
        );

        return apiResponse(res, 200, true, null, 'If an account exists, a reset link has been sent');
    } catch (error) {
        console.error('Forgot password error:', error);
        return apiResponse(res, 500, false, null, 'Error processing request');
    }
});

// POST /api/auth/reset-password
app.post('/api/auth/reset-password', async (req, res) => {
    try {
        const { token, password } = req.body;

        if (!token || !password) {
            return apiResponse(res, 400, false, null, 'Token and new password are required');
        }

        if (password.length < 8) {
            return apiResponse(res, 400, false, null, 'Password must be at least 8 characters');
        }

        const hashedToken = crypto.createHash('sha256').update(token).digest('hex');

        const user = await User.findOne({
            resetPasswordToken: hashedToken,
            resetPasswordExpires: { $gt: Date.now() }
        });

        if (!user) {
            return apiResponse(res, 400, false, null, 'Invalid or expired reset token');
        }

        user.password = password;
        user.resetPasswordToken = undefined;
        user.resetPasswordExpires = undefined;
        await user.save();

        return apiResponse(res, 200, true, null, 'Password reset successfully');
    } catch (error) {
        console.error('Reset password error:', error);
        return apiResponse(res, 500, false, null, 'Error resetting password');
    }
});

// GET /api/auth/me - Get current user
app.get('/api/auth/me', authMiddleware, async (req, res) => {
    return apiResponse(res, 200, true, { user: req.user.toJSON() });
});

// ============ PRODUCT ROUTES ============

// GET /api/products - List all active products
app.get('/api/products', async (req, res) => {
    try {
        const { category, search, featured } = req.query;

        let query = { active: true };

        if (category) {
            query.category = category;
        }

        if (featured === 'true') {
            query.featured = true;
        }

        if (search) {
            query.$text = { $search: search };
        }

        const products = await Product.find(query)
            .select('-supplierCost')
            .sort({ name: 1 });

        return apiResponse(res, 200, true, { products });
    } catch (error) {
        console.error('Get products error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching products');
    }
});

// GET /api/products/:id - Get single product
app.get('/api/products/:id', async (req, res) => {
    try {
        const product = await Product.findById(req.params.id).select('-supplierCost');

        if (!product) {
            return apiResponse(res, 404, false, null, 'Product not found');
        }

        return apiResponse(res, 200, true, { product });
    } catch (error) {
        console.error('Get product error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching product');
    }
});

// POST /api/products - Create product (admin only)
app.post('/api/products', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const product = new Product(req.body);
        await product.save();

        return apiResponse(res, 201, true, { product }, 'Product created successfully');
    } catch (error) {
        console.error('Create product error:', error);
        return apiResponse(res, 500, false, null, 'Error creating product');
    }
});

// PUT /api/products/:id - Update product (admin only)
app.put('/api/products/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const product = await Product.findByIdAndUpdate(
            req.params.id,
            { ...req.body, updatedAt: new Date() },
            { new: true, runValidators: true }
        );

        if (!product) {
            return apiResponse(res, 404, false, null, 'Product not found');
        }

        return apiResponse(res, 200, true, { product }, 'Product updated successfully');
    } catch (error) {
        console.error('Update product error:', error);
        return apiResponse(res, 500, false, null, 'Error updating product');
    }
});

// DELETE /api/products/:id - Soft delete product (admin only)
app.delete('/api/products/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const product = await Product.findByIdAndUpdate(
            req.params.id,
            { active: false, updatedAt: new Date() },
            { new: true }
        );

        if (!product) {
            return apiResponse(res, 404, false, null, 'Product not found');
        }

        return apiResponse(res, 200, true, null, 'Product deleted successfully');
    } catch (error) {
        console.error('Delete product error:', error);
        return apiResponse(res, 500, false, null, 'Error deleting product');
    }
});

// ============ ORDER ROUTES ============

// POST /api/orders - Create order with Stripe payment
app.post('/api/orders', authMiddleware, async (req, res) => {
    try {
        const { items, shippingAddress } = req.body;

        if (!items || !items.length) {
            return apiResponse(res, 400, false, null, 'Order items are required');
        }

        // Validate items and calculate totals
        let subtotal = 0;
        const orderItems = [];

        for (const item of items) {
            const product = await Product.findById(item.productId);
            if (!product) {
                return apiResponse(res, 400, false, null, `Product not found: ${item.productId}`);
            }

            const packageSize = product.packageSizes[item.packageSizeIndex];
            if (!packageSize) {
                return apiResponse(res, 400, false, null, `Invalid package size for ${product.name}`);
            }

            if (packageSize.inventory < item.quantity) {
                return apiResponse(res, 400, false, null, `Insufficient inventory for ${product.name}`);
            }

            const totalPrice = packageSize.price * item.quantity;
            subtotal += totalPrice;

            orderItems.push({
                product: product._id,
                productName: product.name,
                packageSize: {
                    volume: packageSize.volume,
                    unit: packageSize.unit,
                    price: packageSize.price
                },
                quantity: item.quantity,
                unitPrice: packageSize.price,
                totalPrice
            });

            // Decrement inventory
            product.packageSizes[item.packageSizeIndex].inventory -= item.quantity;
            await product.save();
        }

        const tax = Math.round(subtotal * 0.07 * 100) / 100; // 7% tax
        const shipping = subtotal >= 500 ? 0 : 25; // Free shipping over $500
        const total = subtotal + tax + shipping;

        // Create order
        const order = new Order({
            customerId: req.user._id,
            items: orderItems,
            subtotal,
            tax,
            shipping,
            total,
            shippingAddress: shippingAddress || req.user.address,
            status: 'pending',
            paymentStatus: 'pending'
        });

        await order.save();

        // Create Stripe payment intent
        let paymentIntent = null;
        if (stripe) {
            paymentIntent = await stripe.paymentIntents.create({
                amount: Math.round(total * 100), // Stripe uses cents
                currency: 'usd',
                metadata: {
                    orderId: order._id.toString(),
                    orderNumber: order.orderNumber
                }
            });

            order.stripePaymentIntentId = paymentIntent.id;
            await order.save();
        }

        return apiResponse(res, 201, true, {
            order,
            clientSecret: paymentIntent?.client_secret
        }, 'Order created successfully');
    } catch (error) {
        console.error('Create order error:', error);
        return apiResponse(res, 500, false, null, 'Error creating order');
    }
});

// GET /api/orders/mine - Get customer's orders
app.get('/api/orders/mine', authMiddleware, async (req, res) => {
    try {
        const orders = await Order.find({ customerId: req.user._id })
            .populate('items.product', 'name imageUrl')
            .sort({ createdAt: -1 });

        return apiResponse(res, 200, true, { orders });
    } catch (error) {
        console.error('Get my orders error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching orders');
    }
});

// GET /api/orders/:id - Get single order
app.get('/api/orders/:id', authMiddleware, async (req, res) => {
    try {
        const order = await Order.findById(req.params.id)
            .populate('items.product', 'name imageUrl')
            .populate('customerId', 'name email phone');

        if (!order) {
            return apiResponse(res, 404, false, null, 'Order not found');
        }

        // Check authorization
        if (req.user.role !== 'admin' && order.customerId._id.toString() !== req.user._id.toString()) {
            return apiResponse(res, 403, false, null, 'Not authorized to view this order');
        }

        return apiResponse(res, 200, true, { order });
    } catch (error) {
        console.error('Get order error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching order');
    }
});

// GET /api/orders - Get all orders (admin only)
app.get('/api/orders', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, startDate, endDate, page = 1, limit = 50 } = req.query;

        let query = {};

        if (status) {
            query.status = status;
        }

        if (startDate || endDate) {
            query.createdAt = {};
            if (startDate) query.createdAt.$gte = new Date(startDate);
            if (endDate) query.createdAt.$lte = new Date(endDate);
        }

        const skip = (parseInt(page) - 1) * parseInt(limit);

        const [orders, total] = await Promise.all([
            Order.find(query)
                .populate('customerId', 'name email phone')
                .sort({ createdAt: -1 })
                .skip(skip)
                .limit(parseInt(limit)),
            Order.countDocuments(query)
        ]);

        return apiResponse(res, 200, true, {
            orders,
            pagination: {
                page: parseInt(page),
                limit: parseInt(limit),
                total,
                pages: Math.ceil(total / parseInt(limit))
            }
        });
    } catch (error) {
        console.error('Get all orders error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching orders');
    }
});

// PUT /api/orders/:id - Update order status (admin only)
app.put('/api/orders/:id', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { status, trackingNumber, notes } = req.body;

        const updateData = { updatedAt: new Date() };
        if (status) updateData.status = status;
        if (trackingNumber) updateData.trackingNumber = trackingNumber;
        if (notes) updateData.notes = notes;

        const order = await Order.findByIdAndUpdate(
            req.params.id,
            updateData,
            { new: true }
        ).populate('customerId', 'name email');

        if (!order) {
            return apiResponse(res, 404, false, null, 'Order not found');
        }

        // Send email notification for status changes
        if (status && order.customerId.email) {
            await sendEmail(
                order.customerId.email,
                `Order ${order.orderNumber} - ${status.charAt(0).toUpperCase() + status.slice(1)}`,
                `
                <h2>Order Update</h2>
                <p>Hello ${order.customerId.name},</p>
                <p>Your order <strong>${order.orderNumber}</strong> status has been updated to: <strong>${status}</strong></p>
                ${trackingNumber ? `<p>Tracking Number: ${trackingNumber}</p>` : ''}
                <p>Thank you for choosing AcreProfit!</p>
                `
            );
        }

        return apiResponse(res, 200, true, { order }, 'Order updated successfully');
    } catch (error) {
        console.error('Update order error:', error);
        return apiResponse(res, 500, false, null, 'Error updating order');
    }
});

// ============ INVENTORY ROUTES ============

// POST /api/inventory/purchase - Log supplier purchase (admin only)
app.post('/api/inventory/purchase', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const { productId, packageSizeIndex, quantity, costPerUnit, supplier, invoiceNumber, notes } = req.body;

        if (!productId || packageSizeIndex === undefined || !quantity || !costPerUnit || !supplier) {
            return apiResponse(res, 400, false, null, 'Missing required fields');
        }

        const product = await Product.findById(productId);
        if (!product) {
            return apiResponse(res, 404, false, null, 'Product not found');
        }

        if (!product.packageSizes[packageSizeIndex]) {
            return apiResponse(res, 400, false, null, 'Invalid package size');
        }

        // Update inventory
        product.packageSizes[packageSizeIndex].inventory += quantity;
        product.supplierCost = costPerUnit;
        await product.save();

        // Log purchase
        const purchase = new SupplierPurchase({
            product: productId,
            productName: product.name,
            packageSizeIndex,
            quantity,
            costPerUnit,
            totalCost: quantity * costPerUnit,
            supplier,
            invoiceNumber,
            notes,
            createdBy: req.user._id
        });
        await purchase.save();

        return apiResponse(res, 201, true, { purchase, product }, 'Purchase logged and inventory updated');
    } catch (error) {
        console.error('Log purchase error:', error);
        return apiResponse(res, 500, false, null, 'Error logging purchase');
    }
});

// GET /api/inventory - Get full inventory (admin only)
app.get('/api/inventory', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const products = await Product.find({})
            .sort({ name: 1 });

        const inventory = products.map(p => ({
            _id: p._id,
            name: p.name,
            category: p.category,
            active: p.active,
            supplierCost: p.supplierCost,
            packageSizes: p.packageSizes.map(ps => ({
                volume: ps.volume,
                unit: ps.unit,
                price: ps.price,
                inventory: ps.inventory,
                sku: ps.sku
            }))
        }));

        return apiResponse(res, 200, true, { inventory });
    } catch (error) {
        console.error('Get inventory error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching inventory');
    }
});

// GET /api/inventory/purchases - Get purchase history (admin only)
app.get('/api/inventory/purchases', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const purchases = await SupplierPurchase.find({})
            .populate('product', 'name')
            .populate('createdBy', 'name')
            .sort({ purchaseDate: -1 })
            .limit(100);

        return apiResponse(res, 200, true, { purchases });
    } catch (error) {
        console.error('Get purchases error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching purchases');
    }
});

// ============ CUSTOMER PROFILE ROUTES ============

// GET /api/profile - Get customer profile
app.get('/api/profile', authMiddleware, async (req, res) => {
    try {
        let profile = await CustomerProfile.findOne({ userId: req.user._id });

        if (!profile) {
            profile = new CustomerProfile({
                userId: req.user._id,
                acres: 0,
                crops: []
            });
            await profile.save();
        }

        return apiResponse(res, 200, true, { profile, user: req.user.toJSON() });
    } catch (error) {
        console.error('Get profile error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching profile');
    }
});

// PUT /api/profile - Update customer profile
app.put('/api/profile', authMiddleware, async (req, res) => {
    try {
        const { acres, crops, name, phone, address } = req.body;

        // Update user info if provided
        if (name || phone || address) {
            const userUpdate = {};
            if (name) userUpdate.name = name;
            if (phone) userUpdate.phone = phone;
            if (address) userUpdate.address = address;

            await User.findByIdAndUpdate(req.user._id, userUpdate);
        }

        // Update profile
        let profile = await CustomerProfile.findOne({ userId: req.user._id });

        if (!profile) {
            profile = new CustomerProfile({ userId: req.user._id });
        }

        if (acres !== undefined) profile.acres = acres;
        if (crops) profile.crops = crops;
        profile.updatedAt = new Date();

        await profile.save();

        const user = await User.findById(req.user._id);

        return apiResponse(res, 200, true, { profile, user: user.toJSON() }, 'Profile updated successfully');
    } catch (error) {
        console.error('Update profile error:', error);
        return apiResponse(res, 500, false, null, 'Error updating profile');
    }
});

// GET /api/calculator - Get recommended products based on profile
app.get('/api/calculator', authMiddleware, async (req, res) => {
    try {
        const profile = await CustomerProfile.findOne({ userId: req.user._id });

        if (!profile || !profile.acres || !profile.crops.length) {
            return apiResponse(res, 400, false, null, 'Please set up your acres and crops first');
        }

        const products = await Product.find({ active: true }).select('-supplierCost');

        const recommendations = [];

        for (const crop of profile.crops) {
            const cropProducts = products.filter(p =>
                p.crops.includes(crop.name) || p.crops.length === 0
            );

            for (const product of cropProducts) {
                if (product.applicationRate && product.applicationRate.min) {
                    const rate = (product.applicationRate.min + (product.applicationRate.max || product.applicationRate.min)) / 2;
                    const totalNeeded = rate * crop.acres;

                    // Find best package size
                    const sortedPackages = [...product.packageSizes]
                        .filter(ps => ps.inventory > 0)
                        .sort((a, b) => {
                            const aOz = convertToOz(a.volume, a.unit);
                            const bOz = convertToOz(b.volume, b.unit);
                            return (a.price / aOz) - (b.price / bOz);
                        });

                    if (sortedPackages.length > 0) {
                        const bestPackage = sortedPackages[0];
                        const packageOz = convertToOz(bestPackage.volume, bestPackage.unit);
                        const unitsNeeded = Math.ceil(totalNeeded / packageOz);

                        recommendations.push({
                            crop: crop.name,
                            cropAcres: crop.acres,
                            product: {
                                _id: product._id,
                                name: product.name,
                                category: product.category,
                                applicationRate: product.applicationRate
                            },
                            recommendedPackage: bestPackage,
                            unitsNeeded,
                            totalOzNeeded: totalNeeded,
                            estimatedCost: unitsNeeded * bestPackage.price
                        });
                    }
                }
            }
        }

        return apiResponse(res, 200, true, {
            profile: {
                totalAcres: profile.acres,
                crops: profile.crops
            },
            recommendations
        });
    } catch (error) {
        console.error('Calculator error:', error);
        return apiResponse(res, 500, false, null, 'Error calculating recommendations');
    }
});

function convertToOz(volume, unit) {
    const conversions = {
        'oz': 1,
        'pt': 16,
        'qt': 32,
        'gal': 128,
        '2.5gal': 320,
        '30gal': 3840,
        'lb': 16,
        'case': 1
    };
    return volume * (conversions[unit] || 1);
}

// ============ ADMIN STATS ROUTES ============

// GET /api/admin/stats - Dashboard statistics
app.get('/api/admin/stats', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const today = new Date();
        const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
        const startOfYear = new Date(today.getFullYear(), 0, 1);

        const [
            totalCustomers,
            totalProducts,
            monthlyOrders,
            yearlyRevenue,
            recentOrders,
            lowStockProducts
        ] = await Promise.all([
            User.countDocuments({ role: 'customer' }),
            Product.countDocuments({ active: true }),
            Order.countDocuments({ createdAt: { $gte: startOfMonth } }),
            Order.aggregate([
                { $match: { createdAt: { $gte: startOfYear }, paymentStatus: 'paid' } },
                { $group: { _id: null, total: { $sum: '$total' } } }
            ]),
            Order.find({})
                .populate('customerId', 'name email')
                .sort({ createdAt: -1 })
                .limit(10),
            Product.find({
                active: true,
                'packageSizes.inventory': { $lt: 10 }
            }).select('name packageSizes')
        ]);

        return apiResponse(res, 200, true, {
            totalCustomers,
            totalProducts,
            monthlyOrders,
            yearlyRevenue: yearlyRevenue[0]?.total || 0,
            recentOrders,
            lowStockProducts
        });
    } catch (error) {
        console.error('Get admin stats error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching statistics');
    }
});

// GET /api/admin/customers - List all customers (admin only)
app.get('/api/admin/customers', authMiddleware, adminMiddleware, async (req, res) => {
    try {
        const customers = await User.find({ role: 'customer' })
            .sort({ createdAt: -1 });

        const customerData = await Promise.all(
            customers.map(async (customer) => {
                const profile = await CustomerProfile.findOne({ userId: customer._id });
                const orderStats = await Order.aggregate([
                    { $match: { customerId: customer._id } },
                    { $group: {
                        _id: null,
                        totalOrders: { $sum: 1 },
                        totalSpent: { $sum: '$total' }
                    }}
                ]);

                return {
                    ...customer.toJSON(),
                    profile,
                    orderStats: orderStats[0] || { totalOrders: 0, totalSpent: 0 }
                };
            })
        );

        return apiResponse(res, 200, true, { customers: customerData });
    } catch (error) {
        console.error('Get customers error:', error);
        return apiResponse(res, 500, false, null, 'Error fetching customers');
    }
});

// ============ STRIPE WEBHOOK ============

app.post('/api/payments/webhook', async (req, res) => {
    if (!stripe) {
        return res.status(500).send('Stripe not configured');
    }

    const sig = req.headers['stripe-signature'];
    let event;

    try {
        event = stripe.webhooks.constructEvent(
            req.body,
            sig,
            process.env.STRIPE_WEBHOOK_SECRET
        );
    } catch (err) {
        console.error('Webhook signature verification failed:', err.message);
        return res.status(400).send(`Webhook Error: ${err.message}`);
    }

    switch (event.type) {
        case 'payment_intent.succeeded':
            const paymentIntent = event.data.object;
            const orderId = paymentIntent.metadata.orderId;

            if (orderId) {
                await Order.findByIdAndUpdate(orderId, {
                    paymentStatus: 'paid',
                    status: 'processing',
                    stripePaymentMethodId: paymentIntent.payment_method
                });
                console.log(`Order ${orderId} marked as paid`);
            }
            break;

        case 'payment_intent.payment_failed':
            const failedIntent = event.data.object;
            const failedOrderId = failedIntent.metadata.orderId;

            if (failedOrderId) {
                // Restore inventory
                const order = await Order.findById(failedOrderId);
                if (order) {
                    for (const item of order.items) {
                        const product = await Product.findById(item.product);
                        if (product) {
                            const psIndex = product.packageSizes.findIndex(
                                ps => ps.volume === item.packageSize.volume && ps.unit === item.packageSize.unit
                            );
                            if (psIndex !== -1) {
                                product.packageSizes[psIndex].inventory += item.quantity;
                                await product.save();
                            }
                        }
                    }
                }

                await Order.findByIdAndUpdate(failedOrderId, {
                    paymentStatus: 'failed',
                    status: 'cancelled'
                });
            }
            break;
    }

    res.json({ received: true });
});

// POST /api/payments/create-intent - Create payment intent for existing order
app.post('/api/payments/create-intent', authMiddleware, async (req, res) => {
    try {
        const { orderId } = req.body;

        if (!stripe) {
            return apiResponse(res, 500, false, null, 'Payment system not configured');
        }

        const order = await Order.findById(orderId);
        if (!order) {
            return apiResponse(res, 404, false, null, 'Order not found');
        }

        if (order.customerId.toString() !== req.user._id.toString()) {
            return apiResponse(res, 403, false, null, 'Not authorized');
        }

        if (order.paymentStatus === 'paid') {
            return apiResponse(res, 400, false, null, 'Order already paid');
        }

        const paymentIntent = await stripe.paymentIntents.create({
            amount: Math.round(order.total * 100),
            currency: 'usd',
            metadata: {
                orderId: order._id.toString(),
                orderNumber: order.orderNumber
            }
        });

        order.stripePaymentIntentId = paymentIntent.id;
        await order.save();

        return apiResponse(res, 200, true, {
            clientSecret: paymentIntent.client_secret,
            amount: order.total
        });
    } catch (error) {
        console.error('Create payment intent error:', error);
        return apiResponse(res, 500, false, null, 'Error creating payment intent');
    }
});

// ============ SEED DATA ============

const seedInitialAdmin = async () => {
    try {
        const adminExists = await User.findOne({ role: 'admin' });
        if (!adminExists) {
            const admin = new User({
                email: 'kyle@acreprofit.com',
                password: 'ChangeThisPassword123!',
                name: 'Kyle McConnell',
                role: 'admin'
            });
            await admin.save();
            console.log('Admin user created: kyle@acreprofit.com');
        }
    } catch (error) {
        console.error('Seed admin error:', error);
    }
};

// ============ DATABASE CONNECTION ============

const connectDB = async () => {
    try {
        if (process.env.MONGODB_URI) {
            await mongoose.connect(process.env.MONGODB_URI);
            console.log('MongoDB connected successfully');
            await seedInitialAdmin();
        } else {
            console.log('No MongoDB URI provided - running in demo mode');
        }
    } catch (error) {
        console.error('MongoDB connection error:', error.message);
        process.exit(1);
    }
};

// ============ SERVER START ============

connectDB().then(() => {
    app.listen(PORT, () => {
        console.log(`AcreProfit API running on port ${PORT}`);
        console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
    });
});

module.exports = app;

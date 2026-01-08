require('dotenv').config();
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');

const app = express();
const PORT = process.env.PORT || 3001;

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

// Models
const farmerSchema = new mongoose.Schema({
    name: { type: String, required: true },
    email: { type: String, required: true, unique: true },
    phone: String,
    acres: { type: Number, required: true },
    crops: [String],
    location: {
        state: String,
        county: String
    },
    createdAt: { type: Date, default: Date.now }
});

const orderSchema = new mongoose.Schema({
    farmerId: { type: mongoose.Schema.Types.ObjectId, ref: 'Farmer' },
    product: { type: String, required: true },
    quantity: { type: Number, required: true },
    unit: String,
    acres: Number,
    status: {
        type: String,
        enum: ['pending', 'confirmed', 'shipped', 'delivered'],
        default: 'pending'
    },
    bulkOrderId: { type: mongoose.Schema.Types.ObjectId, ref: 'BulkOrder' },
    createdAt: { type: Date, default: Date.now }
});

const bulkOrderSchema = new mongoose.Schema({
    product: { type: String, required: true },
    targetQuantity: { type: Number, required: true },
    currentQuantity: { type: Number, default: 0 },
    pricePerUnit: Number,
    bulkPricePerUnit: Number,
    unit: String,
    status: {
        type: String,
        enum: ['open', 'filled', 'ordered', 'delivered'],
        default: 'open'
    },
    deadline: Date,
    createdAt: { type: Date, default: Date.now }
});

const Farmer = mongoose.model('Farmer', farmerSchema);
const Order = mongoose.model('Order', orderSchema);
const BulkOrder = mongoose.model('BulkOrder', bulkOrderSchema);

// Routes

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', message: 'Acre Profit API is running' });
});

// Farmer routes
app.post('/api/farmers', async (req, res) => {
    try {
        const farmer = new Farmer(req.body);
        await farmer.save();
        res.status(201).json(farmer);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.get('/api/farmers/:id', async (req, res) => {
    try {
        const farmer = await Farmer.findById(req.params.id);
        if (!farmer) {
            return res.status(404).json({ error: 'Farmer not found' });
        }
        res.json(farmer);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Bulk order routes
app.get('/api/bulk-orders', async (req, res) => {
    try {
        const orders = await BulkOrder.find({ status: 'open' });
        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

app.post('/api/bulk-orders', async (req, res) => {
    try {
        const bulkOrder = new BulkOrder(req.body);
        await bulkOrder.save();
        res.status(201).json(bulkOrder);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Join a bulk order
app.post('/api/orders', async (req, res) => {
    try {
        const { farmerId, bulkOrderId, quantity, acres } = req.body;

        const bulkOrder = await BulkOrder.findById(bulkOrderId);
        if (!bulkOrder) {
            return res.status(404).json({ error: 'Bulk order not found' });
        }

        const order = new Order({
            farmerId,
            product: bulkOrder.product,
            quantity,
            unit: bulkOrder.unit,
            acres,
            bulkOrderId
        });
        await order.save();

        // Update bulk order quantity
        bulkOrder.currentQuantity += quantity;
        if (bulkOrder.currentQuantity >= bulkOrder.targetQuantity) {
            bulkOrder.status = 'filled';
        }
        await bulkOrder.save();

        res.status(201).json(order);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Get orders for a farmer
app.get('/api/orders/farmer/:farmerId', async (req, res) => {
    try {
        const orders = await Order.find({ farmerId: req.params.farmerId })
            .populate('bulkOrderId');
        res.json(orders);
    } catch (error) {
        res.status(400).json({ error: error.message });
    }
});

// Calculate product needs based on acres
app.post('/api/calculate', (req, res) => {
    const { acres, product } = req.body;

    // Basic product calculations (these would be refined based on actual data)
    const calculations = {
        herbicide: { rate: 1.5, unit: 'gallons' },
        fertilizer: { rate: 200, unit: 'lbs' },
        insecticide: { rate: 0.5, unit: 'gallons' },
        fungicide: { rate: 1, unit: 'gallons' }
    };

    const calc = calculations[product] || { rate: 1, unit: 'units' };
    const amount = acres * calc.rate;

    res.json({
        acres,
        product,
        amount,
        unit: calc.unit,
        message: `For ${acres} acres, you'll need approximately ${amount} ${calc.unit} of ${product}`
    });
});

// Start server
connectDB().then(() => {
    app.listen(PORT, () => {
        console.log(`Acre Profit API running on port ${PORT}`);
    });
});

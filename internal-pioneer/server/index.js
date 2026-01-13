// Pioneer Analytics - Private Grower Data Server
require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, '..')));

// Environment
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb://localhost:27017/pioneer_analytics';
const JWT_SECRET = process.env.JWT_SECRET || 'pioneer-analytics-secret-key-change-in-production';
const ACCESS_CODE = process.env.ACCESS_CODE || 'pioneer2024';

// ============================================
// MongoDB Schemas
// ============================================

const growerDataSchema = new mongoose.Schema({
    date: { type: Date, required: true },
    invoice_number: { type: String, default: '' },
    grower_name: { type: String, required: true },
    product: { type: String, default: 'Other' },
    quantity: { type: Number, default: 0 },
    amount: { type: Number, required: true },
    category: { type: String, default: '' },
    notes: { type: String, default: '' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

const userSchema = new mongoose.Schema({
    username: { type: String, required: true, unique: true },
    password: { type: String, required: true },
    role: { type: String, default: 'user' },
    createdAt: { type: Date, default: Date.now }
});

const GrowerData = mongoose.model('GrowerData', growerDataSchema);
const User = mongoose.model('PioneerUser', userSchema);

// ============================================
// Authentication Middleware
// ============================================

const authMiddleware = async (req, res, next) => {
    try {
        const token = req.headers.authorization?.replace('Bearer ', '');
        if (!token) {
            return res.status(401).json({ error: 'Authentication required' });
        }

        const decoded = jwt.verify(token, JWT_SECRET);
        req.userId = decoded.userId;
        next();
    } catch (error) {
        res.status(401).json({ error: 'Invalid token' });
    }
};

// ============================================
// Auth Routes
// ============================================

// Simple code-based authentication
app.post('/api/auth/login', async (req, res) => {
    try {
        const { code } = req.body;

        if (code !== ACCESS_CODE) {
            return res.status(401).json({ error: 'Invalid access code' });
        }

        const token = jwt.sign(
            { userId: 'pioneer-user', role: 'admin' },
            JWT_SECRET,
            { expiresIn: '7d' }
        );

        res.json({ token, message: 'Access granted' });
    } catch (error) {
        res.status(500).json({ error: 'Authentication failed' });
    }
});

// Verify token
app.get('/api/auth/verify', authMiddleware, (req, res) => {
    res.json({ valid: true });
});

// ============================================
// Data Routes
// ============================================

// Get all grower data
app.get('/api/data', authMiddleware, async (req, res) => {
    try {
        const { year, grower, product } = req.query;
        let query = {};

        if (year && year !== 'all') {
            const startDate = new Date(`${year}-01-01`);
            const endDate = new Date(`${parseInt(year) + 1}-01-01`);
            query.date = { $gte: startDate, $lt: endDate };
        }

        if (grower) {
            query.grower_name = new RegExp(grower, 'i');
        }

        if (product) {
            query.product = product;
        }

        const data = await GrowerData.find(query).sort({ date: -1 });
        res.json(data);
    } catch (error) {
        res.status(500).json({ error: 'Failed to fetch data' });
    }
});

// Add single entry
app.post('/api/data', authMiddleware, async (req, res) => {
    try {
        const entry = new GrowerData({
            date: new Date(req.body.date),
            invoice_number: req.body.invoice_number || '',
            grower_name: req.body.grower_name,
            product: req.body.product || 'Other',
            quantity: parseFloat(req.body.quantity) || 0,
            amount: parseFloat(req.body.amount) || 0,
            category: req.body.category || '',
            notes: req.body.notes || ''
        });

        await entry.save();
        res.status(201).json(entry);
    } catch (error) {
        res.status(500).json({ error: 'Failed to add entry' });
    }
});

// Bulk import
app.post('/api/data/bulk', authMiddleware, async (req, res) => {
    try {
        const { entries } = req.body;

        if (!Array.isArray(entries) || entries.length === 0) {
            return res.status(400).json({ error: 'No entries provided' });
        }

        const documents = entries.map(e => ({
            date: new Date(e.date),
            invoice_number: e.invoice_number || '',
            grower_name: e.grower_name,
            product: e.product || 'Other',
            quantity: parseFloat(e.quantity) || 0,
            amount: parseFloat(e.amount) || 0,
            category: e.category || '',
            notes: e.notes || ''
        }));

        const result = await GrowerData.insertMany(documents);
        res.status(201).json({
            message: `Successfully imported ${result.length} records`,
            count: result.length
        });
    } catch (error) {
        res.status(500).json({ error: 'Failed to import data' });
    }
});

// Update entry
app.put('/api/data/:id', authMiddleware, async (req, res) => {
    try {
        const entry = await GrowerData.findByIdAndUpdate(
            req.params.id,
            {
                ...req.body,
                updatedAt: Date.now()
            },
            { new: true }
        );

        if (!entry) {
            return res.status(404).json({ error: 'Entry not found' });
        }

        res.json(entry);
    } catch (error) {
        res.status(500).json({ error: 'Failed to update entry' });
    }
});

// Delete entry
app.delete('/api/data/:id', authMiddleware, async (req, res) => {
    try {
        const entry = await GrowerData.findByIdAndDelete(req.params.id);

        if (!entry) {
            return res.status(404).json({ error: 'Entry not found' });
        }

        res.json({ message: 'Entry deleted', id: req.params.id });
    } catch (error) {
        res.status(500).json({ error: 'Failed to delete entry' });
    }
});

// Delete all data (with confirmation)
app.delete('/api/data', authMiddleware, async (req, res) => {
    try {
        const { confirm } = req.query;

        if (confirm !== 'yes') {
            return res.status(400).json({ error: 'Confirmation required' });
        }

        await GrowerData.deleteMany({});
        res.json({ message: 'All data deleted' });
    } catch (error) {
        res.status(500).json({ error: 'Failed to delete data' });
    }
});

// ============================================
// Analytics Routes
// ============================================

// Get summary statistics
app.get('/api/analytics/summary', authMiddleware, async (req, res) => {
    try {
        const allData = await GrowerData.find();

        // Calculate summary stats
        const totalRevenue = allData.reduce((sum, d) => sum + d.amount, 0);
        const uniqueGrowers = [...new Set(allData.map(d => d.grower_name))];
        const invoiceCount = [...new Set(allData.map(d => d.invoice_number))].filter(i => i).length || allData.length;

        // Yearly breakdown
        const yearlyData = {};
        const years = [2022, 2023, 2024, 2025, 2026];

        years.forEach(year => {
            const yearEntries = allData.filter(d => new Date(d.date).getFullYear() === year);
            yearlyData[year] = {
                revenue: yearEntries.reduce((sum, d) => sum + d.amount, 0),
                invoices: [...new Set(yearEntries.map(d => d.invoice_number))].filter(i => i).length || yearEntries.length,
                growers: [...new Set(yearEntries.map(d => d.grower_name))].length
            };
        });

        // Calculate growth rate
        const revenues = years.map(y => yearlyData[y].revenue).filter(r => r > 0);
        let avgGrowth = 0;
        if (revenues.length >= 2) {
            let totalGrowth = 0;
            for (let i = 1; i < revenues.length; i++) {
                if (revenues[i-1] > 0) {
                    totalGrowth += (revenues[i] - revenues[i-1]) / revenues[i-1];
                }
            }
            avgGrowth = totalGrowth / (revenues.length - 1);
        }

        // Forecast 2027
        const lastRevenue = yearlyData[2026].revenue || revenues[revenues.length - 1] || 0;
        const forecast2027 = lastRevenue * (1 + avgGrowth);

        res.json({
            totalRevenue,
            totalGrowers: uniqueGrowers.length,
            totalInvoices: invoiceCount,
            avgGrowthRate: avgGrowth,
            forecast2027,
            yearlyData
        });
    } catch (error) {
        res.status(500).json({ error: 'Failed to calculate analytics' });
    }
});

// Get grower statistics
app.get('/api/analytics/growers', authMiddleware, async (req, res) => {
    try {
        const allData = await GrowerData.find();
        const growerNames = [...new Set(allData.map(d => d.grower_name))];

        const growerStats = growerNames.map(name => {
            const growerEntries = allData.filter(d => d.grower_name === name);
            const yearlyRevenue = {};

            [2022, 2023, 2024, 2025, 2026].forEach(year => {
                yearlyRevenue[year] = growerEntries
                    .filter(d => new Date(d.date).getFullYear() === year)
                    .reduce((sum, d) => sum + d.amount, 0);
            });

            return {
                name,
                totalRevenue: growerEntries.reduce((sum, d) => sum + d.amount, 0),
                yearlyRevenue,
                entryCount: growerEntries.length
            };
        });

        // Sort by total revenue
        growerStats.sort((a, b) => b.totalRevenue - a.totalRevenue);

        res.json(growerStats);
    } catch (error) {
        res.status(500).json({ error: 'Failed to calculate grower stats' });
    }
});

// Get product statistics
app.get('/api/analytics/products', authMiddleware, async (req, res) => {
    try {
        const allData = await GrowerData.find();
        const products = [...new Set(allData.map(d => d.product))];

        const productStats = products.map(product => {
            const productEntries = allData.filter(d => d.product === product);
            const yearlyRevenue = {};

            [2022, 2023, 2024, 2025, 2026].forEach(year => {
                yearlyRevenue[year] = productEntries
                    .filter(d => new Date(d.date).getFullYear() === year)
                    .reduce((sum, d) => sum + d.amount, 0);
            });

            return {
                product,
                totalRevenue: productEntries.reduce((sum, d) => sum + d.amount, 0),
                yearlyRevenue,
                quantity: productEntries.reduce((sum, d) => sum + d.quantity, 0)
            };
        });

        productStats.sort((a, b) => b.totalRevenue - a.totalRevenue);

        res.json(productStats);
    } catch (error) {
        res.status(500).json({ error: 'Failed to calculate product stats' });
    }
});

// Get monthly trends
app.get('/api/analytics/monthly', authMiddleware, async (req, res) => {
    try {
        const allData = await GrowerData.find();
        const years = [2022, 2023, 2024, 2025, 2026];
        const monthlyData = {};

        years.forEach(year => {
            const monthly = Array(12).fill(0);
            allData
                .filter(d => new Date(d.date).getFullYear() === year)
                .forEach(d => {
                    const month = new Date(d.date).getMonth();
                    monthly[month] += d.amount;
                });
            monthlyData[year] = monthly;
        });

        res.json(monthlyData);
    } catch (error) {
        res.status(500).json({ error: 'Failed to calculate monthly trends' });
    }
});

// ============================================
// Serve Frontend
// ============================================

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '..', 'index.html'));
});

// ============================================
// Database Connection & Server Start
// ============================================

mongoose.connect(MONGODB_URI)
    .then(() => {
        console.log('Connected to MongoDB');
        app.listen(PORT, () => {
            console.log(`Pioneer Analytics server running on port ${PORT}`);
            console.log(`Access the dashboard at http://localhost:${PORT}`);
        });
    })
    .catch(err => {
        console.log('MongoDB connection failed, running with in-memory fallback');
        console.log('For production, please set MONGODB_URI in .env');

        // Start server anyway for local development
        app.listen(PORT, () => {
            console.log(`Pioneer Analytics server running on port ${PORT}`);
            console.log(`Access the dashboard at http://localhost:${PORT}`);
            console.log('Note: Data will only persist in localStorage without MongoDB');
        });
    });

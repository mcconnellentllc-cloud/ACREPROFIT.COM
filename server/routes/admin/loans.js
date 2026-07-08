// server/routes/admin/loans.js
// Distributor loan tracking - loans/product credits from distributors to Acre Profit LLC
// Access: Distributors and SuperAdmins only (middleware applied at mount point)

const express = require('express');
const router = express.Router();
const mongoose = require('mongoose');

const getModels = () => ({
    DistributorLoan: mongoose.model('DistributorLoan'),
    User: mongoose.model('User')
});

// GET / - List all loans and payments, with running balance per lender
router.get('/', async (req, res) => {
    try {
        const { DistributorLoan } = getModels();
        const { lenderId, category } = req.query;

        const query = {};
        if (lenderId) query.lenderId = lenderId;
        if (category) query.category = category;

        const entries = await DistributorLoan.find(query)
            .populate('lenderId', 'name email')
            .populate('createdBy', 'name')
            .sort({ date: -1, createdAt: -1 })
            .lean();

        res.json({ entries });
    } catch (error) {
        console.error('GET /api/admin/loans error:', error);
        res.status(500).json({ error: error.message });
    }
});

// GET /summary - Summary: total owed to each lender, total repaid, net balance
router.get('/summary', async (req, res) => {
    try {
        const { DistributorLoan, User } = getModels();

        // Get all distributors (potential lenders)
        const distributors = await User.find({ role: { $in: ['distributor', 'superadmin'] } })
            .select('name email')
            .lean();

        const summaries = [];

        for (const dist of distributors) {
            const entries = await DistributorLoan.find({ lenderId: dist._id })
                .sort({ date: -1, createdAt: -1 })
                .lean();

            if (entries.length === 0) continue;

            let totalLoaned = 0;
            let totalRepaid = 0;

            for (const entry of entries) {
                const amt = Number(entry.amount || 0);
                if (entry.category === 'loan' || entry.category === 'product_credit') {
                    totalLoaned += amt;
                } else if (entry.category === 'repayment') {
                    totalRepaid += amt;
                } else if (entry.category === 'adjustment') {
                    // Adjustments can go either direction; positive = increases debt
                    totalLoaned += amt;
                }
            }

            const netBalance = Math.round((totalLoaned - totalRepaid) * 100) / 100;
            const latestEntry = entries[0]; // already sorted desc

            summaries.push({
                lenderId: dist._id,
                lenderName: dist.name,
                lenderEmail: dist.email,
                totalLoaned: Math.round(totalLoaned * 100) / 100,
                totalRepaid: Math.round(totalRepaid * 100) / 100,
                netBalance,
                entryCount: entries.length,
                lastActivity: latestEntry ? latestEntry.date : null,
                currentRunningBalance: latestEntry ? latestEntry.runningBalance : 0
            });
        }

        const grandTotalLoaned = summaries.reduce((s, x) => s + x.totalLoaned, 0);
        const grandTotalRepaid = summaries.reduce((s, x) => s + x.totalRepaid, 0);
        const grandNetBalance = summaries.reduce((s, x) => s + x.netBalance, 0);

        res.json({
            summaries,
            totals: {
                totalLoaned: Math.round(grandTotalLoaned * 100) / 100,
                totalRepaid: Math.round(grandTotalRepaid * 100) / 100,
                netBalance: Math.round(grandNetBalance * 100) / 100
            }
        });
    } catch (error) {
        console.error('GET /api/admin/loans/summary error:', error);
        res.status(500).json({ error: error.message });
    }
});

// POST / - Create new loan or repayment record
router.post('/', async (req, res) => {
    try {
        const { DistributorLoan, User } = getModels();
        const { lenderId, amount, description, category, date, referenceNumber, notes } = req.body;

        if (!lenderId) return res.status(400).json({ error: 'Lender is required.' });
        if (!description) return res.status(400).json({ error: 'Description is required.' });
        if (!category) return res.status(400).json({ error: 'Category is required.' });

        const amt = Number(amount);
        if (!(amt > 0)) return res.status(400).json({ error: 'Amount must be a positive number.' });

        const lender = await User.findById(lenderId).select('name');
        if (!lender) return res.status(404).json({ error: 'Lender not found.' });

        // Compute running balance: find the most recent entry for this lender
        const lastEntry = await DistributorLoan.findOne({ lenderId })
            .sort({ date: -1, createdAt: -1 })
            .lean();

        const previousBalance = lastEntry ? Number(lastEntry.runningBalance || 0) : 0;

        // loan / product_credit / adjustment increase what business owes
        // repayment decreases what business owes
        let balanceChange;
        if (category === 'repayment') {
            balanceChange = -amt;
        } else {
            balanceChange = amt;
        }
        const newBalance = Math.round((previousBalance + balanceChange) * 100) / 100;

        const entry = new DistributorLoan({
            lenderId,
            lenderName: lender.name,
            borrower: 'Acre Profit LLC',
            amount: amt,
            description,
            category,
            date: date ? new Date(date) : new Date(),
            referenceNumber: referenceNumber || '',
            notes: notes || '',
            runningBalance: newBalance,
            createdBy: req.user._id
        });

        await entry.save();

        const populated = await DistributorLoan.findById(entry._id)
            .populate('lenderId', 'name email')
            .populate('createdBy', 'name')
            .lean();

        res.status(201).json(populated);
    } catch (error) {
        console.error('POST /api/admin/loans error:', error);
        res.status(400).json({ error: error.message });
    }
});

// GET /:lenderId - Get loan history for a specific lender
router.get('/:lenderId', async (req, res) => {
    try {
        const { DistributorLoan, User } = getModels();
        const { lenderId } = req.params;

        if (!mongoose.Types.ObjectId.isValid(lenderId)) {
            return res.status(400).json({ error: 'Invalid lender ID.' });
        }

        const lender = await User.findById(lenderId).select('name email').lean();
        if (!lender) return res.status(404).json({ error: 'Lender not found.' });

        const entries = await DistributorLoan.find({ lenderId })
            .populate('createdBy', 'name')
            .sort({ date: -1, createdAt: -1 })
            .lean();

        let totalLoaned = 0;
        let totalRepaid = 0;

        for (const entry of entries) {
            const amt = Number(entry.amount || 0);
            if (entry.category === 'loan' || entry.category === 'product_credit') {
                totalLoaned += amt;
            } else if (entry.category === 'repayment') {
                totalRepaid += amt;
            } else if (entry.category === 'adjustment') {
                totalLoaned += amt;
            }
        }

        const netBalance = Math.round((totalLoaned - totalRepaid) * 100) / 100;
        const currentRunningBalance = entries.length > 0 ? entries[0].runningBalance : 0;

        res.json({
            lender,
            entries,
            totalLoaned: Math.round(totalLoaned * 100) / 100,
            totalRepaid: Math.round(totalRepaid * 100) / 100,
            netBalance,
            currentRunningBalance
        });
    } catch (error) {
        console.error('GET /api/admin/loans/:lenderId error:', error);
        res.status(500).json({ error: error.message });
    }
});

// DELETE /:id - Delete a loan entry and recalculate running balances
router.delete('/:id', async (req, res) => {
    try {
        const { DistributorLoan } = getModels();
        const entry = await DistributorLoan.findById(req.params.id);
        if (!entry) return res.status(404).json({ error: 'Loan entry not found.' });

        const lenderId = entry.lenderId;
        await DistributorLoan.findByIdAndDelete(req.params.id);

        // Recalculate running balances for this lender
        const remaining = await DistributorLoan.find({ lenderId })
            .sort({ date: 1, createdAt: 1 })
            .exec();

        let runningBalance = 0;
        for (const e of remaining) {
            const amt = Number(e.amount || 0);
            if (e.category === 'repayment') {
                runningBalance -= amt;
            } else {
                runningBalance += amt;
            }
            e.runningBalance = Math.round(runningBalance * 100) / 100;
            await e.save();
        }

        res.json({ message: 'Entry deleted and balances recalculated.' });
    } catch (error) {
        console.error('DELETE /api/admin/loans/:id error:', error);
        res.status(500).json({ error: error.message });
    }
});

module.exports = router;

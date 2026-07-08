// server/routes/admin/supplier-invoices.js
// Supplier invoice management - tracks goods received from suppliers
// Access: Distributors and SuperAdmins only

const express = require('express');
const router = express.Router();
const mongoose = require('mongoose');

// Models will be accessed via mongoose.model() since they're defined in index.js
const getModels = () => ({
    SupplierInvoice: mongoose.model('SupplierInvoice'),
    InventoryBatch: mongoose.model('InventoryBatch'),
    Inventory: mongoose.model('Inventory'),
    InventoryTransaction: mongoose.model('InventoryTransaction'),
    PurchaseOrder: mongoose.model('PurchaseOrder'),
    Chemical: mongoose.model('Chemical')
});

// GET /api/admin/supplier-invoices - List all supplier invoices
router.get('/', async (req, res) => {
    try {
        const { SupplierInvoice } = getModels();
        const { status, paymentStatus, supplierId, startDate, endDate, limit = 50, skip = 0 } = req.query;

        const query = {};
        if (status) query.status = status;
        if (paymentStatus) query.paymentStatus = paymentStatus;
        if (supplierId) query.supplierId = supplierId;
        if (startDate || endDate) {
            query.receivedDate = {};
            if (startDate) query.receivedDate.$gte = new Date(startDate);
            if (endDate) query.receivedDate.$lte = new Date(endDate);
        }

        const [invoices, total] = await Promise.all([
            SupplierInvoice.find(query)
                .sort({ receivedDate: -1 })
                .skip(parseInt(skip))
                .limit(parseInt(limit))
                .lean(),
            SupplierInvoice.countDocuments(query)
        ]);

        res.json({ invoices, total, limit: parseInt(limit), skip: parseInt(skip) });
    } catch (err) {
        console.error('GET /supplier-invoices error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/supplier-invoices/:id - Get single supplier invoice
router.get('/:id', async (req, res) => {
    try {
        const { SupplierInvoice } = getModels();
        const invoice = await SupplierInvoice.findById(req.params.id).lean();
        if (!invoice) {
            return res.status(404).json({ error: 'Supplier invoice not found' });
        }
        res.json(invoice);
    } catch (err) {
        console.error('GET /supplier-invoices/:id error:', err);
        res.status(500).json({ error: err.message });
    }
});

// POST /api/admin/supplier-invoices - Create supplier invoice and receive goods
// This is the main "Receive Shipment" endpoint - creates invoice, batches, and updates inventory
router.post('/', async (req, res) => {
    const session = await mongoose.startSession();
    session.startTransaction();

    try {
        const { SupplierInvoice, InventoryBatch, Inventory, InventoryTransaction, Chemical } = getModels();

        const {
            invoiceNumber,
            purchaseOrderId,
            poNumber,
            supplierId,
            supplierName,
            supplierContact,
            items,
            subtotal,
            freight = 0,
            otherFees = 0,
            taxAmount = 0,
            totalAmount,
            paymentDueDate,
            paymentTerms,
            invoiceDate,
            receivingLocation = 'main',
            receivingNotes,
            documentUrl
        } = req.body;

        // Validate required fields
        if (!invoiceNumber || !supplierName || !items || items.length === 0) {
            await session.abortTransaction();
            return res.status(400).json({ error: 'Invoice number, supplier name, and items are required' });
        }

        // Create the supplier invoice
        const supplierInvoice = new SupplierInvoice({
            invoiceNumber,
            purchaseOrderId,
            poNumber,
            supplierId,
            supplierName,
            supplierContact,
            items: items.map(item => ({
                chemicalId: item.chemicalId,
                productName: item.productName,
                packSize: item.packSize,
                unit: item.unit,
                quantityOrdered: item.quantityOrdered || item.quantityReceived,
                quantityReceived: item.quantityReceived,
                quantityDamaged: item.quantityDamaged || 0,
                unitCost: item.unitCost,
                totalCost: item.unitCost * item.quantityReceived,
                lotNumber: item.lotNumber,
                expirationDate: item.expirationDate,
                manufactureDate: item.manufactureDate
            })),
            subtotal: subtotal || items.reduce((sum, i) => sum + (i.unitCost * i.quantityReceived), 0),
            freight,
            otherFees,
            taxAmount,
            totalAmount: totalAmount || (subtotal + freight + otherFees + taxAmount),
            paymentDueDate,
            paymentTerms,
            invoiceDate: invoiceDate || new Date(),
            receivedDate: new Date(),
            receivedBy: req.user._id,
            receivedByName: req.user.name || req.user.email,
            receivingLocation,
            receivingNotes,
            documentUrl,
            status: 'received',
            createdBy: req.user._id
        });

        await supplierInvoice.save({ session });

        // Create inventory batches and update inventory for each item
        const batchesCreated = [];
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const netQuantity = item.quantityReceived - (item.quantityDamaged || 0);

            if (netQuantity <= 0) continue;

            // Create inventory batch
            const batch = new InventoryBatch({
                chemicalId: item.chemicalId,
                productName: item.productName,
                packSize: item.packSize,
                unit: item.unit,
                poNumber: poNumber || `SI-${supplierInvoice.internalReference}`,
                poId: purchaseOrderId,
                lotNumber: item.lotNumber,
                quantityReceived: netQuantity,
                quantityRemaining: netQuantity,
                quantitySold: 0,
                costPerUnit: item.unitCost,
                totalCost: item.unitCost * netQuantity,
                location: receivingLocation,
                status: 'active',
                receivedDate: new Date(),
                expirationDate: item.expirationDate,
                manufactureDate: item.manufactureDate,
                supplierId,
                supplierName,
                supplierInvoiceId: supplierInvoice._id
            });

            await batch.save({ session });
            batchesCreated.push(batch);

            // Update the supplier invoice item with batch reference
            supplierInvoice.items[i].inventoryBatchId = batch._id;

            // Update or create inventory record
            let inventory = await Inventory.findOne({
                chemicalId: item.chemicalId,
                location: receivingLocation
            }).session(session);

            const previousQty = inventory ? inventory.quantityOnHand : 0;

            if (inventory) {
                // Update existing inventory - recalculate average cost
                const totalOldValue = inventory.quantityOnHand * inventory.averageCost;
                const totalNewValue = netQuantity * item.unitCost;
                const newTotalQty = inventory.quantityOnHand + netQuantity;
                const newAvgCost = newTotalQty > 0 ? (totalOldValue + totalNewValue) / newTotalQty : item.unitCost;

                inventory.quantityOnHand = newTotalQty;
                inventory.quantityAvailable = newTotalQty - inventory.quantityReserved;
                inventory.averageCost = Math.round(newAvgCost * 10000) / 10000;
                inventory.lastCost = item.unitCost;
                inventory.lastReceivedDate = new Date();
                inventory.updatedAt = new Date();
            } else {
                // Create new inventory record
                const chemical = await Chemical.findById(item.chemicalId).lean();
                inventory = new Inventory({
                    chemicalId: item.chemicalId,
                    productName: item.productName || (chemical ? chemical.productName : 'Unknown'),
                    packSize: item.packSize,
                    unit: item.unit,
                    quantityOnHand: netQuantity,
                    quantityReserved: 0,
                    quantityAvailable: netQuantity,
                    location: receivingLocation,
                    averageCost: item.unitCost,
                    lastCost: item.unitCost,
                    lastReceivedDate: new Date()
                });
            }

            await inventory.save({ session });

            // Create inventory transaction for audit trail
            const transaction = new InventoryTransaction({
                inventoryId: inventory._id,
                chemicalId: item.chemicalId,
                productName: item.productName,
                type: 'receive',
                quantityChange: netQuantity,
                previousQuantity: previousQty,
                newQuantity: inventory.quantityOnHand,
                unitCost: item.unitCost,
                totalCost: item.unitCost * netQuantity,
                referenceType: 'PurchaseOrder',
                referenceId: supplierInvoice._id,
                referenceNumber: supplierInvoice.internalReference,
                location: receivingLocation,
                notes: `Received from ${supplierName} - Invoice #${invoiceNumber}${item.lotNumber ? ` - Lot: ${item.lotNumber}` : ''}`,
                createdBy: req.user._id
            });

            await transaction.save({ session });
        }

        // Save supplier invoice with batch references
        await supplierInvoice.save({ session });

        await session.commitTransaction();

        res.status(201).json({
            success: true,
            supplierInvoice,
            batchesCreated: batchesCreated.length,
            message: `Received ${batchesCreated.length} items into inventory`
        });
    } catch (err) {
        await session.abortTransaction();
        console.error('POST /supplier-invoices error:', err);
        res.status(500).json({ error: err.message });
    } finally {
        session.endSession();
    }
});

// PATCH /api/admin/supplier-invoices/:id/payment - Record payment to supplier
router.patch('/:id/payment', async (req, res) => {
    try {
        const { SupplierInvoice } = getModels();
        const { amount, method, checkNumber, reference, date } = req.body;

        if (!amount || amount <= 0) {
            return res.status(400).json({ error: 'Payment amount is required and must be positive' });
        }

        const invoice = await SupplierInvoice.findById(req.params.id);
        if (!invoice) {
            return res.status(404).json({ error: 'Supplier invoice not found' });
        }

        // Add payment record
        invoice.payments.push({
            amount,
            date: date || new Date(),
            method: method || 'check',
            checkNumber,
            reference,
            recordedBy: req.user._id
        });

        // Update amount paid and status
        invoice.amountPaid = (invoice.amountPaid || 0) + amount;

        if (invoice.amountPaid >= invoice.totalAmount) {
            invoice.paymentStatus = 'paid';
        } else if (invoice.amountPaid > 0) {
            invoice.paymentStatus = 'partial';
        }

        invoice.updatedAt = new Date();
        await invoice.save();

        res.json({
            success: true,
            invoice,
            message: `Payment of $${amount.toFixed(2)} recorded`
        });
    } catch (err) {
        console.error('PATCH /supplier-invoices/:id/payment error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/supplier-invoices/unpaid - Get unpaid supplier invoices
router.get('/status/unpaid', async (req, res) => {
    try {
        const { SupplierInvoice } = getModels();

        const invoices = await SupplierInvoice.find({
            paymentStatus: { $in: ['unpaid', 'partial'] }
        })
            .sort({ paymentDueDate: 1, receivedDate: -1 })
            .lean();

        const totalOwed = invoices.reduce((sum, inv) => sum + (inv.totalAmount - (inv.amountPaid || 0)), 0);

        res.json({
            invoices,
            count: invoices.length,
            totalOwed
        });
    } catch (err) {
        console.error('GET /supplier-invoices/status/unpaid error:', err);
        res.status(500).json({ error: err.message });
    }
});

// DELETE /api/admin/supplier-invoices/:id - Void a supplier invoice (superadmin only)
router.delete('/:id', async (req, res) => {
    // This should be restricted to superadmin - checked in index.js middleware
    if (req.user.role !== 'superadmin') {
        return res.status(403).json({ error: 'Only super admins can void supplier invoices' });
    }

    const session = await mongoose.startSession();
    session.startTransaction();

    try {
        const { SupplierInvoice, InventoryBatch, Inventory, InventoryTransaction } = getModels();

        const invoice = await SupplierInvoice.findById(req.params.id).session(session);
        if (!invoice) {
            await session.abortTransaction();
            return res.status(404).json({ error: 'Supplier invoice not found' });
        }

        // Check if any batches have been sold
        const batches = await InventoryBatch.find({ supplierInvoiceId: invoice._id }).session(session);
        const hasSales = batches.some(b => b.quantitySold > 0);

        if (hasSales) {
            await session.abortTransaction();
            return res.status(400).json({ error: 'Cannot void invoice - some items have already been sold' });
        }

        // Reverse inventory for each batch
        for (const batch of batches) {
            const inventory = await Inventory.findOne({
                chemicalId: batch.chemicalId,
                location: batch.location
            }).session(session);

            if (inventory) {
                const previousQty = inventory.quantityOnHand;
                inventory.quantityOnHand -= batch.quantityRemaining;
                inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
                inventory.updatedAt = new Date();
                await inventory.save({ session });

                // Create reversal transaction
                const transaction = new InventoryTransaction({
                    inventoryId: inventory._id,
                    chemicalId: batch.chemicalId,
                    productName: batch.productName,
                    type: 'adjustment',
                    quantityChange: -batch.quantityRemaining,
                    previousQuantity: previousQty,
                    newQuantity: inventory.quantityOnHand,
                    referenceType: 'Manual',
                    referenceId: invoice._id,
                    referenceNumber: invoice.internalReference,
                    location: batch.location,
                    reason: 'Supplier invoice voided',
                    notes: `Voided SI ${invoice.internalReference} - reversed ${batch.quantityRemaining} ${batch.unit}`,
                    createdBy: req.user._id
                });
                await transaction.save({ session });
            }

            // Mark batch as returned/voided
            batch.status = 'returned';
            batch.quantityRemaining = 0;
            batch.updatedAt = new Date();
            await batch.save({ session });
        }

        // Mark invoice as voided
        invoice.status = 'voided';
        invoice.updatedAt = new Date();
        await invoice.save({ session });

        await session.commitTransaction();

        res.json({
            success: true,
            message: `Supplier invoice ${invoice.internalReference} voided and inventory reversed`
        });
    } catch (err) {
        await session.abortTransaction();
        console.error('DELETE /supplier-invoices/:id error:', err);
        res.status(500).json({ error: err.message });
    } finally {
        session.endSession();
    }
});

module.exports = router;

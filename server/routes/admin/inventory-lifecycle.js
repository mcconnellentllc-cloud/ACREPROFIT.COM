// server/routes/admin/inventory-lifecycle.js
// Inventory lifecycle management - allocation, fulfillment, payment verification
// Access: Distributors and SuperAdmins only

const express = require('express');
const router = express.Router();
const mongoose = require('mongoose');

const getModels = () => ({
    ChemicalOrder: mongoose.model('ChemicalOrder'),
    InventoryBatch: mongoose.model('InventoryBatch'),
    Inventory: mongoose.model('Inventory'),
    InventoryTransaction: mongoose.model('InventoryTransaction'),
    Invoice: mongoose.model('Invoice'),
    Chemical: mongoose.model('Chemical')
});

// POST /api/admin/inventory/allocate-order/:orderId
// Allocate inventory batches to an order using FIFO
router.post('/allocate-order/:orderId', async (req, res) => {
    const session = await mongoose.startSession();
    session.startTransaction();

    try {
        const { ChemicalOrder, InventoryBatch, Inventory, InventoryTransaction } = getModels();

        const order = await ChemicalOrder.findById(req.params.orderId).session(session);
        if (!order) {
            await session.abortTransaction();
            return res.status(404).json({ error: 'Order not found' });
        }

        if (order.inventoryStatus === 'fully_allocated' || order.inventoryStatus === 'delivered') {
            await session.abortTransaction();
            return res.status(400).json({ error: 'Order already allocated or delivered' });
        }

        const location = req.body.location || 'main';
        const allocations = [];
        let allItemsFullyAllocated = true;

        for (let i = 0; i < order.items.length; i++) {
            const item = order.items[i];
            if (!item.chemicalId) continue;

            let remainingToAllocate = item.quantity;
            item.batchAllocations = item.batchAllocations || [];

            // Find active batches for this product, sorted by received date (FIFO)
            const batches = await InventoryBatch.find({
                chemicalId: item.chemicalId,
                location,
                status: 'active',
                quantityRemaining: { $gt: 0 }
            })
                .sort({ receivedDate: 1 })
                .session(session);

            for (const batch of batches) {
                if (remainingToAllocate <= 0) break;

                const allocateQty = Math.min(remainingToAllocate, batch.quantityRemaining);

                // Reserve from batch
                batch.quantityRemaining -= allocateQty;
                batch.updatedAt = new Date();

                if (batch.quantityRemaining === 0) {
                    batch.status = 'depleted';
                }

                await batch.save({ session });

                // Update inventory reserved quantity
                const inventory = await Inventory.findOne({
                    chemicalId: item.chemicalId,
                    location
                }).session(session);

                if (inventory) {
                    const previousQty = inventory.quantityReserved;
                    inventory.quantityReserved += allocateQty;
                    inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
                    inventory.updatedAt = new Date();
                    await inventory.save({ session });

                    // Create reserve transaction
                    const transaction = new InventoryTransaction({
                        inventoryId: inventory._id,
                        chemicalId: item.chemicalId,
                        productName: item.productName,
                        type: 'reserve',
                        quantityChange: allocateQty,
                        previousQuantity: previousQty,
                        newQuantity: inventory.quantityReserved,
                        referenceType: 'ChemicalOrder',
                        referenceId: order._id,
                        referenceNumber: order.orderNumber,
                        location,
                        notes: `Reserved for order ${order.orderNumber} from batch ${batch.poNumber}`,
                        createdBy: req.user._id
                    });
                    await transaction.save({ session });
                }

                // Add allocation to order item
                item.batchAllocations.push({
                    batchId: batch._id,
                    poNumber: batch.poNumber,
                    lotNumber: batch.lotNumber,
                    quantityFromBatch: allocateQty,
                    allocatedAt: new Date()
                });

                allocations.push({
                    productName: item.productName,
                    batchPO: batch.poNumber,
                    lotNumber: batch.lotNumber,
                    quantity: allocateQty
                });

                remainingToAllocate -= allocateQty;
            }

            // Update item inventory status
            if (remainingToAllocate === 0) {
                item.inventoryStatus = 'allocated';
            } else if (remainingToAllocate < item.quantity) {
                item.inventoryStatus = 'allocated'; // Partial but allocated what we have
                allItemsFullyAllocated = false;
            } else {
                item.inventoryStatus = 'backordered';
                allItemsFullyAllocated = false;
            }
        }

        // Update order-level inventory status
        order.inventoryStatus = allItemsFullyAllocated ? 'fully_allocated' : 'partially_allocated';
        order.inventoryAllocatedAt = new Date();
        order.updatedAt = new Date();

        await order.save({ session });
        await session.commitTransaction();

        res.json({
            success: true,
            order,
            allocations,
            message: allItemsFullyAllocated
                ? 'All items fully allocated'
                : 'Some items partially allocated or backordered'
        });
    } catch (err) {
        await session.abortTransaction();
        console.error('POST /allocate-order error:', err);
        res.status(500).json({ error: err.message });
    } finally {
        session.endSession();
    }
});

// POST /api/admin/inventory/fulfill-order/:orderId
// Mark order as picked/delivered - decrements inventory
router.post('/fulfill-order/:orderId', async (req, res) => {
    const session = await mongoose.startSession();
    session.startTransaction();

    try {
        const { ChemicalOrder, InventoryBatch, Inventory, InventoryTransaction, Invoice } = getModels();

        const order = await ChemicalOrder.findById(req.params.orderId).session(session);
        if (!order) {
            await session.abortTransaction();
            return res.status(404).json({ error: 'Order not found' });
        }

        // Check payment verification if required
        const requirePaymentVerification = req.body.requirePaymentVerification !== false;
        if (requirePaymentVerification && !order.paymentVerified && order.paymentStatus !== 'paid') {
            await session.abortTransaction();
            return res.status(400).json({
                error: 'Payment must be verified before fulfillment',
                paymentStatus: order.paymentStatus,
                paymentVerified: order.paymentVerified
            });
        }

        const location = req.body.location || 'main';

        for (const item of order.items) {
            if (!item.batchAllocations || item.batchAllocations.length === 0) continue;

            for (const allocation of item.batchAllocations) {
                const batch = await InventoryBatch.findById(allocation.batchId).session(session);
                if (!batch) continue;

                // Add to sales allocations on batch
                batch.salesAllocations = batch.salesAllocations || [];
                batch.salesAllocations.push({
                    orderId: order._id,
                    orderNumber: order.orderNumber,
                    invoiceId: order.invoiceId,
                    invoiceNumber: order.invoiceId ? (await Invoice.findById(order.invoiceId).select('invoiceNumber').lean())?.invoiceNumber : null,
                    quantitySold: allocation.quantityFromBatch,
                    saleDate: new Date(),
                    customerId: order.userId,
                    customerName: order.contactInfo?.name,
                    paymentVerified: order.paymentVerified || order.paymentStatus === 'paid'
                });

                batch.quantitySold += allocation.quantityFromBatch;
                batch.updatedAt = new Date();
                await batch.save({ session });

                // Update inventory - release reservation and decrement on-hand
                const inventory = await Inventory.findOne({
                    chemicalId: item.chemicalId,
                    location
                }).session(session);

                if (inventory) {
                    const previousOnHand = inventory.quantityOnHand;
                    inventory.quantityOnHand -= allocation.quantityFromBatch;
                    inventory.quantityReserved -= allocation.quantityFromBatch;
                    inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
                    inventory.lastSoldDate = new Date();
                    inventory.updatedAt = new Date();
                    await inventory.save({ session });

                    // Create sale transaction
                    const transaction = new InventoryTransaction({
                        inventoryId: inventory._id,
                        chemicalId: item.chemicalId,
                        productName: item.productName,
                        type: 'sale',
                        quantityChange: -allocation.quantityFromBatch,
                        previousQuantity: previousOnHand,
                        newQuantity: inventory.quantityOnHand,
                        unitCost: batch.costPerUnit,
                        totalCost: batch.costPerUnit * allocation.quantityFromBatch,
                        referenceType: 'ChemicalOrder',
                        referenceId: order._id,
                        referenceNumber: order.orderNumber,
                        location,
                        notes: `Sold to ${order.contactInfo?.name || 'customer'} - Order ${order.orderNumber}`,
                        createdBy: req.user._id
                    });
                    await transaction.save({ session });
                }
            }

            item.inventoryStatus = 'delivered';
        }

        // Update order status
        order.inventoryStatus = 'delivered';
        order.inventoryPickedAt = new Date();
        order.pickedBy = req.user._id;
        order.status = 'delivered';
        order.deliveredAt = new Date();
        order.updatedAt = new Date();

        await order.save({ session });
        await session.commitTransaction();

        res.json({
            success: true,
            order,
            message: 'Order fulfilled and inventory updated'
        });
    } catch (err) {
        await session.abortTransaction();
        console.error('POST /fulfill-order error:', err);
        res.status(500).json({ error: err.message });
    } finally {
        session.endSession();
    }
});

// POST /api/admin/inventory/verify-payment/:orderId
// Mark payment as verified - allows inventory release
router.post('/verify-payment/:orderId', async (req, res) => {
    try {
        const { ChemicalOrder } = getModels();

        const order = await ChemicalOrder.findById(req.params.orderId);
        if (!order) {
            return res.status(404).json({ error: 'Order not found' });
        }

        order.paymentVerified = true;
        order.paymentVerifiedAt = new Date();
        order.paymentVerifiedBy = req.user._id;
        order.paymentVerificationNotes = req.body.notes || '';

        // Optionally update payment status
        if (req.body.paymentStatus) {
            order.paymentStatus = req.body.paymentStatus;
        }

        order.updatedAt = new Date();
        await order.save();

        res.json({
            success: true,
            order,
            message: 'Payment verified - order ready for fulfillment'
        });
    } catch (err) {
        console.error('POST /verify-payment error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/inventory/orders/pending-payment
// Get orders pending payment verification
router.get('/orders/pending-payment', async (req, res) => {
    try {
        const { ChemicalOrder } = getModels();

        const orders = await ChemicalOrder.find({
            paymentVerified: { $ne: true },
            paymentStatus: { $ne: 'paid' },
            status: { $nin: ['draft', 'cancelled', 'delivered'] }
        })
            .sort({ submittedAt: -1 })
            .lean();

        res.json({
            orders,
            count: orders.length
        });
    } catch (err) {
        console.error('GET /orders/pending-payment error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/inventory/orders/pending-allocation
// Get orders pending inventory allocation
router.get('/orders/pending-allocation', async (req, res) => {
    try {
        const { ChemicalOrder } = getModels();

        const orders = await ChemicalOrder.find({
            inventoryStatus: { $in: ['pending_allocation', 'partially_allocated'] },
            status: { $nin: ['draft', 'cancelled', 'delivered'] }
        })
            .sort({ submittedAt: -1 })
            .lean();

        res.json({
            orders,
            count: orders.length
        });
    } catch (err) {
        console.error('GET /orders/pending-allocation error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/inventory/lifecycle/:chemicalId
// Full lifecycle view for a specific product
router.get('/lifecycle/:chemicalId', async (req, res) => {
    try {
        const { Chemical, InventoryBatch, Inventory, InventoryTransaction, SupplierInvoice } = getModels();
        const SupplierInvoiceModel = mongoose.models.SupplierInvoice;

        const chemicalId = req.params.chemicalId;
        const location = req.query.location || 'main';

        // Get product info
        const chemical = await Chemical.findById(chemicalId).lean();
        if (!chemical) {
            return res.status(404).json({ error: 'Product not found' });
        }

        // Get current inventory
        const inventory = await Inventory.findOne({ chemicalId, location }).lean();

        // Get all batches for this product
        const batches = await InventoryBatch.find({ chemicalId, location })
            .sort({ receivedDate: -1 })
            .lean();

        // Get recent transactions
        const transactions = await InventoryTransaction.find({ chemicalId })
            .sort({ createdAt: -1 })
            .limit(50)
            .lean();

        // Calculate summary stats
        const activeBatches = batches.filter(b => b.status === 'active');
        const totalReceived = batches.reduce((sum, b) => sum + b.quantityReceived, 0);
        const totalSold = batches.reduce((sum, b) => sum + (b.quantitySold || 0), 0);
        const totalRemaining = batches.reduce((sum, b) => sum + (b.quantityRemaining || 0), 0);

        // Get expiring soon (within 90 days)
        const ninetyDaysFromNow = new Date();
        ninetyDaysFromNow.setDate(ninetyDaysFromNow.getDate() + 90);
        const expiringSoon = batches.filter(b =>
            b.status === 'active' &&
            b.expirationDate &&
            new Date(b.expirationDate) <= ninetyDaysFromNow
        );

        res.json({
            product: {
                _id: chemical._id,
                productName: chemical.productName,
                tradeName: chemical.tradeName,
                packSize: chemical.packSize,
                unit: chemical.unit,
                category: chemical.category
            },
            inventory: inventory || { quantityOnHand: 0, quantityReserved: 0, quantityAvailable: 0 },
            summary: {
                totalReceived,
                totalSold,
                totalRemaining,
                activeBatchCount: activeBatches.length,
                averageCost: inventory?.averageCost || 0,
                lastCost: inventory?.lastCost || 0
            },
            batches,
            expiringSoon,
            recentTransactions: transactions
        });
    } catch (err) {
        console.error('GET /lifecycle/:chemicalId error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/inventory/expiring
// Get products expiring within X days
router.get('/expiring', async (req, res) => {
    try {
        const { InventoryBatch } = getModels();

        const days = parseInt(req.query.days) || 90;
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() + days);

        const expiringBatches = await InventoryBatch.find({
            status: 'active',
            quantityRemaining: { $gt: 0 },
            expirationDate: { $lte: cutoffDate }
        })
            .sort({ expirationDate: 1 })
            .lean();

        // Group by product
        const byProduct = {};
        for (const batch of expiringBatches) {
            const key = batch.chemicalId.toString();
            if (!byProduct[key]) {
                byProduct[key] = {
                    chemicalId: batch.chemicalId,
                    productName: batch.productName,
                    batches: [],
                    totalExpiring: 0
                };
            }
            byProduct[key].batches.push(batch);
            byProduct[key].totalExpiring += batch.quantityRemaining;
        }

        res.json({
            withinDays: days,
            products: Object.values(byProduct),
            totalBatches: expiringBatches.length
        });
    } catch (err) {
        console.error('GET /expiring error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/inventory/low-stock
// Get products below reorder point
router.get('/low-stock', async (req, res) => {
    try {
        const { Inventory } = getModels();

        const lowStock = await Inventory.find({
            $expr: { $lte: ['$quantityAvailable', '$reorderPoint'] },
            reorderPoint: { $gt: 0 }
        })
            .sort({ quantityAvailable: 1 })
            .lean();

        res.json({
            products: lowStock,
            count: lowStock.length
        });
    } catch (err) {
        console.error('GET /low-stock error:', err);
        res.status(500).json({ error: err.message });
    }
});

// POST /api/admin/inventory/release-allocation/:orderId
// Release allocated inventory (cancel allocation)
router.post('/release-allocation/:orderId', async (req, res) => {
    const session = await mongoose.startSession();
    session.startTransaction();

    try {
        const { ChemicalOrder, InventoryBatch, Inventory, InventoryTransaction } = getModels();

        const order = await ChemicalOrder.findById(req.params.orderId).session(session);
        if (!order) {
            await session.abortTransaction();
            return res.status(404).json({ error: 'Order not found' });
        }

        if (order.inventoryStatus === 'delivered') {
            await session.abortTransaction();
            return res.status(400).json({ error: 'Cannot release allocation - order already delivered' });
        }

        const location = req.body.location || 'main';

        for (const item of order.items) {
            if (!item.batchAllocations || item.batchAllocations.length === 0) continue;

            for (const allocation of item.batchAllocations) {
                const batch = await InventoryBatch.findById(allocation.batchId).session(session);
                if (!batch) continue;

                // Return quantity to batch
                batch.quantityRemaining += allocation.quantityFromBatch;
                if (batch.status === 'depleted') {
                    batch.status = 'active';
                }
                batch.updatedAt = new Date();
                await batch.save({ session });

                // Update inventory
                const inventory = await Inventory.findOne({
                    chemicalId: item.chemicalId,
                    location
                }).session(session);

                if (inventory) {
                    const previousReserved = inventory.quantityReserved;
                    inventory.quantityReserved -= allocation.quantityFromBatch;
                    inventory.quantityAvailable = inventory.quantityOnHand - inventory.quantityReserved;
                    inventory.updatedAt = new Date();
                    await inventory.save({ session });

                    // Create release transaction
                    const transaction = new InventoryTransaction({
                        inventoryId: inventory._id,
                        chemicalId: item.chemicalId,
                        productName: item.productName,
                        type: 'release',
                        quantityChange: -allocation.quantityFromBatch,
                        previousQuantity: previousReserved,
                        newQuantity: inventory.quantityReserved,
                        referenceType: 'ChemicalOrder',
                        referenceId: order._id,
                        referenceNumber: order.orderNumber,
                        location,
                        notes: `Released allocation for order ${order.orderNumber}`,
                        createdBy: req.user._id
                    });
                    await transaction.save({ session });
                }
            }

            // Clear allocations
            item.batchAllocations = [];
            item.inventoryStatus = 'pending_allocation';
        }

        order.inventoryStatus = 'pending_allocation';
        order.inventoryAllocatedAt = null;
        order.updatedAt = new Date();

        await order.save({ session });
        await session.commitTransaction();

        res.json({
            success: true,
            order,
            message: 'Inventory allocation released'
        });
    } catch (err) {
        await session.abortTransaction();
        console.error('POST /release-allocation error:', err);
        res.status(500).json({ error: err.message });
    } finally {
        session.endSession();
    }
});

module.exports = router;

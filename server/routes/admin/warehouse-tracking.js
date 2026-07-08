// server/routes/admin/warehouse-tracking.js
// Warehouse tracking & customer accounting - inventory lifecycle visibility
// Access: Distributors and SuperAdmins only (middleware applied at mount point)

const express = require('express');
const router = express.Router();
const mongoose = require('mongoose');

const getModels = () => ({
    Inventory: mongoose.model('Inventory'),
    InventoryBatch: mongoose.model('InventoryBatch'),
    InventoryTransaction: mongoose.model('InventoryTransaction'),
    PurchaseOrder: mongoose.model('PurchaseOrder'),
    Invoice: mongoose.model('Invoice'),
    Chemical: mongoose.model('Chemical'),
    User: mongoose.model('User')
});

// GET /api/admin/warehouse/warehouse-summary
// Returns all products grouped by warehouse location with totals
router.get('/warehouse-summary', async (req, res) => {
    try {
        const { Inventory } = getModels();

        const inventoryRecords = await Inventory.find().lean();

        // Group by location
        const locationMap = {};
        for (const record of inventoryRecords) {
            const loc = record.location || 'main';
            if (!locationMap[loc]) {
                locationMap[loc] = {
                    location: loc,
                    products: [],
                    totals: {
                        totalQuantityOnHand: 0,
                        totalQuantityReserved: 0,
                        totalQuantityAvailable: 0,
                        totalValue: 0
                    }
                };
            }

            const value = (record.quantityOnHand || 0) * (record.averageCost || 0);

            locationMap[loc].products.push({
                chemicalId: record.chemicalId,
                productName: record.productName,
                packSize: record.packSize,
                unit: record.unit,
                quantityOnHand: record.quantityOnHand || 0,
                quantityReserved: record.quantityReserved || 0,
                quantityAvailable: record.quantityAvailable || 0,
                averageCost: record.averageCost || 0,
                value
            });

            locationMap[loc].totals.totalQuantityOnHand += record.quantityOnHand || 0;
            locationMap[loc].totals.totalQuantityReserved += record.quantityReserved || 0;
            locationMap[loc].totals.totalQuantityAvailable += record.quantityAvailable || 0;
            locationMap[loc].totals.totalValue += value;
        }

        const locations = Object.values(locationMap);

        res.json({ locations });
    } catch (err) {
        console.error('GET /warehouse-summary error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/warehouse/purchases-by-warehouse
// Returns PurchaseOrders enriched with warehouse location from InventoryBatch records
router.get('/purchases-by-warehouse', async (req, res) => {
    try {
        const { PurchaseOrder, InventoryBatch } = getModels();

        const [purchaseOrders, batches] = await Promise.all([
            PurchaseOrder.find().sort({ orderDate: -1 }).lean(),
            InventoryBatch.find({}, 'poNumber chemicalId location').lean()
        ]);

        // Build a lookup: poNumber -> { chemicalId -> location }
        const batchLocationMap = {};
        for (const batch of batches) {
            if (!batchLocationMap[batch.poNumber]) {
                batchLocationMap[batch.poNumber] = {};
            }
            const chemId = batch.chemicalId ? batch.chemicalId.toString() : 'unknown';
            batchLocationMap[batch.poNumber][chemId] = batch.location || 'main';
        }

        // Enrich PO items with location and group by location
        const locationMap = {};
        for (const po of purchaseOrders) {
            const poLocations = batchLocationMap[po.poNumber] || {};

            // Determine the primary location for this PO
            // Use deliveryLocation fallback, then batch locations, then 'main'
            const enrichedItems = (po.items || []).map(item => {
                const chemId = item.chemicalId ? item.chemicalId.toString() : 'unknown';
                const location = poLocations[chemId] || po.deliveryLocation || 'main';
                return { ...item, location };
            });

            // Group by location - a PO may span multiple locations
            const poByLocation = {};
            for (const item of enrichedItems) {
                const loc = item.location;
                if (!poByLocation[loc]) {
                    poByLocation[loc] = [];
                }
                poByLocation[loc].push(item);
            }

            for (const [loc, items] of Object.entries(poByLocation)) {
                if (!locationMap[loc]) {
                    locationMap[loc] = { location: loc, purchaseOrders: [] };
                }
                locationMap[loc].purchaseOrders.push({
                    _id: po._id,
                    poNumber: po.poNumber,
                    supplier: po.supplier,
                    status: po.status,
                    orderDate: po.orderDate,
                    receivedDate: po.receivedDate,
                    paymentStatus: po.paymentStatus,
                    totalCost: po.totalCost,
                    items
                });
            }
        }

        const locations = Object.values(locationMap);

        res.json({ locations });
    } catch (err) {
        console.error('GET /purchases-by-warehouse error:', err);
        res.status(500).json({ error: err.message });
    }
});

// POST /api/admin/warehouse/transfer
// Transfer product between warehouses with full audit trail
router.post('/transfer', async (req, res) => {
    const session = await mongoose.startSession();
    session.startTransaction();

    try {
        const { Inventory, InventoryBatch, InventoryTransaction } = getModels();

        const { chemicalId, productName, quantity, fromLocation, toLocation, notes } = req.body;

        // Validate required fields
        if (!chemicalId || !productName || !quantity || !fromLocation || !toLocation) {
            await session.abortTransaction();
            session.endSession();
            return res.status(400).json({
                error: 'Missing required fields: chemicalId, productName, quantity, fromLocation, toLocation'
            });
        }

        if (quantity <= 0) {
            await session.abortTransaction();
            session.endSession();
            return res.status(400).json({ error: 'Quantity must be greater than zero' });
        }

        if (fromLocation === toLocation) {
            await session.abortTransaction();
            session.endSession();
            return res.status(400).json({ error: 'From and to locations must be different' });
        }

        // Verify sufficient quantity at source
        const sourceInventory = await Inventory.findOne({
            chemicalId,
            location: fromLocation
        }).session(session);

        if (!sourceInventory) {
            await session.abortTransaction();
            session.endSession();
            return res.status(404).json({ error: `No inventory found for this product at ${fromLocation}` });
        }

        if (sourceInventory.quantityAvailable < quantity) {
            await session.abortTransaction();
            session.endSession();
            return res.status(400).json({
                error: `Insufficient available quantity at ${fromLocation}. Available: ${sourceInventory.quantityAvailable}, Requested: ${quantity}`
            });
        }

        // Decrement source inventory
        const previousSourceQty = sourceInventory.quantityOnHand;
        sourceInventory.quantityOnHand -= quantity;
        sourceInventory.quantityAvailable -= quantity;
        sourceInventory.updatedAt = new Date();
        await sourceInventory.save({ session });

        // Increment destination inventory (create if doesn't exist)
        let destInventory = await Inventory.findOne({
            chemicalId,
            location: toLocation
        }).session(session);

        let previousDestQty = 0;
        if (destInventory) {
            previousDestQty = destInventory.quantityOnHand;
            destInventory.quantityOnHand += quantity;
            destInventory.quantityAvailable += quantity;
            // Recalculate average cost with the transferred units
            if (destInventory.quantityOnHand > 0) {
                const existingValue = previousDestQty * (destInventory.averageCost || 0);
                const transferValue = quantity * (sourceInventory.averageCost || 0);
                destInventory.averageCost = (existingValue + transferValue) / destInventory.quantityOnHand;
            }
            destInventory.updatedAt = new Date();
            await destInventory.save({ session });
        } else {
            destInventory = new Inventory({
                chemicalId,
                productName,
                packSize: sourceInventory.packSize,
                unit: sourceInventory.unit,
                quantityOnHand: quantity,
                quantityReserved: 0,
                quantityAvailable: quantity,
                location: toLocation,
                averageCost: sourceInventory.averageCost || 0,
                lastCost: sourceInventory.lastCost || 0
            });
            await destInventory.save({ session });
        }

        // Move quantity between InventoryBatch records (FIFO - oldest first)
        const sourceBatches = await InventoryBatch.find({
            chemicalId,
            location: fromLocation,
            status: 'active',
            quantityRemaining: { $gt: 0 }
        })
            .sort({ receivedDate: 1 })
            .session(session);

        let remainingToTransfer = quantity;
        const batchTransfers = [];

        for (const batch of sourceBatches) {
            if (remainingToTransfer <= 0) break;

            const transferFromBatch = Math.min(remainingToTransfer, batch.quantityRemaining);

            // Reduce source batch
            batch.quantityRemaining -= transferFromBatch;
            batch.updatedAt = new Date();
            if (batch.quantityRemaining === 0) {
                batch.status = 'depleted';
            }
            await batch.save({ session });

            // Create or update destination batch
            let destBatch = await InventoryBatch.findOne({
                chemicalId,
                poNumber: batch.poNumber,
                location: toLocation,
                status: 'active'
            }).session(session);

            if (destBatch) {
                destBatch.quantityRemaining += transferFromBatch;
                destBatch.quantityReceived += transferFromBatch;
                destBatch.updatedAt = new Date();
                await destBatch.save({ session });
            } else {
                destBatch = new InventoryBatch({
                    chemicalId,
                    productName: batch.productName,
                    packSize: batch.packSize,
                    unit: batch.unit,
                    poNumber: batch.poNumber,
                    poId: batch.poId,
                    lotNumber: batch.lotNumber,
                    quantityReceived: transferFromBatch,
                    quantityRemaining: transferFromBatch,
                    quantitySold: 0,
                    costPerUnit: batch.costPerUnit,
                    totalCost: transferFromBatch * batch.costPerUnit,
                    location: toLocation,
                    status: 'active',
                    receivedDate: batch.receivedDate,
                    expirationDate: batch.expirationDate,
                    supplierId: batch.supplierId,
                    supplierName: batch.supplierName
                });
                await destBatch.save({ session });
            }

            batchTransfers.push({
                poNumber: batch.poNumber,
                quantity: transferFromBatch,
                costPerUnit: batch.costPerUnit
            });

            remainingToTransfer -= transferFromBatch;
        }

        if (remainingToTransfer > 0) {
            await session.abortTransaction();
            session.endSession();
            return res.status(400).json({
                error: `Insufficient batch quantity at ${fromLocation}. Could only source ${quantity - remainingToTransfer} of ${quantity} from active batches.`
            });
        }

        // Create InventoryTransaction records for both sides
        const outboundTransaction = new InventoryTransaction({
            inventoryId: sourceInventory._id,
            chemicalId,
            productName,
            type: 'transfer',
            quantityChange: -quantity,
            previousQuantity: previousSourceQty,
            newQuantity: sourceInventory.quantityOnHand,
            location: fromLocation,
            fromLocation,
            toLocation,
            referenceType: 'Transfer',
            notes: notes || `Transfer to ${toLocation}`,
            createdBy: req.user._id
        });
        await outboundTransaction.save({ session });

        const inboundTransaction = new InventoryTransaction({
            inventoryId: destInventory._id,
            chemicalId,
            productName,
            type: 'transfer',
            quantityChange: quantity,
            previousQuantity: previousDestQty,
            newQuantity: destInventory.quantityOnHand,
            location: toLocation,
            fromLocation,
            toLocation,
            referenceType: 'Transfer',
            notes: notes || `Transfer from ${fromLocation}`,
            createdBy: req.user._id
        });
        await inboundTransaction.save({ session });

        await session.commitTransaction();
        session.endSession();

        res.json({
            message: 'Transfer completed successfully',
            transfer: {
                chemicalId,
                productName,
                quantity,
                fromLocation,
                toLocation,
                batchTransfers,
                outboundTransactionId: outboundTransaction._id,
                inboundTransactionId: inboundTransaction._id
            }
        });
    } catch (err) {
        await session.abortTransaction();
        session.endSession();
        console.error('POST /transfer error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/warehouse/transfers
// Returns transfer history with optional filters
router.get('/transfers', async (req, res) => {
    try {
        const { InventoryTransaction } = getModels();

        const { startDate, endDate, location, chemicalId } = req.query;

        const query = { type: 'transfer' };

        if (startDate || endDate) {
            query.createdAt = {};
            if (startDate) query.createdAt.$gte = new Date(startDate);
            if (endDate) query.createdAt.$lte = new Date(endDate);
        }

        if (location) {
            // Match transfers involving this location (either side)
            query.$or = [
                { fromLocation: location },
                { toLocation: location }
            ];
        }

        if (chemicalId) {
            query.chemicalId = chemicalId;
        }

        const transfers = await InventoryTransaction.find(query)
            .sort({ createdAt: -1 })
            .populate('createdBy', 'name email')
            .lean();

        res.json({ transfers, total: transfers.length });
    } catch (err) {
        console.error('GET /transfers error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/warehouse/customer-accounting
// Returns customer-level product accounting with invoice details
router.get('/customer-accounting', async (req, res) => {
    try {
        const { Invoice, User } = getModels();

        const invoices = await Invoice.find({
            status: { $ne: 'cancelled' }
        })
            .sort({ invoiceDate: -1 })
            .lean();

        // Group by customer
        const customerMap = {};

        for (const invoice of invoices) {
            const custId = invoice.customerId ? invoice.customerId.toString() : 'unknown';

            if (!customerMap[custId]) {
                customerMap[custId] = {
                    customerId: invoice.customerId,
                    customerName: invoice.customerName || 'Unknown',
                    customerEmail: invoice.customerEmail || '',
                    products: {},
                    invoices: [],
                    totalInvoiced: 0,
                    totalPaid: 0,
                    balanceDue: 0
                };
            }

            const customer = customerMap[custId];

            // Aggregate product-level data
            for (const item of (invoice.items || [])) {
                const productKey = item.productName || 'Unknown Product';
                if (!customer.products[productKey]) {
                    customer.products[productKey] = {
                        productName: productKey,
                        chemicalId: item.chemicalId,
                        totalQuantity: 0,
                        totalPrice: 0,
                        entries: []
                    };
                }
                customer.products[productKey].totalQuantity += item.quantity || 0;
                customer.products[productKey].totalPrice += item.totalPrice || 0;
                customer.products[productKey].entries.push({
                    invoiceNumber: invoice.invoiceNumber,
                    invoiceId: invoice._id,
                    quantity: item.quantity || 0,
                    unitPrice: item.unitPrice || 0,
                    totalPrice: item.totalPrice || 0,
                    unit: item.unit,
                    packSize: item.packSize
                });
            }

            // Add invoice summary
            customer.invoices.push({
                _id: invoice._id,
                invoiceNumber: invoice.invoiceNumber,
                invoiceDate: invoice.invoiceDate,
                total: invoice.total || 0,
                amountPaid: invoice.amountPaid || 0,
                amountDue: invoice.amountDue || 0,
                paymentStatus: invoice.paymentStatus,
                status: invoice.status
            });

            customer.totalInvoiced += invoice.total || 0;
            customer.totalPaid += invoice.amountPaid || 0;
        }

        // Finalize customer records
        const customers = Object.values(customerMap).map(customer => ({
            ...customer,
            products: Object.values(customer.products),
            balanceDue: customer.totalInvoiced - customer.totalPaid,
            paymentStatus: customer.totalPaid >= customer.totalInvoiced
                ? 'paid'
                : customer.totalPaid > 0 ? 'partial' : 'unpaid'
        }));

        // Sort by balance due descending
        customers.sort((a, b) => b.balanceDue - a.balanceDue);

        res.json({ customers, total: customers.length });
    } catch (err) {
        console.error('GET /customer-accounting error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/warehouse/customer-accounting/:customerId
// Detailed customer accounting for a single customer
router.get('/customer-accounting/:customerId', async (req, res) => {
    try {
        const { Invoice, User } = getModels();

        const customerId = req.params.customerId;

        const [customer, invoices] = await Promise.all([
            User.findById(customerId, 'name email phone').lean(),
            Invoice.find({
                customerId,
                status: { $ne: 'cancelled' }
            })
                .sort({ invoiceDate: -1 })
                .lean()
        ]);

        if (!customer && invoices.length === 0) {
            return res.status(404).json({ error: 'Customer not found or has no invoices' });
        }

        // Build product-level aggregation
        const productMap = {};
        let totalInvoiced = 0;
        let totalPaid = 0;

        const invoiceDetails = invoices.map(invoice => {
            for (const item of (invoice.items || [])) {
                const productKey = item.productName || 'Unknown Product';
                if (!productMap[productKey]) {
                    productMap[productKey] = {
                        productName: productKey,
                        chemicalId: item.chemicalId,
                        totalQuantity: 0,
                        totalPrice: 0,
                        lineItems: []
                    };
                }
                productMap[productKey].totalQuantity += item.quantity || 0;
                productMap[productKey].totalPrice += item.totalPrice || 0;
                productMap[productKey].lineItems.push({
                    invoiceNumber: invoice.invoiceNumber,
                    invoiceId: invoice._id,
                    invoiceDate: invoice.invoiceDate,
                    quantity: item.quantity || 0,
                    unitPrice: item.unitPrice || 0,
                    totalPrice: item.totalPrice || 0,
                    unit: item.unit,
                    packSize: item.packSize,
                    description: item.description
                });
            }

            totalInvoiced += invoice.total || 0;
            totalPaid += invoice.amountPaid || 0;

            return {
                _id: invoice._id,
                invoiceNumber: invoice.invoiceNumber,
                orderNumber: invoice.orderNumber,
                invoiceDate: invoice.invoiceDate,
                dueDate: invoice.dueDate,
                items: invoice.items,
                subtotal: invoice.subtotal,
                discount: invoice.discount,
                tax: invoice.tax,
                shipping: invoice.shipping,
                total: invoice.total || 0,
                amountPaid: invoice.amountPaid || 0,
                amountDue: invoice.amountDue || 0,
                paymentStatus: invoice.paymentStatus,
                paymentMethod: invoice.paymentMethod,
                paymentDate: invoice.paymentDate,
                status: invoice.status
            };
        });

        const balanceDue = totalInvoiced - totalPaid;

        res.json({
            customer: {
                _id: customer ? customer._id : customerId,
                name: customer ? customer.name : (invoices[0] ? invoices[0].customerName : 'Unknown'),
                email: customer ? customer.email : (invoices[0] ? invoices[0].customerEmail : '')
            },
            products: Object.values(productMap),
            invoices: invoiceDetails,
            summary: {
                totalInvoiced,
                totalPaid,
                balanceDue,
                paymentStatus: totalPaid >= totalInvoiced
                    ? 'paid'
                    : totalPaid > 0 ? 'partial' : 'unpaid',
                invoiceCount: invoices.length
            }
        });
    } catch (err) {
        console.error('GET /customer-accounting/:customerId error:', err);
        res.status(500).json({ error: err.message });
    }
});

// GET /api/admin/warehouse/inventory-value
// Returns total inventory value by location and by product
router.get('/inventory-value', async (req, res) => {
    try {
        const { Inventory } = getModels();

        const inventoryRecords = await Inventory.find().lean();

        // By location
        const byLocation = {};
        // By product
        const byProduct = {};
        // Summary
        let totalValue = 0;
        let totalUnits = 0;
        const productSet = new Set();

        for (const record of inventoryRecords) {
            const loc = record.location || 'main';
            const value = (record.quantityOnHand || 0) * (record.averageCost || 0);
            const units = record.quantityOnHand || 0;

            // By location
            if (!byLocation[loc]) {
                byLocation[loc] = {
                    location: loc,
                    totalValue: 0,
                    totalUnits: 0,
                    productCount: 0,
                    products: []
                };
            }
            byLocation[loc].totalValue += value;
            byLocation[loc].totalUnits += units;
            byLocation[loc].productCount += 1;
            byLocation[loc].products.push({
                chemicalId: record.chemicalId,
                productName: record.productName,
                packSize: record.packSize,
                unit: record.unit,
                quantityOnHand: units,
                averageCost: record.averageCost || 0,
                value
            });

            // By product (aggregate across locations)
            const productKey = record.chemicalId ? record.chemicalId.toString() : record.productName;
            if (!byProduct[productKey]) {
                byProduct[productKey] = {
                    chemicalId: record.chemicalId,
                    productName: record.productName,
                    totalQuantity: 0,
                    totalValue: 0,
                    locations: []
                };
            }
            byProduct[productKey].totalQuantity += units;
            byProduct[productKey].totalValue += value;
            byProduct[productKey].locations.push({
                location: loc,
                quantityOnHand: units,
                averageCost: record.averageCost || 0,
                value
            });

            totalValue += value;
            totalUnits += units;
            productSet.add(productKey);
        }

        res.json({
            summary: {
                totalValue,
                totalUnits,
                productCount: productSet.size,
                locationCount: Object.keys(byLocation).length
            },
            byLocation: Object.values(byLocation),
            byProduct: Object.values(byProduct)
        });
    } catch (err) {
        console.error('GET /inventory-value error:', err);
        res.status(500).json({ error: err.message });
    }
});

module.exports = router;

#!/usr/bin/env node
// Read-only audit: buckets Invoice documents by payment state, cross-tabbed
// against status + paymentStatus so we can see where stored fields disagree
// with stored amountPaid. Also spotlights the three named invoices from the
// manual-payment-recording pre-work (INV-*-00006, 00007, 00008) so the live
// DB state of the "$-0.00 Voided" and "Paid Paid" anomalies is visible. No
// writes.
//
// Usage:
//   node -r dotenv/config server/scripts/amountPaid-buckets.js
//
// Exit codes:
//   0 - ran to completion (output on stdout)
//   non-zero - connection or query error

'use strict';

const mongoose = require('mongoose');

(async () => {
    await mongoose.connect(process.env.MONGODB_URI);
    const col = mongoose.connection.db.collection('invoices');

    const pipeline = [
        {
            $project: {
                invoiceNumber: 1, total: 1, amountPaid: 1, amountDue: 1,
                status: 1, paymentStatus: 1, createdAt: 1,
                bucket: {
                    $switch: {
                        branches: [
                            { case: { $or: [{ $eq: ['$amountPaid', null] }, { $not: ['$amountPaid'] }] }, then: 'zero_null' },
                            { case: { $eq: ['$amountPaid', 0] }, then: 'zero' },
                            { case: { $lt: ['$amountPaid', 0] }, then: 'negative' },
                            { case: { $lt: ['$amountPaid', '$total'] }, then: 'partial' },
                            { case: { $eq: ['$amountPaid', '$total'] }, then: 'fully_paid' },
                            { case: { $gt: ['$amountPaid', '$total'] }, then: 'overpaid' },
                        ],
                        default: 'unknown',
                    },
                },
            },
        },
        { $group: { _id: { bucket: '$bucket', status: '$status', paymentStatus: '$paymentStatus' }, n: { $sum: 1 } } },
        { $sort: { '_id.bucket': 1, '_id.status': 1 } },
    ];
    const rows = await col.aggregate(pipeline).toArray();
    console.table(rows.map(r => ({ bucket: r._id.bucket, status: r._id.status, paymentStatus: r._id.paymentStatus, n: r.n })));

    const negatives = await col.find({ amountPaid: { $lt: 0 } }, { projection: { invoiceNumber: 1, total: 1, amountPaid: 1, creditedAmount: 1, status: 1, paymentStatus: 1 } }).toArray();
    const overpaid = await col.find({ $expr: { $gt: ['$amountPaid', '$total'] } }, { projection: { invoiceNumber: 1, total: 1, amountPaid: 1, status: 1, paymentStatus: 1 } }).toArray();
    const inv8 = await col.findOne({ invoiceNumber: /00008/ });
    const inv6 = await col.findOne({ invoiceNumber: /00006/ });
    const inv7 = await col.findOne({ invoiceNumber: /00007/ });
    console.log('NEGATIVES:', negatives);
    console.log('OVERPAID:', overpaid);
    console.log('INV-00006:', inv6);
    console.log('INV-00007:', inv7);
    console.log('INV-00008:', inv8);

    await mongoose.disconnect();
})();

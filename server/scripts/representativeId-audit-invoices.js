#!/usr/bin/env node
// Read-only audit: counts Invoice documents by the runtime BSON type of
// representativeId (objectId vs string vs nullish vs other) and samples up
// to 10 string-typed rows so a backfill plan can be sized. No writes.
//
// Usage:
//   node -r dotenv/config server/scripts/representativeId-audit-invoices.js
//
// Exit codes:
//   0 - ran to completion (output on stdout as JSON)
//   non-zero - connection or query error

'use strict';

const mongoose = require('mongoose');

(async () => {
    await mongoose.connect(process.env.MONGODB_URI);
    const col = mongoose.connection.db.collection('invoices');
    const total = await col.countDocuments({});
    const objId = await col.countDocuments({ representativeId: { $type: 'objectId' } });
    const str = await col.countDocuments({ representativeId: { $type: 'string' } });
    const nullish = await col.countDocuments({ $or: [{ representativeId: null }, { representativeId: { $exists: false } }] });
    const other = total - objId - str - nullish;
    const sampleStrings = await col.find(
        { representativeId: { $type: 'string' } },
        { projection: { invoiceNumber: 1, representativeId: 1, createdAt: 1 } }
    ).limit(10).toArray();
    console.log(JSON.stringify({ total, objId, str, nullish, other, sampleStrings }, null, 2));
    await mongoose.disconnect();
})();

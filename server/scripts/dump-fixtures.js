#!/usr/bin/env node
// Read-only fixture dump for C3b integration tests. Produces anonymized
// per-collection JSON files in server/test/fixtures/ that the Decimal128
// migration test (server/test/migration.test.js, ships in C3b) seeds into
// mongodb-memory-server.
//
// Anonymization policy:
//   - PII (customer names, emails, phones, addresses) replaced with
//     deterministic placeholders ("Customer A", "customer-a@example.test",
//     555-0100, "100 Main St", "Anytown", state preserved, zip "00000").
//     Mapping is by User._id sort order so the same prod customer maps to
//     the same fixture label across re-runs — test failures stay reproducible.
//   - Money fields, dates, ObjectIds, and product details (chemical names,
//     pack sizes, units) are NOT anonymized. They're what the migration
//     tests against. Anonymizing them would defeat the purpose.
//   - User collection itself is NOT dumped. Tests don't populate refs;
//     dangling userId pointers in fixture docs are fine. Skips also avoids
//     leaking password hashes or license URLs.
//
// Usage (run in Render shell against prod):
//   node -r dotenv/config server/scripts/dump-fixtures.js
//
// Output: writes one JSON file per collection in server/test/fixtures/.
// Kyle commits those files manually after review. C3b proper (the schema
// flip + migration script + integration tests) reads those committed
// fixtures.
//
// Sample size: SAMPLE_LIMIT docs per collection, sorted by _id descending
// (newest first). Set high enough (50) that small collections like
// invoices (12 docs), chemicalorders (8 docs), orders (0 docs) are
// captured in full, including the float-trap reference docs INV-2026-00006,
// INV-2026-00007, INV-2026-00008 that anchor the regression suite.

'use strict';

const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');

const SAMPLE_LIMIT = 50;

// 23 collections with money fields. Dependency order matches the C3b
// migration script. Mongoose default lowercased-pluralized collection names.
const COLLECTIONS = [
    'invoices',
    'creditnotes',
    'ledgerentries',
    'chemicalorders',
    'orders',
    'programquotes',
    'quoterequests',
    'purchaseorders',
    'purchaseordersplits',
    'supplierbidsheets',
    'repcommissions',
    'chemicalpricehistories',
    'distributorpricings',
    'chemicalquotes',
    'inventories',
    'inventorytransactions',
    'inventorybatches',
    'chemicals',
    'auditlogs',
    'rupsalerecords',
    'chemicalrequests',
    'usersubmittedchemicals',
    'sprayprograms',
];

// Collection name → document field path map for embedded customer PII.
// Walked by anonymize() at the top of every fetched doc. Empty paths are
// no-ops so we can list collections without customer fields here too.
const CUSTOMER_FIELDS = {
    invoices: { id: 'customerId', name: 'customerName', email: 'customerEmail', phone: 'customerPhone', address: 'customerAddress' },
    chemicalorders: { id: 'userId' },
    orders: { id: 'userId' },
    rupsalerecords: { id: 'purchaserId', name: 'purchaserName', email: 'purchaserEmail', phone: 'purchaserPhone', address: 'purchaserAddress' },
    quoterequests: { id: 'customerId', name: 'customerName', email: 'customerEmail', phone: 'customerPhone' },
    programquotes: { id: 'customerId', name: 'customerName', email: 'customerEmail' },
    creditnotes: { id: 'customerId', name: 'customerName', email: 'customerEmail' },
};

function anonLabelFor(idx) {
    // 0 → A, 1 → B, ..., 25 → Z, 26 → A1, 27 → B1, ...
    const letter = String.fromCharCode(65 + (idx % 26));
    const suffix = Math.floor(idx / 26);
    return `Customer ${letter}${suffix > 0 ? suffix : ''}`;
}

function anonRecordFor(label, idx) {
    const slug = label.toLowerCase().replace(/\s+/g, '-');
    return {
        name: label,
        email: `${slug}@example.test`,
        phone: `555-${String(100 + idx).padStart(4, '0')}`,
        address: { street: '100 Main St', city: 'Anytown', state: '', zip: '00000' },
    };
}

function applyAnon(doc, fields, anon) {
    if (!fields || !anon) return;
    if (fields.name && doc[fields.name] !== undefined) doc[fields.name] = anon.name;
    if (fields.email && doc[fields.email] !== undefined) doc[fields.email] = anon.email;
    if (fields.phone && doc[fields.phone] !== undefined) doc[fields.phone] = anon.phone;
    if (fields.address && doc[fields.address] !== undefined) {
        // Preserve state if present, replace everything else
        const origState = doc[fields.address] && doc[fields.address].state;
        doc[fields.address] = { ...anon.address, state: origState || '' };
    }
}

(async () => {
    if (!process.env.MONGODB_URI) {
        console.error('MONGODB_URI not set. Source from .env or pass via env.');
        process.exit(2);
    }
    await mongoose.connect(process.env.MONGODB_URI);
    const db = mongoose.connection.db;

    const fixturesDir = path.join(__dirname, '..', 'test', 'fixtures');
    if (!fs.existsSync(fixturesDir)) fs.mkdirSync(fixturesDir, { recursive: true });

    // Build deterministic User → anon-label map. Sort by _id ascending so the
    // earliest-created customer is "Customer A" across re-runs.
    const customers = await db.collection('users')
        .find({ role: 'customer' })
        .project({ _id: 1, address: 1 })
        .sort({ _id: 1 })
        .toArray();

    const userMap = {};
    customers.forEach((u, i) => {
        const label = anonLabelFor(i);
        const anon = anonRecordFor(label, i);
        // Preserve real state since it's non-PII regional info the migration
        // doesn't touch but tests might want to assert against.
        if (u.address && u.address.state) anon.address.state = u.address.state;
        userMap[u._id.toString()] = anon;
    });

    // Fallback for orphaned refs (deleted users, system actors, etc.)
    const ORPHAN_ANON = {
        name: 'Customer (orphaned ref)',
        email: 'orphan@example.test',
        phone: '555-0199',
        address: { street: '100 Main St', city: 'Anytown', state: '', zip: '00000' },
    };

    const summary = [];
    for (const col of COLLECTIONS) {
        const fields = CUSTOMER_FIELDS[col];
        const docs = await db.collection(col)
            .find({})
            .sort({ _id: -1 })
            .limit(SAMPLE_LIMIT)
            .toArray();

        for (const doc of docs) {
            if (fields && fields.id) {
                const id = doc[fields.id];
                const anon = (id && userMap[id.toString()]) || ORPHAN_ANON;
                applyAnon(doc, fields, anon);
            }
        }

        const outPath = path.join(fixturesDir, `${col}.json`);
        fs.writeFileSync(outPath, JSON.stringify(docs, null, 2));
        summary.push({ collection: col, dumped: docs.length });
        console.log(`${col}: ${docs.length} docs → server/test/fixtures/${col}.json`);
    }

    console.log('---');
    console.log(JSON.stringify(summary, null, 2));
    console.log('done. Review the fixtures, commit them to server/test/fixtures/.');

    await mongoose.disconnect();
})();

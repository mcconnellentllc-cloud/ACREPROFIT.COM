#!/usr/bin/env node
//
// scripts/backfill-rep-attribution.js
// Stage 2 of pr-rep-attribution.
//
// Backfills EXISTING records so customers (and their orders) attribute to the
// correct distributor _id instead of the creating admin. Uses the same
// slug -> distributor resolution as the Stage 1 write-path fix.
//
// PREREQUISITE: Stage 1 must be merged + deployed first. Stage 1 seeds the rep
// slug onto each distributor User record (initializeAdmins); this script reads
// those slugs to build the slug -> _id map. If no distributor slugs are found,
// the script aborts.
//
// USAGE (run from Render Shell, which has MONGODB_URI in the environment):
//   node scripts/backfill-rep-attribution.js            # DRY RUN - reports only, writes nothing
//   node scripts/backfill-rep-attribution.js --commit   # applies the changes
//
// SAFETY PROPERTIES:
//   - Dry-run by default. Writes NOTHING unless --commit is passed.
//   - Idempotent. Already-correct records are skipped, not re-touched. Safe to
//     run twice.
//   - Never overwrites a correct distributor _id with a fallback. If a slug
//     can't be resolved to a distributor, the record is left UNTOUCHED and
//     logged for manual review.
//   - Read-then-write per record; every change is logged.
//
// NOTE ON COLLECTION NAMES: this uses the Mongoose default collection names
// (lowercased + pluralized model names): User -> "users",
// ChemicalOrder -> "chemicalorders", Order -> "orders". These match the models
// in server/index.js (no custom collection option). Verify in the dry-run
// output (scanned counts should be non-zero) before committing.

const mongoose = require('mongoose');

const COMMIT = process.argv.includes('--commit');
const MONGODB_URI = process.env.MONGODB_URI;

if (!MONGODB_URI) {
    console.error('FATAL: MONGODB_URI environment variable is not set. Run this from the Render Shell.');
    process.exit(1);
}

async function main() {
    await mongoose.connect(MONGODB_URI);
    const db = mongoose.connection.db;
    const Users = db.collection('users');
    const collections = [
        ['chemicalorders', db.collection('chemicalorders')],
        ['orders', db.collection('orders')],
    ];

    console.log(`\n=== rep-attribution backfill (${COMMIT ? 'COMMIT — WILL WRITE' : 'DRY RUN — no writes'}) ===\n`);

    // 1) Build slug -> distributor _id map from the (Stage-1-seeded) distributor
    //    records, plus the set of valid distributor ids.
    const distributors = await Users
        .find({ role: 'distributor' })
        .project({ _id: 1, name: 1, representativeId: 1 })
        .toArray();

    const slugToDist = new Map();
    const distributorIds = new Set();
    for (const d of distributors) {
        distributorIds.add(d._id.toString());
        if (d.representativeId) slugToDist.set(String(d.representativeId).toLowerCase(), d._id);
    }

    if (slugToDist.size === 0) {
        console.error('ABORT: No distributor slugs found. Deploy Stage 1 first (it seeds distributor');
        console.error('       slugs in initializeAdmins), then re-run. Nothing was changed.');
        await mongoose.disconnect();
        process.exit(1);
    }
    console.log(`Distributors with slugs: ${[...slugToDist.keys()].sort().join(', ')}\n`);

    // 2) Customers: correct `representative` to the resolved distributor _id.
    const customers = await Users
        .find({ role: 'customer' })
        .project({ _id: 1, name: 1, representative: 1, representativeId: 1 })
        .toArray();

    let cScanned = 0, cChange = 0, cAlready = 0, cUnresolved = 0;
    const cSamples = [];
    // customerId -> corrected distributor _id (used to fix that customer's orders below)
    const correctedCustomerRep = new Map();

    for (const c of customers) {
        cScanned++;
        const slug = c.representativeId ? String(c.representativeId).toLowerCase() : null;
        const distId = slug ? slugToDist.get(slug) : null;

        if (!distId) {
            cUnresolved++;
            console.log(`  [SKIP unresolved] customer ${c._id} "${c.name || ''}" slug=${slug ?? '(none)'} — left untouched, review manually`);
            continue;
        }

        const currentRep = c.representative ? c.representative.toString() : null;

        // Already attributing to a distributor? Only treat as "already correct"
        // when it matches the slug-resolved distributor (idempotent skip).
        if (currentRep === distId.toString()) {
            cAlready++;
            correctedCustomerRep.set(c._id.toString(), distId); // record for order pass
            continue;
        }

        // representative is missing or points elsewhere (typically the creating
        // admin). Correct it to the resolved distributor.
        cChange++;
        correctedCustomerRep.set(c._id.toString(), distId);
        if (cSamples.length < 10) {
            cSamples.push(`    customer ${c._id} "${c.name || ''}": representative ${currentRep ?? '(none)'} -> ${distId} (slug "${slug}")`);
        }
        if (COMMIT) {
            await Users.updateOne({ _id: c._id }, { $set: { representative: distId } });
        }
    }

    console.log(`Customers: scanned=${cScanned} wouldChange=${cChange} alreadyCorrect=${cAlready} unresolved=${cUnresolved}`);
    if (cSamples.length) { console.log('  sample changes (before -> after):'); cSamples.forEach(s => console.log(s)); }
    console.log('');

    // 3) Orders: re-derive representativeId from the customer's corrected
    //    representative, but only when the order is NOT already attributed to a
    //    real distributor (never overwrite a correct attribution).
    for (const [name, coll] of collections) {
        const orders = await coll.find({}).project({ _id: 1, userId: 1, representativeId: 1 }).toArray();
        let oScanned = 0, oChange = 0, oAlready = 0, oSkipped = 0;
        const oSamples = [];

        for (const o of orders) {
            oScanned++;
            const custId = o.userId ? o.userId.toString() : null;

            // Resolve the customer's correct distributor: prefer the value we
            // just computed this run; else read the customer's current
            // representative if it already points at a real distributor.
            let newRep = custId ? correctedCustomerRep.get(custId) : null;
            if (!newRep && custId) {
                const cust = await Users.findOne({ _id: o.userId }, { projection: { representative: 1 } });
                if (cust && cust.representative && distributorIds.has(cust.representative.toString())) {
                    newRep = cust.representative;
                }
            }
            if (!newRep) { oSkipped++; continue; } // can't resolve to a distributor — leave untouched

            const currentRep = o.representativeId ? o.representativeId.toString() : null;
            if (currentRep === newRep.toString()) { oAlready++; continue; } // idempotent skip
            // Never overwrite an attribution that's already a real distributor.
            if (currentRep && distributorIds.has(currentRep)) { oSkipped++; continue; }

            oChange++;
            if (oSamples.length < 10) {
                oSamples.push(`    ${name} ${o._id}: representativeId ${currentRep ?? '(none)'} -> ${newRep}`);
            }
            if (COMMIT) {
                await coll.updateOne({ _id: o._id }, { $set: { representativeId: newRep } });
            }
        }

        console.log(`${name}: scanned=${oScanned} wouldChange=${oChange} alreadyCorrect=${oAlready} skipped=${oSkipped}`);
        if (oSamples.length) { console.log('  sample changes (before -> after):'); oSamples.forEach(s => console.log(s)); }
        console.log('');
    }

    console.log(COMMIT
        ? '=== COMMIT complete. Changes written. Safe to re-run (idempotent) to verify zero further changes. ==='
        : '=== DRY RUN complete. Nothing written. Re-run with --commit to apply. ===');

    await mongoose.disconnect();
}

main().catch(async (err) => {
    console.error('Backfill error:', err);
    try { await mongoose.disconnect(); } catch (_) {}
    process.exit(1);
});

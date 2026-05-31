#!/usr/bin/env node
//
// server/scripts/dedupe-inventory.js
//
// Merge TRUE duplicate Inventory rows: same chemicalId AND same distributorId.
// Per the production classification, only 4 of the 13 duplicate chemicalIds are
// true dupes (same distributor); the other 9 are different-distributor splits
// and are LEFT UNTOUCHED (they land in their own single-row groups here).
//
// Lives under server/ so it resolves mongoose from server/node_modules.
//
// USAGE (from repo root in Render Shell, which has MONGODB_URI set):
//   node server/scripts/dedupe-inventory.js            # DRY RUN - reports only, writes nothing
//   node server/scripts/dedupe-inventory.js --commit   # applies the merges
//
// SAFETY:
//   - Dry-run by default; writes NOTHING without --commit.
//   - Touches ONLY groups with >1 row sharing the same chemicalId + distributorId.
//     Different-distributor rows are never grouped together, so the 9 legit
//     splits are never merged.
//   - Idempotent: after a commit, every group is size 1, so a re-run reports 0.
//   - Preserves references: InventoryTransaction rows pointing at a removed
//     Inventory row are repointed to the kept row before deletion.
//   - Keeps the most-complete/priced row; sums quantities into it.

const mongoose = require('mongoose');

const COMMIT = process.argv.includes('--commit');
const MONGODB_URI = process.env.MONGODB_URI;

if (!MONGODB_URI) {
    console.error('FATAL: MONGODB_URI is not set. Run this from the Render Shell.');
    process.exit(1);
}

// Higher score = more complete/priced; used to choose the row to keep.
function completeness(r) {
    let s = 0;
    if (Number(r.averageCost || 0) > 0) s += 4;
    if (Number(r.lastCost || 0) > 0) s += 2;
    if (Number(r.quantityOnHand || 0) > 0) s += 1;
    return s;
}

async function main() {
    await mongoose.connect(MONGODB_URI);
    const db = mongoose.connection.db;
    const Inventories = db.collection('inventories');
    const Transactions = db.collection('inventorytransactions');

    console.log(`\n=== inventory dedupe (${COMMIT ? 'COMMIT - WILL WRITE' : 'DRY RUN - no writes'}) ===\n`);

    const all = await Inventories.find({}).toArray();

    // Group by chemicalId + distributorId. Rows with different distributorId
    // (the 9 legit splits) fall into separate groups and are never merged.
    const groups = new Map();
    for (const r of all) {
        const chem = (r.chemicalId || '').toString();
        if (!chem) continue; // ignore rows with no product reference
        const dist = r.distributorId ? r.distributorId.toString() : 'none';
        const key = `${chem}|${dist}`;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(r);
    }

    let mergeGroups = 0, rowsRemoved = 0, txRepointed = 0;

    for (const [key, rows] of groups) {
        if (rows.length < 2) continue; // single row = nothing to merge (includes the 9 splits)
        mergeGroups++;

        const sorted = [...rows].sort((a, b) =>
            completeness(b) - completeness(a)
            || Number(b.quantityOnHand || 0) - Number(a.quantityOnHand || 0)
            || new Date(b.updatedAt || 0) - new Date(a.updatedAt || 0)
        );
        const keeper = sorted[0];
        const redundant = sorted.slice(1);
        const redundantIds = redundant.map(r => r._id);

        const sumOnHand = rows.reduce((s, r) => s + Number(r.quantityOnHand || 0), 0);
        const sumReserved = rows.reduce((s, r) => s + Number(r.quantityReserved || 0), 0);
        const txCount = await Transactions.countDocuments({ inventoryId: { $in: redundantIds } });

        console.log(`Group: chemicalId=${key.split('|')[0]} distributorId=${key.split('|')[1]}`);
        console.log(`  product: ${keeper.productName || '(unknown)'}  rows: ${rows.length}`);
        rows.forEach(r => console.log(
            `    row ${r._id} qtyOnHand=${r.quantityOnHand || 0} reserved=${r.quantityReserved || 0} ` +
            `avgCost=${r.averageCost || 0} location=${r.location || '-'} ${String(r._id) === String(keeper._id) ? '<- KEEP' : '(remove)'}`
        ));
        console.log(`  -> keeper ${keeper._id}: qtyOnHand ${keeper.quantityOnHand || 0} -> ${sumOnHand}, reserved -> ${sumReserved}`);
        if (txCount) console.log(`  -> ${txCount} inventory transaction(s) ${COMMIT ? 'repointed' : 'would be repointed'} to keeper`);
        console.log('');

        rowsRemoved += redundant.length;
        txRepointed += txCount;

        if (COMMIT) {
            await Inventories.updateOne({ _id: keeper._id }, { $set: {
                quantityOnHand: sumOnHand,
                quantityReserved: sumReserved,
                quantityAvailable: sumOnHand - sumReserved,
                updatedAt: new Date()
            }});
            if (redundantIds.length) {
                await Transactions.updateMany({ inventoryId: { $in: redundantIds } }, { $set: { inventoryId: keeper._id } });
                await Inventories.deleteMany({ _id: { $in: redundantIds } });
            }
        }
    }

    console.log(`Summary: ${mergeGroups} duplicate group(s); ${rowsRemoved} redundant row(s) ${COMMIT ? 'removed' : 'would be removed'}; ${txRepointed} transaction(s) ${COMMIT ? 'repointed' : 'would be repointed'}.`);
    console.log(COMMIT
        ? '=== COMMIT complete. Re-run (idempotent) to confirm 0 remaining groups. ==='
        : '=== DRY RUN complete. Nothing written. Re-run with --commit to apply. ===');

    await mongoose.disconnect();
}

main().catch(async (err) => {
    console.error('dedupe error:', err);
    try { await mongoose.disconnect(); } catch (_) {}
    process.exit(1);
});

// ============ SUPPLIER SEED (Step 0 of Price Mining refactor) ============
// Creates supplier User docs for JABCO, SIMS, CORBET, CPD. Normalizes existing
// free-text Chemical.sourceSupplier values to clean codes. Backfills
// Chemical.supplierId so the bid sheet system can query by supplier User.
//
// Idempotent:
//  - Finds each supplier by email; skips creation if exists
//  - Syncs role/companyName/supplierCode if they drifted
//  - Never touches an existing supplier's password
//  - Safe to re-run on every deploy
//
// First-run output in Render console:
//   ✅ Normalized N chemicals: <aliases> → <CODE>
//   ✅ Created supplier <CODE> (<companyName>)
//      [INITIAL PASSWORD] <email>: <temp> — MUST CHANGE ON FIRST LOGIN
//      Chemicals linked to <CODE>: N
//
// Subsequent runs print "already seeded, skipping" or "synced drift".

const crypto = require('crypto');

const SUPPLIERS = [
    {
        supplierCode: 'JABCO',
        companyName: 'JABCO LLC',
        name: 'JABCO Sales',
        email: 'jabco@acreprofit.com',  // Placeholder - mail forwarding rule at Office 365
        phone: '',
        bidEligible: true,
        aliases: ['Jabco', 'jabco', 'JABCO LLC']
    },
    {
        supplierCode: 'SIMS',
        companyName: 'Sims Fertilizer & Chemical',
        name: 'Sims Sales',
        email: 'sims@acreprofit.com',
        phone: '',
        bidEligible: true,
        aliases: ['Sims Fertilizer & Chemical', 'Sims', 'sims']
    },
    {
        supplierCode: 'CORBET',
        companyName: 'Corbet Scientific, LLC',
        name: 'Corbet Sales',
        email: 'corbet@acreprofit.com',
        phone: '',
        // Direct-purchase only (Hydrovant). Never included in competitive bids.
        bidEligible: false,
        aliases: ['Corbet Scientific, LLC', 'Corbet Scientific', 'Corbet']
    },
    {
        supplierCode: 'CPD',
        companyName: 'Crop Protect Direct',
        name: 'CPD Sales',
        email: 'cpd@acreprofit.com',
        phone: '',
        bidEligible: false, // Tier-2 supplier (CPD -> JABCO -> AcreProfit). Not a direct AP supplier.
        aliases: ['Crop Protect Direct', 'cpd']
    }
];

async function initializeSuppliers(User, Chemical) {
    try {
        let totalLinked = 0;
        let totalCreated = 0;
        let totalUpdated = 0;
        let totalNormalized = 0;

        for (const spec of SUPPLIERS) {
            // --- Step 1: Normalize free-text sourceSupplier → clean code ---
            // Only non-canonical aliases - skip the code itself to avoid no-op update
            const aliasesToRewrite = spec.aliases.filter(a => a !== spec.supplierCode);
            if (aliasesToRewrite.length > 0) {
                const normResult = await Chemical.updateMany(
                    { sourceSupplier: { $in: aliasesToRewrite } },
                    { $set: { sourceSupplier: spec.supplierCode } }
                );
                if (normResult.modifiedCount > 0) {
                    console.log(`  Normalized ${normResult.modifiedCount} chemicals: ${aliasesToRewrite.join(', ')} → ${spec.supplierCode}`);
                    totalNormalized += normResult.modifiedCount;
                }
            }

            // --- Step 2: Find or create supplier User doc ---
            const existing = await User.findOne({ email: spec.email.toLowerCase() });
            let supplier;

            if (existing) {
                // Sync drift on role / companyName / supplierCode / bidEligible.
                // Never touches password.
                let changed = false;
                if (existing.role !== 'supplier') {
                    existing.role = 'supplier';
                    changed = true;
                }
                if (existing.companyName !== spec.companyName) {
                    existing.companyName = spec.companyName;
                    changed = true;
                }
                if (existing.supplierCode !== spec.supplierCode) {
                    existing.supplierCode = spec.supplierCode;
                    changed = true;
                }
                if (existing.bidEligible !== spec.bidEligible) {
                    existing.bidEligible = spec.bidEligible;
                    // markModified: Mongoose can skip persisting a boolean that
                    // matches the schema default. Force the dirty flag so the
                    // write actually lands.
                    existing.markModified('bidEligible');
                    changed = true;
                }
                if (changed) {
                    await existing.save();
                    console.log(`  Updated existing supplier: ${spec.supplierCode} (bidEligible=${spec.bidEligible})`);
                    totalUpdated++;
                } else {
                    console.log(`  Supplier ${spec.supplierCode} already seeded, skipping`);
                }
                supplier = existing;
            } else {
                // Create new supplier with random 12-char temp password.
                // mustChangePassword=true enforces rotation if the supplier ever
                // logs into /supplier-dashboard.html. Same pattern as initializeAdmins().
                const tempPassword = crypto.randomBytes(9).toString('base64').replace(/[+/=]/g, '').slice(0, 12);
                supplier = new User({
                    name: spec.name,
                    email: spec.email.toLowerCase(),
                    password: tempPassword,
                    role: 'supplier',
                    companyName: spec.companyName,
                    supplierCode: spec.supplierCode,
                    phone: spec.phone,
                    bidEligible: spec.bidEligible,
                    mustChangePassword: true
                });
                await supplier.save();
                console.log(`  Created supplier ${spec.supplierCode} (${spec.companyName})`);
                console.log(`  [INITIAL PASSWORD] ${spec.email}: ${tempPassword} — MUST CHANGE ON FIRST LOGIN`);
                totalCreated++;
            }

            // --- Step 3: Backfill Chemical.supplierId for unlinked products ---
            const linkResult = await Chemical.updateMany(
                {
                    sourceSupplier: spec.supplierCode,
                    $or: [{ supplierId: { $exists: false } }, { supplierId: null }]
                },
                { $set: { supplierId: supplier._id } }
            );
            if (linkResult.modifiedCount > 0) {
                console.log(`  Chemicals linked to ${spec.supplierCode}: ${linkResult.modifiedCount}`);
                totalLinked += linkResult.modifiedCount;
            }
        }

        console.log(`Supplier seed complete: created=${totalCreated}, updated=${totalUpdated}, normalized=${totalNormalized}, newly-linked=${totalLinked}`);
    } catch (err) {
        console.error('Supplier seed error:', err.message);
    }
}

module.exports = { initializeSuppliers };

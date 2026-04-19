// server/migrations/002_applications_to_passes.js
//
// One-shot migration: SprayProgram field rename applications[] → passes[]
// Stamps schemaVersion: 2 on every migrated doc so future migrations target
// precisely.
//
// Decision logic per doc (from Kyle's locked spec):
//   - Has applications[] with {name, timing, chemicals[]}   → rename lossless to passes[]
//   - Has top-level flat chemicals[], no applications[]     → wrap as passes[0]
//       name: "Preplant", timing: doc.timing || ""
//   - Already has passes[] or schemaVersion === 2           → skip
//
// Idempotent. Safe to re-run.
//
// Target docs: Milo Hardy, Milo Economical, and any legacy corn programs that
// existed before the corn-program replacement (those will be deleted by the
// seed route regardless — but migrate first so the schema read path doesn't
// explode on the admin page during the brief window between migration and
// seed replacement).
//
// Usage: node server/migrations/002_applications_to_passes.js
//   --dry-run    print intended changes without writing
//   --verbose    log each doc

const mongoose = require('mongoose');
require('dotenv').config();

async function run({ dryRun = false, verbose = false } = {}) {
  if (!process.env.MONGODB_URI) {
    throw new Error('MONGODB_URI not set');
  }
  await mongoose.connect(process.env.MONGODB_URI);

  // Load via raw collection to avoid schema-layer field filtering — the old
  // docs don't match the new schema so mongoose would strip applications[]
  // during hydration.
  const db = mongoose.connection.db;
  const coll = db.collection('sprayprograms');
  const all = await coll.find({}).toArray();

  console.log(`[002] ${all.length} spray programs found`);

  let renamed = 0;
  let wrapped = 0;
  let skipped = 0;
  let addedPassNumber = 0;

  for (const doc of all) {
    const _id = doc._id;
    const name = doc.name || '<unnamed>';

    // Skip: already migrated
    if (doc.schemaVersion === 2 && Array.isArray(doc.passes)) {
      skipped++;
      if (verbose) console.log(`  SKIP  ${name} (already v2)`);
      continue;
    }

    let passes;
    let action;

    if (Array.isArray(doc.applications) && doc.applications.length > 0) {
      // Case A: field rename, preserve everything. Top-level doc.timing (if any)
      // stays where it was — v2 schema ignores it, but we don't unset here since
      // some Case A docs may have meaningful top-level timing we'd rather leave
      // intact than destroy.
      passes = doc.applications.map((app, idx) => ({
        passNumber: idx + 1,
        name: app.name || `Pass ${idx + 1}`,
        timing: app.timing || '',
        chemicals: Array.isArray(app.chemicals) ? app.chemicals.map(c => ({
          chemicalId: c.chemicalId || c.chemical_id || null,
          productName: c.productName || c.name || null,
          rate: c.rate,
          rateUnit: c.rateUnit || c.unit,
          optional: c.optional || false,
          defaultOn: c.defaultOn !== undefined ? c.defaultOn : true,
          conditionNote: c.conditionNote || null,
        })) : [],
      }));
      addedPassNumber += passes.length;
      renamed++;
      action = 'RENAME';
    } else if (Array.isArray(doc.chemicals) && doc.chemicals.length > 0) {
      // Case B: flat chemicals[], wrap as single pass. Pull top-level timing
      // INTO passes[0].timing before unsetting (lossless preservation).
      passes = [{
        passNumber: 1,
        name: 'Preplant',
        timing: doc.timing || '',
        chemicals: doc.chemicals.map(c => ({
          chemicalId: c.chemicalId || c.chemical_id || null,
          productName: c.productName || c.name || null,
          rate: c.rate,
          rateUnit: c.rateUnit || c.unit,
          optional: false,
          defaultOn: true,
          conditionNote: null,
        })),
      }];
      wrapped++;
      action = 'WRAP ';
    } else {
      skipped++;
      if (verbose) console.log(`  SKIP  ${name} (no applications or chemicals)`);
      continue;
    }

    if (verbose) {
      console.log(`  ${action} ${name}  → ${passes.length} pass(es), ${passes.reduce((n, p) => n + p.chemicals.length, 0)} chems`);
    }

    if (!dryRun) {
      // Unset applications[] + flat chemicals[] + top-level timing (now preserved
      // in passes[0].timing for Case B; Case A never had top-level timing as
      // authoritative).
      await coll.updateOne(
        { _id },
        {
          $set: {
            schemaVersion: 2,
            passes,
            updatedAt: new Date(),
          },
          $unset: {
            applications: '',
            chemicals: '',
            timing: '',
          },
        }
      );
    }
  }

  console.log(`[002] complete${dryRun ? ' (DRY RUN)' : ''}`);
  console.log(`       renamed:  ${renamed}`);
  console.log(`       wrapped:  ${wrapped}`);
  console.log(`       skipped:  ${skipped}`);
  console.log(`       passes stamped: ${addedPassNumber + wrapped}`);

  await mongoose.disconnect();
  return { renamed, wrapped, skipped };
}

if (require.main === module) {
  const args = process.argv.slice(2);
  run({
    dryRun: args.includes('--dry-run'),
    verbose: args.includes('--verbose'),
  }).catch(err => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = run;

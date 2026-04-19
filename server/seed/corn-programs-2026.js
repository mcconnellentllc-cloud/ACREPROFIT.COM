// server/seed/corn-programs-2026.js
//
// Seeds the 5 corn programs locked with Kyle:
//   1. Corn Dryland Standard     (2-pass)
//   2. Corn Dryland Heavy        (3-pass)
//   3. Corn Irrigated Standard   (2-pass)
//   4. Corn Irrigated Heavy      (3-pass)
//   5. Corn Post-Wheat Rotation  (4-pass, optional flumi on Pass 4)
//
// Behavior:
//   - Deletes ALL existing corn programs (clean replacement per locked spec).
//     Milo Hardy + Milo Economical are untouched (migration renames their
//     applications[] to passes[] in place).
//   - Resolves chemical references by tradeName. Throws if no match.
//   - Runs atrazine cap assertion BEFORE insert. Throws on violation → whole
//     seed aborts, no partial writes.
//   - Stamps capCheckStatus='passed' and atzSeasonLbPerAcre on each seeded doc.
//
// Usage:
//   node server/seed/corn-programs-2026.js
//   node server/seed/corn-programs-2026.js --dry-run

const mongoose = require('mongoose');
require('dotenv').config();

const Chemical = require('../models/Chemical');
const SprayProgram = require('../models/SprayProgram');
const { assertAtzCaps, computeAtrazineTotal } = require('../lib/atrazineCap');
const { PROGRAMS } = require('./corn-programs-data');

async function resolveChemicals(programs) {
  const needed = new Map();
  for (const p of programs) {
    for (const pass of p.passes) {
      for (const row of pass.chemicals) {
        const key = `${row.tradeName}||${row.manufacturer || ''}`;
        needed.set(key, { tradeName: row.tradeName, manufacturer: row.manufacturer || null });
      }
    }
  }

  const tradeNames = Array.from(new Set(Array.from(needed.values()).map(n => n.tradeName)));
  const chems = await Chemical.find({ tradeName: { $in: tradeNames } }).lean();

  const byKey = new Map();
  const byTradeName = new Map();
  chems.forEach(c => {
    byKey.set(`${c.tradeName}||${c.manufacturer || ''}`, c);
    const list = byTradeName.get(c.tradeName) || [];
    list.push(c);
    byTradeName.set(c.tradeName, list);
  });

  const resolver = new Map();
  const missing = [];
  for (const [key, { tradeName, manufacturer }] of needed) {
    let chem = byKey.get(key);
    if (!chem) {
      const candidates = byTradeName.get(tradeName) || [];
      chem = candidates.find(c => c.status === 'approved') || candidates[0];
    }
    if (!chem) {
      missing.push({ tradeName, manufacturer });
    } else {
      resolver.set(String(chem._id), chem);
      resolver.set(`TRADENAME:${tradeName}`, chem);
    }
  }

  if (missing.length > 0) {
    throw new Error(
      `Seed cannot proceed — missing chemicals from catalog. Import MAINCHEM first or verify tradeName spelling:\n` +
      missing.map(m => `  - "${m.tradeName}"${m.manufacturer ? ` (${m.manufacturer})` : ''}`).join('\n')
    );
  }

  return resolver;
}

async function run({ dryRun = false } = {}) {
  if (!process.env.MONGODB_URI) throw new Error('MONGODB_URI not set');
  await mongoose.connect(process.env.MONGODB_URI);

  console.log('[seed corn] resolving chemicals...');
  const resolver = await resolveChemicals(PROGRAMS);
  console.log(`[seed corn] resolved ${resolver.size / 2} chemicals`);

  console.log('[seed corn] running atrazine cap assertion on all programs...');
  const capResults = [];
  for (const p of PROGRAMS) {
    const programForCheck = {
      name: p.name,
      passes: p.passes.map(pass => ({
        passNumber: pass.passNumber,
        chemicals: pass.chemicals.map(row => {
          const chem = resolver.get(`TRADENAME:${row.tradeName}`);
          return { ...row, chemicalId: chem._id };
        }),
      })),
    };
    const checkResolver = new Map();
    programForCheck.passes.forEach(pass => {
      pass.chemicals.forEach(row => {
        checkResolver.set(String(row.chemicalId), resolver.get(String(row.chemicalId)));
      });
    });
    const result = assertAtzCaps(programForCheck, checkResolver);
    capResults.push({ name: p.name, ...result });
    console.log(`  ✓ ${p.name}  atz=${result.totalLbPerAcre} lb ai/A  pass breakdown=${JSON.stringify(result.perPass)}`);
  }

  if (dryRun) {
    console.log('[seed corn] DRY RUN — no writes performed');
    await mongoose.disconnect();
    return { programs: PROGRAMS.length, capResults, dryRun: true };
  }

  const deleted = await SprayProgram.deleteMany({ crop: 'corn' });
  console.log(`[seed corn] deleted ${deleted.deletedCount} existing corn programs`);

  for (let i = 0; i < PROGRAMS.length; i++) {
    const p = PROGRAMS[i];
    const capResult = capResults[i];
    const doc = new SprayProgram({
      schemaVersion: 2,
      name: p.name,
      crop: p.crop,
      tier: p.tier,
      description: p.description,
      rotationNotes: p.rotationNotes,
      grazingNotes: p.grazingNotes,
      notes: p.notes,
      passes: p.passes.map(pass => ({
        passNumber: pass.passNumber,
        name: pass.name,
        timing: pass.timing,
        chemicals: pass.chemicals.map(row => {
          const chem = resolver.get(`TRADENAME:${row.tradeName}`);
          return {
            chemicalId: chem._id,
            productName: chem.tradeName,
            rate: row.rate,
            rateUnit: row.rateUnit,
            optional: row.optional || false,
            defaultOn: row.defaultOn !== undefined ? row.defaultOn : true,
            conditionNote: row.conditionNote || null,
          };
        }),
      })),
      containsPendingChemicals: false,
      atzSeasonLbPerAcre: capResult.totalLbPerAcre,
      capCheckStatus: 'passed',
      active: true,
    });
    await doc.save();
    console.log(`  + inserted ${p.name}`);
  }

  console.log(`[seed corn] complete — ${PROGRAMS.length} programs seeded`);
  await mongoose.disconnect();
  return { programs: PROGRAMS.length, capResults };
}

if (require.main === module) {
  const args = process.argv.slice(2);
  run({ dryRun: args.includes('--dry-run') }).catch(err => {
    console.error('[seed corn] FAILED:', err.message);
    process.exit(1);
  });
}

module.exports = { run };

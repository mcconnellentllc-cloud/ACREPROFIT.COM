// server/lib/atrazineCap.js
//
// Shared atrazine cap-math. Used by:
//   - Seed-time assertion (corn-programs-2026.js fails fast if any program
//     exceeds 2.2 lb ai/A spring or 2.5 lb ai/A annual)
//   - Runtime validator on program save (future use)
//
// Label caps (Aatrex 4L):
//   Single application:     2.0 lb ai/A
//   Annual (12-month):      2.5 lb ai/A
//   Kyle's op cap:          2.2 lb ai/A (spring) — tighter than label, per user spec
//
// Math strategy: sum per-pass atrazine contribution across all chemicals in the
// program. For each pass chemical, look up its active ingredients, find any
// named 'atrazine', use its lbPerGal for liquid rates or percentByWeight for
// dry rates. Optional chemicals are included ONLY if defaultOn === true.
//
// Rate-unit normalization to gallons (for liquid × lbPerGal math):
//   1 gal = 1 gal
//   1 qt  = 0.25 gal
//   1 pt  = 0.125 gal
//   1 fl oz = 1/128 gal  (0.0078125)
//
// Rate-unit normalization to pounds (for dry × percentByWeight math):
//   1 lb  = 1 lb
//   1 oz  = 0.0625 lb
//   1 dry oz = 0.0625 lb
//
// Caller supplies a resolver function that maps chemicalId → Chemical doc.
// Seed code passes a Map<chemicalId, chemDoc> built at seed time; runtime
// code passes an async DB lookup. Cap-checker stays sync by requiring the
// caller to pre-resolve.

const GAL_PER_UNIT = {
  gal: 1,
  qt: 0.25,
  pt: 0.125,
  'fl oz': 1 / 128,
  'fl_oz': 1 / 128,
  floz: 1 / 128,
  oz: 1 / 128,  // ambiguous — treated as fl oz for liquids. Dry uses 'dry oz'.
};

const LB_PER_UNIT = {
  lb: 1,
  lbs: 1,
  'dry oz': 0.0625,
  'dry_oz': 0.0625,
};

/**
 * Compute atrazine lb ai/A contributed by a single pass chemical row.
 *
 * @param {Object} row               pass chemical — { chemicalId, rate, rateUnit, optional, defaultOn }
 * @param {Object} chem              resolved Chemical doc with activeIngredients[]
 * @returns {number}                 lb ai/A of atrazine (0 if none, or optional+off)
 */
function atzLbPerAcreForRow(row, chem) {
  if (!chem || !Array.isArray(chem.activeIngredients)) return 0;
  if (row.optional && row.defaultOn === false) return 0;

  const atzAi = chem.activeIngredients.find(ai => ai.name === 'atrazine');
  if (!atzAi) return 0;

  const rateUnit = String(row.rateUnit || '').toLowerCase();

  // Liquid path: rate × gal/unit × lbPerGal
  if (atzAi.lbPerGal != null) {
    const galPerUnit = GAL_PER_UNIT[rateUnit];
    if (galPerUnit == null) {
      throw new Error(`unknown liquid rate unit "${row.rateUnit}" for ${chem.tradeName}`);
    }
    return row.rate * galPerUnit * atzAi.lbPerGal;
  }

  // Dry path: rate × lb/unit × (percentByWeight / 100)
  if (atzAi.percentByWeight != null) {
    const lbPerUnit = LB_PER_UNIT[rateUnit];
    if (lbPerUnit == null) {
      throw new Error(`unknown dry rate unit "${row.rateUnit}" for ${chem.tradeName}`);
    }
    return row.rate * lbPerUnit * (atzAi.percentByWeight / 100);
  }

  // AI present but concentration unknown — can't compute, treat as 0 but flag
  // at caller level (seed assertion fails for concentration-unknown atrazine).
  throw new Error(
    `chemical ${chem.tradeName} has atrazine AI but no concentration data — cap math impossible`
  );
}

/**
 * Sum atrazine across all passes in a program.
 *
 * @param {Object} program           { passes: [{ chemicals: [...] }] }
 * @param {Map}    chemResolver      Map<chemicalIdString, chemDoc>
 * @returns {Object}                 { totalLbPerAcre, perPass: [{passNumber, lbPerAcre}] }
 */
function computeAtrazineTotal(program, chemResolver) {
  const perPass = [];
  let total = 0;

  for (const pass of program.passes || []) {
    let passTotal = 0;
    for (const row of pass.chemicals || []) {
      const key = String(row.chemicalId);
      const chem = chemResolver.get(key);
      const contrib = atzLbPerAcreForRow(row, chem);
      passTotal += contrib;
    }
    perPass.push({ passNumber: pass.passNumber, lbPerAcre: round4(passTotal) });
    total += passTotal;
  }

  return { totalLbPerAcre: round4(total), perPass };
}

/**
 * Seed-time assertion. Throws with detail if any program violates caps.
 *
 * @param {Object} program
 * @param {Map}    chemResolver
 * @param {Object} caps              { singleApp: 2.0, annual: 2.2 }
 */
function assertAtzCaps(program, chemResolver, caps = { singleApp: 2.0, annual: 2.2 }) {
  const { totalLbPerAcre, perPass } = computeAtrazineTotal(program, chemResolver);

  // Single-application cap: no pass may exceed caps.singleApp lb ai/A
  const over = perPass.find(p => p.lbPerAcre > caps.singleApp + 0.001);
  if (over) {
    throw new Error(
      `[${program.name}] pass ${over.passNumber} atrazine = ${over.lbPerAcre} lb ai/A — exceeds single-app cap ${caps.singleApp}`
    );
  }

  // Annual cap: sum across passes
  if (totalLbPerAcre > caps.annual + 0.001) {
    throw new Error(
      `[${program.name}] total atrazine = ${totalLbPerAcre} lb ai/A — exceeds annual cap ${caps.annual}. per-pass: ${JSON.stringify(perPass)}`
    );
  }

  return { totalLbPerAcre, perPass };
}

function round4(n) { return Math.round(n * 10000) / 10000; }

module.exports = {
  atzLbPerAcreForRow,
  computeAtrazineTotal,
  assertAtzCaps,
  GAL_PER_UNIT,
  LB_PER_UNIT,
};

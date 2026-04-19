#!/usr/bin/env node
/**
 * parse-mainchem.js
 *
 * One-shot parser. Reads MAINCHEM xlsx sheet, emits:
 *   - server/seed/mainchem-source.json   (input for POST /api/admin/mainchem/import)
 *
 * Exclusions:
 *   - use === 'FERTILIZER'    (skip, reason: 'fertilizer_excluded')
 *   - missing tradeName + AI  (skip, reason: 'missing_required_field')
 *
 * AI concentration parsing pipeline (ordered):
 *   1. Parenthetical "(X.XX lb/gal)" or "(X.XX lb)"  → canonical lbPerGal
 *   2. Premix "X.XX + Y.YY lb/gal"                   → two AI entries, flag 'premix_needs_review'
 *   3. Liquid "X.XX L/SL/SC/EC/E/S/AC"               → lbPerGal (if no % sign)
 *   4. Dry or percent-by-weight "X.XX%? WDG/DF/..."  → percentByWeight
 *   5. Explicit "X.XX lb[ ae]/gal"                   → lbPerGal
 *   6. None match                                     → aiParseStatus: 'unparseable'
 *
 * Dedupe pass at end merges duplicate (tradeName, manufacturer) rows with
 * non-null-preference policy. Prevents unique-index violations on import.
 *
 * Usage: node scripts/parse-mainchem.js <input.xlsx> <output.json>
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function normalizeAi(raw) {
  if (!raw) return null;
  let s = String(raw).trim().toLowerCase().replace(/\s+/g, ' ');
  s = s.replace(/\s*-?\s*methyl\b/g, '-methyl');
  const fixes = {
    'two4d': '2,4-d',
    '2,4-d': '2,4-d',
    '2,4d': '2,4-d',
    's-metolachlor': 's-metolachlor',
    's-metolachlor ': 's-metolachlor',
    'metolachlor': 'metolachlor',
    'atrazine': 'atrazine',
    'glyphosate': 'glyphosate',
    'flumioxazin': 'flumioxazin',
    'mesotrione': 'mesotrione',
    'dicamba': 'dicamba',
    'pyroxasulfone': 'pyroxasulfone',
    'fluthiacet-methyl': 'fluthiacet-methyl',
    'thiencarboazone-methyl': 'thiencarbazone-methyl',
    'isoxoflutole': 'isoxaflutole',
  };
  return fixes[s] || s;
}

const MOA_CLASS = {
  '1': 'ACCase', '2': 'ALS', '3': 'microtubule', '4': 'auxin', '5': 'triazine',
  '9': 'EPSPS', '10': 'glutamine-synthetase', '13': 'DOXP', '14': 'PPO',
  '15': 'VLCFA', '19': 'auxin-transport', '22': 'PS-I', '27': 'HPPD',
};

const AI_TO_CLASS = {
  '2,4-d': 'auxin',
  'acetochlor': 'VLCFA',
  'atrazine': 'triazine',
  'bicyclopyrone': 'HPPD',
  'bromoxynil': 'PS-II',
  'carentrazone': 'PPO',
  'clopyralid': 'auxin',
  'dicamba': 'auxin',
  'diflufenzopyr': 'auxin-transport',
  'dimethenamid': 'VLCFA',
  'fluroxypyr': 'auxin',
  'flumetsulam': 'ALS',
  'flumioxazin': 'PPO',
  'fluthiacet-methyl': 'PPO',
  'glufosinate': 'glutamine-synthetase',
  'glyphosate': 'EPSPS',
  'isoxaflutole': 'HPPD',
  'mesotrione': 'HPPD',
  'metolachlor': 'VLCFA',
  'nicosulfuron': 'ALS',
  'paraquat': 'PS-I',
  'pendimethalin': 'microtubule',
  'prometryn': 'triazine',
  'pyroxasulfone': 'VLCFA',
  'rimsulfuron': 'ALS',
  's-metolachlor': 'VLCFA',
  'saflufenacil': 'PPO',
  'tembotrione': 'HPPD',
  'thiencarbazone-methyl': 'ALS',
  'topramezone': 'HPPD',
};

function classForAi(aiName) {
  if (!aiName) return null;
  return AI_TO_CLASS[aiName] || null;
}

function parseMoaNumbers(moa) {
  if (!moa) return [];
  return String(moa).split(/[+,\s]+/).map(s => s.trim()).filter(Boolean)
    .map(n => (MOA_CLASS[n] ? parseInt(n, 10) : null))
    .filter(n => n !== null);
}

const CHEMISTRY_CLASS_DESCRIPTORS = new Set([
  'chloroacetamide', 'triazine', 'sulfonylurea', 'dinitroaniline',
  'cyclohexanedione', 'benzoic acid', 'bipyridylium', 'diphenylether',
  'imidazolinone', 'triazinone', 'phenoxy-carboxylic acid', 'organoarsenical',
  'urea', 'glycine', 'n-phenylphthalimide', 'pyridine', 'triketone',
  'thiocarbamate', 'benzamide', 'pyridinecarboxylic acid', 'pyridine carboxylic acid',
  'quinaline carboxylic acid', '6-dichloro-o-anisic acid', 'unclassified',
  'aquatic plant control', 'granular', 'clopyralid', '2,4d', 'benzoylpyrazole',
  'replaces crop oil (higher concentration)',
  'several',
]);

function isChemistryClassDescriptor(s) {
  if (!s) return false;
  const lower = s.toLowerCase().trim();
  if (CHEMISTRY_CLASS_DESCRIPTORS.has(lower)) return true;
  if (lower.includes(' + ')) {
    const parts = lower.split(' + ').map(p => p.trim());
    return parts.every(p => CHEMISTRY_CLASS_DESCRIPTORS.has(p));
  }
  return false;
}

function parseConcentration(concStr, aiList) {
  if (!concStr || typeof concStr !== 'string') {
    return {
      entries: aiList.map(() => ({ lbPerGal: null, percentByWeight: null })),
      status: 'no_concentration_data',
      reason: 'source row had null concentration',
    };
  }

  const s = concStr.trim();

  if (isChemistryClassDescriptor(s)) {
    return {
      entries: aiList.map(() => ({ lbPerGal: null, percentByWeight: null })),
      status: 'no_concentration_data',
      reason: `chemistry class descriptor: "${s}"`,
    };
  }

  // Step 1: Premix "(X.XX + Y.YY [+ Z.ZZ] lb[/gal])"
  const premixMatch = s.match(/\(\s*(\d+\.?\d*)\s*\+\s*(\d+\.?\d*)\s*(?:\+\s*(\d+\.?\d*)\s*)?(?:lb(?:\/gal)?)\s*\)/i);
  if (premixMatch) {
    const vals = [premixMatch[1], premixMatch[2], premixMatch[3]]
      .filter(v => v !== undefined).map(v => parseFloat(v));
    const entries = aiList.map((_, i) => ({
      lbPerGal: vals[i] ?? null,
      percentByWeight: null,
    }));
    return {
      entries,
      status: vals.length === aiList.length ? 'premix_parsed' : 'premix_needs_review',
      reason: vals.length === aiList.length ? null
        : `premix has ${vals.length} concentrations for ${aiList.length} AIs — verify alignment`,
    };
  }

  // Step 2: Parenthetical single "(X.XX lb[/gal])"
  const parenMatch = s.match(/\(\s*(\d+\.?\d*)\s*lb(?:\/gal)?\s*\)/i);
  if (parenMatch) {
    const val = parseFloat(parenMatch[1]);
    const entries = aiList.map((_, i) => ({
      lbPerGal: i === 0 ? val : null,
      percentByWeight: null,
    }));
    return {
      entries,
      status: aiList.length === 1 ? 'parsed' : 'premix_needs_review',
      reason: aiList.length > 1 ? 'single concentration value for multi-AI product' : null,
    };
  }

  // Step 3: Liquid formulation (no % sign — % flips interpretation to percent-by-weight)
  const liquidMatch = s.match(/(\d+\.?\d*)\s*(L|SL|SC|EC|E|S|AC)\b(?!BS)/i);
  const parenHasLbGal = /\(\s*\d+\.?\d*.*lb(?:\/gal)?\s*\)/i.test(s);
  const hasPercent = /%/.test(s);
  if (liquidMatch && !parenHasLbGal && !hasPercent) {
    const val = parseFloat(liquidMatch[1]);
    const entries = aiList.map((_, i) => ({
      lbPerGal: i === 0 ? val : null,
      percentByWeight: null,
    }));
    return {
      entries,
      status: aiList.length === 1 ? 'parsed' : 'premix_needs_review',
      reason: aiList.length > 1 ? 'liquid shorthand for multi-AI — verify' : null,
    };
  }

  // Step 4: Dry OR percent-prefixed-liquid.
  // "40% SL" → 40 percent by weight (NOT 40 lb/gal) — percent sign disambiguates.
  const dryMatch = s.match(/(\d+\.?\d*)\s*%\s*(WDG|DF|WG|DG|SG|WP|G|SL|SC|EC|L|E|S|AC)\b/i)
    || s.match(/(\d+\.?\d*)\s*(WDG|DF|WG|DG|SG|WP|G)\b/i);
  if (dryMatch) {
    const val = parseFloat(dryMatch[1]);
    const entries = aiList.map((_, i) => ({
      lbPerGal: null,
      percentByWeight: i === 0 ? val : null,
    }));
    return {
      entries,
      status: aiList.length === 1 ? 'parsed' : 'premix_needs_review',
      reason: aiList.length > 1 ? 'single percent for multi-AI product' : null,
    };
  }

  // Step 5: Explicit "X.XX lb[ ae]/gal"
  const explicitMatch = s.match(/(\d+\.?\d*)\s*lb\s*(?:ae\s*)?\/gal\b/i);
  if (explicitMatch) {
    const val = parseFloat(explicitMatch[1]);
    const entries = aiList.map((_, i) => ({
      lbPerGal: i === 0 ? val : null,
      percentByWeight: null,
    }));
    return {
      entries,
      status: aiList.length === 1 ? 'parsed' : 'premix_needs_review',
      reason: aiList.length > 1 ? 'single lb/gal for multi-AI product' : null,
    };
  }

  return {
    entries: aiList.map(() => ({ lbPerGal: null, percentByWeight: null })),
    status: 'unparseable',
    reason: `concentration string "${s}" did not match any known format`,
  };
}

function main() {
  const inPath = process.argv[2];
  const outPath = process.argv[3];
  if (!inPath || !outPath) {
    console.error('Usage: node parse-mainchem.js <input.xlsx> <output.json>');
    process.exit(1);
  }

  const wb = XLSX.readFile(inPath);
  const ws = wb.Sheets['MAINCHEM'];
  if (!ws) { console.error('MAINCHEM sheet not found'); process.exit(1); }

  // MAINCHEM: headers on row 21 (1-indexed). Data starts row 22.
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, range: 20 });
  const dataRows = rows.slice(1);

  const records = [];
  const skipped = [];

  dataRows.forEach((row, idx) => {
    const rowIndex = 22 + idx;
    const get = (col) => {
      const v = row[col];
      if (v === null || v === undefined) return null;
      const s = String(v).trim();
      return s === '' || s === '.' ? null : s;
    };

    const tradeName = get(0);
    const ai1 = get(2);
    const ai2 = get(3);
    const aiRest = get(4);
    const concentration = get(5);
    const uom = get(6);
    const moa = get(7);
    const wssa = get(8);
    const manufacturer = get(9);
    const use = get(10);
    const control = get(11);
    const purposes = [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]
      .map(c => get(c)).filter(Boolean);
    const pkg = get(1);

    if (use && use.toUpperCase() === 'FERTILIZER') {
      skipped.push({ rowIndex, tradeName, reason: 'fertilizer_excluded',
        rawData: { tradeName, use, manufacturer } });
      return;
    }

    if (!tradeName || !ai1) {
      skipped.push({ rowIndex, tradeName: tradeName || '<missing>',
        reason: 'missing_required_field',
        rawData: { tradeName, ai1, use, manufacturer } });
      return;
    }

    const rawAis = [ai1, ai2, aiRest].filter(Boolean);
    const aiNames = [];
    rawAis.forEach(raw => {
      raw.split(/\s*\+\s*/).forEach(piece => {
        const n = normalizeAi(piece);
        if (n) aiNames.push(n);
      });
    });

    const concResult = parseConcentration(concentration, aiNames);
    const moaNumbers = parseMoaNumbers(moa);

    const activeIngredients = aiNames.map((name, i) => ({
      name,
      class: classForAi(name),
      lbPerGal: concResult.entries[i]?.lbPerGal ?? null,
      percentByWeight: concResult.entries[i]?.percentByWeight ?? null,
    }));

    records.push({
      tradeName, manufacturer, pkg, uom, use, control,
      moaRaw: moa, moaNumbers, wssa,
      activeIngredients, purposes,
      rawConcentration: concentration,
      aiParseStatus: concResult.status,
      aiParseReason: concResult.reason,
      sourceRow: rowIndex,
    });

    if (concResult.status === 'premix_needs_review') {
      skipped.push({ rowIndex, tradeName, reason: 'premix_needs_review',
        rawData: { concentration, aiNames, parsed: concResult.entries } });
    } else if (concResult.status === 'unparseable') {
      skipped.push({ rowIndex, tradeName, reason: 'ai_concentration_unparseable',
        rawData: { concentration, aiNames, note: concResult.reason } });
    }
  });

  // Dedupe by (tradeName, manufacturer) — merge non-null from either row.
  // Known source dupes: FOAM BUSTER/HELENA rows 313-314, ROZOL PD BAIT/LIPHAT rows 607-608.
  const dedupedMap = new Map();
  const dupeCount = { merged: 0 };
  for (const r of records) {
    const key = `${r.tradeName}||${r.manufacturer || ''}`;
    const existing = dedupedMap.get(key);
    if (!existing) {
      dedupedMap.set(key, r);
    } else {
      const merged = {};
      for (const k of Object.keys(r)) {
        const eVal = existing[k];
        const nVal = r[k];
        if (eVal === null || eVal === undefined || eVal === '') {
          merged[k] = nVal;
        } else if (nVal === null || nVal === undefined || nVal === '') {
          merged[k] = eVal;
        } else if (Array.isArray(eVal) && Array.isArray(nVal)) {
          merged[k] = Array.from(new Set([...eVal, ...nVal]));
        } else {
          merged[k] = nVal;
        }
      }
      dedupedMap.set(key, merged);
      dupeCount.merged++;
    }
  }
  const dedupedRecords = Array.from(dedupedMap.values());

  const out = {
    sourceFile: path.basename(inPath),
    generatedAt: new Date().toISOString(),
    totalRows: dataRows.filter(r => r.some(c => c !== null)).length,
    validRecords: dedupedRecords.length,
    duplicatesMerged: dupeCount.merged,
    skippedCount: skipped.length,
    records: dedupedRecords,
    skipped,
  };

  fs.writeFileSync(outPath, JSON.stringify(out, null, 2));
  console.log(`parsed ${dedupedRecords.length} chemicals (merged ${dupeCount.merged} dupes), skipped ${skipped.length}`);
  console.log(`  by reason:`);
  const byReason = {};
  skipped.forEach(s => { byReason[s.reason] = (byReason[s.reason] || 0) + 1; });
  Object.entries(byReason).forEach(([r, n]) => console.log(`    ${r}: ${n}`));
  console.log(`wrote ${outPath}`);
}

if (require.main === module) main();
module.exports = { normalizeAi, classForAi, parseMoaNumbers, parseConcentration };

// server/models/Chemical.js
//
// Chemical catalog entry. Path A (AI-first) schema for cross-brand cap math.
// Each trade name is its own doc (EPA reg numbers differ across manufacturers,
// and the pack selector UI already handles supplier variants).
//
// Commit 1 additions:
//   - activeIngredients[]           per-AI name, class, concentration
//   - status                        approval gate (hard-blocks at order submit)
//   - addedBy                       audit trail
//   - moaNumbers[] + wssa + use + control  queryable classification
//   - purposes[]                    crop/timing tags from MAINCHEM
//   - aiParseStatus + aiParseReason source-row parse outcome (diagnostic)
//   - sourceRow                     MAINCHEM row number for skip-report linking

const mongoose = require('mongoose');

const ActiveIngredientSchema = new mongoose.Schema({
  // Normalized lowercase AI name. Matches across trade variants for cap math.
  // e.g. "atrazine" matches whether the farmer picked AAtrex, AAtrex 4L, or
  // Drexel Atrazine 4L.
  name: { type: String, required: true, lowercase: true, trim: true, index: true },
  // MOA class from canonical AI→class map (e.g. "triazine", "EPSPS", "HPPD").
  // Computed at parse time from AI name, NOT from positional MOA index on the
  // source row — MAINCHEM AI order and MOA order don't always align.
  class: { type: String, default: null },
  // Liquid formulations only. e.g. Atrazine 4L → 4.0
  lbPerGal: { type: Number, default: null },
  // Dry formulations only. e.g. Valor 51 WDG → 51
  percentByWeight: { type: Number, default: null },
}, { _id: false });

const ChemicalSchema = new mongoose.Schema({
  // Display name — keep verbatim from MAINCHEM. Uniqueness enforced by
  // (tradeName, manufacturer) pair; same tradeName from different manufacturers
  // are distinct SKUs (e.g. "Atrazine 4L" from Winfield vs Helena).
  tradeName: { type: String, required: true, trim: true, index: true },
  manufacturer: { type: String, default: null, trim: true },
  pkg: { type: String, default: null, trim: true },
  uom: { type: String, default: null, trim: true },

  // Top-level category (HERBICIDE, INSECTICIDE, FUNGICIDE, GLYPHOSATES, etc.)
  use: { type: String, default: null, index: true },
  // What it controls (Broadleaves, Grass, Gra_Broad, Insect, Additive, etc.)
  control: { type: String, default: null },

  // AI list (always at least 1 for non-adjuvant products). Premixes get multiple.
  activeIngredients: {
    type: [ActiveIngredientSchema],
    default: [],
    validate: v => Array.isArray(v),
  },

  // Source audit — MOA numbers as stored in MAINCHEM (e.g. [5, 14, 15]).
  // These are preserved for traceability but NOT the source of truth for
  // per-AI class assignment (that's AI_TO_CLASS lookup in the parser).
  moaNumbers: { type: [Number], default: [] },
  wssa: { type: String, default: null },

  // Purpose tags from MAINCHEM (e.g. "Corn - Pre-emerge", "Wheat Topdress").
  // Used to filter available chemicals per recipe pass context.
  purposes: { type: [String], default: [] },

  // Approval gate. 'pending' chemicals can appear in draft recipes (soft warn at
  // save) but hard-block at POST /api/chemical-orders submit.
  status: {
    type: String,
    enum: ['approved', 'pending', 'rejected'],
    default: 'pending',
    index: true,
  },
  addedBy: { type: String, default: null },  // 'mainchem-import' | userId | 'farmer-suggestion'

  // Diagnostic — preserved from parser run. If aiParseStatus !== 'parsed' the
  // admin inline editor surfaces this row for review.
  aiParseStatus: {
    type: String,
    enum: ['parsed', 'premix_parsed', 'premix_needs_review', 'no_concentration_data', 'unparseable', null],
    default: null,
  },
  aiParseReason: { type: String, default: null },
  rawConcentration: { type: String, default: null },
  sourceRow: { type: Number, default: null },  // MAINCHEM row number

  createdAt: { type: Date, default: Date.now },
  updatedAt: { type: Date, default: Date.now },
});

// Compound index so the same trade name from different manufacturers can coexist
ChemicalSchema.index({ tradeName: 1, manufacturer: 1 }, { unique: false });
// Cap-math lookups need fast access by AI name across all chemicals
ChemicalSchema.index({ 'activeIngredients.name': 1 });

ChemicalSchema.pre('save', function(next) {
  this.updatedAt = Date.now();
  next();
});

module.exports = mongoose.models.Chemical || mongoose.model('Chemical', ChemicalSchema);

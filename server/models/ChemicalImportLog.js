// server/models/ChemicalImportLog.js
//
// One doc per POST /api/admin/mainchem/import run. Persisted so the admin page
// can list historical runs and link into the skip-report detail view for inline
// editing of unparseable rows.
//
// status workflow:
//   complete   → run finished, skips present, not yet reviewed
//   reviewed   → admin walked the skips, no action required (e.g. all
//                fertilizer_excluded as expected)
//   resolved   → admin hand-edited the unparseable rows, seed data patched

const mongoose = require('mongoose');

const SKIP_REASONS = [
  'fertilizer_excluded',      // use === 'FERTILIZER' — expected exclusion, not a failure
  'missing_required_field',   // no tradeName or no ai1 — can't create doc
  'ai_concentration_unparseable',  // concentration text didn't match any format
  'premix_needs_review',      // premix parsed but AI-to-value alignment needs verification
  'duplicate_reg_number',     // future — when EPA reg # dedup lands
];

const SkippedRowSchema = new mongoose.Schema({
  rowIndex: { type: Number, required: true },  // original MAINCHEM row (1-indexed)
  tradeName: { type: String, default: null },
  reason: { type: String, required: true, enum: SKIP_REASONS },
  // Full source data captured so admin can retry without re-running the xlsx parse
  rawData: { type: mongoose.Schema.Types.Mixed, default: {} },
  // Set to true once admin has hand-edited this row into a valid chemical doc
  resolved: { type: Boolean, default: false },
  resolvedBy: { type: String, default: null },
  resolvedAt: { type: Date, default: null },
  resolvedChemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical', default: null },
}, { _id: true });

const ChemicalImportLogSchema = new mongoose.Schema({
  importRunId: { type: String, required: true, unique: true },  // uuid v4
  timestamp: { type: Date, default: Date.now, index: true },
  importedBy: { type: String, required: true },  // userId (superadmin)
  sourceFile: { type: String, required: true },  // filename for audit

  totalRows: { type: Number, required: true },
  insertedCount: { type: Number, default: 0 },
  updatedCount: { type: Number, default: 0 },
  skippedRows: { type: [SkippedRowSchema], default: [] },

  status: {
    type: String,
    enum: ['complete', 'reviewed', 'resolved'],
    default: 'complete',
    index: true,
  },
  reviewedBy: { type: String, default: null },
  reviewedAt: { type: Date, default: null },
});

ChemicalImportLogSchema.statics.SKIP_REASONS = SKIP_REASONS;

module.exports = mongoose.models.ChemicalImportLog
  || mongoose.model('ChemicalImportLog', ChemicalImportLogSchema);
module.exports.SKIP_REASONS = SKIP_REASONS;

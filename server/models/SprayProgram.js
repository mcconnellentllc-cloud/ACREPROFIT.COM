// server/models/SprayProgram.js
//
// schemaVersion 2: passes[] replaces applications[].
// Field rename migration handled in migrations/002_applications_to_passes.js.
//
// Pass chemical row additions:
//   optional       bool   UI shows toggle; off by default unless defaultOn
//   defaultOn      bool   Initial toggle state for optional chemicals
//   conditionNote  str    Plain-English guidance shown next to the toggle
//
// Cap-checker treats optional chemicals as "if toggled on, include in cap
// totals; if off, exclude." No special-casing at the cap-math layer.
//
// containsPendingChemicals is the soft-warn flag. True if any chemical.chemicalId
// resolves to a Chemical doc with status !== 'approved' at save time. Program
// save is allowed; order submit (POST /api/chemical-orders) hard-blocks with 400.

const mongoose = require('mongoose');

const PassChemicalSchema = new mongoose.Schema({
  chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical', required: true },
  // Display label cached for historical recipes — shown if chemical later renamed
  productName: { type: String, default: null },
  rate: { type: Number, required: true, min: 0 },
  rateUnit: { type: String, required: true },  // 'fl oz', 'qt', 'pt', 'dry oz', etc.

  // Optional chemical toggle. Cap-checker sums rates only for chemicals where
  // (!optional || defaultOn === true || farmer toggled on at build time).
  optional: { type: Boolean, default: false },
  defaultOn: { type: Boolean, default: true },  // ignored unless optional=true
  conditionNote: { type: String, default: null },
}, { _id: false });

const PassSchema = new mongoose.Schema({
  passNumber: { type: Number, required: true, min: 1 },
  name: { type: String, required: true },  // "Preplant", "Early Post V3-V5", etc.
  timing: { type: String, default: '' },   // "End of March-April", "30d pre-plant", etc.
  chemicals: { type: [PassChemicalSchema], default: [] },
}, { _id: false });

const SprayProgramSchema = new mongoose.Schema({
  schemaVersion: { type: Number, default: 2, required: true },

  name: { type: String, required: true, trim: true, unique: true },
  crop: { type: String, required: true, index: true },
  tier: {
    type: String,
    enum: ['standard', 'heavy', 'rotation', null],
    default: null,
  },
  description: { type: String, default: '' },

  passes: {
    type: [PassSchema],
    required: true,
    validate: v => Array.isArray(v) && v.length >= 1,
  },

  // Program-level notes surfaced in UI
  rotationNotes: { type: String, default: '' },
  grazingNotes: { type: String, default: '' },
  notes: { type: String, default: '' },

  // Set by save-hook when any referenced chemical has status !== 'approved'.
  // Hard-blocks at /api/chemical-orders submit.
  containsPendingChemicals: { type: Boolean, default: false, index: true },

  // Cap-checker self-assertion result stamped at seed time. Null for user-built
  // recipes until first save triggers the checker. Failing programs won't seed.
  atzSeasonLbPerAcre: { type: Number, default: null },
  capCheckStatus: {
    type: String,
    enum: ['passed', 'failed', 'not_checked', null],
    default: 'not_checked',
  },

  active: { type: Boolean, default: true, index: true },
  createdAt: { type: Date, default: Date.now },
  updatedAt: { type: Date, default: Date.now },
});

SprayProgramSchema.pre('save', async function(next) {
  this.updatedAt = Date.now();
  if (!Array.isArray(this.passes) || this.passes.length === 0) {
    this.containsPendingChemicals = false;
    return next();
  }
  const ids = [];
  for (const pass of this.passes) {
    for (const c of pass.chemicals || []) {
      if (c.chemicalId) ids.push(c.chemicalId);
    }
  }
  if (ids.length === 0) {
    this.containsPendingChemicals = false;
    return next();
  }
  try {
    const Chemical = mongoose.model('Chemical');
    const nonApprovedCount = await Chemical.countDocuments({
      _id: { $in: ids },
      status: { $ne: 'approved' },
    });
    this.containsPendingChemicals = nonApprovedCount > 0;
    next();
  } catch (err) {
    next(err);
  }
});

module.exports = mongoose.models.SprayProgram || mongoose.model('SprayProgram', SprayProgramSchema);

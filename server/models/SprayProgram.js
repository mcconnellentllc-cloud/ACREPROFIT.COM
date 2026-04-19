// server/models/SprayProgram.js
//
// CONSOLIDATED MODEL — Path 2 merge of the pre-existing two definitions:
//   (a) inline schema formerly at server/index.js:1235-1296 (legacy, applications[])
//   (b) file schema formerly here (v2 design, passes[])
//
// Path 2 approach (same pattern as PR #132 Chemical merge): keep BOTH legacy
// and v2 field names populated so existing readers keep working unchanged.
// A pre-save hook mirrors passes[]<->applications[], rate<->suggestedRate,
// active<->isActive, rotationNotes<->rotationRestrictions, grazingNotes<->
// grazingRestrictions. Whichever name a caller writes, the other side is
// derived before the document persists.
//
// Required constraints relaxed: passes[] + applications[] both default to []
// so legacy docs without passes hydrate clean; name + crop stay required
// because those are identity fields.
//
// containsPendingChemicals pre-save check retained — hard-block on order
// submit still works via server/middleware/rejectPendingChemicalOrders.js.

const mongoose = require('mongoose');

// Unified chemical-row sub-schema. Holds BOTH v2 fields (rate, optional,
// defaultOn, conditionNote) and legacy fields (suggestedRate, packSize,
// unit, isAdjuvant, notes). Mirror handled in the parent pre-save hook.
const ProgramChemicalSchema = new mongoose.Schema({
  chemicalId: { type: mongoose.Schema.Types.ObjectId, ref: 'Chemical', default: null },
  productName: { type: String, default: null },

  rate: { type: Number, min: 0, default: null },
  suggestedRate: { type: Number, default: null },
  rateUnit: { type: String, default: null },

  packSize: { type: String, default: null },
  unit: { type: String, default: null },
  isAdjuvant: { type: Boolean, default: false },
  notes: { type: String, default: null },

  optional: { type: Boolean, default: false },
  defaultOn: { type: Boolean, default: true },
  conditionNote: { type: String, default: null },
}, { _id: false });

const ApplicationSchema = new mongoose.Schema({
  name: { type: String, default: '' },
  timing: { type: String, default: '' },
  deliveryWindow: { type: String, default: '' },
  chemicals: { type: [ProgramChemicalSchema], default: [] },
}, { _id: false });

const PassSchema = new mongoose.Schema({
  passNumber: { type: Number, min: 1, default: 1 },
  name: { type: String, default: '' },
  timing: { type: String, default: '' },
  chemicals: { type: [ProgramChemicalSchema], default: [] },
}, { _id: false });

const SprayProgramSchema = new mongoose.Schema({
  schemaVersion: { type: Number, default: 2 },

  name: { type: String, required: true, trim: true },
  crop: { type: String, required: true, index: true },

  tier: { type: String, enum: ['standard', 'heavy', 'rotation', null], default: null },

  roundNumber: { type: Number, default: null },
  type: { type: String, enum: ['suggestion', 'custom', 'template'], default: 'suggestion' },
  isPublic: { type: Boolean, default: false },
  disclaimer: {
    type: String,
    default: 'This is a suggestion only. Each field requires its own evaluation to determine if this chemical program will work for your specific conditions.',
  },

  description: { type: String, default: '' },

  // Dual-populated. passes[] is canonical going forward; applications[]
  // mirrors for the ~14 frontend reads in chemicals.html + calculator.html.
  passes: { type: [PassSchema], default: [] },
  applications: { type: [ApplicationSchema], default: [] },

  rotationNotes: { type: String, default: '' },
  grazingNotes: { type: String, default: '' },
  notes: { type: String, default: '' },

  rotationRestrictions: { type: String, default: '' },
  grazingRestrictions: { type: String, default: '' },
  precautions: { type: [String], default: [] },
  groundType: { type: String, default: '' },

  estimatedCostPerAcre: { type: Number, default: null },
  atzSeasonLbPerAcre: { type: Number, default: null },
  capCheckStatus: {
    type: String,
    enum: ['passed', 'failed', 'not_checked', null],
    default: 'not_checked',
  },
  containsPendingChemicals: { type: Boolean, default: false, index: true },

  isActive: { type: Boolean, default: true },
  active: { type: Boolean, default: true, index: true },

  createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User', default: null },

  createdAt: { type: Date, default: Date.now },
  updatedAt: { type: Date, default: Date.now },
});

// Mirror helper: if caller wrote the v2 name, copy to legacy; vice versa.
// Only mirrors when one side was modified and the other wasn't, so explicit
// writes to both sides win over auto-mirror.
function mirrorIfOneSided(doc, v2Name, legacyName) {
  if (doc.isModified(v2Name) && !doc.isModified(legacyName)) {
    doc[legacyName] = doc[v2Name];
  } else if (doc.isModified(legacyName) && !doc.isModified(v2Name)) {
    doc[v2Name] = doc[legacyName];
  }
}

function mirrorChemicalRates(rows) {
  for (const c of rows || []) {
    if (c.rate != null && c.suggestedRate == null) c.suggestedRate = c.rate;
    else if (c.suggestedRate != null && c.rate == null) c.rate = c.suggestedRate;
    if (c.rateUnit == null && c.unit != null) c.rateUnit = c.unit;
    if (c.unit == null && c.rateUnit != null) c.unit = c.rateUnit;
  }
}

SprayProgramSchema.pre('save', async function(next) {
  this.updatedAt = Date.now();

  mirrorIfOneSided(this, 'active', 'isActive');
  mirrorIfOneSided(this, 'rotationNotes', 'rotationRestrictions');
  mirrorIfOneSided(this, 'grazingNotes', 'grazingRestrictions');

  const hasPasses = Array.isArray(this.passes) && this.passes.length > 0;
  const hasApplications = Array.isArray(this.applications) && this.applications.length > 0;

  if (hasPasses && !hasApplications) {
    this.applications = this.passes.map(p => ({
      name: p.name || '',
      timing: p.timing || '',
      deliveryWindow: '',
      chemicals: (p.chemicals || []).map(c => ({ ...c.toObject ? c.toObject() : c })),
    }));
  } else if (hasApplications && !hasPasses) {
    this.passes = this.applications.map((a, idx) => ({
      passNumber: idx + 1,
      name: a.name || `Pass ${idx + 1}`,
      timing: a.timing || '',
      chemicals: (a.chemicals || []).map(c => ({ ...c.toObject ? c.toObject() : c })),
    }));
  }

  for (const pass of this.passes || []) mirrorChemicalRates(pass.chemicals);
  for (const app of this.applications || []) mirrorChemicalRates(app.chemicals);

  // containsPendingChemicals flag — consulted by order-submit middleware.
  try {
    const ids = [];
    for (const pass of this.passes || []) {
      for (const c of pass.chemicals || []) {
        if (c.chemicalId) ids.push(c.chemicalId);
      }
    }
    if (ids.length === 0) {
      this.containsPendingChemicals = false;
      return next();
    }
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

module.exports = mongoose.models.SprayProgram
  || mongoose.model('SprayProgram', SprayProgramSchema);

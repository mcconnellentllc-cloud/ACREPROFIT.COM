// server/models/Chemical.js
//
// Unified Chemical catalog model. Merges two historical schemas:
//   (1) The inline pricing/compliance catalog from server/index.js:534-723
//       (costPrice/sellPrice/RUP/signalWord/state registrations/etc.)
//   (2) The AI-first MAINCHEM schema introduced in Commit 1
//       (tradeName/activeIngredients[class,lbPerGal]/moaNumbers/status/etc.)
//
// Merge decisions (Path 2, signed off by Kyle):
//   D1 — activeIngredients[] sub-doc UNIONS both shapes. Legacy reads
//        (percentage/poundsPerGallon/casNumber) and AI reads
//        (class/lbPerGal/percentByWeight) coexist. Storage-cheap.
//        Consolidate in a future pass when all call sites are audited.
//   D2 — costPrice/sellPrice/adminPrice are now OPTIONAL. MAINCHEM imports
//        land without pricing; the rejectPendingChemicalOrders middleware
//        blocks orders against any chemical whose status !== 'approved'.
//        Admin workflow: import creates status='pending' docs -> admin adds
//        pricing + flips status='approved' -> orderable. The hard-block
//        middleware is the pricing gate, not a schema-required check.
//   D3 — Dual-name population on import. tradeName/productName, pkg/packSize,
//        uom/unit coexist. Existing 451 references to legacy names unchanged.
//        Planned rename in a future pass with full regression coverage.
//
// Model guard: mongoose.models.Chemical || mongoose.model(...) — safe against
// a second registration attempt during hot-reload or test harness re-imports.

const mongoose = require('mongoose');

const chemicalSchema = new mongoose.Schema({
    // ============ PRODUCT IDENTITY ============
    // Legacy display name. REQUIRED. Populated from MAINCHEM `tradeName` on
    // import (see D3) so existing 451 call sites that read productName continue
    // to work unchanged.
    productName: { type: String, required: true },

    // AI-first trade name. Populated on MAINCHEM import alongside productName.
    // Indexed for new-code lookups.
    tradeName: { type: String, trim: true, index: true },

    sourceSupplier: { type: String }, // D2 relaxed: MAINCHEM imports have no supplier
    manufacturer: { type: String, trim: true },
    supplierId: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },

    // ============ CATEGORY / CROP ============
    category: {
        type: String,
        enum: ['herbicide', 'fungicide', 'insecticide', 'adjuvant', 'fertilizer', 'other'],
        default: 'herbicide'
    },
    crops: [String],

    // AI-first top-level classification (from MAINCHEM "use" column).
    // e.g. HERBICIDE, INSECTICIDE, FUNGICIDE, GLYPHOSATES. Indexed.
    use: { type: String, default: null, index: true },
    // What the product controls — Broadleaves, Grass, Insect, Additive, etc.
    control: { type: String, default: null },

    // ============ PACKAGING ============
    // D3: packSize REQUIRED, populated from MAINCHEM `pkg` on import.
    packSize: { type: String, required: true },
    // D3: unit REQUIRED, populated from MAINCHEM `uom` on import.
    unit: { type: String, required: true },
    // AI-first duplicates. Imported alongside legacy names.
    pkg: { type: String, default: null, trim: true },
    uom: { type: String, default: null, trim: true },
    unitsPerPack: { type: Number },

    // ============ PRICING (3-tier model) ============
    // D2: costPrice/sellPrice/adminPrice are OPTIONAL. Presence gated by the
    // rejectPendingChemicalOrders middleware via status !== 'approved'.
    costPrice: { type: Number },
    adminMarginDollars: { type: Number, default: 0 },
    adminPrice: { type: Number },
    marginDollars: { type: Number, default: 0 },
    sellPrice: { type: Number },
    priceIsSpeculated: { type: Boolean, default: false },
    // Legacy fields (kept for backward compatibility)
    adminMargin: { type: Number, default: 0 },
    regularMargin: { type: Number, default: 0 },
    margin: { type: Number },

    // ============ APPLICATION INFO (program building) ============
    defaultRate: { type: Number },
    rateUnit: { type: String },
    minRate: { type: Number },
    maxRate: { type: Number },

    // ============ VERSION / DATE TRACKING ============
    priceDate: { type: Date, default: Date.now },
    priceVersion: { type: String },

    // ============ COMPARISON / EQUIVALENT DATA ============
    equivalentProduct: String,
    equivalentSupplier: String,
    notes: String,

    // ============ REGULATORY COMPLIANCE ============
    epaRegistrationNumber: String,
    isRestrictedUse: { type: Boolean, default: false },
    rupStates: [String],

    signalWord: {
        type: String,
        enum: ['DANGER', 'DANGER-POISON', 'WARNING', 'CAUTION', 'NONE'],
        default: 'CAUTION'
    },

    hazardClassifications: [{
        type: String,
        enum: [
            'acute_oral_toxicity',
            'acute_dermal_toxicity',
            'acute_inhalation_toxicity',
            'eye_irritant',
            'skin_irritant',
            'skin_sensitizer',
            'carcinogen',
            'reproductive_toxin',
            'environmental_hazard_aquatic',
            'environmental_hazard_bees',
            'groundwater_advisory'
        ]
    }],

    requiredCertifications: [{
        type: String,
        enum: [
            'private_applicator',
            'commercial_applicator',
            'paraquat_training',
            'dicamba_training',
            'fumigant_training'
        ]
    }],

    sdsUrl: String,
    sdsRevisionDate: Date,
    labelUrl: String,
    labelRevisionDate: Date,

    stateRegistrations: [{
        state: { type: String, maxlength: 2 },
        registrationNumber: String,
        expirationDate: Date,
        isRestricted: { type: Boolean, default: false },
        restrictions: String
    }],

    // ============ ACTIVE INGREDIENTS — D1 union sub-doc ============
    // Legacy fields (percentage, poundsPerGallon, casNumber) and AI-first
    // fields (class, lbPerGal, percentByWeight) coexist. Imports populate
    // the AI-first fields; admin UI edits may still write legacy names.
    activeIngredients: [{
        name: { type: String, lowercase: true, trim: true },
        // Legacy shape
        percentage: Number,
        poundsPerGallon: Number,
        casNumber: String,
        // AI-first shape
        class: { type: String, default: null },
        lbPerGal: { type: Number, default: null },
        percentByWeight: { type: Number, default: null }
    }],

    // ============ MAINCHEM SOURCE METADATA ============
    moaNumbers: { type: [Number], default: [] },
    wssa: { type: String, default: null },
    purposes: { type: [String], default: [] },
    aiParseStatus: {
        type: String,
        enum: ['parsed', 'premix_parsed', 'premix_needs_review', 'no_concentration_data', 'unparseable', null],
        default: null
    },
    aiParseReason: { type: String, default: null },
    rawConcentration: { type: String, default: null },
    sourceRow: { type: Number, default: null },

    // ============ APPROVAL GATE ============
    // D2: status is the real pricing/readiness gate. MAINCHEM imports land
    // with 'pending' (no pricing yet). Admin flips to 'approved' after
    // pricing is entered. rejectPendingChemicalOrders middleware blocks
    // order submission against any non-'approved' chemical.
    status: {
        type: String,
        enum: ['approved', 'pending', 'rejected'],
        default: 'pending',
        index: true
    },
    addedBy: { type: String, default: null },

    // ============ DOT / TRANSPORT / STORAGE ============
    dotHazClass: String,
    unNumber: String,
    packingGroup: String,
    storageRequirements: String,
    shelfLifeMonths: Number,

    // ============ MANUFACTURER DETAILS ============
    manufacturerAddress: String,
    manufacturerPhone: String,
    epaEstablishmentNumber: String,

    // ============ COMPLIANCE FLAGS ============
    requiresApplicatorVerification: { type: Boolean, default: false },
    requiresAnnualTraining: { type: Boolean, default: false },
    hasBuyerAgreement: { type: Boolean, default: false },
    isGroundwaterAdvisory: { type: Boolean, default: false },
    hasBufferZoneRequirements: { type: Boolean, default: false },
    bufferZoneDetails: String,
    complianceNotes: String,

    // ============ STATUS / METADATA ============
    isActive: { type: Boolean, default: true },
    availableForOrder: { type: Boolean, default: true },
    createdBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User' },
    createdAt: { type: Date, default: Date.now },
    updatedAt: { type: Date, default: Date.now }
});

// ============ PRE-SAVE HOOK — pricing computation + updatedAt ============
// Preserved verbatim from the original inline schema (server/index.js:683-709).
// updatedAt stamp added from the AI-first schema. Order: updatedAt first,
// pricing computation second so the stamp reflects this save regardless of
// whether pricing fields changed.
chemicalSchema.pre('save', function(next) {
    this.updatedAt = Date.now();

    if (this.costPrice) {
        if (!this.adminPrice && this.adminMarginDollars !== undefined) {
            this.adminPrice = Math.round((this.costPrice + (this.adminMarginDollars || 0)) * 100) / 100;
        } else if (!this.adminPrice) {
            this.adminPrice = this.costPrice;
        }

        if (this.adminMarginDollars === undefined || this.adminMarginDollars === null) {
            this.adminMarginDollars = Math.round((this.adminPrice - this.costPrice) * 100) / 100;
        }

        if ((this.marginDollars === undefined || this.marginDollars === null) && this.sellPrice && this.adminPrice) {
            this.marginDollars = Math.round((this.sellPrice - this.adminPrice) * 100) / 100;
        }

        if (this.sellPrice && this.sellPrice > 0) {
            this.margin = Math.round(((this.sellPrice - this.costPrice) / this.sellPrice) * 100 * 100) / 100;
        }
    }
    next();
});

// ============ INDEXES ============
// 9 legacy indexes preserved verbatim.
chemicalSchema.index({ productName: 1, sourceSupplier: 1, packSize: 1 });
chemicalSchema.index({ sourceSupplier: 1 });
chemicalSchema.index({ supplierId: 1 });
chemicalSchema.index({ category: 1 });
chemicalSchema.index({ crops: 1 });
chemicalSchema.index({ priceDate: -1 });
chemicalSchema.index({ isRestrictedUse: 1 });
chemicalSchema.index({ epaRegistrationNumber: 1 });
chemicalSchema.index({ 'stateRegistrations.state': 1 });
// AI-first indexes — (tradeName, manufacturer) for cross-brand lookups and
// activeIngredients.name for cap-math aggregation.
chemicalSchema.index({ tradeName: 1, manufacturer: 1 });
chemicalSchema.index({ 'activeIngredients.name': 1 });

module.exports = mongoose.models.Chemical || mongoose.model('Chemical', chemicalSchema);

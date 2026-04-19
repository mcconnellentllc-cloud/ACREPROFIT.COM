// server/models/RecipeRating.js
//
// Farmer-facing rating on a spray program (and eventually custom recipes,
// once CustomRecipe model ships — forward-compatible via customRecipeId).
//
// Compound unique index on (ratedBy, sprayProgramId, customRecipeId) enforces
// one rating per farmer per target. Edit-in-place is the flow — duplicate
// submissions return 409 from the router.
//
// Moderation:
//   - profanityFlagged is set by the router's bad-words check; submissions
//     are ALWAYS accepted (false-positive friendly). Superadmin review queue
//     surfaces flagged rows.
//   - isHidden soft-deletes from public view without dropping the row,
//     preserving audit trail (hiddenBy, hiddenAt, hiddenReason).

const mongoose = require('mongoose');

const RecipeRatingSchema = new mongoose.Schema({
  // === target (exactly one populated; pre-validate enforces) ===
  sprayProgramId: { type: mongoose.Schema.Types.ObjectId, ref: 'SprayProgram', index: true, default: null },
  customRecipeId: { type: mongoose.Schema.Types.ObjectId, ref: 'CustomRecipe', index: true, default: null },

  // === rater ===
  ratedBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, index: true },
  ratedByName: { type: String, required: true },

  // === core rating ===
  stars: { type: Number, required: true, min: 1, max: 5 },
  comment: { type: String, trim: true, maxlength: 2000, default: '' },

  // === field conditions ===
  crop: { type: String, trim: true, default: null },
  seasonYear: { type: Number, default: null },
  weatherConditions: { type: String, trim: true, maxlength: 500, default: '' },
  yieldOutcome: { type: String, trim: true, maxlength: 1000, default: '' },

  // === photo (schema reserved; UI disabled in V1) ===
  photoUrl: { type: String, default: null },
  photoFilename: { type: String, default: null },

  // === moderation ===
  isHidden: { type: Boolean, default: false, index: true },
  hiddenBy: { type: mongoose.Schema.Types.ObjectId, ref: 'User', default: null },
  hiddenAt: { type: Date, default: null },
  hiddenReason: { type: String, default: null },

  // === profanity filter result ===
  profanityFlagged: { type: Boolean, default: false, index: true },
  profanityMatches: { type: [String], default: [] },

  // === audit ===
  createdAt: { type: Date, default: Date.now, index: true },
  updatedAt: { type: Date, default: Date.now },
});

// One rating per farmer per target. null sides coexist because MongoDB
// treats missing/null as valid distinct values in compound unique indexes
// at the document level (i.e. rows with the same ratedBy + sprayProgramId +
// null customRecipeId collide, which is the intended dedup behavior).
RecipeRatingSchema.index(
  { ratedBy: 1, sprayProgramId: 1, customRecipeId: 1 },
  { unique: true }
);

RecipeRatingSchema.pre('save', function(next) {
  this.updatedAt = Date.now();
  next();
});

RecipeRatingSchema.pre('validate', function(next) {
  if (!this.sprayProgramId && !this.customRecipeId) {
    return next(new Error('Rating must target a sprayProgramId or customRecipeId'));
  }
  if (this.sprayProgramId && this.customRecipeId) {
    return next(new Error('Rating cannot target both a sprayProgramId and customRecipeId'));
  }
  next();
});

module.exports = mongoose.models.RecipeRating
  || mongoose.model('RecipeRating', RecipeRatingSchema);

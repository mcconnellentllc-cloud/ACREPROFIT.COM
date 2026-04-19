// server/routes/ratings.js
//
// Farmer-facing rating CRUD. Mounted at /api/ratings in server/index.js
// with authMiddleware applied at mount time — every route in this router
// assumes req.user is populated.
//
// Spec: docs/RECIPE_RATING_SPEC (Kyle, 2026-04-19).
//
// Routes:
//   GET    /api/ratings/:sprayProgramId — public list (excludes isHidden)
//                                          + summary + userHasRated flag
//   POST   /api/ratings                  — create (unique per user+target)
//   PUT    /api/ratings/:id              — edit own
//   DELETE /api/ratings/:id              — delete own (or any if superadmin)
//
// Custom-recipe targets intentionally rejected with 400 until CustomRecipe
// model ships — schema is forward-compatible (customRecipeId column exists).

const express = require('express');
const mongoose = require('mongoose');
const RecipeRating = require('../models/RecipeRating');
// Touch the SprayProgram model so the consolidated file registration wins
// even if this router is required before server/index.js's SprayProgram
// require line runs (safe under the model guard).
require('../models/SprayProgram');
const Filter = require('bad-words');

const profanityFilter = new Filter();

function checkProfanity(text) {
  if (!text) return { flagged: false, matches: [] };
  const words = String(text).toLowerCase().match(/\b[a-z]+\b/g) || [];
  const matches = words.filter(w => profanityFilter.isProfane(w));
  return { flagged: matches.length > 0, matches: Array.from(new Set(matches)) };
}

const router = express.Router();

// List ratings for a spray program. Public view — hidden rows excluded.
router.get('/:sprayProgramId', async (req, res) => {
  try {
    const { sprayProgramId } = req.params;
    if (!mongoose.Types.ObjectId.isValid(sprayProgramId)) {
      return res.status(400).json({ ok: false, error: 'invalid_program_id' });
    }

    const ratings = await RecipeRating.find({ sprayProgramId, isHidden: false })
      .sort({ createdAt: -1 })
      .lean();

    const averageStars = ratings.length
      ? Math.round((ratings.reduce((s, r) => s + r.stars, 0) / ratings.length) * 10) / 10
      : null;

    const summary = {
      count: ratings.length,
      averageStars,
      starDistribution: [1, 2, 3, 4, 5].map(n => ({
        stars: n,
        count: ratings.filter(r => r.stars === n).length,
      })),
    };

    const userHasRated = ratings.some(r => String(r.ratedBy) === String(req.user._id));

    res.json({ ok: true, summary, ratings, userHasRated });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Create rating. Unique index on (ratedBy, sprayProgramId, customRecipeId)
// enforces one-per-target; dup insert returns 409.
router.post('/', async (req, res) => {
  try {
    const {
      sprayProgramId,
      customRecipeId,
      stars,
      comment = '',
      crop = null,
      seasonYear = null,
      weatherConditions = '',
      yieldOutcome = '',
      photoUrl = null,
      photoFilename = null,
    } = req.body || {};

    if (customRecipeId) {
      return res.status(400).json({
        ok: false,
        error: 'custom_recipe_ratings_not_yet_supported',
        message: 'Rating custom recipes is coming soon. For now, rate named spray programs.',
      });
    }

    if (!sprayProgramId || !mongoose.Types.ObjectId.isValid(sprayProgramId)) {
      return res.status(400).json({ ok: false, error: 'invalid_program_id' });
    }

    const starsNum = Number(stars);
    if (!Number.isInteger(starsNum) || starsNum < 1 || starsNum > 5) {
      return res.status(400).json({ ok: false, error: 'stars_must_be_1_to_5' });
    }

    const SprayProgram = mongoose.model('SprayProgram');
    const program = await SprayProgram.findById(sprayProgramId).select('_id name').lean();
    if (!program) return res.status(404).json({ ok: false, error: 'program_not_found' });

    const allText = [comment, weatherConditions, yieldOutcome].filter(Boolean).join(' ');
    const { flagged, matches } = checkProfanity(allText);

    const User = mongoose.model('User');
    const userDoc = await User.findById(req.user._id).select('name email').lean();
    const ratedByName = userDoc?.name || userDoc?.email || 'Unknown';

    const rating = await RecipeRating.create({
      sprayProgramId,
      ratedBy: req.user._id,
      ratedByName,
      stars: starsNum,
      comment,
      crop,
      seasonYear: seasonYear ? Number(seasonYear) : null,
      weatherConditions,
      yieldOutcome,
      photoUrl,
      photoFilename,
      profanityFlagged: flagged,
      profanityMatches: matches,
    });

    res.status(201).json({ ok: true, rating });
  } catch (err) {
    if (err.code === 11000) {
      return res.status(409).json({
        ok: false,
        error: 'already_rated',
        message: 'You have already rated this program. Edit your existing rating instead.',
      });
    }
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Edit own rating. 403 if caller is not the owner (admin uses admin router
// for destructive actions; this route is strictly owner-scoped edits).
router.put('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    if (!mongoose.Types.ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, error: 'invalid_rating_id' });
    }
    const rating = await RecipeRating.findById(id);
    if (!rating) return res.status(404).json({ ok: false, error: 'rating_not_found' });
    if (String(rating.ratedBy) !== String(req.user._id)) {
      return res.status(403).json({ ok: false, error: 'can_only_edit_own_rating' });
    }

    const {
      stars,
      comment = rating.comment,
      crop = rating.crop,
      seasonYear = rating.seasonYear,
      weatherConditions = rating.weatherConditions,
      yieldOutcome = rating.yieldOutcome,
      photoUrl = rating.photoUrl,
      photoFilename = rating.photoFilename,
    } = req.body || {};

    if (stars != null) {
      const starsNum = Number(stars);
      if (!Number.isInteger(starsNum) || starsNum < 1 || starsNum > 5) {
        return res.status(400).json({ ok: false, error: 'stars_must_be_1_to_5' });
      }
      rating.stars = starsNum;
    }
    rating.comment = comment;
    rating.crop = crop;
    rating.seasonYear = seasonYear ? Number(seasonYear) : null;
    rating.weatherConditions = weatherConditions;
    rating.yieldOutcome = yieldOutcome;
    rating.photoUrl = photoUrl;
    rating.photoFilename = photoFilename;

    const allText = [comment, weatherConditions, yieldOutcome].filter(Boolean).join(' ');
    const { flagged, matches } = checkProfanity(allText);
    rating.profanityFlagged = flagged;
    rating.profanityMatches = matches;

    await rating.save();
    res.json({ ok: true, rating });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Delete. Owner → 200. Superadmin → 200 on any rating. Everyone else → 403.
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    if (!mongoose.Types.ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, error: 'invalid_rating_id' });
    }
    const rating = await RecipeRating.findById(id);
    if (!rating) return res.status(404).json({ ok: false, error: 'rating_not_found' });

    const isOwnRating = String(rating.ratedBy) === String(req.user._id);
    const isSuperadmin = req.user.role === 'superadmin';

    if (!isOwnRating && !isSuperadmin) {
      return res.status(403).json({ ok: false, error: 'not_authorized' });
    }

    await RecipeRating.deleteOne({ _id: id });
    res.json({ ok: true, deleted: id });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

module.exports = router;

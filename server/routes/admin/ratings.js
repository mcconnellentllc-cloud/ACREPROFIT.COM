// server/routes/admin/ratings.js
//
// Superadmin moderation for farmer ratings. Mounted at /api/admin/ratings
// in server/index.js with authMiddleware + superAdminMiddleware chain —
// scope-specific prefix (same pattern as /api/admin/mainchem) so the
// strict 403 doesn't leak onto the 38 inline /api/admin/* routes that use
// the permissive adminMiddleware.
//
// Routes:
//   GET    /api/admin/ratings               — list (filter: all|flagged|hidden)
//                                             + counts block for nav badge
//   GET    /api/admin/ratings?count=true    — badge-count only (short-circuit)
//   PATCH  /api/admin/ratings/:id/hide      — soft-hide rating
//   PATCH  /api/admin/ratings/:id/unhide    — restore
//   PATCH  /api/admin/ratings/:id/clear-flag — dismiss profanity flag
//   DELETE /api/admin/ratings/:id           — hard-delete (audit preserved
//                                             server-side only via logs)
//
// Spec note reconciled: the draft spec had a duplicate GET '/' handler for
// the count path. Unified here — one GET route that returns the count
// payload when ?count=true, else the full list.

const express = require('express');
const mongoose = require('mongoose');
const RecipeRating = require('../../models/RecipeRating');

const router = express.Router();

// List / count. ?count=true returns the nav-badge payload and nothing else.
router.get('/', async (req, res) => {
  try {
    if (req.query.count === 'true') {
      const flagged = await RecipeRating.countDocuments({ profanityFlagged: true, isHidden: false });
      const hidden = await RecipeRating.countDocuments({ isHidden: true });
      return res.json({ ok: true, flagged, hidden });
    }

    const filter = String(req.query.filter || 'all');
    const limit = Math.min(Number(req.query.limit) || 100, 500);

    const query = {};
    if (filter === 'flagged') query.profanityFlagged = true;
    else if (filter === 'hidden') query.isHidden = true;

    const ratings = await RecipeRating.find(query)
      .populate('sprayProgramId', 'name crop')
      .sort({ createdAt: -1 })
      .limit(limit)
      .lean();

    const flaggedCount = await RecipeRating.countDocuments({ profanityFlagged: true, isHidden: false });
    const hiddenCount = await RecipeRating.countDocuments({ isHidden: true });

    res.json({
      ok: true,
      ratings,
      counts: { flagged: flaggedCount, hidden: hiddenCount, total: ratings.length },
    });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

router.patch('/:id/hide', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ ok: false, error: 'invalid_rating_id' });
    }
    const { reason = null } = req.body || {};
    const rating = await RecipeRating.findByIdAndUpdate(
      req.params.id,
      {
        isHidden: true,
        hiddenBy: req.user._id,
        hiddenAt: new Date(),
        hiddenReason: reason,
      },
      { new: true }
    );
    if (!rating) return res.status(404).json({ ok: false, error: 'rating_not_found' });
    res.json({ ok: true, rating });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

router.patch('/:id/unhide', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ ok: false, error: 'invalid_rating_id' });
    }
    const rating = await RecipeRating.findByIdAndUpdate(
      req.params.id,
      { isHidden: false, hiddenBy: null, hiddenAt: null, hiddenReason: null },
      { new: true }
    );
    if (!rating) return res.status(404).json({ ok: false, error: 'rating_not_found' });
    res.json({ ok: true, rating });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

router.patch('/:id/clear-flag', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ ok: false, error: 'invalid_rating_id' });
    }
    const rating = await RecipeRating.findByIdAndUpdate(
      req.params.id,
      { profanityFlagged: false, profanityMatches: [] },
      { new: true }
    );
    if (!rating) return res.status(404).json({ ok: false, error: 'rating_not_found' });
    res.json({ ok: true, rating });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

router.delete('/:id', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ ok: false, error: 'invalid_rating_id' });
    }
    const result = await RecipeRating.findByIdAndDelete(req.params.id);
    if (!result) return res.status(404).json({ ok: false, error: 'rating_not_found' });
    res.json({ ok: true, deleted: req.params.id });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

module.exports = router;

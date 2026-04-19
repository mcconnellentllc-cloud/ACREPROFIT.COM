// server/routes/admin/spray-programs.js
//
// CRUD router for admin-managed SprayProgram docs. Mounted at
// /api/admin/spray-programs in server/index.js with authMiddleware +
// superAdminMiddleware (scope-specific prefix — same pattern as
// /api/admin/mainchem).
//
// Writes target the canonical v2 shape (passes[]). The SprayProgram pre-save
// hook mirrors passes[]<->applications[] so the 14+ frontend reads of
// .applications[].chemicals[].suggestedRate in chemicals.html + calculator.html
// keep working without a frontend change.
//
// Routes:
//   GET    /api/admin/spray-programs          — list (all programs)
//   GET    /api/admin/spray-programs/:id      — single
//   POST   /api/admin/spray-programs          — create
//   PUT    /api/admin/spray-programs/:id      — update
//   DELETE /api/admin/spray-programs/:id      — delete
//   GET    /api/admin/spray-programs/_chemicals
//          — approved-chemical picker feed for the admin form

const express = require('express');
const mongoose = require('mongoose');

const SprayProgram = require('../../models/SprayProgram');
const Chemical = require('../../models/Chemical');

const router = express.Router();

// List all programs (admin view — no public/isActive filter).
router.get('/', async (req, res) => {
  try {
    const { crop, tier } = req.query;
    const q = {};
    if (crop) q.crop = crop;
    if (tier) q.tier = tier;
    const programs = await SprayProgram.find(q)
      .populate('createdBy', 'name email')
      .sort({ crop: 1, name: 1 });
    res.json({ programs, count: programs.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Approved-chemical picker. Returns the minimal shape the admin form needs
// to populate the chemical-row dropdown. Filters to status='approved' so
// pending MAINCHEM imports don't appear until admin promotes them.
router.get('/_chemicals', async (req, res) => {
  try {
    const chemicals = await Chemical.find({ status: 'approved' })
      .select('_id productName tradeName manufacturer packSize pkg unit uom')
      .sort({ productName: 1 })
      .lean();
    res.json({ chemicals, count: chemicals.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Single program.
router.get('/:id', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ error: 'invalid id' });
    }
    const program = await SprayProgram.findById(req.params.id)
      .populate('createdBy', 'name email');
    if (!program) return res.status(404).json({ error: 'not found' });
    res.json(program);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create. Body accepts both passes[] (v2, preferred) and applications[]
// (legacy) — whichever is sent, the pre-save hook mirrors to the other.
router.post('/', async (req, res) => {
  try {
    const body = req.body || {};
    if (!body.name || !body.crop) {
      return res.status(400).json({ error: 'name and crop are required' });
    }
    const program = new SprayProgram({
      ...body,
      type: body.type || 'suggestion',
      createdBy: req.user?._id || null,
    });
    await program.save();
    res.status(201).json(program);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// Update. Whitelist the fields admin is allowed to overwrite so an
// accidental `createdAt` or `createdBy` in the payload doesn't clobber.
const UPDATABLE_FIELDS = [
  'name', 'description', 'crop', 'tier', 'roundNumber', 'type', 'isPublic',
  'disclaimer', 'passes', 'applications', 'precautions', 'groundType',
  'rotationNotes', 'grazingNotes', 'rotationRestrictions', 'grazingRestrictions',
  'notes', 'estimatedCostPerAcre', 'atzSeasonLbPerAcre', 'capCheckStatus',
  'isActive', 'active',
];

router.put('/:id', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ error: 'invalid id' });
    }
    const program = await SprayProgram.findById(req.params.id);
    if (!program) return res.status(404).json({ error: 'not found' });

    for (const field of UPDATABLE_FIELDS) {
      if (req.body[field] !== undefined) program[field] = req.body[field];
    }
    await program.save();
    res.json(program);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

router.delete('/:id', async (req, res) => {
  try {
    if (!mongoose.Types.ObjectId.isValid(req.params.id)) {
      return res.status(400).json({ error: 'invalid id' });
    }
    const result = await SprayProgram.findByIdAndDelete(req.params.id);
    if (!result) return res.status(404).json({ error: 'not found' });
    res.json({ ok: true, deleted: req.params.id });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;

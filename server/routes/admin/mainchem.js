// server/routes/admin/mainchem.js
//
// POST /api/admin/mainchem/import
//   Reads server/seed/mainchem-source.json (pre-parsed from the xlsx by the
//   one-shot scripts/parse-mainchem.js), upserts chemicals into the catalog,
//   persists a ChemicalImportLog doc with all skip-report rows.
//
// Idempotent. Re-running updates existing docs matched by (tradeName, manufacturer).
//
// Response: inline skip report (for UI toast) AND persistent log doc (for
// history page). Delivery = (c) both, per locked spec.
//
// Auth: requires superadmin. Uses existing middleware chain — mount point
// registered in server/index.js alongside other admin routes.

const express = require('express');
const { randomUUID } = require('crypto');
const fs = require('fs');
const path = require('path');

const Chemical = require('../../models/Chemical');
const ChemicalImportLog = require('../../models/ChemicalImportLog');

const router = express.Router();

// Assumes requireSuperadmin middleware is applied upstream at mount point:
//   app.use('/api/admin', authMiddleware, superAdminMiddleware, adminRouter);
// which matches the existing pattern for admin routes in server/index.js.

router.post('/mainchem/import', async (req, res) => {
  const sourceFile = 'mainchem-source.json';
  const sourcePath = path.join(__dirname, '..', '..', 'seed', sourceFile);

  if (!fs.existsSync(sourcePath)) {
    return res.status(500).json({
      ok: false,
      error: 'seed source not found',
      path: sourcePath,
      hint: 'run `node scripts/parse-mainchem.js <xlsx> server/seed/mainchem-source.json` first',
    });
  }

  let source;
  try {
    source = JSON.parse(fs.readFileSync(sourcePath, 'utf8'));
  } catch (e) {
    return res.status(500).json({ ok: false, error: 'seed source malformed', detail: e.message });
  }

  const importRunId = randomUUID();
  const importedBy = req.user?.id || req.user?._id?.toString() || 'unknown';
  let insertedCount = 0;
  let updatedCount = 0;
  const errors = [];

  for (const rec of (source.records || [])) {
    try {
      const filter = {
        tradeName: rec.tradeName,
        manufacturer: rec.manufacturer || null,
      };
      const update = {
        $set: {
          tradeName: rec.tradeName,
          manufacturer: rec.manufacturer || null,
          pkg: rec.pkg || null,
          uom: rec.uom || null,
          use: rec.use || null,
          control: rec.control || null,
          activeIngredients: rec.activeIngredients || [],
          moaNumbers: rec.moaNumbers || [],
          wssa: rec.wssa || null,
          purposes: rec.purposes || [],
          aiParseStatus: rec.aiParseStatus || null,
          aiParseReason: rec.aiParseReason || null,
          rawConcentration: rec.rawConcentration || null,
          sourceRow: rec.sourceRow || null,
          updatedAt: new Date(),
        },
        $setOnInsert: {
          // Approval state and audit-origin are admin-controlled decisions.
          // They only get set on first-time insert — re-imports refresh
          // pricing/concentration/MOA fields but respect admin decisions on
          // status (e.g. rejecting a discontinued product) and audit origin
          // (e.g. a farmer-suggested chemical later gaining a MAINCHEM entry
          // keeps the suggester's userId, not overwritten to 'mainchem-import').
          createdAt: new Date(),
          status: 'approved',
          addedBy: 'mainchem-import',
        },
      };
      const result = await Chemical.updateOne(filter, update, { upsert: true });
      if (result.upsertedCount && result.upsertedCount > 0) {
        insertedCount++;
      } else if (result.modifiedCount && result.modifiedCount > 0) {
        updatedCount++;
      }
    } catch (err) {
      errors.push({ tradeName: rec.tradeName, error: err.message });
    }
  }

  // Persist skip-report log. Reasons are enum-validated at write time.
  const logDoc = await ChemicalImportLog.create({
    importRunId,
    importedBy,
    sourceFile: source.sourceFile || sourceFile,
    totalRows: source.totalRows || 0,
    insertedCount,
    updatedCount,
    skippedRows: (source.skipped || []).map(s => ({
      rowIndex: s.rowIndex,
      tradeName: s.tradeName || null,
      reason: s.reason,
      rawData: s.rawData || {},
      resolved: false,
    })),
    status: 'complete',
  });

  res.json({
    ok: true,
    importRunId,
    logId: logDoc._id,
    insertedCount,
    updatedCount,
    skippedCount: (source.skipped || []).length,
    errors,
    skippedByReason: countByReason(source.skipped || []),
    skipped: source.skipped || [],  // inline so admin UI can render without a second fetch
  });
});

// GET /api/admin/mainchem/logs — list historical import runs
router.get('/mainchem/logs', async (req, res) => {
  const logs = await ChemicalImportLog.find({})
    .sort({ timestamp: -1 })
    .limit(50)
    .select('-skippedRows.rawData')  // omit heavy raw data from list view
    .lean();
  res.json({ ok: true, logs });
});

// GET /api/admin/mainchem/logs/:id — full detail with skipped rows
router.get('/mainchem/logs/:id', async (req, res) => {
  const log = await ChemicalImportLog.findById(req.params.id).lean();
  if (!log) return res.status(404).json({ ok: false, error: 'log not found' });
  res.json({ ok: true, log });
});

// PATCH /api/admin/mainchem/logs/:id/skipped/:skipId/resolve
// Admin marks a skipped row as resolved (typically after hand-creating the
// Chemical doc via the main catalog editor).
router.patch('/mainchem/logs/:id/skipped/:skipId/resolve', async (req, res) => {
  const { resolvedChemicalId } = req.body;
  const log = await ChemicalImportLog.findById(req.params.id);
  if (!log) return res.status(404).json({ ok: false, error: 'log not found' });
  const skip = log.skippedRows.id(req.params.skipId);
  if (!skip) return res.status(404).json({ ok: false, error: 'skip row not found' });
  skip.resolved = true;
  skip.resolvedBy = req.user?.id || 'unknown';
  skip.resolvedAt = new Date();
  if (resolvedChemicalId) skip.resolvedChemicalId = resolvedChemicalId;
  await log.save();
  res.json({ ok: true, skip });
});

// PATCH /api/admin/mainchem/logs/:id/status
router.patch('/mainchem/logs/:id/status', async (req, res) => {
  const { status } = req.body;
  if (!['complete', 'reviewed', 'resolved'].includes(status)) {
    return res.status(400).json({ ok: false, error: 'invalid status' });
  }
  const log = await ChemicalImportLog.findByIdAndUpdate(
    req.params.id,
    {
      status,
      reviewedBy: req.user?.id || null,
      reviewedAt: new Date(),
    },
    { new: true }
  );
  if (!log) return res.status(404).json({ ok: false, error: 'log not found' });
  res.json({ ok: true, log });
});

function countByReason(skipped) {
  const counts = {};
  skipped.forEach(s => { counts[s.reason] = (counts[s.reason] || 0) + 1; });
  return counts;
}

module.exports = router;

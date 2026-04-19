// server/middleware/rejectPendingChemicalOrders.js
//
// Hard-block at order submit. The locked integrity gate per spec:
//   save-time: soft warn (SprayProgram.containsPendingChemicals flag)
//   submit-time: HARD BLOCK (this middleware)
//
// Applied in server/index.js upstream of the chemical-orders POST handler:
//   app.post('/api/chemical-orders', authMiddleware, rejectPendingChemicalOrders, createOrder);
//
// Rejects with 400 if ANY line item references a Chemical doc whose status is
// not 'approved'. Returns the offending chemicals so the frontend can display
// per-line validation errors instead of a generic "submit failed".

const mongoose = require('mongoose');
const Chemical = require('../models/Chemical');

async function rejectPendingChemicalOrders(req, res, next) {
  const items = req.body?.items || req.body?.lineItems || [];
  if (!Array.isArray(items) || items.length === 0) {
    return next();  // empty payload — let downstream validator handle it
  }

  const ids = items
    .map(i => i.chemicalId || i.chemical_id)
    .filter(Boolean)
    .filter(id => mongoose.Types.ObjectId.isValid(id));

  if (ids.length === 0) {
    return next();
  }

  let chems;
  try {
    chems = await Chemical.find({ _id: { $in: ids } })
      .select('_id tradeName status')
      .lean();
  } catch (err) {
    return next(err);
  }

  const offenders = chems.filter(c => c.status !== 'approved');
  if (offenders.length > 0) {
    return res.status(400).json({
      ok: false,
      error: 'order_contains_pending_chemicals',
      message: 'Order cannot be submitted — contains chemicals that are not yet approved.',
      offenders: offenders.map(c => ({
        chemicalId: c._id,
        tradeName: c.tradeName,
        status: c.status,
      })),
    });
  }

  // Also reject if any line references an ID that doesn't exist in the catalog —
  // would otherwise pass silently and create an order with orphaned chemicalId.
  const foundIds = new Set(chems.map(c => String(c._id)));
  const missing = ids.filter(id => !foundIds.has(String(id)));
  if (missing.length > 0) {
    return res.status(400).json({
      ok: false,
      error: 'order_contains_unknown_chemicals',
      message: 'Order references chemicals that do not exist in the catalog.',
      missingIds: missing,
    });
  }

  next();
}

module.exports = rejectPendingChemicalOrders;

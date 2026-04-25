# Post-C3 Verification Checklist

Items surfaced during the Decimal128 migration (C1a–C3) that aren't blockers
for the migration itself but warrant a deliberate check after C3 deploys.
Run through this list once when the data migration lands, then archive.

## Open items

- **`totalAmount` vs `total` field-name consistency.** `dashboard.html` reduces
  on `o.totalAmount` from the orders endpoint, but the Invoice schema field
  is `total` (`server/index.js:1970`). The orders endpoint may serve
  `totalAmount` as an alias, or the dashboard may be reading a stale field
  that silently returns `undefined` (and falls back to `0` via `|| 0`). Verify
  field naming by:
  1. Calling `GET /api/orders` (admin role) and inspecting the response shape.
  2. Calling `GET /api/orders` (distributor role) and inspecting the response
     shape — same field?
  3. Confirming `dashboard.html` lines `~553`, `~584` produce non-zero
     `totalRevenue`/`totalSales` when orders with non-zero `total` exist.
  Fix path if drift exists: align dashboard to the canonical field name OR
  add an explicit alias on the orders serializer.

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

- **`chemicals.html` site #6 — DOM textContent parseFloat round-trip
  (line ~2308 post-C1f).** `recalcPackageTotals` reads the Hydrovant row's
  rendered currency string out of the DOM, strips `$`/`,` with regex, and
  `parseFloat`s to a Number for accumulator math. Survives post-C3 because
  `fmtMoney`/`fmtMoneyLocale` output matches the strip regex. Out of scope
  for the migration — deliberately left as-is to avoid architectural debt
  cleanup creeping into the forwards-compat work. Future refactor
  opportunity: compute `hydroCost` directly from the closure at the render
  site (line ~2921) and accumulate that value, not a DOM re-parse. One-day
  cleanup, do it when you touch this flow for another reason.

- **`chemicals.html` `packageEdits` Map post-C3 smoke test.** Open a
  program with a chemical priced at a non-round value (e.g. `$12.755/gal`)
  where Decimal128 precision matters more than Number's. Trigger a package
  edit (change qty or SKU). Verify the grand-total in the footer and the
  cost-per-acre cell compute correctly and display the expected rounded
  values. Confirms the seed-time `Number(item.price)` coercion (chemicals.html
  ~line 2848) and the `onSkuChange` Number coerce (~line 2281) both keep
  the Map Number-typed end-to-end so downstream Number arithmetic at
  `recalcPackageTotals` and per-row `lineCost`/`overageCost` stays correct.

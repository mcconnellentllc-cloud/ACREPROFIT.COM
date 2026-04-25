# Post-C3 Verification Checklist

Two sections:
- **C3 prerequisites** — work that MUST land before C3 deploys, otherwise
  precision-loss bugs ship to production. Convert these sites to Decimal
  arithmetic as part of the C3 commit.
- **Verification items** — checks to run once after C3 lands.

## C3 prerequisites — server-side native arithmetic on Decimal128 fields

Three sites do native JS arithmetic on values that will be Decimal128
post-C3. JS coerces Decimal128 via `valueOf()` → string → Number → arithmetic
→ result is a Number with float drift. For typical 2dp money values this is
within-cent accuracy on a single op, but compounds across reduces and
multi-operand expressions. **Rewrite all three in Decimal before C3
deploys.**

1. **`server/index.js:2442-2455`** (`autoGenerateInvoice` invoice items
   build). Currently:
   ```js
   const unitPrice = item.unitPrice || item.pricePerUnit || chem?.sellPrice || 0;
   const totalPrice = qty * unitPrice;
   const margin = (unitPrice - costPrice) * qty;
   ```
   Rewrite both `totalPrice` and `margin` in Decimal. Output stays as a
   Decimal128 instance assigned into the new Invoice doc so the schema
   pre('validate') hook coerces correctly on save.

2. **`server/index.js:5499-5538`** (checkout `normalizedChemicals` map).
   Currently:
   ```js
   totalPrice: Math.round(qty * pricePerUnit * 100) / 100
   ```
   Rewrite in Decimal. Store as Decimal128 → Order doc for save.

3. **`server/index.js:8138-8147`** (margin report). Currently
   `profitPerUnit: Math.round((c.sellPrice - c.costPrice) * 100) / 100`.
   Has a defensive `Number(...)` cast added in C2 so it won't crash post-C3,
   but it loses precision via the cast. Rewrite in Decimal:
   `profitPerUnit: serializeMoney(new Decimal(c.sellPrice).minus(c.costPrice))`.

These three sites cannot deploy after C3 in their current form. Treat as
hard prerequisites, not future-checklist items.

## Verification items

- **C2 manual-pick site coverage.** During C2 implementation, the survey
  under-counted manual-pick leak sites (reported 2, found 10). All 10 are
  now wrapped with `serializeMoney`. After C3 deploys, spot-check the
  responses of the affected routes to confirm money fields render as
  strings (not raw `{$numberDecimal: ...}` extJSON):
  `GET /api/chemicals/public`, `GET /api/chemicals` (admin variant at
  ~7886), the chemicals listing at ~7972, the customer-facing list at
  ~8121, the margin report at ~8138, the chemical-archive search at ~8525,
  the admin chemicals page at ~11955, the inventory placeholder route at
  ~13240, the admin inventory listing at ~14080, and the distributor
  custom-pricing list at ~16155. Survey method: hit each endpoint, grep
  the response for `$numberDecimal` — should return zero matches.

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

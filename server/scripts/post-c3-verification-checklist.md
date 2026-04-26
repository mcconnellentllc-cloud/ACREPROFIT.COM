# Post-C3 Verification Checklist

Two sections:
- **C3 prerequisites** — work that MUST land before C3 deploys, otherwise
  precision-loss bugs ship to production. Convert these sites to Decimal
  arithmetic as part of the C3 commit.
- **Verification items** — checks to run once after C3 lands.

## C3 prerequisites — RESOLVED in C3a

The three native-arithmetic sites flagged in C2 as "must convert before C3
deploys" were converted as part of C3a. Plus the broader sweep of ~67
math sites total. C3a result: server computes in Decimal, writes back as
`Number(dec.toFixed(dp))` at the schema/wire boundary. Schemas remain
Number-typed, wire format unchanged.

When C3b lands, all the math is already correct; C3b only flips the
storage type and seeds the backfill.

## Future cleanup items (not blocking C3)

- **`chemicalQuoteSchema.virtual('packPriceCalculated')` (server/index.js:891)
  is dead code.** Single reference is the definition itself; no callers
  anywhere in the repo (verified by grep). C3a left the body unchanged
  (`Math.round(this.pricePerUnit * this.unitsPerPack * 100) / 100`) per
  the "don't refactor dead code" rule. Remove the virtual entirely in a
  separate cleanup commit after C3 lands. If a future feature needs a
  computed pack price, it can be reintroduced with the proper Decimal
  shape at that time.

- **Sequential-round vs round-once-at-end pattern (lines ~12609-12610
  and ~12648-12649 post-C3a, line numbers will shift again at C3b).**
  Two pairs of sites in chemicals.html submit-order validation round
  the per-package price first, THEN multiply by package count and round
  again. Pre-C3a behavior was sequential rounding; C3a preserves it
  literally for parity safety. The mathematically more accurate approach
  is to compute the full chain in Decimal and round once at the end —
  preserves more precision but produces different penny-level results
  on long product chains. **If round-once-at-end is the better long-term
  math, ship as a separate, intentional commit with documented behavior
  change and stakeholder visibility — not as a hidden side effect of
  the migration.**

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

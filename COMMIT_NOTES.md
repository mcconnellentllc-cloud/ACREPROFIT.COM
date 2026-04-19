# Commit 1 — AI-first schema + MAINCHEM import route + 5 corn programs

Branch: `claude/ai-first-corn-programs`

## What shipped

| File | Purpose |
|---|---|
| `scripts/parse-mainchem.js` | One-shot xlsx parser. Run locally by Kyle — emits `mainchem-source.json`. NOT in this commit; generated post-merge. |
| `server/models/Chemical.js` | `activeIngredients[]` + `status` + audit fields + `aiParseStatus` |
| `server/models/SprayProgram.js` | `passes[]` + `schemaVersion: 2` + optional-chemical toggles + `containsPendingChemicals` pre-save hook |
| `server/models/ChemicalImportLog.js` | Persisted import history with enum'd skip reasons + status workflow |
| `server/migrations/002_applications_to_passes.js` | Idempotent field rename — Milo Hardy/Economical stay functional |
| `server/routes/admin/mainchem.js` | POST `/api/admin/mainchem/import` + list/detail/resolve/status endpoints |
| `server/middleware/rejectPendingChemicalOrders.js` | Hard-block at `/api/chemical-orders` submit |
| `server/lib/atrazineCap.js` | Shared cap-math — unit normalization, per-AI extraction, cap assertion |
| `server/seed/corn-programs-data.js` | Pure data module — 5 corn programs, no DB deps |
| `server/seed/corn-programs-2026.js` | Seed runner — resolves chemicals, asserts caps, writes docs |
| `test/atrazine-cap.test.js` | 15 tests. CI safety rail. |

**Not in this commit:** `server/seed/mainchem-source.json` (parser output, ~500KB of 721 parsed chemical records). Kyle generates locally against source xlsx and commits as a follow-up artifact commit. The import route at `POST /api/admin/mainchem/import` will return 500 with `seed source not found` + instructions until that file is in place. Schema, migration, corn programs seed, and hard-block middleware are all functional without it.

## Seed JSON — deferred to follow-up commit

Reason: the source xlsx lives on Kyle's sandbox as the source-of-truth spreadsheet. Parser runs locally, not on Render. Data artifact shouldn't gate the code merge.

Post-merge step for Kyle:
```bash
npm install --save-dev xlsx
node scripts/parse-mainchem.js /path/to/2025_SPRAY_RECORDS.xlsx server/seed/mainchem-source.json
git add server/seed/mainchem-source.json
git commit -m "data: initial MAINCHEM seed JSON (721 chemicals, 2 dupes merged)"
git push
```

## Atrazine cap verification (run at seed time, fail-fast)

Every program targets exactly 2.0 lb ai/A spring total, safely under the 2.2 annual cap:

```
Corn Dryland Standard        total=2.0  per-pass=[{1:1}, {2:1}]
Corn Dryland Heavy           total=2.0  per-pass=[{1:0}, {2:1}, {3:1}]
Corn Irrigated Standard      total=2.0  per-pass=[{1:1}, {2:1}]
Corn Irrigated Heavy         total=2.0  per-pass=[{1:0}, {2:2}, {3:0}]
Corn Post-Wheat Rotation     total=2.0  per-pass=[{1:0}, {2:0}, {3:1}, {4:1}]
```

`assertAtzCaps()` throws with program name + pass number + computed values if any pass exceeds single-app (2.0) or program total exceeds annual (2.2).

## Deploy sequence (after push + review)

```bash
# 1. Mount new admin route + hard-block middleware in server/index.js (manual)

# 2. Migrate existing SprayProgram docs
node server/migrations/002_applications_to_passes.js --dry-run --verbose
node server/migrations/002_applications_to_passes.js

# 3. Generate + commit seed JSON (Kyle runs locally)
npm install --save-dev xlsx
node scripts/parse-mainchem.js /path/to/spray_records.xlsx server/seed/mainchem-source.json
git add server/seed/mainchem-source.json && git commit -m "data: MAINCHEM seed" && git push

# 4. Import MAINCHEM catalog via admin UI button or curl
curl -X POST https://acreprofit-api.onrender.com/api/admin/mainchem/import -b "cookie=..."

# 5. Seed 5 corn programs
node server/seed/corn-programs-2026.js --dry-run
node server/seed/corn-programs-2026.js
```

## Required manual wiring in server/index.js

```js
// Mount the admin route upstream of existing superadmin chain:
const mainchemRouter = require('./routes/admin/mainchem');
app.use('/api/admin', requireSuperadmin, mainchemRouter);

// Hard-block middleware on order submit:
const rejectPendingChemicalOrders = require('./middleware/rejectPendingChemicalOrders');
app.post('/api/chemical-orders', requireAuth, rejectPendingChemicalOrders, createChemicalOrderHandler);
```

Line numbers will depend on the existing index.js layout — paste the relevant section for an exact diff.

## Review-driven fixes applied this commit

1. MAINCHEM import sets `status: 'approved'` via `$setOnInsert` (admin decisions survive re-imports)
2. `SprayProgram.containsPendingChemicals` pre-save hook resolves chemical statuses async
3. `rejectPendingChemicalOrders.js` wraps `Chemical.find()` in try/catch
4. `Chemical` compound index flipped to `unique: true` — parser dedupes 2 known source dupes (FOAM BUSTER/HELENA, ROZOL PD BAIT/LIPHAT) with non-null-preference merge
5. Migration 002 preserves top-level `doc.timing` into `passes[0].timing` before `$unset`

## Deferred to future commits

- ChemicalImportLog `skippedRows` → `rowReviewQueue` rename (UX, commit 2)
- Runtime cap-check farmer-toggle override (runtime-validator commit)
- `ChemicalImportLog.importedBy` → `ObjectId ref: 'User'` (admin UI commit)
- Frontend import log admin UI with inline edit (commit 2 — confirmed as next priority)
- Suggest-chemical modal + admin pending review page (commit 3)
- Sorghum/milo programs from 2026 B.E. Precise PDF (commit 4)

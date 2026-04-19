// server/test/rating.test.js
//
// Rating system test suite — 15 cases per spec.
//
// Runs under node --test with supertest against a minimal Express app that
// mounts the two rating routers. JWT is replaced by a tiny auth-emulator
// middleware that stubs req.user from a custom header.
//
// Backing MongoDB: prefer MONGO_TEST_URI env var (points at any mongod —
// local dev container, CI service, Atlas test cluster); fall back to
// mongodb-memory-server when no URI is set. mongodb-memory-server downloads
// a mongod binary on first run, which requires outbound internet —
// sandboxes that block fastdl.mongodb.org will see the fallback fail.
// In that case export MONGO_TEST_URI=mongodb://localhost:27017/rating_test
// and re-run.
//
// Run from server/:
//   node --test test/rating.test.js
// or:
//   MONGO_TEST_URI=mongodb://localhost:27017/rating_test node --test test/rating.test.js

const { test, before, after, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const express = require('express');
const mongoose = require('mongoose');
const request = require('supertest');

let mongod;
let app;
let User;
let SprayProgram;
let RecipeRating;

// Test actors. Seeded in before() and reused across tests.
let farmerA, farmerB, superadmin, program;

// Emulated auth. Tests set X-Test-User header to a user _id to impersonate.
function testAuth(req, res, next) {
  const uid = req.header('X-Test-User');
  if (!uid) return res.status(401).json({ ok: false, error: 'unauthenticated' });
  User.findById(uid).lean().then(user => {
    if (!user) return res.status(401).json({ ok: false, error: 'user_not_found' });
    req.user = user;
    next();
  }).catch(err => res.status(500).json({ error: err.message }));
}

function testSuperadmin(req, res, next) {
  if (req.user.role !== 'superadmin') return res.status(403).json({ error: 'forbidden' });
  next();
}

before(async () => {
  let uri;
  if (process.env.MONGO_TEST_URI) {
    uri = process.env.MONGO_TEST_URI;
  } else {
    const { MongoMemoryServer } = require('mongodb-memory-server');
    mongod = await MongoMemoryServer.create();
    uri = mongod.getUri();
  }
  await mongoose.connect(uri, { dbName: 'rating_test' });

  // Minimal User model (real app's User schema is in server/index.js and
  // pulls in too much; we just need _id + role + name for the router).
  User = mongoose.model('User', new mongoose.Schema({
    name: String,
    email: String,
    role: { type: String, default: 'customer' },
  }));

  SprayProgram = require('../models/SprayProgram');
  RecipeRating = require('../models/RecipeRating');

  // Seed actors + a program to rate.
  farmerA = await User.create({ name: 'Farmer A', email: 'a@test.com', role: 'customer' });
  farmerB = await User.create({ name: 'Farmer B', email: 'b@test.com', role: 'customer' });
  superadmin = await User.create({ name: 'Kyle', email: 'k@test.com', role: 'superadmin' });
  program = await SprayProgram.create({
    name: 'Test Corn Standard',
    crop: 'corn',
    passes: [{ passNumber: 1, name: 'Burndown', timing: 'Pre-plant', chemicals: [] }],
  });

  // Wire Express app identical to server/index.js mount chain.
  const ratingsRouter = require('../routes/ratings');
  const adminRatingsRouter = require('../routes/admin/ratings');
  app = express();
  app.use(express.json());
  app.use('/api/ratings', testAuth, ratingsRouter);
  app.use('/api/admin/ratings', testAuth, testSuperadmin, adminRatingsRouter);
});

after(async () => {
  await mongoose.disconnect();
  if (mongod) await mongod.stop();
});

// Clear ratings between tests so the unique-index cases don't leak.
beforeEach(async () => {
  await RecipeRating.deleteMany({});
});

// ============ TEST 1-6 — CREATE ============

test('1. Create rating as logged-in user — success (201)', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 5, comment: 'Worked great on my dryland.' });
  assert.equal(res.status, 201);
  assert.equal(res.body.ok, true);
  assert.equal(res.body.rating.stars, 5);
  assert.equal(res.body.rating.ratedByName, 'Farmer A');
});

test('2. Create rating with stars=6 → 400', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 6 });
  assert.equal(res.status, 400);
  assert.equal(res.body.error, 'stars_must_be_1_to_5');
});

test('3. Create rating with stars=0 → 400', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 0 });
  assert.equal(res.status, 400);
  assert.equal(res.body.error, 'stars_must_be_1_to_5');
});

test('4. Create rating twice for same program → 409 (unique index)', async () => {
  await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 4 });
  const res = await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 3 });
  assert.equal(res.status, 409);
  assert.equal(res.body.error, 'already_rated');
});

test('5. Create rating with customRecipeId → 400 (not yet supported)', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ customRecipeId: new mongoose.Types.ObjectId().toString(), stars: 4 });
  assert.equal(res.status, 400);
  assert.equal(res.body.error, 'custom_recipe_ratings_not_yet_supported');
});

test('6. Create rating for non-existent program → 404', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: new mongoose.Types.ObjectId().toString(), stars: 5 });
  assert.equal(res.status, 404);
  assert.equal(res.body.error, 'program_not_found');
});

// ============ TEST 7-8 — PROFANITY FILTER ============

test('7. Profanity check: input with known profanity → flagged=true', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 4, comment: 'This shit worked fine.' });
  assert.equal(res.status, 201);
  assert.equal(res.body.rating.profanityFlagged, true);
  assert.ok(res.body.rating.profanityMatches.length > 0);
});

test('8. Profanity check: clean input → flagged=false', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 4, comment: 'Solid program, strong early control.' });
  assert.equal(res.status, 201);
  assert.equal(res.body.rating.profanityFlagged, false);
  assert.deepEqual(res.body.rating.profanityMatches, []);
});

// ============ TEST 9 — AUTH ============

test('9. Create rating without auth middleware (no header) → 401', async () => {
  const res = await request(app)
    .post('/api/ratings')
    .send({ sprayProgramId: program._id.toString(), stars: 4 });
  assert.equal(res.status, 401);
});

// ============ TEST 10-11 — EDIT ============

test('10. Edit own rating → 200', async () => {
  const create = await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 3 });
  const id = create.body.rating._id;
  const res = await request(app).put(`/api/ratings/${id}`)
    .set('X-Test-User', farmerA._id.toString())
    .send({ stars: 5, comment: 'Upgraded opinion after harvest.' });
  assert.equal(res.status, 200);
  assert.equal(res.body.rating.stars, 5);
});

test('11. Edit someone else\'s rating → 403', async () => {
  const create = await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 3 });
  const id = create.body.rating._id;
  const res = await request(app).put(`/api/ratings/${id}`)
    .set('X-Test-User', farmerB._id.toString())
    .send({ stars: 1 });
  assert.equal(res.status, 403);
  assert.equal(res.body.error, 'can_only_edit_own_rating');
});

// ============ TEST 12-14 — DELETE ============

test('12. Delete own rating → 200', async () => {
  const create = await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 3 });
  const id = create.body.rating._id;
  const res = await request(app).delete(`/api/ratings/${id}`)
    .set('X-Test-User', farmerA._id.toString());
  assert.equal(res.status, 200);
  assert.equal(res.body.ok, true);
});

test('13. Delete someone else\'s rating as farmer → 403', async () => {
  const create = await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 3 });
  const id = create.body.rating._id;
  const res = await request(app).delete(`/api/ratings/${id}`)
    .set('X-Test-User', farmerB._id.toString());
  assert.equal(res.status, 403);
  assert.equal(res.body.error, 'not_authorized');
});

test('14. Delete someone else\'s rating as superadmin → 200', async () => {
  const create = await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 3 });
  const id = create.body.rating._id;
  const res = await request(app).delete(`/api/ratings/${id}`)
    .set('X-Test-User', superadmin._id.toString());
  assert.equal(res.status, 200);
  assert.equal(res.body.ok, true);
});

// ============ TEST 15 — HIDDEN ROWS EXCLUDED ============

test('15. List ratings excludes hidden ones', async () => {
  const create = await request(app).post('/api/ratings')
    .set('X-Test-User', farmerA._id.toString())
    .send({ sprayProgramId: program._id.toString(), stars: 3 });
  const id = create.body.rating._id;

  // Superadmin hides it via the admin router.
  const hide = await request(app).patch(`/api/admin/ratings/${id}/hide`)
    .set('X-Test-User', superadmin._id.toString())
    .send({ reason: 'test' });
  assert.equal(hide.status, 200);
  assert.equal(hide.body.rating.isHidden, true);

  // Public list for the same program should now be empty.
  const list = await request(app).get(`/api/ratings/${program._id}`)
    .set('X-Test-User', farmerB._id.toString());
  assert.equal(list.status, 200);
  assert.equal(list.body.summary.count, 0);
  assert.equal(list.body.ratings.length, 0);
});

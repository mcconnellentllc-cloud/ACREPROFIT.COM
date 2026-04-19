// test/atrazine-cap.test.js
//
// Unit tests for the atrazine cap-checker. Runs with plain node:test (no jest
// dep added). Usage:
//   node --test test/atrazine-cap.test.js

const { test } = require('node:test');
const assert = require('node:assert');
const {
  atzLbPerAcreForRow,
  computeAtrazineTotal,
  assertAtzCaps,
} = require('../server/lib/atrazineCap');
const { PROGRAMS } = require('../server/seed/corn-programs-data');

const atrazine4L = {
  _id: 'atrazine-4l',
  tradeName: 'Atrazine 4L',
  activeIngredients: [{ name: 'atrazine', class: 'triazine', lbPerGal: 4, percentByWeight: null }],
};

const aatrex90DF = {
  _id: 'aatrex-90df',
  tradeName: 'Aatrex 90-DF',
  activeIngredients: [{ name: 'atrazine', class: 'triazine', lbPerGal: null, percentByWeight: 90 }],
};

const glyphosate = {
  _id: 'glyph',
  tradeName: 'Glyphosate',
  activeIngredients: [{ name: 'glyphosate', class: 'EPSPS', lbPerGal: 4.5, percentByWeight: null }],
};

test('liquid atrazine: 1 qt of 4L = 1.0 lb ai', () => {
  const row = { rate: 1, rateUnit: 'qt', optional: false };
  assert.strictEqual(round4(atzLbPerAcreForRow(row, atrazine4L)), 1.0);
});

test('liquid atrazine: 2 qt of 4L = 2.0 lb ai (single-app cap)', () => {
  const row = { rate: 2, rateUnit: 'qt', optional: false };
  assert.strictEqual(round4(atzLbPerAcreForRow(row, atrazine4L)), 2.0);
});

test('liquid atrazine: 0.5 pt of 4L = 0.25 lb ai', () => {
  const row = { rate: 0.5, rateUnit: 'pt', optional: false };
  assert.strictEqual(round4(atzLbPerAcreForRow(row, atrazine4L)), 0.25);
});

test('liquid atrazine: 16 fl oz of 4L = 0.5 lb ai', () => {
  const row = { rate: 16, rateUnit: 'fl oz', optional: false };
  assert.strictEqual(round4(atzLbPerAcreForRow(row, atrazine4L)), 0.5);
});

test('dry atrazine: 1.11 lb of 90 DF = 1.0 lb ai', () => {
  const row = { rate: 1.111, rateUnit: 'lb', optional: false };
  const result = atzLbPerAcreForRow(row, aatrex90DF);
  assert.ok(Math.abs(result - 1.0) < 0.01, `got ${result}`);
});

test('dry atrazine: 16 dry oz of 90 DF = 0.9 lb ai', () => {
  const row = { rate: 16, rateUnit: 'dry oz', optional: false };
  assert.strictEqual(round4(atzLbPerAcreForRow(row, aatrex90DF)), 0.9);
});

test('non-atrazine chemical contributes 0', () => {
  const row = { rate: 32, rateUnit: 'fl oz', optional: false };
  assert.strictEqual(atzLbPerAcreForRow(row, glyphosate), 0);
});

test('optional + defaultOn=false returns 0 contribution', () => {
  const row = { rate: 1, rateUnit: 'qt', optional: true, defaultOn: false };
  assert.strictEqual(atzLbPerAcreForRow(row, atrazine4L), 0);
});

test('optional + defaultOn=true counts toward cap', () => {
  const row = { rate: 1, rateUnit: 'qt', optional: true, defaultOn: true };
  assert.strictEqual(round4(atzLbPerAcreForRow(row, atrazine4L)), 1.0);
});

test('unknown rate unit throws (not silently 0)', () => {
  const row = { rate: 1, rateUnit: 'barrels', optional: false };
  assert.throws(() => atzLbPerAcreForRow(row, atrazine4L), /unknown liquid rate unit/);
});

test('atrazine AI present but no concentration data throws', () => {
  const noConc = {
    _id: 'x',
    tradeName: 'Mystery',
    activeIngredients: [{ name: 'atrazine', class: 'triazine', lbPerGal: null, percentByWeight: null }],
  };
  const row = { rate: 1, rateUnit: 'qt', optional: false };
  assert.throws(() => atzLbPerAcreForRow(row, noConc), /no concentration data/);
});

test('assertAtzCaps throws on single-app cap violation', () => {
  const prog = {
    name: 'Overloaded',
    passes: [{ passNumber: 1, chemicals: [{ chemicalId: 'atrazine-4l', rate: 2.5, rateUnit: 'qt', optional: false }] }],
  };
  const resolver = new Map([['atrazine-4l', atrazine4L]]);
  assert.throws(() => assertAtzCaps(prog, resolver), /exceeds single-app cap/);
});

test('assertAtzCaps throws on annual cap violation', () => {
  const prog = {
    name: 'Cumulative overload',
    passes: [
      { passNumber: 1, chemicals: [{ chemicalId: 'atrazine-4l', rate: 1.5, rateUnit: 'qt' }] },
      { passNumber: 2, chemicals: [{ chemicalId: 'atrazine-4l', rate: 1.5, rateUnit: 'qt' }] },
    ],
  };
  const resolver = new Map([['atrazine-4l', atrazine4L]]);
  assert.throws(() => assertAtzCaps(prog, resolver), /exceeds annual cap/);
});

test('assertAtzCaps passes at exactly the cap boundary', () => {
  const prog = {
    name: 'At limit',
    passes: [
      { passNumber: 1, chemicals: [{ chemicalId: 'atrazine-4l', rate: 1.1, rateUnit: 'qt' }] },
      { passNumber: 2, chemicals: [{ chemicalId: 'atrazine-4l', rate: 1.1, rateUnit: 'qt' }] },
    ],
  };
  const resolver = new Map([['atrazine-4l', atrazine4L]]);
  const result = assertAtzCaps(prog, resolver);
  assert.strictEqual(round4(result.totalLbPerAcre), 2.2);
});

test('all 5 seeded corn programs pass atrazine caps', () => {
  const mockResolver = new Map();
  PROGRAMS.forEach(p => {
    p.passes.forEach(pass => {
      pass.chemicals.forEach(row => {
        const key = `TRADENAME:${row.tradeName}`;
        if (mockResolver.has(key)) return;
        if (row.tradeName.toLowerCase().includes('atrazine 4l')) {
          mockResolver.set(key, {
            _id: `mock-${row.tradeName}`,
            tradeName: row.tradeName,
            activeIngredients: [{ name: 'atrazine', lbPerGal: 4, percentByWeight: null }],
          });
        } else {
          mockResolver.set(key, {
            _id: `mock-${row.tradeName}`,
            tradeName: row.tradeName,
            activeIngredients: [],
          });
        }
      });
    });
  });

  for (const p of PROGRAMS) {
    const programForCheck = {
      name: p.name,
      passes: p.passes.map(pass => ({
        passNumber: pass.passNumber,
        chemicals: pass.chemicals.map(row => {
          const chem = mockResolver.get(`TRADENAME:${row.tradeName}`);
          return { ...row, chemicalId: chem._id };
        }),
      })),
    };
    const directResolver = new Map();
    programForCheck.passes.forEach(pass => {
      pass.chemicals.forEach(row => {
        directResolver.set(String(row.chemicalId),
          Array.from(mockResolver.values()).find(c => c._id === row.chemicalId));
      });
    });

    assert.doesNotThrow(
      () => assertAtzCaps(programForCheck, directResolver),
      `Program "${p.name}" violates atrazine cap`
    );
  }
});

function round4(n) { return Math.round(n * 10000) / 10000; }

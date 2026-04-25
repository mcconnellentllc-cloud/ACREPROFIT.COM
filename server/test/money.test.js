'use strict';

// server/test/money.test.js
//
// Unit + integration tests for the C2 Decimal128 helpers in
// server/lib/money.js. Runs with node --test (matches atrazine-cap and
// rating test precedent — no test framework added).
//
// Run from server/:
//   node --test test/money.test.js

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const mongoose = require('mongoose');

const {
    isDecimal128,
    toDec,
    serializeMoney,
    decimalToJSONTransform,
    coerceMoneyFields,
} = require('../lib/money');

const D = (s) => mongoose.Types.Decimal128.fromString(s);

// ============ HELPER UNIT TESTS ============

describe('isDecimal128', () => {
    test('true for raw BSON Decimal128', () => {
        assert.equal(isDecimal128(D('12.34')), true);
    });
    test('true for mongoose.Types.Decimal128 (same class)', () => {
        const d = mongoose.Types.Decimal128.fromString('99.99');
        assert.equal(isDecimal128(d), true);
    });
    test('false for Number, string, object, null, undefined', () => {
        assert.equal(isDecimal128(12.34), false);
        assert.equal(isDecimal128('12.34'), false);
        assert.equal(isDecimal128({}), false);
        assert.equal(isDecimal128(null), false);
        assert.equal(isDecimal128(undefined), false);
    });
});

describe('toDec', () => {
    test('null/undefined/empty → null', () => {
        assert.equal(toDec(null), null);
        assert.equal(toDec(undefined), null);
    });
    test('Decimal128 input is re-rounded to dp (not passthrough)', () => {
        // Mongoose may cast 100.5 → Decimal128("100.5") via auto-cast before
        // the pre('validate') hook fires. Re-rounding to dp=2 normalizes to
        // "100.50". Identity is NOT preserved.
        assert.equal(toDec(D('100.5'), 2).toString(), '100.50');
        assert.equal(toDec(D('12.7555'), 4).toString(), '12.7555');
    });
    test('Number rounds to dp (default 2)', () => {
        assert.equal(toDec(12.755).toString(), '12.76');
        assert.equal(toDec(12.754).toString(), '12.75');
    });
    test('Number with dp=4 keeps 4 places', () => {
        assert.equal(toDec(12.755, 4).toString(), '12.7550');
        // Round-up case clear of float-half boundary: 12.75568 has no
        // representation ambiguity at this magnitude, so toFixed rounds
        // the trailing 8 cleanly.
        assert.equal(toDec(12.75568, 4).toString(), '12.7557');
    });
    test('string input parses via Number', () => {
        assert.equal(toDec('12.755').toString(), '12.76');
    });
    test('NaN / Infinity / non-numeric → null', () => {
        assert.equal(toDec(NaN), null);
        assert.equal(toDec(Infinity), null);
        assert.equal(toDec('12abc'), null);
    });
    test('-0 normalizes to 0.00', () => {
        assert.equal(toDec(-0).toString(), '0.00');
    });
    test('INV-00008 float artifact 3649.0499999999997 → 3649.05', () => {
        assert.equal(toDec(3649.0499999999997).toString(), '3649.05');
    });
});

describe('serializeMoney', () => {
    test('null/undefined → null', () => {
        assert.equal(serializeMoney(null), null);
        assert.equal(serializeMoney(undefined), null);
    });
    test('Decimal128 → string', () => {
        assert.equal(serializeMoney(D('12.34')), '12.34');
    });
    test('Number → string', () => {
        assert.equal(serializeMoney(12.34), '12.34');
    });
    test('string passthrough', () => {
        assert.equal(serializeMoney('12.34'), '12.34');
    });
    test('NaN → null', () => {
        assert.equal(serializeMoney(NaN), null);
    });
});

// ============ decimalToJSONTransform ============

describe('decimalToJSONTransform — recursive walk', () => {
    test('top-level Decimal128 fields convert to strings', () => {
        const ret = { total: D('1234.56'), other: 'x' };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.total, '1234.56');
        assert.equal(ret.other, 'x');
    });

    test('Number values pass through unchanged (pre-C3 state)', () => {
        const ret = { total: 1234.56, amount: 99.99 };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.total, 1234.56);
        assert.equal(ret.amount, 99.99);
    });

    test('mixed Number + Decimal128 doc (the during-migration state)', () => {
        const ret = {
            subtotal: D('100.00'),  // already Decimal128
            tax: 8.50,               // still Number
            total: D('108.50'),
        };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.subtotal, '100.00');
        assert.equal(ret.tax, 8.50);
        assert.equal(ret.total, '108.50');
    });

    test('arrays of subdocs', () => {
        const ret = {
            items: [
                { unitPrice: D('12.755'), qty: 5 },
                { unitPrice: D('99.99'), qty: 1 },
            ],
        };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.items[0].unitPrice, '12.755');
        assert.equal(ret.items[0].qty, 5);
        assert.equal(ret.items[1].unitPrice, '99.99');
    });

    test('nested embedded objects', () => {
        const ret = {
            volumeDiscount: { discountedPrice: D('5.00'), notes: 'bulk' },
        };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.volumeDiscount.discountedPrice, '5.00');
        assert.equal(ret.volumeDiscount.notes, 'bulk');
    });

    test('doubly-nested arrays (SupplierBidSheet shape)', () => {
        const ret = {
            supplierBids: [
                {
                    subtotal: D('500.00'),
                    itemPricing: [
                        { pricePerUnit: D('12.50'), totalPrice: D('250.00') },
                        { pricePerUnit: D('25.00'), totalPrice: D('250.00') },
                    ],
                },
            ],
        };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.supplierBids[0].subtotal, '500.00');
        assert.equal(ret.supplierBids[0].itemPricing[0].pricePerUnit, '12.50');
        assert.equal(ret.supplierBids[0].itemPricing[1].totalPrice, '250.00');
    });

    test('Date instances are not traversed', () => {
        const date = new Date('2026-04-25T00:00:00Z');
        const ret = { createdAt: date, total: D('10.00') };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.total, '10.00');
        assert.ok(ret.createdAt instanceof Date, 'Date should not have been mutated');
        assert.equal(ret.createdAt.getTime(), date.getTime());
    });

    test('null fields preserved', () => {
        const ret = { total: null, items: null };
        decimalToJSONTransform(null, ret);
        assert.equal(ret.total, null);
        assert.equal(ret.items, null);
    });
});

// ============ 23-SCHEMA FIXTURE COVERAGE ============
//
// Per the C2 spec: every schema with money fields gets a fixture doc
// constructed in this test, run through decimalToJSONTransform, and
// asserted on output. This exercises the transform across every shape
// the schemas use (top-level fields, array subdocs, embedded objects,
// double-nested arrays).
//
// NOTE: the fixtures here mirror the schema MONEY paths registered in
// server/index.js and server/models/. If a new money field is added,
// add it to both the schema's pathsConfig AND a fixture here.

describe('23-schema fixture coverage', () => {
    const fixtures = {
        Order: {
            orderLines: [{ pricePerPackage: D('100.0000'), lineTotal: D('500.00') }],
            hydrovant: { pricePerPackage: D('165.0000'), lineTotal: D('165.00') },
            chemicals: [{ pricePerPackage: D('50.0000'), totalPrice: D('100.00') }],
            seeds: [{ pricePerBag: D('250.0000'), totalPrice: D('1000.00') }],
            pivotBio: [{ pricePerUnit: D('30.0000'), totalPrice: D('60.00') }],
            additionalProducts: { hydrovant: D('50.00'), multiseal: D('100.00'), pump: D('25.00') },
            totalCost: D('2000.00'), costPerAcre: D('20.0000'),
            repCommission: D('100.00'), adminRevenue: D('200.00'),
            totalCost_cost: D('1500.00'), totalCost_admin: D('1700.00'),
            totalConfirmedPrice: D('2000.00'),
        },
        RepCommission: {
            totalCommissionEarned: D('5000.00'),
            totalCommissionPaid: D('3000.00'),
            commissionBalance: D('2000.00'),
            history: [{ orderTotal: D('1000.00'), commissionAmount: D('50.00') }],
        },
        ChemicalPriceHistory: { costPrice: D('10.0000'), sellPrice: D('15.0000') },
        AuditLog: { amount: D('100.00'), before: { foo: 'bar' }, after: null },
        DistributorPricing: { retailPrice: D('25.0000') },
        ChemicalQuote: {
            pricePerUnit: D('12.7550'),
            packPrice: D('100.0000'),
            volumeDiscount: { discountedPrice: D('11.0000'), minQuantity: 10 },
        },
        RupSaleRecord: { totalAmount: D('500.00') },
        ChemicalOrder: {
            items: [{ unitPrice: D('20.0000'), totalPrice: D('100.00') }],
            subtotal: D('100.00'), discount: D('10.00'),
            marginAdjustment: D('5.0000'),
            freight: D('25.00'), processingFee: D('5.00'), total: D('120.00'),
        },
        LedgerEntry: { amount: D('500.00'), runningBalance: D('1500.00') },
        PurchaseOrder: {
            items: [{ pricePerUnit: D('15.0000'), totalPrice: D('300.00') }],
            subtotal: D('300.00'), freight: D('50.00'),
            otherFees: D('10.00'), superAdminFee: D('5.00'), totalCost: D('365.00'),
        },
        QuoteRequest: {
            items: [{
                costPrice: D('10.0000'), adminPrice: D('12.0000'),
                sellPrice: D('15.0000'), totalPrice: D('150.00'),
            }],
            estimatedTotal: D('150.00'),
        },
        SupplierBidSheet: {
            supplierBids: [{
                subtotal: D('400.00'), freight: D('50.00'), totalBid: D('450.00'),
                itemPricing: [
                    { pricePerUnit: D('10.0000'), totalPrice: D('100.00') },
                    { pricePerUnit: D('30.0000'), totalPrice: D('300.00') },
                ],
            }],
        },
        PurchaseOrderSplit: {
            items: [{ pricePerUnit: D('15.0000'), totalPrice: D('300.00') }],
            subtotal: D('300.00'), freightAllocation: D('25.00'), totalCost: D('325.00'),
        },
        Inventory: { averageCost: D('12.7550'), lastCost: D('12.8000') },
        InventoryTransaction: { unitCost: D('12.7550'), totalCost: D('127.55') },
        InventoryBatch: { costPerUnit: D('12.7550'), totalCost: D('127.55') },
        Invoice: {
            subtotal: D('500.00'), discount: D('50.00'), tax: D('25.00'),
            shipping: D('10.00'), total: D('485.00'),
            amountPaid: D('485.00'), amountDue: D('0.00'),
            creditedAmount: D('0.00'), surchargeAmount: D('0.00'),
            items: [{
                unitPrice: D('25.0000'), costPrice: D('10.0000'),
                adminPrice: D('15.0000'), totalPrice: D('500.00'),
                margin: D('15.0000'),
            }],
        },
        CreditNote: { amount: D('100.00') },
        ProgramQuote: {
            items: [{ unitPrice: D('20.0000'), totalPrice: D('200.00') }],
            subtotal: D('200.00'), total: D('200.00'), costPerAcre: D('10.0000'),
        },
        ChemicalRequest: {
            competitorPrice: D('45.0000'),
            supplierQuotes: [{ price: D('40.0000'), supplierName: 'X' }],
        },
        UserSubmittedChemical: { costPerUnit: D('12.7550') },
        Chemical: {
            costPrice: D('10.0000'), adminMarginDollars: D('2.0000'),
            adminPrice: D('12.0000'), marginDollars: D('3.0000'),
            sellPrice: D('15.0000'),
        },
        SprayProgram: { estimatedCostPerAcre: D('20.0000') },
    };

    // Sanity: 23 schemas exactly. If this fails, somebody added/removed a
    // schema without updating the test fixture map.
    test('fixture count is 23 (matches C2 scope)', () => {
        assert.equal(Object.keys(fixtures).length, 23);
    });

    for (const [name, src] of Object.entries(fixtures)) {
        test(`${name} fixture: every Decimal128 field becomes a string`, () => {
            const ret = JSON.parse(JSON.stringify(src, (_k, v) => {
                // BSON Decimal128 serializes via toJSON to {$numberDecimal:...}
                // by default. Use plain clone preserving Decimal128 instances.
                return v;
            }));
            // Plain clone above doesn't preserve Decimal128 — re-clone manually:
            const cloned = deepCloneFixture(src);
            decimalToJSONTransform(null, cloned);
            // Walk cloned and assert no Decimal128 remains.
            const remaining = findRemainingDecimal128(cloned);
            assert.equal(remaining.length, 0,
                `${name} still has Decimal128 instances at: ${remaining.join(', ')}`);
            // Also assert that money fields are now non-empty strings.
            const moneyStrings = findMoneyStrings(cloned);
            assert.ok(moneyStrings.length > 0,
                `${name} fixture should have at least one money string after transform`);
        });
    }
});

function deepCloneFixture(v) {
    if (v === null || v === undefined) return v;
    if (isDecimal128(v)) return v;  // keep instance
    if (v instanceof Date) return new Date(v.getTime());
    if (Array.isArray(v)) return v.map(deepCloneFixture);
    if (typeof v === 'object') {
        const out = {};
        for (const k of Object.keys(v)) out[k] = deepCloneFixture(v[k]);
        return out;
    }
    return v;
}

function findRemainingDecimal128(node, path = '') {
    const found = [];
    if (node === null || typeof node !== 'object') return found;
    if (Array.isArray(node)) {
        node.forEach((v, i) => {
            if (isDecimal128(v)) found.push(`${path}[${i}]`);
            else if (typeof v === 'object') found.push(...findRemainingDecimal128(v, `${path}[${i}]`));
        });
        return found;
    }
    for (const k of Object.keys(node)) {
        const v = node[k];
        if (isDecimal128(v)) found.push(`${path}.${k}`);
        else if (v !== null && typeof v === 'object' && !(v instanceof Date)) {
            found.push(...findRemainingDecimal128(v, path ? `${path}.${k}` : k));
        }
    }
    return found;
}

function findMoneyStrings(node) {
    const found = [];
    if (node === null || typeof node !== 'object') return found;
    if (Array.isArray(node)) {
        node.forEach((v) => {
            if (typeof v === 'string' && /^\d+\.\d{2,}$/.test(v)) found.push(v);
            else if (typeof v === 'object') found.push(...findMoneyStrings(v));
        });
        return found;
    }
    for (const k of Object.keys(node)) {
        const v = node[k];
        if (typeof v === 'string' && /^\d+\.\d{2,}$/.test(v)) found.push(`${k}=${v}`);
        else if (v !== null && typeof v === 'object' && !(v instanceof Date)) {
            found.push(...findMoneyStrings(v));
        }
    }
    return found;
}

// ============ MANUAL-PICK PARITY TESTS ============
//
// The C2 spec calls these out as "the most important test in C2." For each
// affected route, simulate the .map(c => ({ ... })) extraction in BOTH
// pre-C3 (Number) and post-C3 (Decimal128) state. Wrap money fields with
// serializeMoney. Assert the response shape is identical in both states.

describe('manual-pick parity — Number-state vs Decimal128-state', () => {
    // "Shape parity" here means: same keys present, same types (strings stay
    // strings, Numbers stay Numbers, null stays null), and money strings
    // parseFloat to the same numeric amount. Byte-identical strings are NOT
    // expected — pre-C3 a Number 10 serializes as "10", post-C3 a Decimal128
    // with dp=4 serializes as "10.0000". Both round-trip correctly through
    // frontend fmtMoney (parity-tested in C1a).

    function shapeParity(a, b) {
        const aKeys = Object.keys(a).sort();
        const bKeys = Object.keys(b).sort();
        assert.deepEqual(aKeys, bKeys, 'keys must match');
        for (const k of aKeys) {
            assert.equal(typeof a[k], typeof b[k], `${k}: type must match`);
            if (a[k] === null) assert.equal(b[k], null, `${k}: null in both states`);
            // Money strings: same numeric amount via parseFloat
            if (typeof a[k] === 'string' && /^-?\d+(\.\d+)?$/.test(a[k])) {
                assert.equal(parseFloat(a[k]), parseFloat(b[k]),
                    `${k}: numeric value must match (a=${a[k]}, b=${b[k]})`);
            }
        }
    }

    test('publicChemicals shape (server/index.js:~7700)', () => {
        const numberChem = { _id: 'x', sellPrice: 12.755, productName: 'Foo', unit: 'gal' };
        const decimalChem = { _id: 'x', sellPrice: D('12.755'), productName: 'Foo', unit: 'gal' };

        const extract = (c) => ({
            _id: c._id,
            productName: c.productName,
            unit: c.unit,
            price: serializeMoney(c.sellPrice),
            sellPrice: serializeMoney(c.sellPrice),
        });

        const numberOut = extract(numberChem);
        const decimalOut = extract(decimalChem);

        shapeParity(numberOut, decimalOut);
        assert.equal(typeof numberOut.price, 'string');
        assert.equal(typeof decimalOut.price, 'string');
        assert.equal(parseFloat(numberOut.price), 12.755);
        assert.equal(parseFloat(decimalOut.price), 12.755);
    });

    test('admin chemicals shape (server/index.js:~11955 + ~14080)', () => {
        const numberChem = {
            _id: 'x', costPrice: 10, adminPrice: 12, sellPrice: 15,
            adminMargin: 16.67, margin: 33.33, productName: 'Foo',
        };
        const decimalChem = {
            _id: 'x', costPrice: D('10.0000'), adminPrice: D('12.0000'),
            sellPrice: D('15.0000'), adminMargin: 16.67, margin: 33.33,
            productName: 'Foo',
        };

        const extract = (c) => ({
            _id: c._id,
            productName: c.productName,
            costPrice: serializeMoney(c.costPrice),
            adminPrice: serializeMoney(c.adminPrice),
            adminMargin: c.adminMargin,  // percentage, stays Number
            sellPrice: serializeMoney(c.sellPrice),
            margin: c.margin,            // percentage, stays Number
        });

        const numberOut = extract(numberChem);
        const decimalOut = extract(decimalChem);

        shapeParity(numberOut, decimalOut);
        // Money fields are strings
        assert.equal(typeof numberOut.costPrice, 'string');
        assert.equal(typeof decimalOut.costPrice, 'string');
        // Percentages stay Number
        assert.equal(typeof numberOut.adminMargin, 'number');
        assert.equal(typeof decimalOut.adminMargin, 'number');
        // Numeric amounts match
        assert.equal(parseFloat(numberOut.costPrice), 10);
        assert.equal(parseFloat(decimalOut.costPrice), 10);
    });

    test('null sellPrice handled in both states', () => {
        const numberChem = { sellPrice: null };
        const decimalChem = { sellPrice: null };
        const extract = (c) => ({ price: serializeMoney(c.sellPrice) });
        assert.deepEqual(extract(numberChem), extract(decimalChem));
        assert.equal(extract(numberChem).price, null);
    });
});

// ============ coerceMoneyFields ============
//
// Construct a real mongoose Schema with mixed Number / Decimal128 paths and
// verify the hook coerces Decimal128 paths and no-ops Number paths.

describe('coerceMoneyFields', () => {
    const D128 = mongoose.Schema.Types.Decimal128;

    // Pre-C3 schema: money fields still Number. Hook should be a no-op.
    const NumberSchema = new mongoose.Schema({
        subtotal: Number,
        items: [{ unitPrice: Number, totalPrice: Number }],
    }, { autoCreate: false });
    const NUMBER_PATHS = { 'subtotal': 2, 'items.unitPrice': 4, 'items.totalPrice': 2 };

    // Post-C3 schema: same paths but Decimal128. Hook should coerce.
    const DecimalSchema = new mongoose.Schema({
        subtotal: D128,
        items: [{ unitPrice: D128, totalPrice: D128 }],
    }, { autoCreate: false });
    const DECIMAL_PATHS = { 'subtotal': 2, 'items.unitPrice': 4, 'items.totalPrice': 2 };

    test('pre-C3 (Number schema) — Number values pass through', () => {
        const NumberModel = mongoose.model('TestNumberModel_' + Date.now(), NumberSchema);
        const doc = new NumberModel({
            subtotal: 100,
            items: [{ unitPrice: 12.755, totalPrice: 100 }],
        });
        coerceMoneyFields(doc, NUMBER_PATHS);
        assert.equal(typeof doc.subtotal, 'number');
        assert.equal(doc.subtotal, 100);
        assert.equal(typeof doc.items[0].unitPrice, 'number');
    });

    test('post-C3 (Decimal128 schema) — Number / string inputs coerce to Decimal128', () => {
        const DecimalModel = mongoose.model('TestDecimalModel_' + Date.now(), DecimalSchema);
        const doc = new DecimalModel({
            subtotal: 100.5,
            items: [{ unitPrice: '12.755', totalPrice: 100 }],
        });
        coerceMoneyFields(doc, DECIMAL_PATHS);
        assert.ok(isDecimal128(doc.subtotal),
            `subtotal should be Decimal128, got ${typeof doc.subtotal}`);
        assert.equal(doc.subtotal.toString(), '100.50');
        assert.ok(isDecimal128(doc.items[0].unitPrice));
        assert.equal(doc.items[0].unitPrice.toString(), '12.7550');  // dp=4
    });

    test('post-C3 — Decimal128 values are re-rounded to dp (not identity-preserved)', () => {
        const DecimalModel = mongoose.model('TestDecimalModel2_' + Date.now(), DecimalSchema);
        // Already at dp=2, so re-rounding is a value-preserving no-op.
        const doc = new DecimalModel({ subtotal: D('99.99'), items: [] });
        coerceMoneyFields(doc, DECIMAL_PATHS);
        assert.ok(isDecimal128(doc.subtotal));
        assert.equal(doc.subtotal.toString(), '99.99');
    });

    test('hard error in dev when schema has Decimal128 path missing from pathsConfig', () => {
        const InsufficientSchema = new mongoose.Schema({
            subtotal: D128,
            tax: D128,  // not in pathsConfig
        }, { autoCreate: false });
        const Model = mongoose.model('TestInsufficient_' + Date.now(), InsufficientSchema);
        const doc = new Model({ subtotal: 100, tax: 5 });
        const partialPaths = { 'subtotal': 2 };  // missing tax
        const prevEnv = process.env.NODE_ENV;
        delete process.env.NODE_ENV;  // not 'production'
        try {
            assert.throws(() => coerceMoneyFields(doc, partialPaths), /missing from pathsConfig.*tax/);
        } finally {
            if (prevEnv !== undefined) process.env.NODE_ENV = prevEnv;
        }
    });

    test('production: missing path warns instead of throwing', () => {
        const InsufficientSchema = new mongoose.Schema({
            subtotal: D128,
            tax: D128,
        }, { autoCreate: false });
        const Model = mongoose.model('TestProdMissing_' + Date.now(), InsufficientSchema);
        const doc = new Model({ subtotal: 100, tax: 5 });
        const prevEnv = process.env.NODE_ENV;
        process.env.NODE_ENV = 'production';
        const origWarn = console.warn;
        let warned = '';
        console.warn = (...args) => { warned = args.join(' '); };
        try {
            assert.doesNotThrow(() => coerceMoneyFields(doc, { 'subtotal': 2 }));
            assert.match(warned, /missing pathsConfig entries.*tax/);
        } finally {
            console.warn = origWarn;
            process.env.NODE_ENV = prevEnv;
        }
    });
});

'use strict';

// server/lib/money.js — central helpers for the Decimal128 migration (C2/C3).
//
// Used by:
//   - C2 (this commit): pre('validate') coerce hooks + toJSON transform on
//     every money-bearing schema, plus the two manual-pick response sites
//     in server/index.js (lines ~7503 and ~13878 — see serializeMoney).
//   - C3 (next): backfill script that converts existing Number documents
//     to Decimal128.
//
// Pre-C3 behaviour: schemas are still Number-typed. The pre('validate')
// hook is a no-op because the schema-path-instance check only coerces
// when the field is Decimal128. The toJSON transform is also a no-op on
// Number values (only converts Decimal128 BSON instances).
//
// Post-C3 behaviour: schemas are Decimal128. The pre('validate') hook
// coerces incoming string/Number JSON bodies to Decimal128 BSON before
// save. The toJSON transform converts Decimal128 BSON to "12.34" strings
// on every API response.

const mongoose = require('mongoose');
const Decimal = require('decimal.js-light');

// True for both raw BSON Decimal128 (from mongoose docs) and
// mongoose.Types.Decimal128 instances (constructed in code). They share
// the same _bsontype tag.
function isDecimal128(v) {
    return v != null && typeof v === 'object' && v._bsontype === 'Decimal128';
}

// toDec — normalize any input to a Decimal128 BSON value, rounded to dp
// decimal places. Always re-rounds (no passthrough), because the pre('validate')
// hook fires AFTER mongoose's auto-cast — so a Decimal128 value here may
// have arrived with unwanted precision (e.g. 100.5 cast from a Number) and
// needs normalizing to the schema's intended dp.
//   null/undefined  → null      (no value to coerce)
//   Decimal128      → Decimal128 re-rounded to dp
//   "12.755"        → Decimal128("12.76") (string input, half-up at dp=2)
//   12.755          → Decimal128("12.76") (Number input, same rounding)
//   "12.755abc"     → null (Number(...) returns NaN; rejected)
//   NaN / Infinity  → null
function toDec(v, dp = 2) {
    if (v === null || v === undefined) return null;
    const num = isDecimal128(v) ? parseFloat(v.toString()) : Number(v);
    if (!Number.isFinite(num)) return null;
    return mongoose.Types.Decimal128.fromString(num.toFixed(dp));
}

// serializeMoney — the manual-pick equivalent of the toJSON transform.
// Use it where code extracts money fields from a mongoose doc into a
// plain response object via .map(c => ({ ... })) — the transform doesn't
// fire on those paths.
//   null/undefined  → null
//   Decimal128      → "12.34"
//   Number          → "12.34"
//   "12.34"         → "12.34" (passthrough)
function serializeMoney(v) {
    if (v === null || v === undefined) return null;
    if (isDecimal128(v)) return v.toString();
    // decimal.js-light Decimal instance — has .toString() and an isDecimal flag.
    // Identified by ducktype since the constructor reference may differ across
    // module loads. C3a sites pass these directly: `serializeMoney(new Decimal(a).minus(b))`.
    if (v && typeof v === 'object' && typeof v.toFixed === 'function' && typeof v.cmp === 'function') {
        return v.toString();
    }
    if (typeof v === 'number') {
        if (!Number.isFinite(v)) return null;
        return v.toString();
    }
    if (typeof v === 'string') return v;
    return null;
}

// decimalToJSONTransform — recursive walk applied via schema.set('toJSON',
// { transform }). Mongoose calls this with (doc, ret) where ret is the
// already-cloned plain object form of the document. We mutate ret in
// place, converting any Decimal128 BSON instances to strings. Skips Date
// instances so we don't traverse their internals.
function decimalToJSONTransform(_doc, ret) {
    walkConvert(ret);
    return ret;
}

function walkConvert(node) {
    if (node === null || typeof node !== 'object') return;
    if (Array.isArray(node)) {
        for (let i = 0; i < node.length; i++) {
            const v = node[i];
            if (isDecimal128(v)) node[i] = v.toString();
            else if (v !== null && typeof v === 'object' && !(v instanceof Date)) walkConvert(v);
        }
        return;
    }
    for (const k of Object.keys(node)) {
        const v = node[k];
        if (isDecimal128(v)) node[k] = v.toString();
        else if (v !== null && typeof v === 'object' && !(v instanceof Date)) walkConvert(v);
    }
}

// coerceMoneyFields — pre('validate') hook helper. Walks every path in
// pathsConfig, and for each path where the schema's runtime type is
// Decimal128, coerces the document's current value via toDec(v, dp).
//
//   pathsConfig: { 'subtotal': 2, 'items.unitPrice': 4, 'supplierBids.itemPricing.pricePerUnit': 4 }
//
// Pre-C3 the schema field is Number → instance check fails → no-op. Once
// C3 flips the schema, the same hook starts coercing.
//
// In development NODE_ENV !== 'production', verifies that every Decimal128
// path on the schema appears in pathsConfig and throws a hard error on
// mismatch. Post-C3 this catches the silent-drift case where someone
// adds a new Decimal128 field but forgets to register it.
function coerceMoneyFields(doc, pathsConfig) {
    if (!doc || typeof doc !== 'object') return;
    const schema = doc.schema || (doc.constructor && doc.constructor.schema);
    if (!schema) return;

    if (process.env.NODE_ENV !== 'production') {
        const missing = findMissingDecimalPaths(schema, pathsConfig);
        if (missing.length > 0) {
            throw new Error(
                'coerceMoneyFields: schema has Decimal128 paths missing from pathsConfig: ' +
                missing.join(', ') +
                '. Add them to the schema\'s pathsConfig in server/index.js.'
            );
        }
    } else {
        const missing = findMissingDecimalPaths(schema, pathsConfig);
        if (missing.length > 0) {
            // Production: don't throw, but log loudly.
            console.warn('coerceMoneyFields: missing pathsConfig entries:', missing.join(', '));
        }
    }

    for (const path of Object.keys(pathsConfig)) {
        coerceAtPath(doc, schema, path.split('.'), 0, pathsConfig[path], path);
    }
}

function coerceAtPath(node, schema, parts, i, dp, fullPath) {
    if (node === null || node === undefined) return;
    const part = parts[i];

    if (i === parts.length - 1) {
        // Leaf — coerce only if schema says Decimal128
        const sp = schema.path(fullPath);
        if (!sp || sp.instance !== 'Decimal128') return;
        const cur = node[part];
        if (cur === undefined) return;
        node[part] = toDec(cur, dp);
        return;
    }

    const sub = node[part];
    if (sub === undefined || sub === null) return;

    if (Array.isArray(sub)) {
        // Array of subdocs — find the array's inner schema and recurse
        // per-element with the remaining path segments and a fresh schema.
        const arraySchemaPath = parts.slice(0, i + 1).join('.');
        const arraySp = schema.path(arraySchemaPath);
        if (!arraySp || !arraySp.schema) return;
        const subSchema = arraySp.schema;
        const remaining = parts.slice(i + 1);
        const remainingFull = remaining.join('.');
        for (const item of sub) {
            coerceAtPath(item, subSchema, remaining, 0, dp, remainingFull);
        }
    } else if (typeof sub === 'object') {
        // Embedded object literal (e.g., Order's `additionalProducts` or
        // ChemicalQuote's `volumeDiscount`). Schema path stays the same;
        // we just descend into the object.
        coerceAtPath(sub, schema, parts, i + 1, dp, fullPath);
    }
}

// Returns the list of Decimal128 schema paths that are NOT registered in
// pathsConfig. Walks the schema and any nested subdoc schemas (one and
// two levels — covers SupplierBidSheet's supplierBids.itemPricing).
function findMissingDecimalPaths(schema, pathsConfig) {
    const declared = new Set();
    schema.eachPath((path, sp) => {
        if (sp.instance === 'Decimal128') declared.add(path);
        if (sp.instance === 'Array' && sp.schema) {
            sp.schema.eachPath((subPath, subSp) => {
                if (subSp.instance === 'Decimal128') declared.add(`${path}.${subPath}`);
                if (subSp.instance === 'Array' && subSp.schema) {
                    subSp.schema.eachPath((leafPath, leafSp) => {
                        if (leafSp.instance === 'Decimal128') {
                            declared.add(`${path}.${subPath}.${leafPath}`);
                        }
                    });
                }
            });
        }
    });
    const missing = [];
    for (const p of declared) {
        if (!(p in pathsConfig)) missing.push(p);
    }
    return missing;
}

module.exports = {
    isDecimal128,
    toDec,
    serializeMoney,
    decimalToJSONTransform,
    coerceMoneyFields,
    // Re-exported so callers (C3a math conversions, C3b backfill script)
    // can use the same Decimal class without each one re-requiring the dep.
    Decimal,
};

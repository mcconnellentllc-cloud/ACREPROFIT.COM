// money.js — defensive money formatters for the Decimal128 migration.
//
// CDN dependency, load BEFORE this file:
//   <script src="https://cdn.jsdelivr.net/npm/decimal.js-light@2.5.1/decimal.min.js"></script>
//
// Single-source-of-truth on the version. Bumping decimal.js-light means
// updating that URL in every HTML file that loads it. Do not chase @latest;
// pin intentionally:
//   grep -l 'decimal.js-light@' *.html
//
// fmtMoney(v)        — '$X.XX', no thousands separator. Use for cents-precision
//                       displays, table cells, line-item totals, badges.
// fmtMoneyLocale(v)  — '$X,XXX.XX' with en-US thousands separator. Use for
//                       summary cards, dashboard tiles, big-number rollups.
//
// Both accept Number (legacy, pre-migration), string ("123.45", post-migration
// API contract), Decimal instance (intermediate math), or the raw Mongo
// extended-JSON shape {$numberDecimal: "..."} (defensive). Null/undefined/
// empty render as '$0.00'; callers that need a different empty state (e.g.
// 'Quote Required') must guard with their own conditional before calling.

(function () {
    'use strict';

    function _normalize(v) {
        if (v === null || v === undefined || v === '') return null;
        if (typeof v === 'object' && v !== null && '$numberDecimal' in v) return String(v.$numberDecimal);
        return String(v);
    }

    function fmtMoney(v) {
        var s = _normalize(v);
        if (s === null) return '$0.00';
        return '$' + new Decimal(s).toFixed(2);
    }

    var _localeFmt = (typeof Intl !== 'undefined' && Intl.NumberFormat)
        ? new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
        : null;

    function fmtMoneyLocale(v) {
        var s = _normalize(v);
        if (s === null) return '$0.00';
        var fixed = new Decimal(s).toFixed(2);
        if (!_localeFmt) return '$' + fixed;
        return '$' + _localeFmt.format(Number(fixed));
    }

    window.fmtMoney = fmtMoney;
    window.fmtMoneyLocale = fmtMoneyLocale;
})();


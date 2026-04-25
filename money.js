// money.js — defensive money formatter for the Decimal128 migration.
//
// CDN dependency, load BEFORE this file:
//   <script src="https://cdn.jsdelivr.net/npm/decimal.js-light@2.5.1/decimal.min.js"></script>
//
// Single-source-of-truth on the version. Bumping decimal.js-light means
// updating that URL in every HTML file that loads it. Do not chase @latest;
// pin intentionally:
//   grep -l 'decimal.js-light@' *.html
//
// fmtMoney accepts Number (legacy, pre-migration), string ("123.45",
// post-migration API contract), or the raw Mongo extended-JSON shape
// {$numberDecimal: "..."} (defensive, in case a payload bypasses the
// schema toJSON transform). Output matches the existing convention
// '$' + x.toFixed(2) — no thousands separator. Null/undefined/empty
// render as '$0.00'; callers that need a different empty state (e.g.
// 'Quote Required') must guard with their own conditional before calling.

(function () {
    'use strict';

    function fmtMoney(v) {
        if (v === null || v === undefined || v === '') return '$0.00';
        var s = (typeof v === 'object' && v !== null && '$numberDecimal' in v)
            ? String(v.$numberDecimal)
            : String(v);
        return '$' + new Decimal(s).toFixed(2);
    }

    window.fmtMoney = fmtMoney;
})();

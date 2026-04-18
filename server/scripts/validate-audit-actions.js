#!/usr/bin/env node
// Validates that every logAudit({ action: '<literal>' }) call site in
// server/index.js uses an action value that is declared in the AuditLog
// schema's action enum. Exists because logAudit() swallows Mongoose
// validation errors into a non-blocking console.error, so a typo or a
// forgotten enum addition fails silently in production with zero visible
// impact until an audit rolls around and discovers the missing records.
//
// History: Five silent failures accumulated before this check existed.
// See PR #118.
//
// Usage:
//   node server/scripts/validate-audit-actions.js
//   npm run --prefix server validate:audit
//
// Exit codes:
//   0 - all static action literals are present in the enum
//   1 - at least one call-site action is missing from the enum
//   2 - could not locate the enum block or the source file

'use strict';

const fs = require('fs');
const path = require('path');

const SERVER_PATH = path.join(__dirname, '..', 'index.js');

function die(code, msg) {
    console.error(msg);
    process.exit(code);
}

if (!fs.existsSync(SERVER_PATH)) {
    die(2, `Cannot find ${SERVER_PATH}`);
}
const source = fs.readFileSync(SERVER_PATH, 'utf8');

// Extract the AuditLog action enum. We anchor on `auditLogSchema` so we
// never accidentally pick up a different schema's action enum.
const enumMatch = source.match(
    /auditLogSchema\s*=\s*new\s+mongoose\.Schema\s*\(\s*\{[\s\S]*?action\s*:\s*\{[\s\S]*?enum\s*:\s*\[([\s\S]*?)\]/
);
if (!enumMatch) {
    die(2, 'Could not locate the AuditLog action enum in server/index.js (auditLogSchema.action.enum).');
}
const enumValues = new Set(
    [...enumMatch[1].matchAll(/'([^']+)'/g)].map(m => m[1])
);

// Find every logAudit( call site. Use balanced-paren walking so nested
// object literals (e.g. `after: { ... }`) don't confuse a naive regex.
function findCallSites(text) {
    const sites = [];
    const re = /\blogAudit\s*\(/g;
    let m;
    while ((m = re.exec(text)) !== null) {
        // Skip the function declaration itself (`async function logAudit(...)`).
        // We only care about call sites.
        const preceding = text.slice(Math.max(0, m.index - 20), m.index);
        if (/function\s+$/.test(preceding)) continue;

        const argsStart = m.index + m[0].length;
        let depth = 1;
        let i = argsStart;
        let stringChar = null;
        while (i < text.length && depth > 0) {
            const ch = text[i];
            if (stringChar) {
                if (ch === '\\') { i += 2; continue; }
                if (ch === stringChar) stringChar = null;
                i++;
                continue;
            }
            if (ch === "'" || ch === '"' || ch === '`') {
                stringChar = ch;
            } else if (ch === '(') {
                depth++;
            } else if (ch === ')') {
                depth--;
                if (depth === 0) break;
            }
            i++;
        }
        const argsSlice = text.slice(argsStart, i);
        const actionLiteral = argsSlice.match(/(?:^|[,{\s])action\s*:\s*'([^']+)'/);
        const line = text.slice(0, m.index).split('\n').length;
        sites.push({
            line,
            action: actionLiteral ? actionLiteral[1] : null
        });
    }
    return sites;
}

const callSites = findCallSites(source);
const staticSites = callSites.filter(c => c.action !== null);
const dynamicSites = callSites.filter(c => c.action === null);
const callSiteActions = new Set(staticSites.map(c => c.action));
const missing = staticSites.filter(c => !enumValues.has(c.action));
const unused = [...enumValues].filter(e => !callSiteActions.has(e));

console.log(`AuditLog action enum values:      ${enumValues.size}`);
console.log(`logAudit() call sites found:      ${callSites.length}`);
console.log(`  - with static action literal:   ${staticSites.length}`);
console.log(`  - with dynamic action (skipped):${dynamicSites.length}`);

if (dynamicSites.length) {
    console.log('\nDynamic action values (cannot be validated statically):');
    dynamicSites.forEach(c => console.log(`  server/index.js:${c.line}`));
}

if (unused.length) {
    console.log('\nEnum values declared but not used at any call site (informational only):');
    unused.forEach(a => console.log(`  - ${a}`));
}

if (missing.length) {
    console.error('\nFAIL: logAudit() action literals NOT in the enum:');
    missing.forEach(c => {
        console.error(`  - '${c.action}'  (server/index.js:${c.line})`);
    });
    console.error(
        '\nFix: add the missing values to the auditLogSchema action enum ' +
        'in server/index.js. Without this, logAudit() at the flagged call ' +
        'sites will silently fail AuditLog.create() and no compliance ' +
        'record will be written.'
    );
    process.exit(1);
}

console.log('\nOK: all static logAudit() action literals are declared in the enum.');
process.exit(0);

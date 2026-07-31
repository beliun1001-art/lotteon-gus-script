'use strict';
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const source = fs.readFileSync('PR25_Production_Continuation_R13.gs', 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(source, sandbox);

const sample = "var LOTTEON_PATCH_BOOTSTRAP_URLS = ['A.gs','B.gs','A.gs'];";
assert.deepStrictEqual(Array.from(sandbox.pr25r13_parseBootstrapPatchNames_(sample)), ['A.gs','B.gs']);
assert.ok(!/function resumePr25ProductionVatR13\([\s\S]*?deleteProperty\(PR25_R13_VAT_STATE_KEY\)/.test(source));
assert.ok(source.includes("PR25_R13_HANDLER = 'continuePr25ProductionVatR13'"));
assert.ok(source.includes('branch ZIP 1회'));

const goodSheet = { getLastRow: () => 1106 };
const goodSs = { getSheetByName: name => name === '부가세_신고자료' ? goodSheet : null };
sandbox.pr25r13_assertResumeSafe_(goodSs, { status:'running', phase:'detail', writtenRows:1105, sourceRow:1502, sourceLastRow:2056 });
assert.throws(() => sandbox.pr25r13_assertResumeSafe_(goodSs, { status:'running', phase:'detail', writtenRows:1104, sourceRow:1502, sourceLastRow:2056 }), /안전검증 실패/);
console.log('R13 cached production resume tests PASS');

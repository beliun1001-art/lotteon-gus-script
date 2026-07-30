'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const ctx={console};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync('Patch_v6_60_vat_business_half_summary.gs','utf8'),ctx);
vm.runInContext(fs.readFileSync('Patch_v6_60_vat_card_input_runtime_fix.gs','utf8'),ctx);

assert.strictEqual(ctx.normalizeDateText_v660_('2026.06.30'),'2026-06-30');
const date=vm.runInContext("new Date(2026, 6, 1)",ctx);
assert.strictEqual(ctx.normalizeDateText_v660_(date),'2026-07-01');
assert.strictEqual(typeof ctx.prepareVatCardInputs_v660_,'function');
assert.strictEqual(typeof ctx.rebuildVatPurchaseCardSummary_v660_,'function');

const formatted=[];
const range={setValues(){return this;},setBackground(){return this;},setFontWeight(){return this;},setHorizontalAlignment(){return this;},setNumberFormat(fmt){formatted.push(fmt);return this;}};
const sheet={getLastRow:()=>0,getMaxRows:()=>100,getRange:()=>range,setFrozenRows:()=>{}};
const ss={getSheetByName:()=>sheet,insertSheet:()=>sheet};
ctx.ensureSheetHeaders_v660_(ss,'카드_마스터',ctx.vatCardMasterHeaders_v660_());
assert(formatted.includes('@'));

console.log('v6.60 card input runtime fix tests PASS');

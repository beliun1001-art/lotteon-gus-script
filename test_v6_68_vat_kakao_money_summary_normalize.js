'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const compact=v=>String(v==null?'':v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');
const sandbox={
  compact_v660_:compact,
  matchVatOrderCardCanonical_v664_:()=>({
    status:'NON_CARD',company:'비카드',alias:'',
    cardName:'카카오페이 페이머니',evidenceType:'카카오페이 페이머니'
  })
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_68_vat_kakao_money_summary_normalize.gs','utf8'),sandbox);
let r=sandbox.matchVatOrderCardCanonical_v664_({},[],[],{});
assert.strictEqual(r.company,'비카드');
assert.strictEqual(r.alias,'신한은행 계좌결제');
assert.strictEqual(r.cardName,'카카오페이 페이머니');
assert.strictEqual(r.cardNumber,'');
assert.strictEqual(r.cardEnd4,'');
r=sandbox.normalizeVatKakaoMoneyMatch_v668_({status:'MATCHED',company:'우리카드',cardName:'카카오페이 카드'});
assert.strictEqual(r.company,'우리카드');
console.log('v6.68 KakaoPay Money summary normalization tests PASS');

'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');

const events=[];
class Range {
  constructor(row,col,rows,cols,sheet){this.row=row;this.col=col;this.rows=rows;this.cols=cols;this.sheet=sheet;}
  setNumberFormat(fmt){events.push(['format',this.row,this.col,this.rows,this.cols,fmt]);return this;}
  getValues(){
    if(this.col===16)return [['Sun Dec 31 1899 07:22:52 GMT+0827 (한국 표준시)']];
    if(this.col===27)return [['롯데카드 / 036 / 2026-05-29 / Sun Dec 31 1899 09:28:52 GMT+0827 (한국 표준시) / 30720']];
    return Array.from({length:this.rows},()=>Array(this.cols).fill(''));
  }
  setValues(v){events.push(['values',this.row,this.col,JSON.parse(JSON.stringify(v))]);return this;}
}
class Sheet {getRange(r,c,rs=1,cs=1){return new Range(r,c,rs,cs,this);}}
const sheet=new Sheet();
const ss={getSheetByName(){return sheet;},insertSheet(){return sheet;}};
const sandbox={
  console,
  LOTTEON_V660_CARD_DIAG_SHEET:'부가세_카드매칭검증',
  writeVatCardMatchDiagnostic_v660_(){events.push(['baseDiag']);return 1;},
  prependVatBusinessCardHalfSummary_v660_(){events.push(['baseSummary']);return {summaryRows:1};}
};
vm.createContext(sandbox);vm.runInContext(fs.readFileSync('Patch_v6_65_vat_card_text_safe_output.gs','utf8'),sandbox);
assert.strictEqual(sandbox.normalizeVatCardDisplayTime_v665_('Sun Dec 31 1899 07:22:52 GMT+0827 (한국 표준시)'),'07:22:52');
assert.strictEqual(sandbox.normalizeVatCardDisplayTime_v665_('7:05'),'07:05:00');
assert.ok(sandbox.normalizeVatCardDisplayText_v665_('x Sun Dec 31 1899 09:28:52 GMT+0827 (한국 표준시) y').includes('09:28:52'));

sandbox.writeVatCardMatchDiagnostic_v660_(ss,[{}]);
const baseDiagIndex=events.findIndex(e=>e[0]==='baseDiag');
for(const col of [4,5,6,13,14,16,17,23]){
  const idx=events.findIndex(e=>e[0]==='format'&&e[1]===2&&e[2]===col&&e[5]==='@');
  assert.ok(idx>=0 && idx<baseDiagIndex,'text format must precede diagnostic base write for col '+col);
}
const timeWrite=events.find(e=>e[0]==='values'&&e[1]===2&&e[2]===16);
assert.strictEqual(timeWrite[3][0][0],'07:22:52');
const summaryWrite=events.find(e=>e[0]==='values'&&e[1]===2&&e[2]===27);
assert.ok(summaryWrite[3][0][0].includes('09:28:52'));
assert.ok(!summaryWrite[3][0][0].includes('1899'));

events.length=0;
sandbox.prependVatBusinessCardHalfSummary_v660_(sheet,[[]],[['2026']]);
const baseSummaryIndex=events.findIndex(e=>e[0]==='baseSummary');
const preSummary=events.findIndex(e=>e[0]==='format'&&e[1]===3&&e[2]===1&&e[4]===11&&e[5]==='@');
assert.ok(preSummary>=0 && preSummary<baseSummaryIndex,'summary text format must precede base setValues');
console.log('v6.65 VAT card text-safe output tests PASS');

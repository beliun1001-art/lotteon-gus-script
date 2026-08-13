/**
 * Issue #59 v1.0 guarded production apply.
 * Copies the fully adjudicated Issue54 corrected card-rematch preview to
 * production `부가세_카드매칭검증` only.
 *
 * Safety:
 * - exact pre-apply guards for corrected VAT, old production, Issue54 preview,
 *   and Issue55-58 diagnostic chain
 * - backup before destructive write
 * - write flag set immediately before production clear
 * - full matrix verification after copy
 * - protected-sheet signatures verified unchanged
 * - automatic rollback on any post-write error
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE59-v1.0-20260813',
  title: 'corrected VAT 1,355주문 카드매칭검증 운영 반영',
  enabled: true,
  statusSheet: 'ISSUE59_운영반영상태'
};

var I59_VERSION = 'v1.0-ISSUE59-CORRECTED-CARD-VERIFY-PRODUCTION';
var I59_PREVIEW = 'ISSUE54_카드매칭전체PREVIEW';
var I59_PREVIEW_STATUS = 'ISSUE54_실행상태';
var I59_DELTA_STATUS = 'ISSUE55_진단상태';
var I59_ADJ_STATUS = 'ISSUE56_판정상태';
var I59_DEEP_STATUS = 'ISSUE57_진단상태';
var I59_CAUSE_STATUS = 'ISSUE58_진단상태';
var I59_PROD = '부가세_카드매칭검증';
var I59_VAT = '부가세_신고자료';
var I59_PERIOD = '부가세_기간별';
var I59_HISTORY = '카드사용내역_붙여넣기';
var I59_MASTER = '카드_마스터';
var I59_BACKUP = 'ISSUE59_백업_부가세카드매칭검증';

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');

  var statusSh = i59Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  var prior = i59Kv_(statusSh);
  if (i59Text_(prior['상태']) === 'PASS') {
    var prodDone = ss.getSheetByName(I59_PROD);
    var previewDone = ss.getSheetByName(I59_PREVIEW);
    if (prodDone && previewDone && i59MatrixEqual_(previewDone, prodDone)) {
      return {ok:true, done:true, reason:'ALREADY_DONE'};
    }
    throw new Error('Issue59 상태는 PASS이나 운영 카드검증이 preview와 다릅니다. 자동 재실행 금지.');
  }

  i59WriteStatus_(statusSh, 'RUNNING', 'PRECHECK', '운영 반영 사전검증 시작', {
    operatingChange:'0', rollback:'0'
  });

  var wrote = false;
  var oldProdSig = '';
  var protectedBefore = null;

  try {
    var required = [
      I59_PREVIEW, I59_PREVIEW_STATUS, I59_DELTA_STATUS, I59_ADJ_STATUS,
      I59_DEEP_STATUS, I59_CAUSE_STATUS, I59_PROD, I59_VAT,
      I59_PERIOD, I59_HISTORY, I59_MASTER
    ];
    required.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 1) throw new Error('필수 시트 없음: ' + name);
    });

    var prod = ss.getSheetByName(I59_PROD);
    var preview = ss.getSheetByName(I59_PREVIEW);
    var vat = ss.getSheetByName(I59_VAT);
    var period = ss.getSheetByName(I59_PERIOD);
    var history = ss.getSheetByName(I59_HISTORY);
    var master = ss.getSheetByName(I59_MASTER);

    var staleBackup = ss.getSheetByName(I59_BACKUP);
    if (staleBackup) {
      var bst = i59DiagStats_(staleBackup);
      i59AssertDiagStats_('기존 Issue59 백업', bst, {
        orders:1355, matched:810, nonCard:494, ambiguous:1, noMatch:50, purchase:54807644
      });
      i59Restore_(staleBackup, prod);
      SpreadsheetApp.flush();
      ss.deleteSheet(staleBackup);
      prod = ss.getSheetByName(I59_PROD);
    }

    i59ExpectKv_(ss, I59_PREVIEW_STATUS, {
      '버전':'v1.1-ISSUE54-CORRECTED-CARD-REMATCH-PREVIEW-REBUILD',
      '상태':'PASS',
      '운영시트 변경':'0',
      '현재VAT주문':1355,
      'preview주문':1355,
      'canonical증빙행':1990,
      'MATCHED':808,
      'NON_CARD':498,
      'AMBIGUOUS':0,
      'NO_MATCH':49,
      'v6.69 2차귀속':1161,
      'v6.70 3차귀속':81,
      '주문매입금액합계':105762969,
      '잘못된카드identity':0,
      'fallback증빙필드오류':0,
      '부가세_카드매칭검증 변경':0
    });
    i59ExpectKv_(ss, I59_DELTA_STATUS, {
      '버전':'v1.0-ISSUE55-CARD-DELTA-DIAGNOSTIC',
      '상태':'PASS',
      '운영시트 변경':0,
      '기존운영주문':1355,
      'Issue54주문':1355,
      '정규화overlap':1355,
      'oldOnly':0,
      'newOnly':0,
      '상태변경주문':12,
      '상태동일주문':1343,
      '이동_AMBIGUOUS -> MATCHED':1,
      '이동_MATCHED -> NO_MATCH':5,
      '이동_NO_MATCH -> MATCHED':2,
      '이동_NO_MATCH -> NON_CARD':4,
      '부가세_카드매칭검증 변경':0,
      '부가세_신고자료 변경':0
    });
    i59ExpectKv_(ss, I59_ADJ_STATUS, {
      '버전':'v1.0-ISSUE56-CHANGED12-EVIDENCE-ADJUDICATION',
      '상태':'PASS',
      '운영시트 변경':0,
      '상태변경주문':12,
      'AUTO_SAFE':3,
      'REVIEW_REQUIRED':5,
      'INVALID':4,
      '신규MATCHED_AUTO_SAFE':3,
      '운영반영자동승인':'NO',
      '부가세_카드매칭검증 변경':0,
      '부가세_신고자료 변경':0
    });
    i59ExpectKv_(ss, I59_DEEP_STATUS, {
      '버전':'v1.0-ISSUE57-BLOCKED9-DEEP-DIAGNOSTIC',
      '상태':'PASS',
      '운영시트 변경':0,
      '진단대상':9,
      'EXPLAINED_SAFE':5,
      'LIKELY_MATCHER_BUG':4,
      'DATA_GAP_REVIEW':0,
      'INVALID_STATE':0,
      '부가세_카드매칭검증 변경':0,
      '부가세_신고자료 변경':0,
      '카드사용내역_붙여넣기 변경':0,
      '카드_마스터 변경':0
    });
    i59ExpectKv_(ss, I59_CAUSE_STATUS, {
      '버전':'v1.0-ISSUE58-PAYMENT-SOURCE-CAUSE-SPLIT',
      '상태':'PASS',
      '운영시트 변경':0,
      '진단대상':4,
      'DIAGNOSTIC_FALSE_POSITIVE':4,
      'UNRESOLVED':0,
      'OLD_PAYMENT_REUSE_BUG':0,
      'MATCHER_PAYMENT_PRIORITY_BUG':0,
      'BOTH_PAYMENT_AND_MATCHER':0,
      'ALLOCATION_CONFLICT':0,
      'CURRENT_VAT에서MATCHED':0,
      'NO_PAYMENT에서MATCHED':0,
      'PHYSICAL_ONLY에서MATCHED':0,
      'exactPhysical미사용후보존재':0,
      '부가세_카드매칭검증 변경':0,
      '부가세_신고자료 변경':0,
      '카드사용내역_붙여넣기 변경':0,
      '카드_마스터 변경':0
    });

    var vatStats = i59VatStats_(vat);
    if (vatStats.detailRows !== 2752) throw new Error('corrected VAT 상세행 불일치: ' + vatStats.detailRows);
    if (vatStats.orders !== 1355) throw new Error('corrected VAT 주문수 불일치: ' + vatStats.orders);
    if (Math.round(vatStats.purchase) !== 105762969) throw new Error('corrected VAT 매입합계 불일치: ' + vatStats.purchase);
    if (vatStats.unmappedBusiness !== 0) throw new Error('corrected VAT 사업자등록번호 미매핑: ' + vatStats.unmappedBusiness);

    var oldStats = i59DiagStats_(prod);
    i59AssertDiagStats_('기존 운영 카드검증', oldStats, {
      orders:1355, matched:810, nonCard:494, ambiguous:1, noMatch:50, purchase:54807644
    });

    var newStats = i59DiagStats_(preview);
    i59AssertDiagStats_('Issue54 preview', newStats, {
      orders:1355, matched:808, nonCard:498, ambiguous:0, noMatch:49,
      v669:1161, v670:81, purchase:105762969
    });
    if (newStats.duplicateKeys !== 0) throw new Error('Issue54 정규화 주문키 중복: ' + newStats.duplicateKeys);
    if (newStats.duplicateCanonicalMatched !== 0) {
      throw new Error('Issue54 MATCHED canonical key 중복: ' + newStats.duplicateCanonicalMatched);
    }
    if (oldStats.duplicateKeys !== 0) throw new Error('기존 운영 정규화 주문키 중복: ' + oldStats.duplicateKeys);
    var overlap = i59Overlap_(oldStats.keys, newStats.keys);
    if (overlap !== 1355) throw new Error('기존 운영/Issue54 정규화 overlap 불일치: ' + overlap);

    oldProdSig = i59SheetSig_(prod);
    protectedBefore = {
      vat:i59SheetSig_(vat),
      period:i59SheetSig_(period),
      history:i59SheetSig_(history),
      master:i59SheetSig_(master),
      preview:i59SheetSig_(preview),
      issue55:i59SheetSig_(ss.getSheetByName(I59_DELTA_STATUS)),
      issue56:i59SheetSig_(ss.getSheetByName(I59_ADJ_STATUS)),
      issue57:i59SheetSig_(ss.getSheetByName(I59_DEEP_STATUS)),
      issue58:i59SheetSig_(ss.getSheetByName(I59_CAUSE_STATUS))
    };

    i59WriteStatus_(statusSh, 'RUNNING', 'BACKUP', '사전검증 PASS; 운영 카드검증 백업 생성', {
      operatingChange:'0', rollback:'0',
      oldMatched:oldStats.matched, oldNonCard:oldStats.nonCard,
      oldAmbiguous:oldStats.ambiguous, oldNoMatch:oldStats.noMatch,
      oldPurchase:oldStats.purchase
    });

    var backup = prod.copyTo(ss).setName(I59_BACKUP);
    SpreadsheetApp.flush();
    if (i59SheetSig_(backup) !== oldProdSig) throw new Error('운영 카드검증 백업 signature 불일치');

    wrote = true;
    i59WriteStatus_(statusSh, 'RUNNING', 'WRITE', '백업 PASS; corrected 카드검증 운영 반영 시작', {
      operatingChange:'WRITE_IN_PROGRESS', rollback:'0'
    });

    i59CopySheetMatrix_(preview, prod);
    SpreadsheetApp.flush();

    if (!i59MatrixEqual_(preview, prod)) throw new Error('운영 카드검증과 Issue54 preview 전 셀 비교 불일치');

    var finalStats = i59DiagStats_(prod);
    i59AssertDiagStats_('최종 운영 카드검증', finalStats, {
      orders:1355, matched:808, nonCard:498, ambiguous:0, noMatch:49,
      v669:1161, v670:81, purchase:105762969
    });
    if (finalStats.duplicateKeys !== 0) throw new Error('최종 운영 정규화 주문키 중복: ' + finalStats.duplicateKeys);
    if (finalStats.duplicateCanonicalMatched !== 0) {
      throw new Error('최종 운영 MATCHED canonical key 중복: ' + finalStats.duplicateCanonicalMatched);
    }
    if (i59Overlap_(finalStats.keys, newStats.keys) !== 1355) throw new Error('최종 운영/preview 주문 overlap 불일치');

    i59AssertProtected_(ss, protectedBefore);

    i59WriteStatus_(statusSh, 'PASS', 'DONE', 'corrected VAT 카드매칭검증 운영 반영 및 검증 완료', {
      operatingChange:'부가세_카드매칭검증 1개 재작성',
      backup:I59_BACKUP,
      rollback:'0',
      orders:finalStats.orders,
      matched:finalStats.matched,
      nonCard:finalStats.nonCard,
      ambiguous:finalStats.ambiguous,
      noMatch:finalStats.noMatch,
      v669:finalStats.v669,
      v670:finalStats.v670,
      purchase:finalStats.purchase,
      overlap:1355,
      changedOrders:12,
      vatChange:'0',
      periodChange:'0',
      historyChange:'0',
      masterChange:'0'
    });
    return {ok:true, done:true, status:'PASS', stats:finalStats};
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    if (wrote) {
      try {
        var prodRb = ss.getSheetByName(I59_PROD);
        var backupRb = ss.getSheetByName(I59_BACKUP);
        if (!prodRb || !backupRb) throw new Error('롤백 대상/백업 시트 없음');
        i59Restore_(backupRb, prodRb);
        SpreadsheetApp.flush();
        var rbSig = i59SheetSig_(prodRb);
        if (oldProdSig && rbSig !== oldProdSig) throw new Error('롤백 후 운영 signature 불일치');
        var rbStats = i59DiagStats_(prodRb);
        i59AssertDiagStats_('롤백 운영 카드검증', rbStats, {
          orders:1355, matched:810, nonCard:494, ambiguous:1, noMatch:50, purchase:54807644
        });
        if (protectedBefore) i59AssertProtected_(ss, protectedBefore);
        i59WriteStatus_(statusSh, 'ROLLED_BACK', 'DONE', '운영 반영 오류로 자동 롤백 완료', {
          operatingChange:'0 (자동 롤백)',
          backup:I59_BACKUP,
          rollback:'1',
          error:msg,
          oldMatched:rbStats.matched,
          oldNonCard:rbStats.nonCard,
          oldAmbiguous:rbStats.ambiguous,
          oldNoMatch:rbStats.noMatch,
          oldPurchase:rbStats.purchase
        });
        return {ok:false, done:true, status:'ROLLED_BACK', error:msg};
      } catch (rb) {
        var rbMsg = String(rb && rb.message ? rb.message : rb);
        i59WriteStatus_(statusSh, 'ROLLBACK_ERROR', 'FAILED', '운영 반영 실패 후 자동 롤백도 실패', {
          operatingChange:'UNKNOWN',
          backup:I59_BACKUP,
          rollback:'ERROR',
          error:msg + ' / ROLLBACK: ' + rbMsg
        });
        throw new Error(msg + ' / ROLLBACK_ERROR: ' + rbMsg);
      }
    }

    i59WriteStatus_(statusSh, 'ERROR', 'FAILED', '운영 반영 전 사전검증 실패', {
      operatingChange:'0', rollback:'0', error:msg
    });
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemoteTaskStartRemote_();
}

function i59VatStats_(sheet) {
  var v = sheet.getDataRange().getValues();
  if (v.length < 2) throw new Error('부가세_신고자료가 비어 있습니다.');
  var h = v[0].map(i59Text_);
  var account = i59Find_(h, ['쿠팡계정ID']);
  var order = i59Find_(h, ['주문번호','마켓주문번호']);
  var purchase = i59Find_(h, ['매입금액']);
  var business = i59Find_(h, ['사업자등록번호']);
  if (account < 0 || order < 0 || purchase < 0 || business < 0) throw new Error('부가세_신고자료 필수 헤더 누락');
  var keys = {}, sum = 0, unmapped = 0;
  for (var r=1;r<v.length;r++) {
    var k = i59Key_(v[r][account], v[r][order]);
    if (!k) throw new Error('부가세_신고자료 주문키 공란 R' + (r+1));
    keys[k] = true;
    sum += i59Num_(v[r][purchase]);
    if (!i59Text_(v[r][business])) unmapped++;
  }
  return {detailRows:v.length-1, orders:Object.keys(keys).length, purchase:sum, unmappedBusiness:unmapped};
}

function i59DiagStats_(sheet) {
  var v = sheet.getDataRange().getValues();
  if (v.length < 2) throw new Error(sheet.getName() + ' 데이터가 없습니다.');
  var h = v[0].map(i59Text_);
  var ix = {
    account:i59Find_(h,['쿠팡계정ID']),
    order:i59Find_(h,['주문번호']),
    purchase:i59Find_(h,['주문매입금액','매입금액']),
    status:i59Find_(h,['카드매칭상태']),
    v669:i59Find_(h,['v6.69 2차귀속']),
    v670:i59Find_(h,['v6.70 3차귀속']),
    canonical:i59Find_(h,['canonicalEvidenceKey'])
  };
  if (ix.account<0 || ix.order<0 || ix.purchase<0 || ix.status<0) {
    throw new Error(sheet.getName() + ' 카드검증 필수 헤더 누락');
  }
  var st={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v669:0,v670:0,purchase:0,
    duplicateKeys:0,duplicateCanonicalMatched:0,keys:{}};
  var canonicalOwners={};
  for (var r=1;r<v.length;r++) {
    var row=v[r], key=i59Key_(row[ix.account],row[ix.order]);
    if (!key) throw new Error(sheet.getName() + ' 주문키 공란 R' + (r+1));
    if (st.keys[key]) st.duplicateKeys++;
    st.keys[key]=true;
    st.orders++;
    var s=i59Text_(row[ix.status]);
    if (s==='MATCHED' || s==='MASTER_MATCHED') st.matched++;
    else if (s==='NON_CARD') st.nonCard++;
    else if (s==='AMBIGUOUS') st.ambiguous++;
    else st.noMatch++;
    if (ix.v669>=0 && i59Text_(row[ix.v669])==='Y') st.v669++;
    if (ix.v670>=0 && i59Text_(row[ix.v670])==='Y') st.v670++;
    st.purchase += i59Num_(row[ix.purchase]);
    if ((s==='MATCHED' || s==='MASTER_MATCHED') && ix.canonical>=0) {
      var ck=i59Text_(row[ix.canonical]);
      if (ck) {
        canonicalOwners[ck]=(canonicalOwners[ck]||0)+1;
        if (canonicalOwners[ck]===2) st.duplicateCanonicalMatched++;
      }
    }
  }
  return st;
}

function i59AssertDiagStats_(label, actual, expected) {
  Object.keys(expected).forEach(function(k) {
    var a = Number(actual[k] || 0), w = Number(expected[k] || 0);
    if (Math.round(a) !== Math.round(w)) {
      throw new Error(label + ' ' + k + ' 불일치: 실제 ' + a + ' / 기대 ' + w);
    }
  });
}

function i59Overlap_(a, b) {
  var n=0;
  Object.keys(a || {}).forEach(function(k){ if (b && b[k]) n++; });
  return n;
}

function i59ExpectKv_(ss, sheetName, expected) {
  var sh=ss.getSheetByName(sheetName);
  if (!sh) throw new Error('상태 시트 없음: '+sheetName);
  var kv=i59Kv_(sh);
  Object.keys(expected).forEach(function(key){
    var want=expected[key], actual=kv[key];
    if (typeof want==='number') {
      if (Math.round(i59Num_(actual))!==want) throw new Error(sheetName+' '+key+' 불일치: '+actual+' / 기대 '+want);
    } else if (i59Text_(actual)!==String(want)) {
      throw new Error(sheetName+' '+key+' 불일치: '+actual+' / 기대 '+want);
    }
  });
}

function i59Kv_(sheet) {
  var out={};
  if (!sheet || sheet.getLastRow()<1) return out;
  var vals=sheet.getRange(1,1,sheet.getLastRow(),Math.min(2,sheet.getLastColumn())).getValues();
  vals.forEach(function(r){var k=i59Text_(r[0]); if(k)out[k]=r[1];});
  return out;
}

function i59AssertProtected_(ss, before) {
  var now = {
    vat:i59SheetSig_(ss.getSheetByName(I59_VAT)),
    period:i59SheetSig_(ss.getSheetByName(I59_PERIOD)),
    history:i59SheetSig_(ss.getSheetByName(I59_HISTORY)),
    master:i59SheetSig_(ss.getSheetByName(I59_MASTER)),
    preview:i59SheetSig_(ss.getSheetByName(I59_PREVIEW)),
    issue55:i59SheetSig_(ss.getSheetByName(I59_DELTA_STATUS)),
    issue56:i59SheetSig_(ss.getSheetByName(I59_ADJ_STATUS)),
    issue57:i59SheetSig_(ss.getSheetByName(I59_DEEP_STATUS)),
    issue58:i59SheetSig_(ss.getSheetByName(I59_CAUSE_STATUS))
  };
  Object.keys(before || {}).forEach(function(k){
    if (before[k] !== now[k]) throw new Error('보호시트 signature 변경: '+k);
  });
}

function i59CopySheetMatrix_(src, dst) {
  var rows=src.getLastRow(), cols=src.getLastColumn();
  if (rows<1 || cols<1) throw new Error('복사 원본이 비어 있습니다: '+src.getName());
  i59EnsureGrid_(dst, rows, cols);
  var clearRows=Math.max(dst.getLastRow(),rows);
  var clearCols=Math.max(dst.getLastColumn(),cols);
  if (clearRows>0 && clearCols>0) dst.getRange(1,1,clearRows,clearCols).clear();
  src.getRange(1,1,rows,cols).copyTo(dst.getRange(1,1,rows,cols), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  dst.setFrozenRows(src.getFrozenRows());
  dst.setFrozenColumns(src.getFrozenColumns());
  for(var c=1;c<=cols;c++) {
    try { dst.setColumnWidth(c,src.getColumnWidth(c)); } catch(ignore) {}
  }
}

function i59Restore_(backup, prod) {
  i59CopySheetMatrix_(backup, prod);
}

function i59EnsureGrid_(sheet, rows, cols) {
  if (sheet.getMaxRows()<rows) sheet.insertRowsAfter(sheet.getMaxRows(),rows-sheet.getMaxRows());
  if (sheet.getMaxColumns()<cols) sheet.insertColumnsAfter(sheet.getMaxColumns(),cols-sheet.getMaxColumns());
}

function i59MatrixEqual_(a, b) {
  if (!a || !b) return false;
  var ar=a.getLastRow(), ac=a.getLastColumn(), br=b.getLastRow(), bc=b.getLastColumn();
  if (ar!==br || ac!==bc) return false;
  var av=a.getRange(1,1,ar,ac).getValues();
  var bv=b.getRange(1,1,br,bc).getValues();
  for(var r=0;r<ar;r++) for(var c=0;c<ac;c++) {
    if (i59Cell_(av[r][c])!==i59Cell_(bv[r][c])) return false;
  }
  return true;
}

function i59SheetSig_(sheet) {
  if (!sheet) return 'MISSING';
  var rows=sheet.getLastRow(), cols=sheet.getLastColumn();
  if (!rows || !cols) return 'EMPTY|'+rows+'|'+cols;
  var vals=sheet.getRange(1,1,rows,cols).getValues();
  var h1=2166136261, h2=0;
  for(var r=0;r<vals.length;r++) {
    for(var c=0;c<vals[r].length;c++) {
      var s=i59Cell_(vals[r][c])+'\u001f';
      for(var i=0;i<s.length;i++) {
        h1 ^= s.charCodeAt(i);
        h1 = Math.imul(h1,16777619);
        h2 = (Math.imul(h2,31)+s.charCodeAt(i))|0;
      }
    }
    h2=(Math.imul(h2,131)+13)|0;
  }
  return rows+'x'+cols+'|'+(h1>>>0).toString(16)+'|'+(h2>>>0).toString(16);
}

function i59Cell_(v) {
  if (Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v.getTime())) return 'D:'+v.toISOString();
  if (typeof v==='number') return 'N:'+String(v);
  if (typeof v==='boolean') return 'B:'+String(v);
  return 'T:'+i59Text_(v);
}

function i59WriteStatus_(sheet, status, stage, message, x) {
  x=x||{};
  var rows=[
    ['항목','값'],
    ['버전',I59_VERSION],
    ['상태',status],
    ['단계',stage],
    ['메시지',message],
    ['운영시트 변경',x.operatingChange||'0'],
    ['백업시트',x.backup||''],
    ['운영주문',x.orders||0],
    ['MATCHED',x.matched||0],
    ['NON_CARD',x.nonCard||0],
    ['AMBIGUOUS',x.ambiguous||0],
    ['NO_MATCH',x.noMatch||0],
    ['v6.69 2차귀속',x.v669||0],
    ['v6.70 3차귀속',x.v670||0],
    ['주문매입금액합계',x.purchase||0],
    ['정규화overlap',x.overlap||0],
    ['상태변경검증',x.changedOrders||0],
    ['기존_MATCHED',x.oldMatched||0],
    ['기존_NON_CARD',x.oldNonCard||0],
    ['기존_AMBIGUOUS',x.oldAmbiguous||0],
    ['기존_NO_MATCH',x.oldNoMatch||0],
    ['기존매입합계',x.oldPurchase||0],
    ['부가세_신고자료 변경',x.vatChange||'0'],
    ['부가세_기간별 변경',x.periodChange||'0'],
    ['카드사용내역_붙여넣기 변경',x.historyChange||'0'],
    ['카드_마스터 변경',x.masterChange||'0'],
    ['롤백',x.rollback||'0'],
    ['오류',x.error||''],
    ['완료시각',(status==='PASS'||status==='ROLLED_BACK'||status==='ERROR'||status==='ROLLBACK_ERROR')?new Date().toISOString():''],
    ['갱신시각',new Date().toISOString()]
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setFontWeight('bold');
  sheet.setColumnWidth(1,230);
  sheet.setColumnWidth(2,650);
}

function i59Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function i59Find_(headers,names){for(var i=0;i<names.length;i++){var x=headers.indexOf(names[i]);if(x>=0)return x;}return -1;}
function i59Text_(v){return String(v==null?'':v).trim();}
function i59Num_(v){var n=Number(typeof v==='number'?v:i59Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}
function i59NormOrder_(v){
  if(typeof v==='number'&&isFinite(v))return String(Math.trunc(v));
  var s=i59Text_(v).replace(/[,\s]/g,'');
  if(/^\d+\.0+$/.test(s))s=s.replace(/\.0+$/,'');
  return s;
}
function i59Key_(account,order){
  var a=i59Text_(account).toLowerCase(),o=i59NormOrder_(order);
  return a&&o?a+'|'+o:'';
}

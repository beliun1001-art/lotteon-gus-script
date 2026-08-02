/**
 * PR #33 operating preview smoke v1.0.
 * Reads current production card diagnostic only and writes PR33_* preview sheets.
 * Runs in 150-row automatic continuation batches.
 */
const PR33_PREVIEW_VERSION = 'v1.0-PR33-V670-DIAG-BATCHED';
const PR33_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-32-vat-date-window-fallback/Patch_v6_70_vat_date_window_fallback.gs';
const PR33_SOURCE_SHEET = '부가세_카드매칭검증';
const PR33_PREVIEW_SHEET = 'PR33_카드매칭검증';
const PR33_SUMMARY_SHEET = 'PR33_귀속요약';
const PR33_STATUS_SHEET = 'PR33_실행상태';
const PR33_BATCH_SIZE = 150;
const PR33_CONTINUE_HANDLER = 'runPr33PreviewSmokeContinue';
const PR33_SPREADSHEET_ID_KEY = 'PR33_PREVIEW_SPREADSHEET_ID';

function runPr33PreviewSmoke() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  const source = ss.getSheetByName(PR33_SOURCE_SHEET);
  if (!source || source.getLastRow() < 2) throw new Error(PR33_SOURCE_SHEET + ' 시트가 없습니다.');

  pr33DeleteTriggers_();
  PropertiesService.getScriptProperties().setProperty(PR33_SPREADSHEET_ID_KEY, ss.getId());
  [PR33_PREVIEW_SHEET, PR33_SUMMARY_SHEET, PR33_STATUS_SHEET].forEach(function(name) {
    const old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
  });

  const headers = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0].map(pr33Text_);
  pr33RequireHeaders_(headers, ['신고연도','반기','주문일','사업자등록번호','주문번호','롯데결제수단','주문매입금액','카드매칭상태','카드매칭근거','카드번호끝4']);
  if (headers.indexOf('v6.70 3차귀속') < 0) headers.push('v6.70 3차귀속');

  const preview = ss.insertSheet(PR33_PREVIEW_SHEET);
  preview.getRange(1, 1, 1, headers.length).setValues([headers]);
  preview.setFrozenRows(1);
  preview.getRange(1, 1, 1, headers.length).setBackground('#d9eaf7').setFontWeight('bold');

  pr33WriteStatus_(ss, {
    status:'RUNNING', stage:'BATCH', message:'초기화 완료; 1차 배치 실행',
    processed:0, target:source.getLastRow()-1, startedAt:new Date().toISOString()
  });
  return runPr33PreviewSmokeContinue();
}

function runPr33PreviewSmokeContinue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return {ok:false,reason:'LOCK_BUSY'};
  try {
    const id = PropertiesService.getScriptProperties().getProperty(PR33_SPREADSHEET_ID_KEY);
    const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
    if (!ss) throw new Error('PR33 대상 스프레드시트를 찾지 못했습니다.');
    const source = ss.getSheetByName(PR33_SOURCE_SHEET);
    const preview = ss.getSheetByName(PR33_PREVIEW_SHEET);
    if (!source || !preview) throw new Error('PR33 원본 또는 미리보기 시트가 없습니다.');

    pr33LoadPatch_();
    const sourceHeaders = source.getRange(1,1,1,source.getLastColumn()).getValues()[0].map(pr33Text_);
    const previewHeaders = preview.getRange(1,1,1,preview.getLastColumn()).getValues()[0].map(pr33Text_);
    const ix = pr33HeaderIndex_(sourceHeaders);
    const processed = Math.max(0, preview.getLastRow()-1);
    const target = source.getLastRow()-1;
    const size = Math.min(PR33_BATCH_SIZE, target-processed);

    if (size > 0) {
      const values = source.getRange(processed+2,1,size,sourceHeaders.length).getValues();
      const out = values.map(function(row) { return pr33TransformRow_(row, sourceHeaders, previewHeaders, ix); });
      preview.getRange(preview.getLastRow()+1,1,out.length,previewHeaders.length).setValues(out);
    }

    const next = processed + Math.max(0,size);
    pr33WriteStatus_(ss, {
      status:next >= target ? 'FINALIZING' : 'RUNNING',
      stage:next >= target ? 'FINALIZE' : 'BATCH',
      message:next >= target ? '전체 배치 완료; 최종 검증 중' : '다음 배치 자동 예약',
      processed:next,target:target,startedAt:pr33ReadStatus_(ss,'시작시각') || new Date().toISOString()
    });
    SpreadsheetApp.flush();

    if (next >= target) {
      const result = pr33Finalize_(ss);
      pr33DeleteTriggers_();
      PropertiesService.getScriptProperties().deleteProperty(PR33_SPREADSHEET_ID_KEY);
      return result;
    }
    pr33Schedule_();
    return {ok:true,done:false,processed:next,target:target};
  } catch (error) {
    try {
      const id = PropertiesService.getScriptProperties().getProperty(PR33_SPREADSHEET_ID_KEY);
      const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
      if (ss) pr33WriteStatus_(ss, {
        status:'ERROR',stage:'ERROR',message:String(error && error.message ? error.message : error),
        processed:Math.max(0,(ss.getSheetByName(PR33_PREVIEW_SHEET)||{getLastRow:function(){return 1;}}).getLastRow()-1),
        target:Math.max(0,(ss.getSheetByName(PR33_SOURCE_SHEET)||{getLastRow:function(){return 1;}}).getLastRow()-1),
        startedAt:pr33ReadStatus_(ss,'시작시각') || ''
      });
    } catch (ignore) {}
    pr33DeleteTriggers_();
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function pr33TransformRow_(sourceRow, sourceHeaders, previewHeaders, ix) {
  const map = {};
  sourceHeaders.forEach(function(h,i){map[h]=sourceRow[i];});
  const status = pr33Text_(map['카드매칭상태']);
  const payment = pr33Text_(map['롯데결제수단']);
  const orderDate = pr33DateKey_(map['주문일']);
  const rule = typeof vatDateWindowRule_v670_ === 'function'
    ? vatDateWindowRule_v670_({orderDate:orderDate,lottePayment:payment},{status:status})
    : null;

  if (rule) {
    const card = pr33Card_(rule.end4);
    if (!card) throw new Error('PR33 알 수 없는 카드 식별자: ' + rule.end4);
    map['구매카드사'] = card.company;
    map['구매카드별칭'] = '';
    map['구매카드명'] = card.name;
    map['카드번호'] = card.number;
    map['카드번호끝4'] = card.end4;
    map['승인일'] = '';
    map['승인시각'] = '';
    map['승인번호'] = '';
    map['승인금액'] = 0;
    map['카드매칭상태'] = 'MATCHED';
    map['카드매칭근거'] = '트래킹번호_일자구간단일카드_3차귀속_' + rule.code + '_금액비교없음';
    map['후보수'] = 1;
    map['가맹점명'] = '';
    map['가맹점주문번호'] = '';
    map['증빙유형'] = '일자구간단일카드_트래킹귀속';
    map['취소/부분취소메모'] = '';
    map['원본파일'] = card.source;
    map['후보요약'] = card.company + ' / ' + card.name + ' / ' + card.end4 + ' / ' + rule.code;
    map['v6.70 3차귀속'] = 'Y';
  } else {
    map['v6.70 3차귀속'] = '';
  }

  return previewHeaders.map(function(h){return map[h] === undefined ? '' : map[h];});
}

function pr33Finalize_(ss) {
  const source = ss.getSheetByName(PR33_SOURCE_SHEET);
  const preview = ss.getSheetByName(PR33_PREVIEW_SHEET);
  const values = preview.getDataRange().getValues();
  const headers = values[0].map(pr33Text_);
  const ix = pr33HeaderIndex_(headers);
  const stats = {orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,v670:0,invalidIdentity:0,invalidEvidence:0,purchase:0};
  const groups = {};

  values.slice(1).forEach(function(row) {
    const status = pr33Text_(row[ix['카드매칭상태']]);
    const end4 = pr33Text_(row[ix['카드번호끝4']]).replace(/\D/g,'').padStart(4,'0').slice(-4);
    const company = pr33Text_(row[ix['구매카드사']]);
    const cardName = pr33Text_(row[ix['구매카드명']]);
    const purchase = pr33Number_(row[ix['주문매입금액']]);
    const v670 = pr33Text_(row[ix['v6.70 3차귀속']]) === 'Y';
    stats.orders++; stats.purchase += purchase;
    if (status === 'MATCHED' || status === 'MASTER_MATCHED') stats.matched++;
    else if (status === 'NON_CARD') stats.nonCard++;
    else if (status === 'AMBIGUOUS') stats.ambiguous++;
    else stats.noMatch++;

    if (v670) {
      stats.v670++;
      if ((company === 'KB국민카드' && end4 !== '4091') ||
          (company === '우리카드' && end4 !== '7680') ||
          (company === '롯데카드' && end4 !== '0036')) stats.invalidIdentity++;
      if (pr33Text_(row[ix['승인일']]) || pr33Text_(row[ix['승인번호']]) || pr33Number_(row[ix['승인금액']]) !== 0 ||
          pr33Text_(row[ix['카드매칭근거']]).indexOf('3차귀속') < 0) stats.invalidEvidence++;
    }

    const business = pr33Text_(row[ix['사업자등록번호']]);
    const key = [business,company,cardName,end4,status].join('|');
    if (!groups[key]) groups[key] = {business:business,company:company,cardName:cardName,end4:end4,status:status,orders:0,purchase:0};
    groups[key].orders++; groups[key].purchase += purchase;
  });

  const expected = {orders:1355,matched:810,nonCard:494,ambiguous:1,noMatch:50,v670:44,purchase:54807644};
  Object.keys(expected).forEach(function(key) {
    const actual = Math.round(Number(stats[key]||0));
    if (actual !== expected[key]) throw new Error('PR33 검증 실패: ' + key + ' 실제 ' + actual + ' / 기대 ' + expected[key]);
  });
  if (stats.invalidIdentity) throw new Error('PR33 잘못된 카드 식별자 ' + stats.invalidIdentity + '건');
  if (stats.invalidEvidence) throw new Error('PR33 3차귀속 증빙필드 오류 ' + stats.invalidEvidence + '건');
  if (source.getLastRow()-1 !== stats.orders) throw new Error('PR33 원본/미리보기 행 수 불일치');

  const summaryHeaders = ['사업자등록번호','구매카드사','구매카드명','카드번호끝4','카드매칭상태','주문건수','주문매입금액'];
  const summaryRows = Object.keys(groups).sort().map(function(k){
    const g=groups[k]; return [g.business,g.company,g.cardName,g.end4,g.status,g.orders,g.purchase];
  });
  const old = ss.getSheetByName(PR33_SUMMARY_SHEET);
  if (old) ss.deleteSheet(old);
  const summary = ss.insertSheet(PR33_SUMMARY_SHEET);
  summary.getRange(1,1,1,summaryHeaders.length).setValues([summaryHeaders]);
  if (summaryRows.length) summary.getRange(2,1,summaryRows.length,summaryHeaders.length).setValues(summaryRows);
  summary.setFrozenRows(1);
  summary.getRange(1,1,1,summaryHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');
  if (summaryRows.length) {
    summary.getRange(2,4,summaryRows.length,1).setNumberFormat('@');
    summary.getRange(2,6,summaryRows.length,2).setNumberFormat('#,##0');
  }

  pr33WriteFinalStatus_(ss, stats);
  SpreadsheetApp.flush();
  return {ok:true,done:true,version:PR33_PREVIEW_VERSION,stats:stats};
}

function pr33Card_(end4) {
  const cards = {
    '4091':{company:'KB국민카드',name:'HERITAGE Smart(할인형)',number:'5598-69**-****-4091',end4:'4091',source:'카드이용내역_KB 4091(1).xls'},
    '7680':{company:'우리카드',name:'카드의정석 EVERY POINT',number:'7680',end4:'7680',source:'우리카드(1).xls'},
    '0036':{company:'롯데카드',name:'LOCA LIKIT 1.2',number:'3762-776436-56036',end4:'0036',source:'카드이용내역__LOCA LIKIT 1.2(506)(1).xls'}
  };
  return cards[String(end4||'')] || null;
}

function pr33LoadPatch_() {
  const res = UrlFetchApp.fetch(PR33_PATCH_URL + '?ts=' + new Date().getTime(), {muteHttpExceptions:true,followRedirects:true});
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error('PR33 patch load HTTP ' + res.getResponseCode());
  eval(res.getContentText('UTF-8'));
  if (typeof vatDateWindowRule_v670_ !== 'function') throw new Error('v6.70 rule function load 실패');
}

function pr33WriteFinalStatus_(ss, stats) {
  const s = ss.getSheetByName(PR33_STATUS_SHEET) || ss.insertSheet(PR33_STATUS_SHEET);
  const rows = [
    ['항목','값'],['버전',PR33_PREVIEW_VERSION],['상태','PASS'],['단계','DONE'],
    ['메시지','3차귀속 배치 미리보기 및 검증 완료'],['운영시트 변경','없음'],
    ['상반기 주문',stats.orders],['MATCHED',stats.matched],['NON_CARD',stats.nonCard],
    ['AMBIGUOUS',stats.ambiguous],['NO_MATCH',stats.noMatch],['v6.70 3차귀속',stats.v670],
    ['주문매입금액',stats.purchase],['잘못된 카드 식별자',stats.invalidIdentity],
    ['3차귀속 증빙필드 오류',stats.invalidEvidence],['완료시각',new Date().toISOString()]
  ];
  s.clearContents(); s.getRange(1,1,rows.length,2).setValues(rows); s.setFrozenRows(1);
  s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  s.setColumnWidth(1,220); s.setColumnWidth(2,520);
}

function pr33WriteStatus_(ss, state) {
  const s = ss.getSheetByName(PR33_STATUS_SHEET) || ss.insertSheet(PR33_STATUS_SHEET);
  const rows = [
    ['항목','값'],['버전',PR33_PREVIEW_VERSION],['상태',state.status||'RUNNING'],
    ['단계',state.stage||'BATCH'],['메시지',state.message||''],
    ['처리행',Number(state.processed||0)],['대상행',Number(state.target||0)],
    ['오류',state.status==='ERROR' ? state.message||'' : ''],
    ['시작시각',state.startedAt||''],['갱신시각',new Date().toISOString()]
  ];
  s.clearContents(); s.getRange(1,1,rows.length,2).setValues(rows); s.setFrozenRows(1);
  s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  s.setColumnWidth(1,220); s.setColumnWidth(2,520);
}

function pr33ReadStatus_(ss, key) {
  const s=ss.getSheetByName(PR33_STATUS_SHEET); if(!s||s.getLastRow()<2)return'';
  const v=s.getDataRange().getValues(); for(let i=1;i<v.length;i++)if(pr33Text_(v[i][0])===key)return v[i][1]; return'';
}
function pr33HeaderIndex_(headers) { const out={}; headers.forEach(function(h,i){out[pr33Text_(h)]=i;}); return out; }
function pr33RequireHeaders_(headers, required) { required.forEach(function(h){if(headers.indexOf(h)<0)throw new Error('필수 헤더 없음: '+h);}); }
function pr33Text_(v) { return String(v==null?'':v).trim(); }
function pr33Number_(v) { const n=Number(typeof v==='number'?v:pr33Text_(v).replace(/[,원\s]/g,'')); return isNaN(n)?0:n; }
function pr33DateKey_(v) {
  if (Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v.getTime())) return Utilities.formatDate(v, Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');
  const s=pr33Text_(v); const iso=s.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(iso)return iso[1]+'-'+('0'+iso[2]).slice(-2)+'-'+('0'+iso[3]).slice(-2);
  const d=new Date(s); return isNaN(d.getTime())?'':Utilities.formatDate(d,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');
}
function pr33Schedule_() { pr33DeleteTriggers_(); ScriptApp.newTrigger(PR33_CONTINUE_HANDLER).timeBased().after(60*1000).create(); }
function pr33DeleteTriggers_() { ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()===PR33_CONTINUE_HANDLER)ScriptApp.deleteTrigger(t);}); }

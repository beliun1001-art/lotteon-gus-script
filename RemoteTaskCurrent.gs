/**
 * Issue #54 v1.1 corrected VAT 1,355-order read-only card rematch preview REBUILD.
 *
 * Reuses the previously operating-tested Issue #51 v1.0 batch runner from an
 * immutable historical commit and applies only the corrected VAT hard guards:
 * 2,752 detail rows / 1,355 orders / purchase 105,762,969 / VAT-only 0.
 *
 * v1.1 uses a new task id so a stale prior DONE state cannot cause ALREADY_DONE
 * when the ISSUE54 preview sheet has been removed.
 *
 * Production sheets are not modified. Only ISSUE54_* preview/status sheets are written.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE54-v1.1-REBUILD-20260813',
  title: 'corrected VAT 1,355주문 카드매칭 전체 read-only 재검증 REBUILD',
  enabled: true,
  outputSheet: 'ISSUE54_카드매칭전체PREVIEW',
  statusSheet: 'ISSUE54_실행상태'
};

var ISSUE54_BASE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/819fb655be0672ddc88c241e4a45247c7679f670/RemoteTaskCurrent.gs';

function runLotteonRemoteTaskStartRemote_() {
  var src = issue54LoadTransformedBase_();
  return eval(src + '\n;issue54RemoteStartImpl_();');
}

function runLotteonRemoteTaskContinueRemote_() {
  var src = issue54LoadTransformedBase_();
  return eval(src + '\n;issue54RemoteContinueImpl_();');
}

function issue54LoadTransformedBase_() {
  var res = UrlFetchApp.fetch(ISSUE54_BASE_URL + '?ts=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  var code = res.getResponseCode();
  var src = res.getContentText('UTF-8');
  if (code < 200 || code >= 300) throw new Error('Issue54 base runner 로드 실패 HTTP ' + code);

  issue54Require_(src.indexOf('Issue #51 v1.0 read-only full card rematch preview') >= 0, 'Issue51 base runner 식별 실패');
  issue54Require_(src.indexOf('pre.vatRows === 3894') >= 0, 'Issue51 VAT row guard 원본 불일치');
  issue54Require_(src.indexOf('pre.vatOrders === 1893') >= 0, 'Issue51 VAT order guard 원본 불일치');
  issue54Require_(src.indexOf('106707957') >= 0, 'Issue51 purchase guard 원본 불일치');
  issue54Require_(src.indexOf('pre.currentOnly === 538') >= 0, 'Issue51 currentOnly guard 원본 불일치');
  issue54Require_(src.indexOf('stats.oldOrders===1355 && stats.newOrders===538') >= 0, 'Issue51 old/new final guard 원본 불일치');

  src = issue54ReplaceAll_(src, 'runLotteonRemoteTaskStartRemote_', 'issue54RemoteStartImpl_');
  src = issue54ReplaceAll_(src, 'runLotteonRemoteTaskContinueRemote_', 'issue54RemoteContinueImpl_');

  src = issue54ReplaceAll_(src, 'ISSUE51', 'ISSUE54');
  src = issue54ReplaceAll_(src, 'issue51', 'issue54');

  src = issue54ReplaceAll_(src, 'v1.0-ISSUE54-CARD-REMATCH-PREVIEW', 'v1.1-ISSUE54-CORRECTED-CARD-REMATCH-PREVIEW-REBUILD');
  src = issue54ReplaceAll_(src, '현재 VAT 1,893주문 카드매칭 전체 preview', 'corrected VAT 1,355주문 카드매칭 전체 preview REBUILD');
  src = issue54ReplaceAll_(src, '현재 VAT 1,893주문 전체 카드매칭 preview 사전검증 시작', 'corrected VAT 1,355주문 카드매칭 preview 재생성 사전검증 시작');
  src = issue54ReplaceAll_(src, '현재 VAT 1,893주문 카드매칭 전체 preview 재계산 완료', 'corrected VAT 1,355주문 카드매칭 전체 preview 재생성 완료');

  src = issue54ReplaceAll_(src, '3894', '2752');
  src = issue54ReplaceAll_(src, '1893', '1355');
  src = issue54ReplaceAll_(src, '106707957', '105762969');
  src = issue54ReplaceAll_(src, 'pre.currentOnly === 538', 'pre.currentOnly === 0');
  src = issue54ReplaceAll_(src, "['신규주문',538]", "['신규주문',0]");
  src = issue54ReplaceAll_(src, 'stats.oldOrders===1355 && stats.newOrders===538', 'stats.oldOrders===1355 && stats.newOrders===0');
  src = issue54ReplaceAll_(src, '신규538', 'VATonly');

  issue54Require_(src.indexOf("id: 'ISSUE54-v1.0-20260812'") >= 0, 'Issue54 transformed task id 확인 실패');
  src = issue54ReplaceAll_(src, "id: 'ISSUE54-v1.0-20260812'", "id: 'ISSUE54-v1.1-REBUILD-20260813'");

  issue54Require_(src.indexOf('pre.vatRows === 2752') >= 0, 'Issue54 VAT row guard 변환 실패');
  issue54Require_(src.indexOf('pre.vatOrders === 1355') >= 0, 'Issue54 VAT order guard 변환 실패');
  issue54Require_(src.indexOf('105762969') >= 0, 'Issue54 purchase guard 변환 실패');
  issue54Require_(src.indexOf('pre.currentOnly === 0') >= 0, 'Issue54 VAT-only guard 변환 실패');
  issue54Require_(src.indexOf('stats.oldOrders===1355 && stats.newOrders===0') >= 0, 'Issue54 final old/new guard 변환 실패');
  issue54Require_(src.indexOf('ISSUE54_카드매칭전체PREVIEW') >= 0, 'Issue54 preview sheet 변환 실패');
  issue54Require_(src.indexOf('ISSUE54_V1_STATE') >= 0, 'Issue54 state key 변환 실패');
  issue54Require_(src.indexOf("id: 'ISSUE54-v1.1-REBUILD-20260813'") >= 0, 'Issue54 rebuild task id 변환 실패');

  return src;
}

function issue54ReplaceAll_(src, from, to) {
  return String(src).split(from).join(to);
}

function issue54Require_(ok, message) {
  if (!ok) throw new Error(message);
}

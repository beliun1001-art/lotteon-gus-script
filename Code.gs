/**
 * LOTTEON 롯백 Google Sheets 자동 분석 v6.00 SAFE CHANGE APPLY NO DASHBOARD FINAL
 *
 * v6.00 변경사항 반영 초안전화 + 필터갱신 1페이지 안전실행:
 *  - ① 변경사항 반영 실행 시간초과 방지
 *    1) 매출데이터 정리/매출분석/핵심_브랜드요약 중심으로 실행
 *    2) 대시보드는 시간 여유가 있을 때만 실행하고, 길어지면 ⑤ 메뉴로 분리
 *    3) 수동입력/API검증, 필터_대시보드, 필터별_상품수 API, 검수리포트는 ①에서 자동 제외
 *  - v5.97 운영 시트 7개 표시/전체 시트 복구 메뉴 반영
 *  - v5.63 신규: 검수리포트 시트/메뉴 추가
 *    1) 매출데이터_붙여넣기, 매출데이터_정리, 핵심_브랜드요약, 쿠팡재전송_로그 간 브랜드별 매출 차이 자동 검수
 *    2) 엄브로처럼 원본 매출은 있으나 요약/로그 매출이 0 또는 불일치하는 브랜드 자동 탐지
 *    3) 02 필터 잔존, 수동입력 누락, 수동/API 검증 문제행을 한 시트에 요약
 *
 *  - 매출 브랜드 매칭 보강
 *    1) LOTTEON_상품목록에 없는 브랜드라도 매출데이터_붙여넣기 '브랜드' 컬럼을 신뢰하여 브랜드 매출에 반영
 *    2) 필터별_상품수와 쿠팡전송수_수동입력의 브랜드도 매칭 후보에 포함
 *    3) 상품명/마켓상품명에 포함된 브랜드명으로 최종 추정 매칭 추가
 *  - v5.60의 필터/상품수 빠른 갱신 분할 실행 구조 유지
 *  - v5.57의 대시보드 조치요약판 구조 유지
 *  - v5.56의 수동입력 시트 시간초과 방지 구조 유지
 *  - v5.55의 수동입력/API 검증 문제행 출력 구조 유지
 */


const CONFIG = {
  BASE_URL_DEFAULT: 'https://tmg2007.cafe24.com',
  API_GATEWAY_BASE_URL: 'https://apihub.theapi.biz',
  USE_API_GATEWAY_FALLBACK: false,
  SITE_ID: 'LOTTEON',
  TARGET_SOURCE: 'LOTTEON.com',
  FILTER_PREFIX: '롯백',
  TIMEZONE: 'Asia/Seoul',
  MAX_RUN_MS: 180000,
  MAX_PAGES_PER_RUN: 60,
  MAX_FILTER_PAGES_PER_RUN: 20,
  REQUEST_DELAY_MS: 360,
  FAST_COUNT_BATCH_SIZE: 25,
  FAST_COUNT_CHUNK_SIZE: 3,
  TRIGGER_AFTER_MS: 30 * 1000,
  TRANSIENT_RETRY_AFTER_MS: 2 * 60 * 1000,
  MAX_TRANSIENT_URLFETCH_ERRORS: 10,
  ANALYSIS_AFTER_MS: 2 * 60 * 1000,
  MAX_PAGE_PER_QUERY: 500,

  API_TOKEN_DAILY_LIMIT: 10000,
  API_TOKEN_MONTHLY_LIMIT: 100000,
  API_TOKEN_BASELINE_DAILY_DATE: '2026-05-18',
  API_TOKEN_BASELINE_DAILY_USED: 3,
  API_TOKEN_BASELINE_MONTH: '2026-05',
  API_TOKEN_BASELINE_MONTH_USED: 4504,
  API_TOKEN_COST_MAP: { filterList: 1, productList: 1, marketList: 1 },

  MIN_OBSERVE_DAYS: 30,
  DELETE_REVIEW_DAYS: 45,
  LARGE_PRODUCT_COUNT: 300,
  VERY_LARGE_PRODUCT_COUNT: 1000,
  GOOD_REVENUE_30D: 1000000,
  GOOD_ORDER_30D: 10,
  MIN_SELL_RATE_KEEP: 0.005,

  ROTATION_STAGE1_DAYS: 7,
  ROTATION_STAGE2_DAYS: 30,
  ROTATION_STAGE3_DAYS: 60,

  ACCOUNT_MAP: {
    '01': { accountNo: 1, accountId: 'beliun1021' },
    '02': { accountNo: 2, accountId: 'beliun1024' },
    '03': { accountNo: 3, accountId: 'beliun1023' },
    '04': { accountNo: 4, accountId: 'beliun1024' }
  },
  ACCOUNT_ID_PREFIX_MAP: { 'beliun1021': 1, 'beliun1024': 2, 'beliun1023': 3 },

  LOTTEON_SITE_REGEX: /lotteon/i,
  CANCEL_PATTERN: /취소|반품|교환|환불|취소신청|반품신청|반품\/교환\/취소|cancel|return|refund/i,
  LOTTEON_PRODUCT_CODE_REGEX: /(LE\d{6,}|PD\d{6,})/i,
  PERSONAL_INFO_BLOCKLIST: ['수령인명','배송주소','상세주소','우편번호','휴대폰번호','통관고유부호','택배사요청메시지'],

  SHEETS: {
    DASHBOARD: '대시보드', PRODUCTS: 'LOTTEON_상품목록', PRODUCTS_TMP: 'LOTTEON_상품목록_TMP',
    FILTERS: '필터별_상품수', MARKETS: '마켓목록_API', SALES_IN: '매출데이터_붙여넣기',
    SALES_CLEAN: '매출데이터_정리', MONTHLY_BRAND: '월별_브랜드성과',
    BRAND: '브랜드별_성과분석', BRAND_SUMMARY: '핵심_브랜드요약',
    PRODUCT_PERF: '상품별_성과분석', DELETE: '삭제후보_자동추출', ACCOUNT: '신규계정_브랜드배치',
    SENT_CHECK: '쿠팡전송_확인', SENT_MANUAL: '전송일_수동입력',
    COUPANG_SENT_MANUAL: '쿠팡전송수_수동입력', MANUAL_API_CHECK: '수동입력_API검증',
    RETRANSMIT_LOG: '쿠팡재전송_로그', FILTER_RAW: 'API_필터원본확인',
    BRAND_ALIAS: '브랜드명_매칭표', MATCH_DIAG: '매칭진단',
    SETTINGS: '설정_사용법', LOG: '업데이트로그', SYNC_STATUS: '동기화상태',
    SELF_DIAGNOSTIC: '자가진단',
    AUDIT_REPORT: '검수리포트'
  },

  HEADERS: {
    PRODUCTS: ['업데이트일시','최초확인일','상품ID','사이트ID','상품명','모델명','이미지URL','롯데온상품번호_추출','필터ID','검색필터명','브랜드명','계정번호','쿠팡계정ID','API_등록일','API_수집일','API_수정일','API_마켓전송일','마켓전송일_수동','기준등록일','판매기간일수','원본_scope'],
    TMP_PRODUCTS: ['업데이트일시','상품ID','사이트ID','상품명','모델명','이미지URL','필터ID','검색필터명','브랜드명','계정번호','쿠팡계정ID','API_등록일','API_수집일','API_수정일','API_마켓전송일','원본_scope'],
    FILTERS: ['검색필터명','브랜드명','계정번호','쿠팡계정ID','API_totalCount','API_totalPage','이번조회_행수','필터코드','API_최근수집일자','API_필터생성일','API_최근수집일자_필드','API_필터생성일_필드','메모'],
    MARKETS: ['marketId','marketName','비고'],
    SALES_CLEAN: ['주문월','주문일시','마켓명','마켓주문번호','원본계정ID','분석계정ID','계정번호','마켓상품번호','사이트상품번호','브랜드명_원본','브랜드명_매칭','마켓상품명','원문상품명','상품URL','상품이미지','구매사이트명','결제수량','결제금액합계','결제배송비','정산예정금액','마켓주문상태','더망고주문상태','유효여부','분석대상여부','매칭방식','매칭상품ID','원본행','취소여부','취소판정기준'],
    MONTHLY_BRAND: ['주문월','브랜드명','계정번호','쿠팡계정ID','해당월말_더망고등록상품수','쿠팡전송확인상품수','매출상품수','쿠팡전송수대비_매출상품률','주문건수','매출액','쿠팡전송상품당매출','쿠팡전송상품당주문건수','월판정','취소수량','취소매출액'],
    BRAND: ['브랜드명','대표검색필터명','계정번호','쿠팡계정ID','더망고기준등록일','대표쿠팡전송일','최초쿠팡전송일','더망고등록후경과일수','쿠팡전송후판매기간일수','LOTTEON_더망고등록상품수','쿠팡전송확인상품수','전송일확인상태','전체_매출상품수','쿠팡전송수대비_매출상품률','전체_주문건수','전체_매출액','최근월','최근월_매출상품수','최근월_주문건수','최근월_매출액','쿠팡전송상품당매출','일평균매출_쿠팡전송기준','30일환산매출_쿠팡전송기준','판정','판정사유','브랜드매출_상품매칭','브랜드매출_상품미매칭','상품매칭_주문건수','상품미매칭_주문건수','상품매칭률','매출매칭률','정리재수집여부','최근정리재수집일','정리재수집후경과일수','정리재수집메모','재전송여부','최근재전송일','재전송후경과일수','재전송전_일평균매출','재전송후_일평균매출','재전송효과','재전송메모','추가수집여부','최근추가수집일','추가수집후경과일수','추가수집상품수','추가수집전_일평균매출','추가수집후_일평균매출','추가수집효과','추가수집메모','취소수량','취소매출액'],
    BRAND_SUMMARY: ['우선\n순위','운영\n분류','브랜드명','대표\n검색필터명','쿠팡\n계정ID','전송수','매출\n상품수','매출\n상품률','주문\n건수','전체\n매출','판매\n일수','30일\n환산매출','판매\n기준','액션','요약\n메모','다음\n상품갈이','1단계\n추가수집','2단계\n재전송','3단계\n정리재수집','상품갈이\n메모','정리재수집\n여부','최근\n정리재수집일','정리재수집\n메모','재전송\n여부','최근\n재전송일','재전송\n효과','재전송\n메모','추가수집\n여부','최근\n추가수집일','추가수집\n효과','추가수집\n메모','취소\n수량','취소\n매출'],
    PRODUCT_PERF: ['상품ID','롯데온상품번호_추출','상품명','브랜드명','검색필터명','계정번호','쿠팡계정ID','기준등록일','판매기간일수','매출상품여부','주문건수','매출액','최근주문월','판정메모'],
    DELETE: ['삭제우선순위','브랜드명','계정번호','쿠팡계정ID','더망고등록후경과일수','쿠팡전송후판매기간일수','더망고등록상품수','쿠팡전송확인상품수','매출상품수','쿠팡전송수대비_매출상품률','주문건수','매출액','30일환산매출_쿠팡전송기준','판정','판정사유','대표검색필터명','전송일확인상태'],
    ACCOUNT: ['현재필터계정번호','현재쿠팡계정ID','브랜드명','대표검색필터명','더망고등록상품수','쿠팡전송확인상품수','매출상품수','주문건수','매출액','쿠팡전송후판매기간일수','30일환산매출_쿠팡전송기준','추천액션','운영메모'],
    SENT_CHECK: ['브랜드명','검색필터명','계정번호','쿠팡계정ID','더망고등록상품수','쿠팡전송확인상품수','수동입력_쿠팡전체전송수','API_마켓전송일_확인상품수','상품별수동전송일_확인상품수','주문발생상품수','대표쿠팡전송일','전송일확인상태','전송확인상태','메모'],
    SENT_MANUAL: ['상품ID','브랜드명','쿠팡계정ID','마켓전송일','비고','설명: 더망고/API에서 상품별 전송일을 못 가져올 때 여기에 붙여넣기'],
    COUPANG_SENT_MANUAL: ['확인일','검색필터명','브랜드명','쿠팡전체전송수','더망고 수집수','beliun1023','beliun1024','대표쿠팡전송일','확인메모'],
    MANUAL_API_CHECK: ['상태입력','검색필터명','브랜드명','수동_쿠팡전송수','API_마켓전송확인수','전송수차이','수동_더망고수집수','API_더망고수집수','수집수차이','판정','메모'],
    RETRANSMIT_LOG: ['검색필터명','브랜드명','쿠팡\n계정ID','최초\n전송일','누적\n매출액','월평균\n매출액','작업\n유형','1단계_\n추가수집일','2단계_\n재전송일','3단계_\n정리재수집일','최종_\n더망고수집수','최종_\n쿠팡전송수','추가수집\n상품수','작업\n사유','비고'],
    FILTER_RAW: ['조회일시','페이지','검색필터명','필터ID','브랜드명','계정번호','쿠팡계정ID','최근수집일자_후보','필터생성일_후보','최근수집일자_필드','필터생성일_필드','전체원본JSON'],
    BRAND_ALIAS: ['매출브랜드명','더망고브랜드명','메모'],
    MATCH_DIAG: ['항목','값','비중','설명'],
    DASHBOARD: ['구분','항목','값1','값2','값3','값4','값5','값6','값7','값8','메모'],
    SETTINGS: ['항목','값','설명'],
    SYNC_STATUS: ['항목','값','메모'],
    LOG: ['일시','구분','내용'],
    SELF_DIAGNOSTIC: ['점검항목','결과','상세내용','점검시각'],
    AUDIT_REPORT: ['구분','점검항목','판정','브랜드명','검색필터명','기준값1','기준값2','차이','조치','메모']
  },

  PHASE: { IDLE: 'IDLE', COLLECT_FILTERS: 'COLLECT_FILTERS', COLLECT_PRODUCTS: 'COLLECT_PRODUCTS', FINALIZE: 'FINALIZE', DONE: 'DONE', ERROR: 'ERROR' }
};


function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const mainMenu = ui.createMenu('LOTTEON 자동화');

  // v5.98: 평소 운영에서 자주 쓰는 메뉴만 상단에 둡니다.
  mainMenu
    .addItem('① 변경사항 반영 실행', 'runPendingChangesApproval')
    .addItem('② 필터별_상품수 갱신/이어실행', 'refreshFilterCountsFast')
    .addItem('③ 쿠팡재전송_로그 갱신', 'createRetransmitLogSheet')
    .addItem('④ 핵심요약+대시보드 갱신', 'refreshCoreSummaryAndDashboardWithRetransmitLogDates')
    .addItem('⑤ 대시보드만 빠른 갱신', 'refreshDashboardFastOnly')
    .addItem('⑥ 시트 정리: 운영 시트만 표시', 'showOnlyMainSheets')
    .addSeparator();

  const checkMenu = ui.createMenu('점검/검수')
    .addItem('전체 검수 리포트 생성', 'generateAuditReport')
    .addItem('자동 갱신 상태 확인', 'showStableAutoRefreshStatus')
    .addItem('변경감지 상태 확인', 'showChangeDetectionStatus')
    .addItem('수동입력/API 검증 시트 갱신', 'refreshManualApiCheckOnly')
    .addItem('필터_대시보드 갱신', 'refreshFilterDashboardFastOnly');

  const setupMenu = ui.createMenu('설정/초기화')
    .addItem('변경감지 기능 시작', 'startChangeDetectionApproval')
    .addItem('변경감지 기능 중지', 'stopChangeDetectionApproval')
    .addItem('변경감지 플래그 초기화', 'resetChangeDetectionFlags')
    .addItem('필터별_상품수 이어실행 초기화', 'resetFilterListResumeProgress')
    .addItem('자동/예약 트리거 전체 정리', 'cleanupAllAutoRefreshTriggers')
    .addSeparator()
    .addItem('API 인증값 저장', 'saveApiCredentials')
    .addItem('API 연결 테스트', 'testLotteonApiConnection')
    .addSeparator()
    .addItem('시트 복구: 전체 시트 표시', 'showAllSheets');

  const troubleMenu = ui.createMenu('문제해결')
    .addItem('K2 필터일자 진단(조회만)', 'diagnoseRetransmitLogFilterDateK2')
    .addItem('K2 날짜 직접반영(실제수정)', 'patchK2RetransmitLogDateFromApi')
    .addItem('쿠팡재전송_로그 날짜 직접패치', 'patchRetransmitLogDatesFromFilterSummary')
    .addItem('쿠팡재전송_로그 미래날짜 정리', 'cleanFutureDatesInCoupangWorkLog')
    .addItem('초초경량 자동 갱신 1회 실행', 'runStableAutoRefreshOnce');

  mainMenu
    .addSubMenu(checkMenu)
    .addSubMenu(setupMenu)
    .addSubMenu(troubleMenu)
    .addToUi();
}

/**
 * v5.98 필터갱신 1페이지 안전실행
 * - 편집 시 즉시 무거운 작업을 실행하지 않습니다.
 * - 변경 플래그만 저장하고, 사용자가 메뉴에서 "변경사항 반영 실행"을 누를 때 처리합니다.
 */
function startChangeDetectionApproval() {
  const ui = SpreadsheetApp.getUi();
  try {
    const deleted = deleteChangeDetectionTriggers_();
    ScriptApp.newTrigger('handleWatchedSheetEdit')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    const props = PropertiesService.getScriptProperties();
    props.setProperty('CHANGE_DETECTION_ENABLED', 'Y');
    props.setProperty('CHANGE_DETECTION_STARTED_AT', now_());
    props.setProperty('CHANGE_DETECTION_LAST_MESSAGE', '변경감지 기능 시작');
    props.setProperty('CHANGE_DETECTION_LAST_ERROR', '');

    writeChangeDetectionStatus_('ENABLED', '변경감지 기능 시작', '기존 감지 트리거 정리 수: ' + deleted);

    ui.alert(
      '변경감지 기능을 시작했습니다.\n\n' +
      '감지 대상:\n' +
      '1. 매출데이터_붙여넣기\n' +
      '2. 쿠팡전송수_수동입력\n\n' +
      '시트가 변경되면 자동으로 무거운 작업을 실행하지 않고,\n' +
      '자동갱신_상태에 실행 필요 상태만 표시합니다.\n\n' +
      '실제 반영은 LOTTEON 자동화 → ① 변경사항 반영 실행을 눌러주세요.'
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    PropertiesService.getScriptProperties().setProperty('CHANGE_DETECTION_LAST_ERROR', msg);
    writeChangeDetectionStatus_('ERROR', '변경감지 기능 시작 실패', msg);
    ui.alert('변경감지 기능 시작 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function stopChangeDetectionApproval() {
  const ui = SpreadsheetApp.getUi();
  const deleted = deleteChangeDetectionTriggers_();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CHANGE_DETECTION_ENABLED', 'N');
  props.setProperty('CHANGE_DETECTION_LAST_MESSAGE', '변경감지 기능 중지');
  props.setProperty('CHANGE_DETECTION_LAST_ERROR', '');
  writeChangeDetectionStatus_('DISABLED', '변경감지 기능 중지', '삭제한 감지 트리거 수: ' + deleted);
  ui.alert('변경감지 기능을 중지했습니다.\n\n삭제한 감지 트리거 수: ' + deleted);
}

function handleWatchedSheetEdit(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('CHANGE_DETECTION_ENABLED') !== 'Y') return;
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet ? sheet.getName() : '';
    const a1 = e.range.getA1Notation ? e.range.getA1Notation() : '';

    const salesSheetName = CONFIG.SHEETS.SALES_IN;
    const manualSheetName = CONFIG.SHEETS.COUPANG_SENT_MANUAL;

    if (sheetName !== salesSheetName && sheetName !== manualSheetName) return;

    const editedAt = now_();
    props.setProperty('CHANGE_LAST_EDIT_SHEET', sheetName);
    props.setProperty('CHANGE_LAST_EDIT_RANGE', a1);
    props.setProperty('CHANGE_LAST_EDIT_AT', editedAt);
    props.setProperty('CHANGE_LAST_MESSAGE', sheetName + ' 변경 감지');

    if (sheetName === salesSheetName) {
      props.setProperty('CHANGE_SALES_DIRTY', 'Y');
      props.setProperty('CHANGE_SALES_DIRTY_AT', editedAt);
      props.setProperty('CHANGE_SALES_DIRTY_RANGE', a1);
    }

    if (sheetName === manualSheetName) {
      props.setProperty('CHANGE_MANUAL_DIRTY', 'Y');
      props.setProperty('CHANGE_MANUAL_DIRTY_AT', editedAt);
      props.setProperty('CHANGE_MANUAL_DIRTY_RANGE', a1);
    }

    writeChangeDetectionStatus_('DIRTY', sheetName + ' 변경 감지', 'LOTTEON 자동화 → ① 변경사항 반영 실행 필요');

    try {
      SpreadsheetApp.getActive().toast(
        sheetName + ' 변경 감지됨. LOTTEON 자동화 → ① 변경사항 반영 실행을 눌러주세요.',
        'LOTTEON 자동화',
        7
      );
    } catch (toastErr) {}
  } catch (err) {
    try {
      const msg = String(err && err.message ? err.message : err);
      PropertiesService.getScriptProperties().setProperty('CHANGE_DETECTION_LAST_ERROR', msg);
      writeChangeDetectionStatus_('ERROR', '변경감지 처리 오류', msg);
    } catch (ignore) {}
  }
}

function runPendingChangesApproval() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const salesDirty = props.getProperty('CHANGE_SALES_DIRTY') === 'Y';
  const manualDirty = props.getProperty('CHANGE_MANUAL_DIRTY') === 'Y';

  if (!salesDirty && !manualDirty) {
    writeChangeDetectionStatus_('NO_PENDING', '반영할 변경사항 없음', '');
    ui.alert(
      '반영할 변경사항이 없습니다.\n\n' +
      '감지 대상:\n' +
      '- 매출데이터_붙여넣기\n' +
      '- 쿠팡전송수_수동입력'
    );
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    ui.alert('다른 갱신 작업이 실행 중입니다.\n\n잠시 후 다시 실행해주세요.');
    return;
  }

  const started = Date.now();

  try {
    writeChangeDetectionStatus_('APPLY_RUNNING', '변경사항 반영 실행 중', 'v6.00 초안전 모드: 빠른 매출정리 + 빠른 핵심요약');

    // v6.00:
    // v5.99에서도 CHANGE_APPLY_ANALYSIS 단계에서 360초 시간초과가 발생했습니다.
    // 따라서 ①은 기존 refreshLotteonAnalysisNoAlert_() 전체 분석을 호출하지 않고,
    // 매출데이터_붙여넣기 원본 브랜드 기준의 빠른 집계 전용 경로만 실행합니다.
    // 대시보드/검수/필터/API/상품목록/복잡한 전후효과 계산은 모두 별도 메뉴로 분리합니다.
    props.setProperty('AUTO_REFRESH_RUNNING', 'Y');
    props.setProperty('AUTO_REFRESH_STAGE', 'CHANGE_APPLY_FAST_SAFE');
    props.setProperty('AUTO_REFRESH_STARTED_AT', now_());
    props.setProperty('AUTO_REFRESH_LAST_ERROR', '');
    props.setProperty('AUTO_REFRESH_LAST_RESULT', '');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v6.00 변경사항 반영 - 빠른 매출정리/핵심요약 갱신 중');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    writeAutoRefreshStatus_('CHANGE_APPLY_FAST_SAFE', 'v6.00 변경사항 반영 - 빠른 매출정리/핵심요약 갱신 중', '');

    const analysisResult = refreshLotteonAnalysisEmergencyNoAlert_();

    if (manualDirty) {
      props.setProperty('CHANGE_FILTERLIST_REFRESH_RECOMMENDED', 'Y');
      props.setProperty('CHANGE_FILTERLIST_REFRESH_RECOMMENDED_AT', now_());
    }

    const elapsedSec = Math.round((Date.now() - started) / 1000);
    const resultMemo =
      '소요초=' + elapsedSec +
      ' / 매출변경=' + (salesDirty ? 'Y' : 'N') +
      ' / 전송수변경=' + (manualDirty ? 'Y' : 'N') +
      ' / 원본처리행=' + (analysisResult && analysisResult.sourceRows ? analysisResult.sourceRows : '') +
      ' / LOTTEON행=' + (analysisResult && analysisResult.salesRows ? analysisResult.salesRows : '') +
      ' / 브랜드=' + (analysisResult && analysisResult.brandRows ? analysisResult.brandRows : '') +
      ' / 처리모드=빠른 원본브랜드 집계' +
      ' / 대시보드=자동실행 제외' +
      ' / filterList=자동실행 제외' +
      ' / 수동API검증=자동실행 제외' +
      ' / 필터대시보드=자동실행 제외' +
      ' / 검수리포트=자동실행 제외';

    props.setProperty('CHANGE_SALES_DIRTY', 'N');
    props.setProperty('CHANGE_MANUAL_DIRTY', 'N');
    props.setProperty('CHANGE_LAST_APPLIED_AT', now_());
    props.setProperty('CHANGE_LAST_RESULT', resultMemo);
    props.setProperty('CHANGE_LAST_MESSAGE', 'v6.00 변경사항 반영 완료');
    props.setProperty('CHANGE_DETECTION_LAST_ERROR', '');

    props.setProperty('AUTO_REFRESH_RUNNING', 'N');
    props.setProperty('AUTO_REFRESH_STAGE', 'DONE');
    props.setProperty('AUTO_REFRESH_LAST_DONE_AT', now_());
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v6.00 변경사항 반영 완료');
    props.setProperty('AUTO_REFRESH_LAST_RESULT', resultMemo);
    props.setProperty('AUTO_REFRESH_LAST_ERROR', '');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());

    writeAutoRefreshStatus_('DONE', 'v6.00 변경사항 반영 완료', '');
    writeChangeDetectionStatus_('APPLIED', 'v6.00 변경사항 반영 완료', resultMemo);
    log_('change_apply_done_v600', resultMemo);

    ui.alert(
      '변경사항 반영 완료\n\n' +
      resultMemo + '\n\n' +
      '갱신된 시트:\n' +
      '- 매출데이터_정리\n' +
      '- 브랜드별_성과분석\n' +
      '- 핵심_브랜드요약\n\n' +
      'v6.00 안전화로 ①에서 자동 제외된 작업:\n' +
      '- 대시보드\n' +
      '- 수동입력/API검증\n' +
      '- 필터_대시보드\n' +
      '- 필터별_상품수 API 갱신\n' +
      '- 검수리포트\n\n' +
      '대시보드는 필요 시 ⑤ 대시보드만 빠른 갱신을 따로 실행하세요.'
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    props.setProperty('AUTO_REFRESH_RUNNING', 'N');
    props.setProperty('AUTO_REFRESH_STAGE', 'ERROR');
    props.setProperty('AUTO_REFRESH_LAST_ERROR', msg);
    props.setProperty('CHANGE_DETECTION_LAST_ERROR', msg);
    writeAutoRefreshStatus_('ERROR', '변경사항 반영 오류', msg);
    writeChangeDetectionStatus_('ERROR', '변경사항 반영 오류', msg);
    ui.alert('변경사항 반영 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  } finally {
    lock.releaseLock();
  }
}


function showChangeDetectionStatus() {
  const props = PropertiesService.getScriptProperties();
  writeChangeDetectionStatus_(
    buildChangeDetectionStatusLabel_(),
    props.getProperty('CHANGE_LAST_MESSAGE') || props.getProperty('CHANGE_DETECTION_LAST_MESSAGE') || '',
    props.getProperty('CHANGE_DETECTION_LAST_ERROR') || ''
  );

  SpreadsheetApp.getUi().alert(
    '변경감지 상태\n\n' +
    '사용여부: ' + (props.getProperty('CHANGE_DETECTION_ENABLED') || 'N') + '\n' +
    '매출데이터 변경: ' + (props.getProperty('CHANGE_SALES_DIRTY') || 'N') + '\n' +
    '쿠팡전송수 변경: ' + (props.getProperty('CHANGE_MANUAL_DIRTY') || 'N') + '\n' +
    '최근 변경 시트: ' + (props.getProperty('CHANGE_LAST_EDIT_SHEET') || '') + '\n' +
    '최근 변경 위치: ' + (props.getProperty('CHANGE_LAST_EDIT_RANGE') || '') + '\n' +
    '최근 변경 시각: ' + (props.getProperty('CHANGE_LAST_EDIT_AT') || '') + '\n' +
    '최근 반영 시각: ' + (props.getProperty('CHANGE_LAST_APPLIED_AT') || '') + '\n' +
    '최근 결과: ' + (props.getProperty('CHANGE_LAST_RESULT') || '') + '\n' +
    '최근 오류: ' + (props.getProperty('CHANGE_DETECTION_LAST_ERROR') || '')
  );
}

function resetChangeDetectionFlags() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CHANGE_SALES_DIRTY', 'N');
  props.setProperty('CHANGE_MANUAL_DIRTY', 'N');
  props.setProperty('CHANGE_LAST_RESULT', '');
  props.setProperty('CHANGE_FILTERLIST_REFRESH_RECOMMENDED', 'N');
  props.setProperty('CHANGE_FILTERLIST_REFRESH_RECOMMENDED_AT', '');
  props.setProperty('CHANGE_DETECTION_LAST_ERROR', '');
  props.setProperty('CHANGE_LAST_MESSAGE', '변경감지 플래그 초기화');
  writeChangeDetectionStatus_('RESET', '변경감지 플래그 초기화', '');
  SpreadsheetApp.getUi().alert('변경감지 플래그를 초기화했습니다.');
}

function deleteChangeDetectionTriggers_() {
  return deleteTriggersByHandlerNames_(['handleWatchedSheetEdit']);
}

function buildChangeDetectionStatusLabel_() {
  const props = PropertiesService.getScriptProperties();
  const enabled = props.getProperty('CHANGE_DETECTION_ENABLED') === 'Y';
  const salesDirty = props.getProperty('CHANGE_SALES_DIRTY') === 'Y';
  const manualDirty = props.getProperty('CHANGE_MANUAL_DIRTY') === 'Y';

  if (!enabled) return 'DISABLED';
  if (salesDirty || manualDirty) return 'DIRTY';
  return 'READY';
}

function writeChangeDetectionStatus_(stage, message, memo) {
  const props = PropertiesService.getScriptProperties();

  // 자동갱신_상태 시트 하단에 변경감지 상태도 함께 기록합니다.
  // 기존 자동갱신 상태와 충돌하지 않도록 별도 구역을 사용합니다.
  const sheet = ensureAutoRefreshStatusSheet_();
  const startRow = 18;
  const rows = [
    ['변경감지_상태', stage || buildChangeDetectionStatusLabel_(), 'READY/DIRTY/APPLY_RUNNING/APPLIED/ERROR'],
    ['변경감지_사용여부', props.getProperty('CHANGE_DETECTION_ENABLED') || 'N', '설치형 onEdit 감지 트리거 사용 여부'],
    ['매출데이터_변경', props.getProperty('CHANGE_SALES_DIRTY') || 'N', CONFIG.SHEETS.SALES_IN + ' 변경 여부'],
    ['매출데이터_변경시각', props.getProperty('CHANGE_SALES_DIRTY_AT') || '', '마지막 감지 시각'],
    ['매출데이터_변경위치', props.getProperty('CHANGE_SALES_DIRTY_RANGE') || '', '마지막 감지 셀/범위'],
    ['쿠팡전송수_변경', props.getProperty('CHANGE_MANUAL_DIRTY') || 'N', CONFIG.SHEETS.COUPANG_SENT_MANUAL + ' 변경 여부'],
    ['쿠팡전송수_변경시각', props.getProperty('CHANGE_MANUAL_DIRTY_AT') || '', '마지막 감지 시각'],
    ['쿠팡전송수_변경위치', props.getProperty('CHANGE_MANUAL_DIRTY_RANGE') || '', '마지막 감지 셀/범위'],
    ['최근변경시트', props.getProperty('CHANGE_LAST_EDIT_SHEET') || '', '감지된 마지막 시트'],
    ['최근변경위치', props.getProperty('CHANGE_LAST_EDIT_RANGE') || '', '감지된 마지막 위치'],
    ['최근반영시각', props.getProperty('CHANGE_LAST_APPLIED_AT') || '', '★ 변경사항 반영 실행 완료 시각'],
    ['최근결과', props.getProperty('CHANGE_LAST_RESULT') || '', '최근 반영 결과'],
    ['최근메시지', message || props.getProperty('CHANGE_LAST_MESSAGE') || props.getProperty('CHANGE_DETECTION_LAST_MESSAGE') || '', '변경감지 진행 설명'],
    ['최근오류', props.getProperty('CHANGE_DETECTION_LAST_ERROR') || '', '오류 발생 시 원문'],
    ['filterList갱신권장', props.getProperty('CHANGE_FILTERLIST_REFRESH_RECOMMENDED') || 'N', 'Y이면 필요 시 필터별_상품수 갱신/이어실행(filterList) 별도 실행'],
    ['filterList갱신권장시각', props.getProperty('CHANGE_FILTERLIST_REFRESH_RECOMMENDED_AT') || '', '쿠팡전송수_수동입력 변경 후 권장 표시'],
    ['실행안내', '변경 감지 후 LOTTEON 자동화 → ① 변경사항 반영 실행', memo || '자동으로 무거운 작업을 실행하지 않음']
  ];

  try {
    sheet.getRange(startRow, 1, rows.length + 1, 3).clearContent();
    sheet.getRange(startRow, 1, 1, 3).setValues([['변경감지 항목', '값', '메모']]);
    sheet.getRange(startRow + 1, 1, rows.length, 3).setValues(rows);
    sheet.getRange(startRow, 1, 1, 3).setFontWeight('bold').setBackground('#fff2cc').setHorizontalAlignment('center');
    sheet.getRange(startRow, 1, rows.length + 1, 3).setVerticalAlignment('middle');
  } catch (e) {}
}

function refreshManualApiCheckOnly() {
  refreshManualInputApiValidation();
}

function refreshManualApiCheckOnlyNoAlert_() {
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.MANUAL_API_CHECK, CONFIG.HEADERS.MANUAL_API_CHECK);

  const filterSummaryMap = buildFilterSummaryMapFromFilterSheet_();
  const result = writeManualInputApiValidationFast_(filterSummaryMap, { issueOnly: true, ultraLight: true });
  log_('manual_api_check_noalert_v583', 'written=' + result.writtenRows + ', total=' + result.totalRows);
  return result;
}

function refreshFilterDashboardFastOnlyNoAlert_() {
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);

  const filterSummaryMap = buildFilterSummaryMapFromFilterSheet_();
  buildDashboardFastFromFilterSummary_(filterSummaryMap);
  log_('filter_dashboard_noalert_v583', '필터_대시보드 갱신 완료');
}


function showOnlyMainSheets() {
  const ss = SpreadsheetApp.getActive();

  // v5.98 운영 표시 시트 7개만 남깁니다.
  const visible = {
    '대시보드': true,
    '핵심_브랜드요약': true,
    '쿠팡재전송_로그': true,
    '쿠팡전송수_수동입력': true,
    '매출데이터_붙여넣기': true,
    '자동갱신_상태': true,
    '필터_대시보드': true
  };

  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    try {
      if (visible[name]) sheet.showSheet();
      else sheet.hideSheet();
    } catch (e) {}
  });

  SpreadsheetApp.getUi().alert(
    '운영 시트만 표시했습니다.\n\n' +
    '표시 시트:\n' +
    '- 대시보드\n' +
    '- 핵심_브랜드요약\n' +
    '- 쿠팡재전송_로그\n' +
    '- 쿠팡전송수_수동입력\n' +
    '- 매출데이터_붙여넣기\n' +
    '- 자동갱신_상태\n' +
    '- 필터_대시보드\n\n' +
    '전체 시트를 다시 보려면 LOTTEON 자동화 → 설정/초기화 → 시트 복구: 전체 시트 표시를 실행하세요.'
  );
}

function showAllSheets() {
  const ss = SpreadsheetApp.getActive();
  let count = 0;

  ss.getSheets().forEach(function(sheet) {
    try {
      sheet.showSheet();
      count++;
    } catch (e) {}
  });

  SpreadsheetApp.getUi().alert('전체 시트를 표시했습니다.\n\n표시 처리 시트 수: ' + count);
}


function authorizeLotteonAutomation() {
  const ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.getActive();
    UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
    ScriptApp.getProjectTriggers();
    PropertiesService.getScriptProperties().setProperty('AUTH_TEST_TIMESTAMP', now_());
    log_('auth', '권한 승인 확인 완료');
    ui.alert('✅ 권한 승인이 완료되었습니다.\n\n이제 "원클릭 준비" 메뉴를 실행해주세요.');
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    log_('auth_error', msg);
    ui.alert('❌ 권한 승인 중 오류가 발생했습니다.\n\n' + msg + '\n\nGoogle이 권한을 요청하면 "허용"을 눌러주세요.');
    throw e;
  }
}

function safeDeleteTrigger_(trigger) {
  let triggerId = 'unknown';
  let handler = 'unknown';
  try {
    if (!trigger) return false;
    try { if (typeof trigger.getUniqueId === 'function') triggerId = trigger.getUniqueId(); } catch (e) {}
    try { if (typeof trigger.getHandlerFunction === 'function') handler = trigger.getHandlerFunction(); } catch (e) {}
    ScriptApp.deleteTrigger(trigger);
    return true;
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    try { log_('trigger_delete_failed', 'handler=' + handler + ', triggerID=' + triggerId + ', error=' + msg); } catch (le) {}
    return false;
  }
}

/**
 * v5.98 참고: 1단계 자동화 호환 함수
 * - 매일 자동 갱신 예약/중지/즉시 실행/상태 확인
 * - 매출데이터 자동 수집은 아직 포함하지 않음
 * - 현재 매출데이터_붙여넣기에 들어 있는 데이터를 기준으로 분석합니다.
 */
function startDailyAutoRefreshSchedule() {
  const ui = SpreadsheetApp.getUi();
  try {
    deleteDailyAutoRefreshScheduleTriggers_();
    ScriptApp.newTrigger('runDailyAutoRefreshStart')
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .nearMinute(30)
      .create();

    const props = PropertiesService.getScriptProperties();
    props.setProperty('AUTO_REFRESH_ENABLED', 'Y');
    props.setProperty('AUTO_REFRESH_SCHEDULE', '매일 06:30 전후');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '매일 자동 갱신 예약 시작');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());

    writeAutoRefreshStatus_('예약됨', '매일 06:30 전후 자동 갱신 예약 완료', '');
    log_('auto_refresh_schedule_start', '매일 06:30 전후 자동 갱신 예약 시작');

    ui.alert(
      '매일 자동 갱신 예약을 시작했습니다.\n\n' +
      '예약 시간: 매일 06:30 전후\n\n' +
      '자동 실행 순서:\n' +
      '1. 매출데이터 정리 + 핵심분석 갱신\n' +
      '2. 검수리포트 생성\n' +
      '3. 대시보드 갱신\n\n' +
      'v5.73부터 이어실행 트리거 없이 한 번에 처리합니다.\n' +
      '느린 필터/상품수 API 갱신은 매일 자동갱신에서 제외했습니다.\n' +
      '필터/상품수는 필요할 때 별도 메뉴로 실행하세요.\n\n' +
      '주의: 매출데이터_붙여넣기 자체를 더망고에서 자동 수집하는 기능은 아직 포함되지 않습니다.'
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    writeAutoRefreshStatus_('오류', '자동 갱신 예약 시작 실패', msg);
    ui.alert('자동 갱신 예약 시작 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function stopDailyAutoRefreshSchedule() {
  cleanupAllAutoRefreshTriggers();
}

function runDailyAutoRefreshOnceManual() {
  const ui = SpreadsheetApp.getUi();
  try {
    // 이전 버전의 FILTER_COUNTS_WAIT 상태나 느린 productList 분할 트리거가 남아 있으면
    // 즉시 실행 전에 잔여 상태를 정리합니다.
    deleteDailyAutoRefreshContinuationTriggers_();
    try { deleteFastFilterCountTriggers_(); } catch (e) {}

    const props = PropertiesService.getScriptProperties();
    props.setProperty('AUTO_REFRESH_RUNNING', 'N');
    props.setProperty('AUTO_REFRESH_STAGE', 'MANUAL_RESTART');
    props.setProperty('AUTO_REFRESH_LAST_ERROR', '');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '수동 즉시 실행을 위해 이전 자동갱신 상태 초기화');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());

    runDailyAutoRefreshStart();

    ui.alert(
      '자동 갱신 1회 실행이 완료되었습니다.\n\n' +
      '실행 범위:\n' +
      '1. 매출데이터 정리 + 핵심분석 갱신\n' +
      '2. 검수리포트 생성\n' +
      '3. 대시보드 갱신\n\n' +
      '필터/상품수 API 갱신은 시간이 오래 걸릴 수 있어 자동갱신에서 제외했습니다.\n' +
      '결과는 자동갱신_상태 / 검수리포트 / 대시보드에서 확인하세요.'
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    writeAutoRefreshStatus_('오류', '자동 갱신 1회 실행 실패', msg);
    ui.alert('자동 갱신 1회 실행 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function showDailyAutoRefreshStatus() {
  const props = PropertiesService.getScriptProperties();
  writeAutoRefreshStatus_(
    props.getProperty('AUTO_REFRESH_STAGE') || '미예약',
    props.getProperty('AUTO_REFRESH_LAST_MESSAGE') || '',
    props.getProperty('AUTO_REFRESH_LAST_ERROR') || ''
  );

  SpreadsheetApp.getUi().alert(
    '자동 갱신 상태\n\n' +
    '사용여부: ' + (props.getProperty('AUTO_REFRESH_ENABLED') || 'N') + '\n' +
    '실행중: ' + (props.getProperty('AUTO_REFRESH_RUNNING') || 'N') + '\n' +
    '현재단계: ' + (props.getProperty('AUTO_REFRESH_STAGE') || '') + '\n' +
    '예약시간: ' + (props.getProperty('AUTO_REFRESH_SCHEDULE') || '없음') + '\n' +
    '최근시작: ' + (props.getProperty('AUTO_REFRESH_STARTED_AT') || '') + '\n' +
    '최근완료: ' + (props.getProperty('AUTO_REFRESH_LAST_DONE_AT') || '') + '\n' +
    '최근메시지: ' + (props.getProperty('AUTO_REFRESH_LAST_MESSAGE') || '') + '\n' +
    '최근오류: ' + (props.getProperty('AUTO_REFRESH_LAST_ERROR') || '')
  );
}

function runDailyAutoRefreshStart() {
  // v5.78: 과거 예약 트리거가 남아 호출될 경우에도 안정 실행만 수행합니다.
  runStableAutoRefreshOnce();
}

function runDailyAutoRefreshAllStepsNoTrigger_() {
  const props = PropertiesService.getScriptProperties();
  const started = Date.now();

  try {
    props.setProperty('AUTO_REFRESH_STAGE', 'ANALYSIS_RUNNING');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '매출데이터 정리 + 핵심분석 갱신 중');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    writeAutoRefreshStatus_('ANALYSIS_RUNNING', '매출데이터 정리 + 핵심분석 갱신 중', '');

    const analysisResult = refreshLotteonAnalysisNoAlert_();

    props.setProperty('AUTO_REFRESH_STAGE', 'AUDIT_RUNNING');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '검수리포트 생성 중');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    writeAutoRefreshStatus_('AUDIT_RUNNING', '검수리포트 생성 중', '');

    const auditResult = generateAuditReportNoAlert_();

    props.setProperty('AUTO_REFRESH_STAGE', 'DASHBOARD_RUNNING');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '대시보드 갱신 중');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    writeAutoRefreshStatus_('DASHBOARD_RUNNING', '대시보드 갱신 중', '');

    if (typeof refreshDashboardFastOnlyNoAlert_ === 'function') {
      refreshDashboardFastOnlyNoAlert_();
    } else {
      ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
      ensureSheet_(SpreadsheetApp.getActive(), CONFIG.SHEETS.DASHBOARD);
      buildDashboard_(null, null, null);
    }

    const elapsedSec = Math.round((Date.now() - started) / 1000);
    props.setProperty('AUTO_REFRESH_STAGE', 'DONE');
    props.setProperty('AUTO_REFRESH_RUNNING', 'N');
    props.setProperty('AUTO_REFRESH_LAST_DONE_AT', now_());
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '자동 갱신 완료');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    props.setProperty('AUTO_REFRESH_LAST_ERROR', '');

    writeAutoRefreshStatus_(
      'DONE',
      '자동 갱신 완료',
      '소요초=' + elapsedSec +
      ' / 매출정리행=' + (analysisResult && analysisResult.salesRows ? analysisResult.salesRows : '') +
      ' / 브랜드=' + (analysisResult && analysisResult.brandRows ? analysisResult.brandRows : '') +
      ' / 검수긴급=' + (auditResult && auditResult.severe !== undefined ? auditResult.severe : '') +
      ' / 검수주의=' + (auditResult && auditResult.warn !== undefined ? auditResult.warn : '')
    );

    log_('auto_refresh_done_v573', '자동 갱신 완료 / elapsedSec=' + elapsedSec);
    return { done: true, elapsedSec: elapsedSec };
  } catch (e) {
    dailyAutoRefreshFail_(e, props.getProperty('AUTO_REFRESH_STAGE') || 'AUTO_REFRESH_ERROR');
    throw e;
  }
}

function runDailyAutoRefreshContinue() {
  // v5.74: 이어실행 트리거 미사용. 과거 잔여 트리거는 정리합니다.
  cleanupAllAutoRefreshTriggersNoAlert_();
  writeAutoRefreshStatus_('CONTINUE_IGNORED', 'v5.98에서는 이어실행 트리거를 사용하지 않음', '안정 자동 갱신 1회 실행을 사용하세요');
}


/**
 * v5.98 안정화 전용 실행부
 * 목적:
 * - 트리거 생성 없음
 * - 필터/상품수 API 갱신 없음
 * - LOTTEON_상품목록 전체 읽기 없음
 * - 매출분석 → 검수리포트 → 대시보드만 한 번에 실행
 */
function runStableAutoRefreshOnce() {
  const ui = SpreadsheetApp.getUi();
  const started = Date.now();

  try {
    cleanupAllAutoRefreshTriggersNoAlert_();

    const props = PropertiesService.getScriptProperties();
    props.setProperty('AUTO_REFRESH_ENABLED', 'N');
    props.setProperty('AUTO_REFRESH_RUNNING', 'Y');
    props.setProperty('AUTO_REFRESH_STAGE', 'ULTRA_START');
    props.setProperty('AUTO_REFRESH_STARTED_AT', now_());
    props.setProperty('AUTO_REFRESH_LAST_DONE_AT', '');
    props.setProperty('AUTO_REFRESH_LAST_ERROR', '');
    props.setProperty('AUTO_REFRESH_LAST_RESULT', '');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v5.98 초초경량 자동 갱신 시작');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());

    writeAutoRefreshStatus_('ULTRA_START', 'v5.98 초초경량 자동 갱신 시작', '');

    props.setProperty('AUTO_REFRESH_STAGE', 'ULTRA_ANALYSIS_RUNNING');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '매출데이터 정리 + 핵심분석 갱신 중');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    writeAutoRefreshStatus_('ULTRA_ANALYSIS_RUNNING', '매출데이터 정리 + 핵심분석 갱신 중', '');

    const analysisResult = refreshLotteonAnalysisNoAlert_();

    props.setProperty('AUTO_REFRESH_STAGE', 'ULTRA_DASHBOARD_RUNNING');
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '대시보드 갱신 중');
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    writeAutoRefreshStatus_('ULTRA_DASHBOARD_RUNNING', '대시보드 갱신 중', '');

    ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
    ensureSheet_(SpreadsheetApp.getActive(), CONFIG.SHEETS.DASHBOARD);
    buildDashboard_(null, null, null);

    const elapsedSec = Math.round((Date.now() - started) / 1000);
    const resultMemo =
      '소요초=' + elapsedSec +
      ' / 매출정리행=' + (analysisResult && analysisResult.salesRows ? analysisResult.salesRows : '') +
      ' / 브랜드=' + (analysisResult && analysisResult.brandRows ? analysisResult.brandRows : '') +
      ' / 제외=검수리포트·쿠팡재전송로그·매칭진단·필터API·서식';

    props.setProperty('AUTO_REFRESH_RUNNING', 'N');
    props.setProperty('AUTO_REFRESH_STAGE', 'DONE');
    props.setProperty('AUTO_REFRESH_LAST_DONE_AT', now_());
    props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v5.98 초초경량 자동 갱신 완료');
    props.setProperty('AUTO_REFRESH_LAST_RESULT', resultMemo);
    props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
    props.setProperty('AUTO_REFRESH_LAST_ERROR', '');

    writeAutoRefreshStatus_('DONE', 'v5.98 초초경량 자동 갱신 완료', '');
    log_('ultra_fast_auto_refresh_done_v582', resultMemo);

    ui.alert(
      '초초경량 자동 갱신 1회 실행 완료\n\n' +
      '실행 범위:\n' +
      '1. 매출데이터 정리 + 핵심분석 갱신\n' +
      '2. 대시보드 갱신\n\n' +
      '시간초과 방지를 위해 제외:\n' +
      '- 검수리포트 생성\n' +
      '- 쿠팡재전송_로그 동기화\n' +
      '- 매칭진단 재작성\n' +
      '- 필터/상품수 API 갱신\n' +
      '- 전체 서식/열너비 갱신\n\n' +
      resultMemo
    );

    return { done: true, elapsedSec: elapsedSec };
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    const props = PropertiesService.getScriptProperties();
    props.setProperty('AUTO_REFRESH_LAST_RESULT', '');
    dailyAutoRefreshFail_(e, 'ULTRA_ERROR');
    SpreadsheetApp.getUi().alert(
      '초초경량 자동 갱신 실행 중 오류가 발생했습니다.\n\n' +
      msg + '\n\n' +
      '자동갱신_상태 시트의 최근단계를 확인하세요.'
    );
    throw e;
  }
}

function showStableAutoRefreshStatus() {
  showDailyAutoRefreshStatus();
}

function cleanupAllAutoRefreshTriggers() {
  const ui = SpreadsheetApp.getUi();
  const deleted = cleanupAllAutoRefreshTriggersNoAlert_();

  const props = PropertiesService.getScriptProperties();
  props.setProperty('AUTO_REFRESH_ENABLED', 'N');
  props.setProperty('AUTO_REFRESH_RUNNING', 'N');
  props.setProperty('AUTO_REFRESH_STAGE', 'CLEANED');
  props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '자동/예약 트리거 전체 정리 완료');
  props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());

  writeAutoRefreshStatus_('CLEANED', '자동/예약 트리거 전체 정리 완료', '삭제한 트리거 수: ' + deleted);
  log_('stable_auto_refresh_cleanup_v574', 'deleted=' + deleted);
  ui.alert('자동/예약 트리거 전체 정리 완료\n\n삭제한 트리거 수: ' + deleted);
}

function cleanupAllAutoRefreshTriggersNoAlert_() {
  let deleted = 0;
  try { deleted += deleteDailyAutoRefreshScheduleTriggers_(); } catch (e) {}
  try { deleted += deleteDailyAutoRefreshContinuationTriggers_(); } catch (e) {}
  try { deleted += deleteTriggersByHandlerNames_(['runDailyAutoRefreshStart', 'runDailyAutoRefreshContinue', 'continueFastFilterCountSync']); } catch (e) {}
  try { deleteFastFilterCountTriggers_(); } catch (e) {}
  return deleted;
}

/**
 * v5.98에서는 매일 예약을 안정화 전까지 사용하지 않습니다.
 * 메뉴에서는 제거했지만, 과거 트리거/외부 호출 방지를 위해 안전하게 비활성 처리합니다.
 */
function startDailyAutoRefreshSchedule() {
  cleanupAllAutoRefreshTriggersNoAlert_();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AUTO_REFRESH_ENABLED', 'N');
  props.setProperty('AUTO_REFRESH_RUNNING', 'N');
  props.setProperty('AUTO_REFRESH_STAGE', 'SCHEDULE_DISABLED');
  props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v5.98에서는 매일 예약을 비활성화했습니다. 안정 1회 실행을 먼저 사용하세요.');
  props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
  writeAutoRefreshStatus_('SCHEDULE_DISABLED', '매일 예약 비활성화', '안정 1회 실행 검증 후 예약 기능 재추가 권장');
  SpreadsheetApp.getUi().alert('v5.98에서는 매일 예약을 잠시 비활성화했습니다.\n\n먼저 "★ 안정 자동 갱신 1회 실행"으로 안정성을 확인하세요.');
}

function runDailyAutoRefreshOnceManual() {
  runStableAutoRefreshOnce();
}

function startFastFilterCountsForAutomation_() {
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();

  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SYNC_STATUS, CONFIG.HEADERS.SYNC_STATUS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.LOG, CONFIG.HEADERS.LOG);
  ensureCoupangWorkLogSheet_();
  ensureFastFilterCountTmpSheet_();

  deleteFastFilterCountTriggers_();

  writeSyncStatus_({
    phase: 'AUTO_FAST_FILTER_COUNT_INIT',
    status: '자동 갱신 - 필터/상품수 분할 갱신 시작',
    currentFilter: 'filterList',
    lastUrl: buildApiUrl_('filterList'),
    tmpRows: ''
  });

  const filterItems = fetchLotteonFilterItemsFast_();
  const filterNames = [];
  const rows = [];

  filterItems.forEach(function(item) {
    const name = findFilterName_(item);
    if (!name) return;

    const filterName = String(name).trim();
    if (!isValidLotteonFilterName_(filterName)) return;
    if (filterNames.indexOf(filterName) >= 0) return;

    filterNames.push(filterName);

    const account = accountFromFilterName_(filterName);
    const brand = brandFromFilterName_(filterName);
    const meta = extractFilterListMeta_(item) || {};

    rows.push([
      filterName,
      brand,
      account.accountNo,
      account.accountId,
      '',
      '',
      filterCodeFromFilterName_(filterName),
      meta.recentDate || '',
      meta.createDate || '',
      meta.recentField || '',
      meta.createField || '',
      'PENDING',
      '',
      now_()
    ]);
  });

  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });

  if (!rows.length) throw new Error('롯백으로 시작하는 유효 검색필터를 찾지 못했습니다.');

  replaceDataFastLimited_(getFastFilterCountTmpSheet_(), getFastFilterCountTmpHeaders_(), rows);
  const result = processFastFilterCountBatch_(started, false);

  log_('auto_fast_filter_count_start', '자동 필터/상품수 갱신 시작 / total=' + rows.length + ', done=' + (result.doneCount || 0));
  return result;
}


/**
 * v6.00 변경사항 반영 전용 초안전 분석 경로
 * - 기존 refreshLotteonAnalysisNoAlert_()가 360초를 초과하는 경우를 피하기 위한 운영용 빠른 경로입니다.
 * - 매출데이터_붙여넣기의 브랜드 컬럼을 우선 신뢰하여 매출데이터_정리/브랜드별_성과분석/핵심_브랜드요약만 갱신합니다.
 * - LOTTEON_상품목록, 대시보드, 검수리포트, 필터/API 검증, 전후효과 상세 계산은 실행하지 않습니다.
 */
function refreshLotteonAnalysisEmergencyNoAlert_() {
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SALES_CLEAN, CONFIG.HEADERS.SALES_CLEAN);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND, CONFIG.HEADERS.BRAND);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND_ALIAS, CONFIG.HEADERS.BRAND_ALIAS);

  props.setProperty('AUTO_REFRESH_STAGE', 'CHANGE_APPLY_FAST_SALES_CLEAN');
  props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v6.00 빠른 매출데이터 정리 중');
  props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
  writeAutoRefreshStatus_('CHANGE_APPLY_FAST_SALES_CLEAN', 'v6.00 빠른 매출데이터 정리 중', '상품목록/복잡매칭 미사용');

  const salesResult = normalizeSalesDataEmergencyFast_();
  const sales = salesResult.salesObjects || [];

  props.setProperty('AUTO_REFRESH_STAGE', 'CHANGE_APPLY_FAST_BRAND_METRICS');
  props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v6.00 빠른 브랜드 성과 집계 중');
  props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
  writeAutoRefreshStatus_('CHANGE_APPLY_FAST_BRAND_METRICS', 'v6.00 빠른 브랜드 성과 집계 중', '원본 브랜드 기준');

  const brandMetrics = buildBrandMetricsEmergencyFast_(sales);
  const brandRows = buildBrandRows_(brandMetrics);
  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND), CONFIG.HEADERS.BRAND, brandRows);

  props.setProperty('AUTO_REFRESH_STAGE', 'CHANGE_APPLY_FAST_SUMMARY');
  props.setProperty('AUTO_REFRESH_LAST_MESSAGE', 'v6.00 빠른 핵심_브랜드요약 생성 중');
  props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
  writeAutoRefreshStatus_('CHANGE_APPLY_FAST_SUMMARY', 'v6.00 빠른 핵심_브랜드요약 생성 중', '대시보드 제외');

  const summaryRows = buildBrandSummaryRowsFromBrandRows_(brandRows);
  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND_SUMMARY), CONFIG.HEADERS.BRAND_SUMMARY, summaryRows);

  const elapsedSec = Math.round((Date.now() - started) / 1000);
  log_('analysis_emergency_fast_v600', 'sourceRows=' + salesResult.sourceRows + ', lotteonRows=' + sales.length + ', brandRows=' + brandRows.length + ', elapsedSec=' + elapsedSec);

  return {
    sourceRows: salesResult.sourceRows,
    salesRows: sales.length,
    brandRows: brandRows.length,
    elapsedSec: elapsedSec
  };
}

function normalizeSalesDataEmergencyFast_() {
  const src = getSheet_(CONFIG.SHEETS.SALES_IN);
  const out = getSheet_(CONFIG.SHEETS.SALES_CLEAN);
  const headersOut = CONFIG.HEADERS.SALES_CLEAN;
  const aliasMap = readBrandAliasMapFast_();

  if (!src || src.getLastRow() < 2) {
    replaceDataFastLimited_(out, headersOut, []);
    return { sourceRows: 0, salesObjects: [] };
  }

  // v6.00 안전 상한: 기존 readTable_ 기본 3,000행보다 조금 넓히되, 시간초과 방지를 위해 무제한 전체 스캔은 하지 않습니다.
  const values = getTrimmedSheetValues_(src, 5000, 80);
  if (!values || values.length < 2) {
    replaceDataFastLimited_(out, headersOut, []);
    return { sourceRows: 0, salesObjects: [] };
  }

  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const idx = buildSalesHeaderIndexesFast_(headers);
  const cleanRows = [];
  const salesObjects = [];
  let sourceRows = 0;

  for (var r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(function(v) { return String(v || '').trim() === ''; })) continue;
    sourceRows++;

    const orderDateRaw = getByIndexFast_(row, idx.orderDate);
    const marketName = getByIndexFast_(row, idx.marketName);
    const marketOrderNo = getByIndexFast_(row, idx.marketOrderNo);
    const acct = normalizeAccountId_(getByIndexFast_(row, idx.accountId));
    const marketProductNo = getByIndexFast_(row, idx.marketProductNo);
    const siteProductNo = getByIndexFast_(row, idx.siteProductNo);
    const brandRaw = String(getByIndexFast_(row, idx.brandRaw) || '').trim();
    const marketProductName = getByIndexFast_(row, idx.marketProductName);
    let originalProductName = getByIndexFast_(row, idx.originalProductName);
    if (!originalProductName) originalProductName = getByIndexFast_(row, idx.productName) || '';
    const productUrl = getByIndexFast_(row, idx.productUrl);
    const productImage = getByIndexFast_(row, idx.productImage);
    const purchaseSite = getByIndexFast_(row, idx.purchaseSite);
    const qty = toNumber_(getByIndexFast_(row, idx.qty)) || 1;
    const paymentAmount = toNumber_(getByIndexFast_(row, idx.paymentAmount));
    const shippingFee = toNumber_(getByIndexFast_(row, idx.shippingFee));
    const settlementAmount = toNumber_(getByIndexFast_(row, idx.settlementAmount));
    let marketStatus = getByIndexFast_(row, idx.marketStatus);
    const mangoStatus = getByIndexFast_(row, idx.mangoStatus);
    if (!marketStatus) marketStatus = getByIndexFast_(row, idx.status);

    const normalizedOrderDate = normalizeDateText_(orderDateRaw);
    const revenue = paymentAmount || settlementAmount;
    const isAnalysisTarget = CONFIG.LOTTEON_SITE_REGEX.test(String(purchaseSite || ''));
    const isCancelledInfo = isCancelledByMangoStatus_(mangoStatus);

    let matchedBrand = brandRaw;
    const aliasKey = normalizeBrandKey_(brandRaw);
    if (aliasKey && aliasMap[aliasKey]) matchedBrand = aliasMap[aliasKey];
    if (!matchedBrand || isGenericInvalidBrand_(matchedBrand)) matchedBrand = brandRaw || '브랜드미확인';

    const matchMethod = isAnalysisTarget && matchedBrand && matchedBrand !== '브랜드미확인' ? 'brandDirectFast' : '미매칭';
    const cleanRow = [
      monthKey_(orderDateRaw),
      normalizedOrderDate,
      marketName,
      marketOrderNo,
      acct.original,
      acct.normalized,
      acct.accountNo,
      marketProductNo,
      siteProductNo,
      brandRaw,
      matchedBrand,
      marketProductName,
      originalProductName,
      productUrl,
      productImage,
      purchaseSite,
      qty,
      revenue,
      shippingFee,
      settlementAmount,
      marketStatus,
      mangoStatus,
      'Y',
      isAnalysisTarget ? 'Y' : 'N',
      matchMethod,
      '',
      r + 1,
      isCancelledInfo ? 'Y' : 'N',
      '더망고주문상태'
    ];

    cleanRows.push(cleanRow);

    if (isAnalysisTarget) {
      const obj = {};
      headersOut.forEach(function(h, i) { obj[h] = cleanRow[i]; });
      salesObjects.push(obj);
    }
  }

  replaceDataFastLimited_(out, headersOut, cleanRows);
  return { sourceRows: sourceRows, salesObjects: salesObjects };
}

function buildSalesHeaderIndexesFast_(headers) {
  function idx(names) { return findHeaderIndexFast_(headers, names); }
  return {
    orderDate: idx(['마켓주문일자','주문일시','주문일','결제일시','결제일','주문등록일시']),
    marketName: idx(['마켓명','판매처','채널','쇼핑몰','판매마켓']),
    marketOrderNo: idx(['마켓주문번호','주문번호','주문ID','ordNo']),
    accountId: idx(['마켓아이디','판매자ID','계정ID','쿠팡계정ID','아이디']),
    marketProductNo: idx(['마켓상품번호','쿠팡상품번호']),
    siteProductNo: idx(['사이트상품번호','lotteon상품번호','롯데온상품번호','원상품번호']),
    brandRaw: idx(['브랜드','브랜드명','상표']),
    marketProductName: idx(['마켓상품명']),
    originalProductName: idx(['상품명(원문)','원문상품명','상품명원문','소싱상품명']),
    productName: idx(['상품명']),
    productUrl: idx(['상품URL','상품링크']),
    productImage: idx(['상품이미지','상품이미지URL','썸네일']),
    purchaseSite: idx(['구매사이트명','구매사이트','소싱사이트','소싱사이트명']),
    qty: idx(['결제수량','수량','주문수량','판매수량']),
    paymentAmount: idx(['결제금액합계(원)','결제금액합계','결제금액','판매금액','매출금액','상품금액','주문금액']),
    shippingFee: idx(['결제배송비(원)','결제배송비','배송비']),
    settlementAmount: idx(['정산예정금액(원)','정산예정금액','정산금액']),
    marketStatus: idx(['마켓주문상태']),
    mangoStatus: idx(['더망고주문상태']),
    status: idx(['주문상태','상태','배송상태'])
  };
}

function findHeaderIndexFast_(headers, names) {
  const normalized = (headers || []).map(function(h) { return normalizeHeaderKey_(h); });
  for (var i = 0; i < (names || []).length; i++) {
    const key = normalizeHeaderKey_(names[i]);
    const found = normalized.indexOf(key);
    if (found >= 0) return found;
  }
  return -1;
}

function getByIndexFast_(row, index) {
  if (index === undefined || index === null || index < 0) return '';
  return row[index];
}

function readBrandAliasMapFast_() {
  const map = {};
  try {
    readTable_(getSheet_(CONFIG.SHEETS.BRAND_ALIAS)).forEach(function(r) {
      const src = normalizeBrandKey_(getObjectValueByHeader_(r, '매출브랜드명'));
      const dst = String(getObjectValueByHeader_(r, '더망고브랜드명') || '').trim();
      if (src && dst) map[src] = dst;
    });
  } catch (e) {}
  return map;
}

function buildBrandMetricsEmergencyFast_(sales) {
  const brandMap = {};
  const soldSetByBrand = {};
  const latestSoldSetByBrand = {};

  (sales || []).forEach(function(s) {
    if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y') return;
    const method = String(s['매칭방식'] || '');
    if (method === '미매칭') return;

    const brand = String(s['브랜드명_매칭'] || s['브랜드명_원본'] || '').trim();
    if (!brand || isGenericInvalidBrand_(brand)) return;

    if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, '', '', '');
    if (!soldSetByBrand[brand]) soldSetByBrand[brand] = {};

    const b = brandMap[brand];
    const qty = toNumber_(s['결제수량']) || 1;
    const revenue = toNumber_(s['결제금액합계']);
    const month = String(s['주문월'] || '');
    const productKey = String(s['사이트상품번호'] || s['마켓상품번호'] || s['원문상품명'] || s['마켓상품명'] || '').trim();

    b.orderCount += qty;
    b.revenue += revenue;
    b.revenueProductUnmatched += revenue;
    b.ordersProductUnmatched += qty;

    if (String(s['취소여부'] || '') === 'Y') {
      b.cancelledOrderCount += qty;
      b.cancelledRevenue += revenue;
    }

    if (productKey) soldSetByBrand[brand][productKey] = true;
    if (month && (!b.latestMonth || month > b.latestMonth)) b.latestMonth = month;
  });

  (sales || []).forEach(function(s) {
    if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y') return;
    const method = String(s['매칭방식'] || '');
    if (method === '미매칭') return;
    const brand = String(s['브랜드명_매칭'] || s['브랜드명_원본'] || '').trim();
    const b = brandMap[brand];
    if (!b || !b.latestMonth || String(s['주문월'] || '') !== b.latestMonth) return;

    const qty = toNumber_(s['결제수량']) || 1;
    const revenue = toNumber_(s['결제금액합계']);
    const productKey = String(s['사이트상품번호'] || s['마켓상품번호'] || s['원문상품명'] || s['마켓상품명'] || '').trim();

    b.latestMonthOrders += qty;
    b.latestMonthRevenue += revenue;
    if (!latestSoldSetByBrand[brand]) latestSoldSetByBrand[brand] = {};
    if (productKey) latestSoldSetByBrand[brand][productKey] = true;
  });

  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand];
    b.soldProductCount = Object.keys(soldSetByBrand[brand] || {}).length;
    b.latestMonthSoldProducts = Object.keys(latestSoldSetByBrand[brand] || {}).length;
  });

  // 상품수/전송수/날짜는 기존 운영 시트의 최신 값을 반영합니다.
  try { applyFastFilterSummaryToBrandMetrics_(brandMap); } catch (e) { log_('v600_apply_filter_summary_skip', String(e && e.message ? e.message : e)); }
  try { applyCoupangSentManualToBrandMetrics_(brandMap); } catch (e) { log_('v600_apply_manual_skip', String(e && e.message ? e.message : e)); }
  try { applyFilterRawDatesToBrandMetrics_(brandMap); } catch (e) {}
  try { applyCleanupRefreshLogToBrandMetrics_(brandMap); } catch (e) {}
  try { applyRetransmitLogToBrandMetrics_(brandMap); } catch (e) {}
  try { applyAdditionalCollectionLogToBrandMetrics_(brandMap); } catch (e) {}

  return brandMap;
}

function refreshLotteonAnalysisNoAlert_() {
  const started = Date.now();

  // v5.98 초초경량:
  // 시간초과 방지를 위해 자동갱신에서는 아래를 제외합니다.
  // - 쿠팡재전송_로그 동기화
  // - 매칭진단 시트 재작성
  // - 전체 서식/열너비 갱신
  // 필요한 경우 각각 별도 메뉴에서 실행합니다.
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.SALES_CLEAN, CONFIG.HEADERS.SALES_CLEAN);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND, CONFIG.HEADERS.BRAND);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND_ALIAS, CONFIG.HEADERS.BRAND_ALIAS);

  const ctx = buildAnalysisContextWithoutProductList_();
  const salesResult = normalizeSalesData_([], ctx);
  const sales = salesResult.rows;

  const brandMetrics = buildBrandMetricsFastWithoutProductList_(sales);
  const brandRows = buildBrandRows_(brandMetrics);
  const summaryRows = buildBrandSummaryRowsFromBrandRows_(brandRows);

  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND), CONFIG.HEADERS.BRAND, brandRows);
  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND_SUMMARY), CONFIG.HEADERS.BRAND_SUMMARY, summaryRows);

  const elapsedSec = Math.round((Date.now() - started) / 1000);
  log_('auto_analysis_ultra_fast_v582', '매출분석 완료 / 매출정리행=' + sales.length + ', 브랜드=' + brandRows.length + ', 소요초=' + elapsedSec);

  return { salesRows: sales.length, brandRows: brandRows.length, elapsedSec: elapsedSec };
}

function generateAuditReportNoAlert_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheetWithHeader_(ss, CONFIG.SHEETS.AUDIT_REPORT, CONFIG.HEADERS.AUDIT_REPORT);
  const generatedAt = now_();

  const matchDiagRows = auditFastRows_(CONFIG.SHEETS.MATCH_DIAG);
  const summaryRows = auditFastRows_(CONFIG.SHEETS.BRAND_SUMMARY);
  const logRows = auditFastRows_(CONFIG.SHEETS.RETRANSMIT_LOG);
  const filterRows = auditFastRows_(CONFIG.SHEETS.FILTERS);
  const manualApiRows = auditFastRows_(CONFIG.SHEETS.MANUAL_API_CHECK);

  const summaryStats = auditBuildSummaryStats_(summaryRows);
  const logStats = auditBuildRetransmitStats_(logRows);
  const summaryTotal = auditSumStatsAmount_(summaryStats);
  const logTotal = auditSumStatsAmount_(logStats);

  const sourceRows = toNumber_(auditMetricValue_(matchDiagRows, '원본 매출행수'));
  const lotteonRows = toNumber_(auditMetricValue_(matchDiagRows, 'LOTTEON 분석대상 행수'));
  const analyzedRows = toNumber_(auditMetricValue_(matchDiagRows, '분석 반영 행수')) || toNumber_(auditMetricValue_(matchDiagRows, '전체매출 행수'));
  const cancelRows = toNumber_(auditMetricValue_(matchDiagRows, '더망고 취소/반품 주문행수'));
  const cancelRevenue = toNumber_(auditMetricValue_(matchDiagRows, '더망고 취소/반품 매출액'));
  const unmatchedRows = toNumber_(auditMetricValue_(matchDiagRows, '미매칭 행수'));
  const unmatchedRevenue = toNumber_(auditMetricValue_(matchDiagRows, '미매칭 매출액'));
  const analyzedRevenue = toNumber_(auditMetricValue_(matchDiagRows, '분석 반영 매출액'));

  const rows = [];
  let severeCount = 0;
  let warnCount = 0;

  const push = function(section, item, result, brand, filterName, v1, v2, diff, action, memo) {
    rows.push([section || '', item || '', result || '', brand || '', filterName || '', v1, v2, diff, action || '', memo || '']);
  };

  push('요약', '검수 생성시각', 'INFO', '', '', generatedAt, '', '', '검수리포트 확인', 'Asia/Seoul');
  push('요약', '실행모드', 'INFO', '', '', '자동 즉시요약', '', '', '정상', 'v5.98 자동 즉시요약');
  push('요약', '매출데이터_붙여넣기 원본행', sourceRows ? 'INFO' : '주의', '', '', sourceRows || auditSheetRowCount_(CONFIG.SHEETS.SALES_IN), '', '', sourceRows ? '정상' : '⑧ 분석/매출 원본 확인', '매칭진단 기준');
  push('요약', 'LOTTEON 분석대상 행수', lotteonRows ? 'INFO' : '주의', '', '', lotteonRows, '', '', lotteonRows ? '정상' : '구매사이트명/⑧ 분석 확인', '매칭진단 기준');
  push('요약', '분석 반영 행수', analyzedRows ? 'INFO' : '주의', '', '', analyzedRows, '', '', analyzedRows ? '정상' : '⑧ 분석 실행 확인', '취소/반품 포함 전체 주문반응 기준');
  push('요약', '분석 반영 매출액', analyzedRevenue ? 'INFO' : '주의', '', '', formatWon_(analyzedRevenue), '', '', analyzedRevenue ? '정상' : '⑧ 분석 실행 확인', '매칭진단 기준');
  push('요약', '핵심요약 전체매출 합계', 'INFO', '', '', formatWon_(summaryTotal), '', '', '참고', '핵심_브랜드요약 기준');
  push('요약', '재전송로그 누적매출 합계', 'INFO', '', '', formatWon_(logTotal), '', '', '참고', '쿠팡재전송_로그 기준');
  push('취소정보', '더망고 취소/반품 주문행수', 'INFO', '', '', cancelRows, '', '', '참고', '더망고주문상태 기준, 주 분석 매출에는 포함');
  push('취소정보', '더망고 취소/반품 매출액', 'INFO', '', '', formatWon_(cancelRevenue), '', '', '참고', '정보성 지표');

  if (unmatchedRows > 0 || unmatchedRevenue > 0) {
    severeCount++;
    push('매칭진단', '미매칭 존재', '긴급', '', '', unmatchedRows, formatWon_(unmatchedRevenue), '', '브랜드명_매칭표/매출정리 확인', '미매칭은 핵심요약 누락 가능성 있음');
  } else {
    push('매칭진단', '미매칭', '정상', '', '', 0, formatWon_(0), '', '조치 없음', '미매칭 행수/매출액 0');
  }

  let zeroApi = 0;
  (filterRows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const count = toNumber_(getObjectValueByHeader_(r, 'API_totalCount'));
    if (filterName && count === 0) zeroApi++;
  });
  if (zeroApi) {
    push('필터관리', 'API 상품수 0 필터 수', '참고', '', '', zeroApi, '', '', '필터별_상품수 시트 확인', '의도한 0개 필터인지 확인');
  }

  if (manualApiRows && manualApiRows.length) {
    const counts = {};
    manualApiRows.forEach(function(r) {
      const result = String(getObjectValueByHeader_(r, '판정') || '').trim() || '미분류';
      counts[result] = (counts[result] || 0) + 1;
    });
    Object.keys(counts).sort().forEach(function(k) {
      const level = /API 미확인|수동입력 없음|차이/.test(k) ? '주의' : '참고';
      if (level === '주의') warnCount += counts[k] ? 1 : 0;
      push('수동/API검증', k, level, '', '', counts[k], '', '', '수동입력_API검증 시트 상세 확인', '판정별 문제행 수');
    });
  } else {
    push('수동/API검증', '검증시트 행 없음', '참고', '', '', 0, '', '', '필요 시 수동입력/API 검증 시트 갱신', '문제행만 출력 방식이면 정상일 수 있음');
  }

  auditAppendMatchDiagLite_(push, matchDiagRows);
  push('요약', '긴급 이슈 수', severeCount ? '긴급' : '정상', '', '', severeCount, '', '', severeCount ? '미매칭부터 처리' : '긴급 이슈 없음', '매출 누락 가능성 중심');
  push('요약', '주의 이슈 수', warnCount ? '주의' : '정상', '', '', warnCount, '', '', warnCount ? '수동/API검증 요약 확인' : '주의 이슈 없음', '상세 문제행은 수동입력_API검증 시트에서 확인');

  replaceAuditReportQuick_(sheet, rows);
  formatAuditReportSheetInstant_(sheet);
  log_('auto_audit_report_done', '자동 검수리포트 생성 완료 / rows=' + rows.length + ', severe=' + severeCount + ', warn=' + warnCount);
  return { rows: rows.length, severe: severeCount, warn: warnCount };
}

function refreshDashboardFastOnlyNoAlert_() {
  // v5.98: 다른 메뉴/과거 호출 호환용 별칭입니다.
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
  ensureSheet_(SpreadsheetApp.getActive(), CONFIG.SHEETS.DASHBOARD);
  buildDashboard_(null, null, null);
  log_('dashboard_noalert_alias_v580', '대시보드 무알림 갱신 완료 / 핵심_브랜드요약 기준');
}

function refreshDashboardFastOnly() {
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
  ensureSheet_(SpreadsheetApp.getActive(), CONFIG.SHEETS.DASHBOARD);

  buildDashboard_(null, null, null);

  SpreadsheetApp.getUi().alert(
    '대시보드 빠른 갱신 완료\n\n' +
    '대시보드는 핵심_브랜드요약 기준의 매출/브랜드 판단표입니다.\n' +
    '- 브랜드별 매출액: 핵심_브랜드요약과 동일\n' +
    '- 수량/비율/금액 단위: 표시 문자열로 통일\n' +
    '- 필터/API 조치 요약은 필터_대시보드에 분리됩니다.'
  );
}

function scheduleDailyAutoRefreshContinuation_(delayMs) {
  deleteDailyAutoRefreshContinuationTriggers_();
  ScriptApp.newTrigger('runDailyAutoRefreshContinue')
    .timeBased()
    .after(Math.max(15 * 1000, delayMs || 60 * 1000))
    .create();
}

function deleteDailyAutoRefreshScheduleTriggers_() {
  return deleteTriggersByHandlerNames_(['runDailyAutoRefreshStart']);
}

function deleteDailyAutoRefreshContinuationTriggers_() {
  return deleteTriggersByHandlerNames_(['runDailyAutoRefreshContinue']);
}

function deleteTriggersByHandlerNames_(handlerNames) {
  const names = {};
  (handlerNames || []).forEach(function(n) { names[n] = true; });
  let count = 0;
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    try {
      const h = t.getHandlerFunction ? t.getHandlerFunction() : '';
      if (names[h]) {
        ScriptApp.deleteTrigger(t);
        count++;
      }
    } catch (e) {}
  });
  return count;
}

function dailyAutoRefreshFail_(err, stage) {
  const msg = String(err && err.message ? err.message : err);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AUTO_REFRESH_RUNNING', 'N');
  props.setProperty('AUTO_REFRESH_STAGE', stage || 'ERROR');
  props.setProperty('AUTO_REFRESH_LAST_ERROR', msg);
  props.setProperty('AUTO_REFRESH_LAST_MESSAGE', '자동 갱신 오류');
  props.setProperty('AUTO_REFRESH_LAST_RESULT', '');
  props.setProperty('AUTO_REFRESH_LAST_UPDATED', now_());
  writeAutoRefreshStatus_(stage || 'ERROR', '자동 갱신 오류', msg);
  log_('auto_refresh_error', 'stage=' + stage + ', error=' + msg);
}

function ensureAutoRefreshStatusSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheetWithHeader_(ss, '자동갱신_상태', ['항목', '값', '메모']);
  try { sheet.setColumnWidth(1, 190); sheet.setColumnWidth(2, 260); sheet.setColumnWidth(3, 420); } catch (e) {}
  return sheet;
}

function writeAutoRefreshStatus_(stage, message, error) {
  const props = PropertiesService.getScriptProperties();
  const sheet = ensureAutoRefreshStatusSheet_();
  const rows = [
    ['사용여부', props.getProperty('AUTO_REFRESH_ENABLED') || 'N', '매일 예약 사용 여부'],
    ['실행중', props.getProperty('AUTO_REFRESH_RUNNING') || 'N', '현재 자동 파이프라인 실행 여부'],
    ['현재단계', stage || props.getProperty('AUTO_REFRESH_STAGE') || '', 'ULTRA_ANALYSIS / ULTRA_DASHBOARD / DONE / ERROR'],
    ['예약시간', props.getProperty('AUTO_REFRESH_SCHEDULE') || '예약 비활성', 'v6.00에서는 수동 1회 실행 우선'],
    ['최근시작', props.getProperty('AUTO_REFRESH_STARTED_AT') || '', '마지막 자동 갱신 시작 시각'],
    ['최근완료', props.getProperty('AUTO_REFRESH_LAST_DONE_AT') || '', '마지막 자동 갱신 완료 시각'],
    ['최근갱신', now_(), '상태 기록 시각'],
    ['최근메시지', message || props.getProperty('AUTO_REFRESH_LAST_MESSAGE') || '', '진행 설명'],
    ['최근결과', props.getProperty('AUTO_REFRESH_LAST_RESULT') || '', '성공 시 처리 결과 요약'],
    ['최근오류', error || props.getProperty('AUTO_REFRESH_LAST_ERROR') || '', '오류 발생 시 원문'],
    ['자동실행범위', '빠른 매출정리 → 빠른 핵심요약', 'v6.00 시간초과 방지: 대시보드/검수리포트/로그/매칭진단/필터API/수동검증/필터대시보드 제외'],
    ['검수리포트 기준', '필요할 때만 별도 메뉴로 실행', 'LOTTEON 자동화 → ★ 전체 검수 리포트 생성(즉시요약)'],
    ['매출데이터 기준', '매출데이터_붙여넣기 현재 내용', '더망고 주문 API 연동 전까지는 수동 붙여넣기 데이터 기준']
  ];

  try { sheet.getRange(1, 1, Math.max(sheet.getLastRow(), rows.length + 1), 3).clearContent(); } catch (e) {}
  sheet.getRange(1, 1, 1, 3).setValues([['항목', '값', '메모']]);
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');
    sheet.getRange(1, 1, rows.length + 1, 3).setVerticalAlignment('middle');
    sheet.setColumnWidth(1, 190);
    sheet.setColumnWidth(2, 420);
    sheet.setColumnWidth(3, 560);
  } catch (e) {}
}



function prepareLotteonOneClick() {
  const result = prepareLotteonCore_();
  if (!result) return;
  SpreadsheetApp.getUi().alert(
    '원클릭 준비 완료\n\nAPI 연결 성공\nfilterList 1페이지 조회 행수: ' + result.rows +
    '\n\n이제 ⑤ LOTTEON 상품목록 분할 업데이트 시작을 실행하면 됩니다.'
  );
}

function prepareAndStartProductSyncOneClick() {
  const result = prepareLotteonCore_();
  if (!result) return;
  startLotteonProductSync();
  SpreadsheetApp.getUi().alert(
    '원클릭 준비 및 상품목록 업데이트 시작 완료\n\nAPI 연결 성공: filterList 1페이지 ' + result.rows + '행\n' +
    '약 30초 뒤부터 상품목록 분할 업데이트가 자동으로 이어집니다.\n\n진행상황은 동기화상태 시트에서 확인하세요.'
  );
}

function prepareLotteonCore_() {
  const ui = SpreadsheetApp.getUi();
  try {
    removeLotteonTriggers();
    setupLotteonSheets(false);
    resetProductSyncStateSilent_();
    const ok = ensureCredentialsInteractive_();
    if (!ok) return null;
    const result = runApiConnectionTestSilent_();
    log_('one_click_prepare', '원클릭 준비 완료 / filterList rows=' + result.rows);
    return result;
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    writeSyncStatus_({ phase: CONFIG.PHASE.ERROR, status: '원클릭 준비 실패', lastError: msg });
    log_('one_click_prepare_error', msg);
    ui.alert('원클릭 준비 실패\n\n' + msg);
    throw e;
  }
}

function resetProductSyncStateSilent_() {
  const props = PropertiesService.getScriptProperties();
  ['SYNC_ACTIVE','SYNC_PHASE','SYNC_FILTER_NAMES','SYNC_FILTER_INDEX','SYNC_PAGE',
   'SYNC_FILTER_LIST_PAGE','SYNC_LAST_ERROR','SYNC_CURRENT_FILTER','SYNC_CURRENT_TOTAL_PAGE',
   'SYNC_LAST_MESSAGE','SYNC_STARTED_AT','SYNC_FINISHED_AT','SYNC_TOKEN_USED_EST','LAST_API_URL'
  ].forEach(function(k) { props.deleteProperty(k); });
  deleteTriggersForHandler_('continueLotteonProductSync');
  deleteTriggersForHandler_('runAnalysisAfterSync');
  writeSyncStatus_({ phase: CONFIG.PHASE.IDLE, status: '원클릭 준비 중 초기화 완료', lastUrl: '', lastError: '' });
  log_('product_sync', '동기화 상태 초기화(원클릭 내부)');
}

function ensureCredentialsInteractive_() {
  const props = PropertiesService.getScriptProperties();
  const savedApiKey = String(props.getProperty('THE_MANGO_API_KEY') || '').trim();
  const savedSender = String(props.getProperty('THE_MANGO_SENDER') || '').trim();
  if (savedApiKey && savedSender) {
    if (!props.getProperty('THE_MANGO_BASE_URL')) props.setProperty('THE_MANGO_BASE_URL', CONFIG.BASE_URL_DEFAULT);
    return true;
  }
  const ui = SpreadsheetApp.getUi();
  const baseUrlResp = ui.prompt('TheMangoClient URL', '기본값: https://tmg2007.cafe24.com', ui.ButtonSet.OK_CANCEL);
  if (baseUrlResp.getSelectedButton() !== ui.Button.OK) return false;
  const apiKeyResp = ui.prompt('API Key 입력', 'API Key를 붙여넣으세요. 실제 값은 시트에 노출하지 않습니다.', ui.ButtonSet.OK_CANCEL);
  if (apiKeyResp.getSelectedButton() !== ui.Button.OK) return false;
  const senderResp = ui.prompt('X-API-SENDER 입력', 'X-API-SENDER를 붙여넣으세요. 실제 값은 시트에 노출하지 않습니다.', ui.ButtonSet.OK_CANCEL);
  if (senderResp.getSelectedButton() !== ui.Button.OK) return false;
  const baseUrl = normalizeBaseUrl_(baseUrlResp.getResponseText() || CONFIG.BASE_URL_DEFAULT);
  const apiKey = String(apiKeyResp.getResponseText() || '').trim();
  const sender = String(senderResp.getResponseText() || '').trim();
  if (!apiKey || !sender) { ui.alert('API Key와 X-API-SENDER는 필수입니다.'); return false; }
  props.setProperty('THE_MANGO_BASE_URL', baseUrl);
  props.setProperty('THE_MANGO_API_KEY', apiKey);
  props.setProperty('THE_MANGO_SENDER', sender);
  log_('credentials', '원클릭 흐름에서 인증값 저장 완료. 실제 값은 미표시');
  return true;
}

function runApiConnectionTestSilent_() {
  const testUrl = buildApiUrl_('filterList');
  writeSyncStatus_({ phase: 'API_TEST', status: 'API 연결 테스트 중', currentFilter: 'filterList', lastUrl: maskSensitiveUrl_(testUrl) });
  const resp = apiRequest_('filterList', 'post', { page: '1', searchQuery: { searchKeyword: '', siteId: '', filterGroup: 'all', sort: 'nameAsc' } });
  const data = extractData_(resp);
  const items = extractItems_(resp, false);
  writeSyncStatus_({ phase: 'API_TEST', status: 'API 연결 성공', currentFilter: 'filterList', totalPage: data.totalPage || '', lastUrl: maskSensitiveUrl_(testUrl) });
  log_('api_test', 'filterList 연결 성공 / rows=' + items.length + ', totalPage=' + (data.totalPage || ''));
  return { rows: items.length, totalPage: data.totalPage || '' };
}


/**
 * v5.49 수정: 필터/상품수만 빠른 갱신
 * 핵심 원칙: 필터별_상품수를 단일 원천으로 사용.
 * 4개 시트 동기화: 필터별_상품수 → 쿠팡재전송_로그 → 수동입력_API검증 → 대시보드
 * 금지 작업: LOTTEON_상품목록 읽기, 매출데이터 분석, 전체 시트 서식 적용, setupLotteonSheets 전체 실행
 */

function findFilterItemCount_(item) {
  const keys = ['itemCount', 'item_count', 'productCount', 'product_count', 'count', 'totalCount', 'total_count'];
  for (var i = 0; i < keys.length; i++) {
    const raw = item && item[keys[i]];
    if (raw !== undefined && raw !== null && String(raw).trim() !== '') return toNumber_(raw);
  }
  for (const k in item || {}) {
    const lk = String(k || '').toLowerCase();
    if ((lk.indexOf('item') >= 0 || lk.indexOf('product') >= 0 || lk.indexOf('total') >= 0) &&
        lk.indexOf('count') >= 0) {
      const raw = item[k];
      if (raw !== undefined && raw !== null && String(raw).trim() !== '') return toNumber_(raw);
    }
  }
  return 0;
}

function findFilterId_(item) {
  const keys = ['filterId', 'filterID', 'filter_id', 'id'];
  for (var i = 0; i < keys.length; i++) {
    const v = item && item[keys[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function findFilterSiteId_(item) {
  const keys = ['siteId', 'siteID', 'site_id', 'site'];
  for (var i = 0; i < keys.length; i++) {
    const v = item && item[keys[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function findFilterGroupName_(item) {
  const keys = ['groupName', 'filterGroupName', 'filterGroup', 'group_name', 'group'];
  for (var i = 0; i < keys.length; i++) {
    const v = item && item[keys[i]];
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function buildFilterSummaryRowsFromFilterListItems_(filterItems) {
  const seen = {};
  const rows = [];

  (filterItems || []).forEach(function(item) {
    const filterName = String(findFilterName_(item) || '').trim();
    if (!filterName) return;
    if (!isValidLotteonFilterName_(filterName)) return;
    if (seen[filterName]) return;
    seen[filterName] = true;

    const account = accountFromFilterName_(filterName);
    const brand = brandFromFilterName_(filterName);
    const meta = extractFilterListMeta_(item) || {};
    const itemCount = findFilterItemCount_(item);
    const filterId = findFilterId_(item);
    const siteId = findFilterSiteId_(item);
    const groupName = findFilterGroupName_(item);

    rows.push([
      filterName,
      brand,
      account.accountNo,
      account.accountId,
      itemCount,
      '',
      itemCount,
      filterCodeFromFilterName_(filterName),
      meta.recentDate || '',
      meta.createDate || '',
      meta.recentField || '',
      meta.createField || '',
      'filterList.itemCount 기준 / filterId=' + filterId + (siteId ? ' / siteId=' + siteId : '') + (groupName ? ' / group=' + groupName : '')
    ]);
  });

  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
  return rows;
}


function extractFilterListItemsFromResponse_(resp) {
  const data = extractData_(resp);
  if (data && Array.isArray(data.filters)) return data.filters.filter(isObject_);
  if (data && data.data && Array.isArray(data.data.filters)) return data.data.filters.filter(isObject_);
  if (Array.isArray(data)) return data.filter(isObject_);
  if (resp && resp.data && resp.data.filters && Array.isArray(resp.data.filters)) return resp.data.filters.filter(isObject_);
  if (resp && resp.filters && Array.isArray(resp.filters)) return resp.filters.filter(isObject_);
  return extractItems_(resp, false).filter(isObject_);
}

function fetchFilterListItemsByKeyword_(keyword) {
  keyword = String(keyword || '').trim();
  if (!keyword) return [];

  const all = [];
  const seen = {};
  const maxPages = 6;

  for (var page = 1; page <= maxPages; page++) {
    const payload = {
      page: String(page),
      searchQuery: {
        searchKeyword: keyword,
        siteId: '',
        filterGroup: 'all',
        sort: 'nameAsc'
      }
    };

    const resp = apiRequest_('filterList', 'post', payload);
    const items = extractFilterListItemsFromResponse_(resp);

    if (!items || !items.length) break;

    items.forEach(function(item) {
      const name = String(findFilterName_(item) || '').trim();
      const key = name || JSON.stringify(item).slice(0, 200);
      if (!seen[key]) {
        seen[key] = true;
        all.push(item);
      }
    });

    if (items.length < 20) break;
    Utilities.sleep(420);
  }

  return all;
}

function findExactFilterListItemByFilterName_(filterName) {
  const target = String(filterName || '').trim();
  if (!target) return null;

  const items = fetchFilterListItemsByKeyword_(target);
  for (var i = 0; i < items.length; i++) {
    const name = String(findFilterName_(items[i]) || '').trim();
    if (name === target) return items[i];
  }

  return null;
}

function findFilterListItemByFilterNameRobust_(filterName) {
  // v5.98: 정확 검색필터명만 허용합니다.
  return findExactFilterListItemByFilterName_(filterName);
}

function buildFilterSearchKeywordCandidates_(filterName) {
  // v5.98: 정확 필터명만 사용
  const f = String(filterName || '').trim();
  return f ? [f] : [];
}

function findFilterListItemByFilterNameRobust_(filterName) {
  // v5.98: 정확 검색필터명만 허용합니다.
  return findExactFilterListItemByFilterName_(filterName);
}

function pickExactFilterListItem_(items, filterName) {
  const target = String(filterName || '').trim();
  if (!target) return null;

  for (var i = 0; i < (items || []).length; i++) {
    const name = String(findFilterName_(items[i]) || '').trim();
    if (name === target) return items[i];
  }
  return null;
}

function collectRetransmitLogTargetFilterNames_(brandRows) {
  const set = {};
  function add(v) {
    const s = String(v || '').trim();
    if (s && isValidLotteonFilterName_(s)) set[s] = true;
  }

  // 1) 브랜드분석 현재 대상
  const headers = CONFIG.HEADERS.BRAND;
  (brandRows || []).forEach(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    add(obj['대표검색필터명']);
  });

  // 2) 쿠팡재전송_로그 현재 대상
  try {
    const logRows = readTable_(getSheet_(CONFIG.SHEETS.RETRANSMIT_LOG));
    (logRows || []).forEach(function(r) {
      add(getObjectValueByHeader_(r, '검색필터명'));
    });
  } catch (e) {}

  // 3) 쿠팡전송수_수동입력 현재 대상
  try {
    const manualRows = readTable_(getSheet_(CONFIG.SHEETS.COUPANG_SENT_MANUAL));
    (manualRows || []).forEach(function(r) {
      add(getObjectValueByHeader_(r, '검색필터명'));
    });
  } catch (e) {}

  return Object.keys(set).sort();
}

function mergeFilterRowsIntoFilterSummarySheet_(newRows) {
  newRows = newRows || [];
  if (!newRows.length) return { updated: 0, total: 0 };

  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);

  const existingRows = [];
  try {
    const sheet = getSheet_(CONFIG.SHEETS.FILTERS);
    if (sheet && sheet.getLastRow() >= 2) {
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.HEADERS.FILTERS.length).getValues();
      values.forEach(function(row) {
        if (String(row[0] || '').trim()) existingRows.push(row);
      });
    }
  } catch (e) {}

  const map = {};
  existingRows.forEach(function(row) {
    const name = String(row[0] || '').trim();
    if (name) map[name] = row;
  });
  newRows.forEach(function(row) {
    const name = String(row[0] || '').trim();
    if (name) map[name] = row;
  });

  const merged = Object.keys(map).sort().map(function(k) { return map[k]; });
  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.FILTERS), CONFIG.HEADERS.FILTERS, merged);

  return { updated: newRows.length, total: merged.length };
}

function refreshFilterSummaryForRetransmitLogTargetsNoAlert_(brandRows) {
  const started = Date.now();
  const filterNames = collectRetransmitLogTargetFilterNames_(brandRows);
  const rows = [];
  const misses = [];
  const errors = [];
  const maxMillis = 210000;

  for (var i = 0; i < filterNames.length; i++) {
    if ((Date.now() - started) > maxMillis) {
      errors.push('시간보호 중단: ' + (i + 1) + '/' + filterNames.length + '번째에서 중단');
      break;
    }

    const filterName = filterNames[i];
    try {
      const item = findFilterListItemByFilterNameRobust_(filterName);
      if (item) {
        const built = buildFilterSummaryRowsFromFilterListItems_([item]);
        if (built && built.length) rows.push(built[0]);
        else misses.push(filterName);
      } else {
        misses.push(filterName);
      }
    } catch (e) {
      errors.push(filterName + ': ' + String(e && e.message ? e.message : e).slice(0, 160));
    }

    Utilities.sleep(420);
  }

  const mergeResult = mergeFilterRowsIntoFilterSummarySheet_(rows);

  writeSyncStatus_({
    phase: errors.length ? 'FILTERLIST_TARGET_PARTIAL' : 'FILTERLIST_TARGET_DONE',
    status: '쿠팡재전송_로그 대상 필터 filterList 개별 최신화 완료',
    filterCount: rows.length,
    filterIndex: filterNames.length,
    currentFilter: '',
    lastUrl: buildApiUrl_('filterList'),
    tmpRows: mergeResult.total,
    lastError: errors.slice(0, 3).join(' | ')
  });

  log_(
    'filterlist_target_refresh_v591',
    'targets=' + filterNames.length +
    ', updated=' + rows.length +
    ', misses=' + misses.length +
    ', errors=' + errors.length
  );

  return {
    targetCount: filterNames.length,
    updatedCount: rows.length,
    missCount: misses.length,
    errors: errors,
    elapsedSec: Math.round((Date.now() - started) / 1000)
  };
}


function refreshFilterCountsByFilterListNoAlert_() {
  // v5.98: 내부 호출도 1페이지 안전 실행만 수행합니다.
  return runFilterListResumeBatch_({ startFresh: false });
}

function refreshFilterCountsFast() {
  const ui = SpreadsheetApp.getUi();
  try {
    deleteFastFilterCountTriggers_();

    const existingState = getFilterListResumeState_();
    const result = runFilterListResumeBatch_({ startFresh: !existingState });

    if (result.done) {
      ui.alert(
        '필터별_상품수 갱신 완료\n\n' +
        '기준: filterList.itemCount\n' +
        '검색어: 롯백\n' +
        '처리 필터 수: ' + result.filterCount + '개\n' +
        '이번 실행: 1페이지 이하\n' +
        '소요초: ' + result.elapsedSec + '초\n\n' +
        '갱신된 시트:\n' +
        '  ① 필터별_상품수\n\n' +
        '다음 단계:\n' +
        '③ 쿠팡재전송_로그 갱신 또는 ④ 핵심요약+대시보드 갱신을 실행하세요.'
      );
    } else {
      ui.alert(
        '필터별_상품수 1페이지 처리 완료\n\n' +
        '기준: filterList.itemCount\n' +
        '검색어: 롯백\n' +
        '현재 누적 필터 수: ' + result.filterCount + '개\n' +
        '다음 시작 페이지: ' + result.nextPage + '\n' +
        '소요초: ' + result.elapsedSec + '초\n\n' +
        '시간초과 방지를 위해 1페이지만 처리했습니다.\n' +
        '완료될 때까지 같은 메뉴를 다시 실행하세요.'
      );
    }
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    writeSyncStatus_({
      phase: 'FILTERLIST_RESUME_ERROR',
      status: 'filterList 1페이지 안전 실행 실패',
      currentFilter: 'filterList',
      lastUrl: buildApiUrl_('filterList'),
      lastError: msg
    });
    ui.alert('필터별_상품수 갱신 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function resetFilterListResumeProgress() {
  resetFilterListResumeWork_();
  writeSyncStatus_({
    phase: 'FILTERLIST_RESUME_RESET',
    status: 'filterList 이어실행 상태 초기화 완료',
    currentFilter: '',
    lastUrl: buildApiUrl_('filterList'),
    tmpRows: 0,
    lastError: ''
  });
  SpreadsheetApp.getUi().alert('필터별_상품수 filterList 이어실행 상태를 초기화했습니다.');
}

/**
 * v5.49 신규: 수동입력_API검증 경량 버전
 * LOTTEON_상품목록을 읽지 않고 필터별_상품수(방금 갱신)와 쿠팡전송수_수동입력만 비교합니다.
 * 비교 항목: 수동_더망고수집수 vs API_더망고수집수(API_totalCount)
 * 참고: 수동_쿠팡전송수 vs API_마켓전송확인수 비교는 LOTTEON_상품목록 필요 → 빠른 갱신에서는 생략
 */
/**
 * v5.51 유지: 필터별_상품수 시트를 경량 검증용 map으로 변환합니다.
 * LOTTEON_상품목록 전체를 읽지 않습니다.
 */
function buildFilterSummaryMapFromFilterSheet_() {
  const rows = readTable_(getSheet_(CONFIG.SHEETS.FILTERS));
  const map = {};

  (rows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    if (!filterName) return;
    if (typeof isValidLotteonFilterName_ === 'function' && !isValidLotteonFilterName_(filterName)) return;

    const account = accountFromFilterName_(filterName);
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim() || brandFromFilterName_(filterName);
    if (!brand) return;

    map[filterName] = {
      '검색필터명': filterName,
      '브랜드명': brand,
      '계정번호': getObjectValueByHeader_(r, '계정번호') || account.accountNo || '',
      '쿠팡계정ID': getObjectValueByHeader_(r, '쿠팡계정ID') || account.accountId || '',
      'API_totalCount': getObjectValueByHeader_(r, 'API_totalCount'),
      'API_totalPage': getObjectValueByHeader_(r, 'API_totalPage'),
      '이번조회_행수': getObjectValueByHeader_(r, '이번조회_행수'),
      '필터코드': getObjectValueByHeader_(r, '필터코드') || filterCodeFromFilterName_(filterName),
      'API_최근수집일자': getObjectValueByHeader_(r, 'API_최근수집일자'),
      'API_필터생성일': getObjectValueByHeader_(r, 'API_필터생성일'),
      'API_최근수집일자_필드': getObjectValueByHeader_(r, 'API_최근수집일자_필드'),
      'API_필터생성일_필드': getObjectValueByHeader_(r, 'API_필터생성일_필드'),
      '메모': getObjectValueByHeader_(r, '메모') || '필터별_상품수 기준'
    };
  });

  return map;
}


function getFilterSummaryDateInfoForLog_(filterName, brand, filterSummaryMap) {
  filterSummaryMap = filterSummaryMap || {};
  const exact = String(filterName || '').trim();
  let row = exact ? filterSummaryMap[exact] : null;

  if (!row) {
    const targetBrandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
    Object.keys(filterSummaryMap).some(function(k) {
      const r = filterSummaryMap[k] || {};
      const b = String(r['브랜드명'] || '').trim() || brandFromFilterName_(k);
      const bk = normalizeCoupangWorkLogBrandMergeKey_(b, k);
      if (bk && bk === targetBrandKey) {
        row = r;
        return true;
      }
      return false;
    });
  }

  // v5.98:
  // 로그 날짜는 API 필터 생성/최근수집일만 신뢰합니다.
  // 오늘 이후 날짜는 잘못 파싱된 값으로 보고 무효 처리합니다.
  const createDate = isValidApiFilterDateForLog_(row ? row['API_필터생성일'] : '');
  const recentDate = isValidApiFilterDateForLog_(row ? row['API_최근수집일자'] : '');

  let additionalDate = '';
  if (createDate && recentDate && recentDate.slice(0, 10) > createDate.slice(0, 10)) {
    additionalDate = recentDate;
  }

  return {
    hasApiRow: !!row,
    hasDateInfo: !!(createDate || recentDate),
    filterName: row ? String(row['검색필터명'] || exact || '').trim() : exact,
    brand: row ? String(row['브랜드명'] || brand || '').trim() : String(brand || '').trim(),
    accountId: row ? String(row['쿠팡계정ID'] || '').trim() : '',
    createDate: createDate,
    recentDate: recentDate,
    additionalDate: additionalDate,
    createDisplay: formatShortDate_(createDate),
    recentDisplay: formatShortDate_(recentDate),
    additionalDisplay: formatShortDate_(additionalDate)
  };
}

function pickLaterDateText_(a, b) {
  const ad = normalizeDateText_(a), bd = normalizeDateText_(b);
  if (!ad) return bd || '';
  if (!bd) return ad || '';
  return bd >= ad ? bd : ad;
}

function pickLaterShortDateFromRaw_(a, b) {
  return formatShortDate_(pickLaterDateText_(a, b));
}


function todayDateKey_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function currentYear_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy');
}

function normalizeDateForComparison_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  }

  const raw = String(v || '').trim();
  if (!raw) return '';

  const full = raw.match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})(?:[ T](\d{1,2}):?(\d{0,2}):?(\d{0,2}))?/);
  if (full) {
    const yy = full[1];
    const mm = ('0' + full[2]).slice(-2);
    const dd = ('0' + full[3]).slice(-2);
    const hh = ('0' + (full[4] || '00')).slice(-2);
    const mi = ('0' + (full[5] || '00')).slice(-2);
    const ss = ('0' + (full[6] || '00')).slice(-2);
    return yy + '-' + mm + '-' + dd + ' ' + hh + ':' + mi + ':' + ss;
  }

  // v5.98:
  // 쿠팡재전송_로그와 필터별_상품수에는 M/dd 형태가 많습니다.
  // 오늘 날짜 기준 연도를 붙여 비교 가능한 yyyy-MM-dd로 변환합니다.
  const md = raw.match(/^(\d{1,2})[\/\.\-\s]+(\d{1,2})$/);
  if (md) {
    const yy = currentYear_();
    const mm = ('0' + md[1]).slice(-2);
    const dd = ('0' + md[2]).slice(-2);
    return yy + '-' + mm + '-' + dd + ' 00:00:00';
  }

  return normalizeDateText_(raw);
}

function isFutureDateText_(v) {
  const d = normalizeDateForComparison_(v);
  if (!d) return false;
  const key = d.slice(0, 10);
  if (!key.match(/^\d{4}-\d{2}-\d{2}$/)) return false;
  return key > todayDateKey_();
}

function isValidApiFilterDateForLog_(v) {
  const d = normalizeDateForComparison_(v);
  if (!d) return '';
  const key = d.slice(0, 10);
  if (!key.match(/^\d{4}-\d{2}-\d{2}$/)) return '';
  if (key > todayDateKey_()) return '';
  return d;
}


function normalizeRetransmitLogDateForDashboard_(v) {
  const d = isValidApiFilterDateForLog_(v);
  return d || '';
}

function retransmitLogDateIsLater_(a, b) {
  const ad = normalizeRetransmitLogDateForDashboard_(a);
  const bd = normalizeRetransmitLogDateForDashboard_(b);
  if (!ad) return !!bd;
  if (!bd) return false;
  return bd.slice(0, 10) > ad.slice(0, 10);
}

function mergeRetransmitLogDateInfo_(base, incoming) {
  base = base || {};
  incoming = incoming || {};

  function earlier(a, b) {
    const ad = normalizeRetransmitLogDateForDashboard_(a);
    const bd = normalizeRetransmitLogDateForDashboard_(b);
    if (!ad) return bd || '';
    if (!bd) return ad || '';
    return ad.slice(0, 10) <= bd.slice(0, 10) ? ad : bd;
  }

  function later(a, b) {
    const ad = normalizeRetransmitLogDateForDashboard_(a);
    const bd = normalizeRetransmitLogDateForDashboard_(b);
    if (!ad) return bd || '';
    if (!bd) return ad || '';
    return ad.slice(0, 10) >= bd.slice(0, 10) ? ad : bd;
  }

  return {
    filterName: incoming.filterName || base.filterName || '',
    brand: incoming.brand || base.brand || '',
    accountId: incoming.accountId || base.accountId || '',
    firstSendDate: earlier(base.firstSendDate, incoming.firstSendDate),
    additionalDate: later(base.additionalDate, incoming.additionalDate),
    retransmitDate: later(base.retransmitDate, incoming.retransmitDate),
    cleanupDate: later(base.cleanupDate, incoming.cleanupDate),
    productCount: toNumber_(incoming.productCount) || toNumber_(base.productCount),
    sentCount: toNumber_(incoming.sentCount) || toNumber_(base.sentCount),
    source: '쿠팡재전송_로그'
  };
}

function readRetransmitLogDateMap_() {
  const result = { byFilter: {}, byBrand: {}, byBrandAccount: {} };
  let rows = [];
  try {
    rows = readTable_(getSheet_(CONFIG.SHEETS.RETRANSMIT_LOG));
  } catch (e) {
    return result;
  }

  (rows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim();
    const accountId = String(getObjectValueByHeader_(r, '쿠팡계정ID') || getObjectValueByHeader_(r, '쿠팡\n계정ID') || '').trim();
    if (!filterName && !brand) return;

    const info = {
      filterName: filterName,
      brand: brand || brandFromFilterName_(filterName),
      accountId: accountId,
      firstSendDate: normalizeRetransmitLogDateForDashboard_(getObjectValueByHeader_(r, '최초전송일') || getObjectValueByHeader_(r, '최초\n전송일')),
      additionalDate: normalizeRetransmitLogDateForDashboard_(getObjectValueByHeader_(r, '1단계_추가수집일') || getObjectValueByHeader_(r, '추가수집일')),
      retransmitDate: normalizeRetransmitLogDateForDashboard_(getObjectValueByHeader_(r, '2단계_재전송일') || getObjectValueByHeader_(r, '재전송일')),
      cleanupDate: normalizeRetransmitLogDateForDashboard_(getObjectValueByHeader_(r, '3단계_정리재수집일') || getObjectValueByHeader_(r, '정리재수집일')),
      productCount: toNumber_(getObjectValueByHeader_(r, '최종_더망고수집수')),
      sentCount: toNumber_(getObjectValueByHeader_(r, '최종_쿠팡전송수')),
      source: '쿠팡재전송_로그'
    };

    if (!info.firstSendDate && !info.additionalDate && !info.retransmitDate && !info.cleanupDate) return;

    if (filterName) {
      result.byFilter[filterName] = mergeRetransmitLogDateInfo_(result.byFilter[filterName], info);
    }

    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(info.brand || brand, filterName);
    if (brandKey) {
      result.byBrand[brandKey] = mergeRetransmitLogDateInfo_(result.byBrand[brandKey], info);
    }

    const baKey = makeBrandAccountMergeKey_(info.brand || brand, accountId);
    if (baKey) {
      result.byBrandAccount[baKey] = mergeRetransmitLogDateInfo_(result.byBrandAccount[baKey], info);
    }
  });

  return result;
}

function resolveRetransmitLogDateInfo_(filterName, brand, accountId, logMap) {
  logMap = logMap || readRetransmitLogDateMap_();

  const f = String(filterName || '').trim();
  if (f && logMap.byFilter && logMap.byFilter[f]) return logMap.byFilter[f];

  const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
  if (brandKey && logMap.byBrand && logMap.byBrand[brandKey]) return logMap.byBrand[brandKey];

  const baKey = makeBrandAccountMergeKey_(brand, accountId);
  if (baKey && logMap.byBrandAccount && logMap.byBrandAccount[baKey]) return logMap.byBrandAccount[baKey];

  return null;
}

function applyRetransmitLogDatesToBrandObj_(b, logMap) {
  b = b || {};
  const info = resolveRetransmitLogDateInfo_(b['대표검색필터명'], b['브랜드명'], b['쿠팡계정ID'], logMap);
  if (!info) return b;

  const out = {};
  Object.keys(b).forEach(function(k) { out[k] = b[k]; });

  // 상품갈이 기준 우선순위:
  // 최근 정리재수집일 → 최근 재전송일 → 1단계_추가수집일 → 최초 전송일/필터 생성일
  if (info.cleanupDate) {
    out['정리재수집여부'] = 'Y';
    out['최근정리재수집일'] = info.cleanupDate;
  }
  if (info.retransmitDate) {
    out['재전송여부'] = 'Y';
    out['최근재전송일'] = info.retransmitDate;
  }
  if (info.additionalDate) {
    out['추가수집여부'] = 'Y';
    out['최근추가수집일'] = info.additionalDate;
  }
  if (info.firstSendDate) {
    out['대표쿠팡전송일'] = info.firstSendDate;
    out['최초쿠팡전송일'] = info.firstSendDate;
    if (!out['더망고기준등록일']) out['더망고기준등록일'] = info.firstSendDate;
  }

  if (info.sentCount > 0) out['쿠팡전송확인상품수'] = info.sentCount;
  if (info.productCount > 0) out['LOTTEON_더망고등록상품수'] = info.productCount;

  out['상품갈이_로그연동여부'] = 'Y';
  return out;
}

function buildRotationPlanFromDashboardItemWithLog_(x, logMap) {
  const info = resolveRetransmitLogDateInfo_(x.filterName, x.brand, x.accountId, logMap);
  if (!info) return x;

  const sentCount = toNumber_(x.sent);
  const soldProducts = toNumber_(x.sold);
  const unsoldCount = Math.max(0, sentCount - soldProducts);

  const baseDate = normalizeRetransmitLogDateForDashboard_(info.cleanupDate || info.retransmitDate || info.additionalDate || info.firstSendDate);
  const stage1Base = normalizeRetransmitLogDateForDashboard_(info.additionalDate || info.cleanupDate || baseDate);
  const stage2Base = normalizeRetransmitLogDateForDashboard_(info.retransmitDate || info.cleanupDate || baseDate);
  const stage3Base = normalizeRetransmitLogDateForDashboard_(info.cleanupDate || baseDate);

  const stage1 = calcRotationStatus_(stage1Base, CONFIG.ROTATION_STAGE1_DAYS, true, false, sentCount);
  const stage2 = calcRotationStatus_(stage2Base, CONFIG.ROTATION_STAGE2_DAYS, unsoldCount > 0, true, sentCount);
  const stage3 = calcRotationStatus_(stage3Base, CONFIG.ROTATION_STAGE3_DAYS, unsoldCount > 0, true, sentCount);

  let nextAction = '대기';
  if (sentCount <= 0) nextAction = '전송확인';
  else if (stage3 === '필요') nextAction = '3단계';
  else if (stage2 === '필요') nextAction = '2단계';
  else if (stage1 === '필요') nextAction = '1단계';

  const out = {};
  Object.keys(x || {}).forEach(function(k) { out[k] = x[k]; });
  out.nextAction = nextAction;
  out.stage1 = stage1;
  out.stage2 = stage2;
  out.stage3 = stage3;
  out.rotationMemo = makeRotationMemo_(nextAction, unsoldCount, baseDate);
  out.rotationMemo = out.rotationMemo ? out.rotationMemo + ' / 로그기준' : '로그기준';
  return out;
}



function writeManualInputApiValidationFast_(filterSummaryMap, options) {
  options = options || {};
  const issueOnly = options.issueOnly !== false; // 기본값: 문제행만
  const manual = readCoupangSentManualMap_();
  const lookup = buildFilterSummaryLookup_(filterSummaryMap, manual);
  const rows = [];
  const usedApiFilters = {};
  let totalChecked = 0;

  function shouldOutput_(result) {
    if (!issueOnly) return true;
    // 시간초과 방지를 위해 정상/비활성행은 기본 출력에서 제외합니다.
    return result !== '일치' && result !== '비활성행';
  }

  // 1) 수동입력 행을 우선 검증
  (manual.rows || []).forEach(function(item) {
    if (!item || (!item.filterName && !item.brand)) return;
    totalChecked++;

    const resolved = resolveManualItemToFilterSummary_(item, lookup);
    const apiData = resolved.apiData;
    const outputFilterName = resolved.filterName || item.filterName || '';
    const apiExists = !!apiData;

    if (apiExists && outputFilterName) usedApiFilters[outputFilterName] = true;

    const brand = (item && item.brand) ||
      (apiData && apiData['브랜드명']) ||
      brandFromFilterName_(outputFilterName) ||
      '';

    const manualSent = item.hasTotal ? item.total : '';
    const manualMango = item.hasMangoCount ? item.mangoCount : '';

    const apiTotalCount = apiData ? toNumber_(apiData['API_totalCount']) : 0;
    const mangoDiff = (item.hasMangoCount && apiExists) ? (toNumber_(manualMango) - apiTotalCount) : '';

    let result = '확인';
    const memos = [];

    if (item.inactive) {
      result = '비활성행';
      memos.push(String(item.statusRaw || '') + ' 표시 - 분석/대표로그 제외');
    } else if (!apiExists) {
      result = 'API 미확인';
      memos.push('필터별_상품수에 해당 필터 없음');
    } else if (resolved.changed) {
      if (item.hasMangoCount && mangoDiff !== 0 && mangoDiff !== '') {
        result = '필터명 변경+수집수 차이';
        memos.push('수동필터 ' + item.filterName + ' → API필터 ' + outputFilterName);
        memos.push('수집수 차이 ' + (mangoDiff > 0 ? '+' : '') + formatCount_(mangoDiff));
      } else {
        result = '필터명 변경반영';
        memos.push('수동필터 ' + item.filterName + ' → API필터 ' + outputFilterName);
      }
    } else if (item.hasMangoCount && mangoDiff !== 0 && mangoDiff !== '') {
      result = '수집수 차이';
      memos.push('수집수 차이 ' + (mangoDiff > 0 ? '+' : '') + formatCount_(mangoDiff));
    } else {
      result = '일치';
      memos.push('수동 더망고수집수와 API_totalCount 일치');
    }

    memos.push('[v5.55 문제행 검증: 필터별_상품수 API_totalCount 기준 / 전송수 비교는 ⑧에서 확인]');

    if (shouldOutput_(result)) {
      rows.push([
        formatManualStatusForDisplay_(item),
        outputFilterName,
        brand,
        manualSent,
        '',
        '',
        manualMango,
        apiExists ? apiTotalCount : '',
        mangoDiff,
        result,
        memos.join(' / ')
      ]);
    }
  });

  // 2) API에만 존재하고 수동입력에 없는 필터
  Object.keys(filterSummaryMap || {}).sort().forEach(function(filterName) {
    if (!isValidLotteonFilterName_(filterName)) return;
    if (usedApiFilters[filterName]) return;
    totalChecked++;

    const apiData = filterSummaryMap[filterName];
    const brand = (apiData && apiData['브랜드명']) || brandFromFilterName_(filterName) || '';
    const apiTotalCount = toNumber_(apiData && apiData['API_totalCount']);
    const result = '수동입력 없음';

    if (shouldOutput_(result)) {
      rows.push([
        'API만존재',
        filterName,
        brand,
        '',
        '',
        '',
        '',
        apiTotalCount,
        '',
        result,
        '쿠팡전송수_수동입력에 해당 필터 없음 / [v5.55 문제행 검증: 필터별_상품수 API_totalCount 기준]'
      ]);
    }
  });

  replaceManualInputApiValidationDataUltraFast_(rows);
  return { writtenRows: rows.length, totalRows: totalChecked };
}

function replaceManualInputApiValidationDataUltraFast_(rows) {
  const sheet = getSheet_(CONFIG.SHEETS.MANUAL_API_CHECK);
  const headers = CONFIG.HEADERS.MANUAL_API_CHECK;
  const data = [headers].concat(rows || []);
  const newRows = data.length;
  const cols = headers.length;

  // v5.55: sheet.clearContents() 금지. 기존 사용 범위 중 필요한 열만 최소 clear.
  const oldLastRow = Math.max(sheet.getLastRow(), 1);
  const clearRows = Math.max(oldLastRow, newRows);
  try {
    sheet.getRange(1, 1, clearRows, cols).clearContent();
  } catch (e) {}

  sheet.getRange(1, 1, newRows, cols).setValues(data);
  formatManualInputApiValidationSheetUltraLight_(sheet, newRows, cols);
}

function formatManualInputApiValidationSheetUltraLight_(sheet, rowCount, colCount) {
  if (!sheet) return;
  const rows = rowCount || Math.max(sheet.getLastRow(), 1);
  const cols = colCount || Math.max(sheet.getLastColumn(), 1);
  const dataRows = Math.max(rows - 1, 0);

  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, cols)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 42);
  } catch (e) {}

  // v5.55: 전체 서식 리셋/General 적용 금지. 주요 숫자 컬럼만 최소 서식.
  if (dataRows > 0) {
    try { sheet.getRange(2, 4, dataRows, 6).setNumberFormat('#,##0'); } catch (e) {}
    try { sheet.getRange(2, 1, dataRows, cols).setWrap(false).setVerticalAlignment('middle'); } catch (e) {}
  }

  const widths = [80, 190, 130, 90, 90, 80, 95, 105, 80, 110, 320];
  widths.forEach(function(w, i) {
    try { sheet.setColumnWidth(i + 1, w); } catch (e) {}
  });
}

function buildFilterSummaryLookup_(filterSummaryMap, manual) {
  const byFilter = {};
  const byBrand = {};

  Object.keys(filterSummaryMap || {}).forEach(function(filterName) {
    if (!isValidLotteonFilterName_(filterName)) return;

    const row = filterSummaryMap[filterName] || {};
    const brand = String(row['브랜드명'] || '').trim() || brandFromFilterName_(filterName);
    if (!brand) return;

    const item = {
      filterName: filterName,
      brand: brand,
      row: row,
      code: toNumber_(filterCodeFromFilterName_(filterName)),
      totalCount: toNumber_(row['API_totalCount'])
    };

    byFilter[filterName] = item;

    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
    if (!brandKey) return;

    const prev = byBrand[brandKey];
    if (!prev || preferFilterSummaryLookupItem_(item, prev, manual)) {
      byBrand[brandKey] = item;
    }
  });

  return { byFilter: byFilter, byBrand: byBrand };
}

function resolveManualItemToFilterSummary_(item, lookup) {
  const filterName = String(item.filterName || '').trim();

  if (filterName && lookup.byFilter[filterName]) {
    return { filterName: filterName, apiData: lookup.byFilter[filterName].row, changed: false };
  }

  if (item.inactive) {
    return { filterName: filterName, apiData: null, changed: false };
  }

  const brandKey = normalizeCoupangWorkLogBrandMergeKey_(item.brand || brandFromFilterName_(filterName), filterName);
  const candidate = brandKey ? lookup.byBrand[brandKey] : null;

  if (!candidate) {
    return { filterName: filterName, apiData: null, changed: false };
  }

  // v5.53: 02→04 같은 계정번호 변경은 자동 매칭하되, 04→00처럼 운영코드가 아닌 필터로 내려가는 건 허용하지 않습니다.
  // isValidLotteonFilterName_에서 롯백_00은 이미 제외되므로 candidate는 01~04만 남습니다.
  return { filterName: candidate.filterName, apiData: candidate.row, changed: candidate.filterName !== filterName };
}

function preferFilterSummaryLookupItem_(candidate, current, manual) {
  const cManual = manual && manual.byFilter ? manual.byFilter[candidate.filterName] : null;
  const pManual = manual && manual.byFilter ? manual.byFilter[current.filterName] : null;

  const cPriority = cManual ? 3 : 1;
  const pPriority = pManual ? 3 : 1;
  if (cPriority !== pPriority) return cPriority > pPriority;

  // 같은 브랜드가 여러 계정에 있으면 더 높은 운영코드(예: 04)를 최신 대표 후보로 둡니다.
  if (candidate.code !== current.code) return candidate.code > current.code;

  const cLen = String(candidate.filterName || '').length;
  const pLen = String(current.filterName || '').length;
  if (cLen !== pLen) return cLen < pLen;

  return String(candidate.filterName || '') > String(current.filterName || '');
}

function formatManualStatusForDisplay_(item) {
  if (!item) return '';
  if (item.statusType === 'NORMAL' && item.confirmDate) return formatShortDate_(item.confirmDate);
  if (item.confirmDate && normalizeDateText_(item.statusRaw)) return formatShortDate_(item.confirmDate);
  return item.statusRaw || formatShortDate_(item.confirmDate) || '';
}

/**
 * v5.49 신규: 필터/상품수 기반 경량 대시보드 갱신
 * 매출데이터 없이 필터별_상품수(방금 갱신)만으로 대시보드를 구성합니다.
 * 매출 분석 포함 전체 대시보드는 ⑧ 메뉴에서 buildDashboard_()로 재생성합니다.
 */
function buildDashboardActionIssues_(filterSummaryMap) {
  const manual = readCoupangSentManualMap_();
  const lookup = buildFilterSummaryLookup_(filterSummaryMap, manual);
  const issues = [];
  const usedApiFilters = {};
  const summary = {
    totalChecked: 0,
    issueCount: 0,
    apiMissing: 0,
    manualMissing: 0,
    collectDiff: 0,
    collectDiffBig: 0,
    filterChanged: 0,
    zeroApiCount: 0
  };

  function pushIssue_(priority, result, filterName, brand, manualMango, apiMango, diff, action, memo) {
    const d = diff === '' || diff === null || diff === undefined ? '' : toNumber_(diff);
    if (result === 'API 미확인') summary.apiMissing++;
    if (result === '수동입력 없음') summary.manualMissing++;
    if (String(result).indexOf('수집수 차이') >= 0) summary.collectDiff++;
    if (Math.abs(toNumber_(d)) >= 100) summary.collectDiffBig++;
    if (String(result).indexOf('필터명 변경') >= 0) summary.filterChanged++;
    if (toNumber_(apiMango) === 0) summary.zeroApiCount++;

    issues.push({
      priority: priority,
      result: result,
      filterName: filterName || '',
      brand: brand || '',
      manualMango: manualMango === undefined || manualMango === null ? '' : manualMango,
      apiMango: apiMango === undefined || apiMango === null ? '' : apiMango,
      diff: d,
      action: action || '',
      memo: memo || ''
    });
  }

  (manual.rows || []).forEach(function(item) {
    if (!item || (!item.filterName && !item.brand)) return;
    summary.totalChecked++;

    const resolved = resolveManualItemToFilterSummary_(item, lookup);
    const apiData = resolved.apiData;
    const outputFilterName = resolved.filterName || item.filterName || '';
    const apiExists = !!apiData;
    if (apiExists && outputFilterName) usedApiFilters[outputFilterName] = true;

    const brand = item.brand || (apiData && apiData['브랜드명']) || brandFromFilterName_(outputFilterName) || '';
    const manualMango = item.hasMangoCount ? item.mangoCount : '';
    const apiMango = apiExists ? toNumber_(apiData['API_totalCount']) : '';
    const diff = (item.hasMangoCount && apiExists) ? (toNumber_(manualMango) - toNumber_(apiMango)) : '';

    if (item.inactive) return;

    if (!apiExists) {
      pushIssue_(1, 'API 미확인', outputFilterName, brand, manualMango, '', '', '더망고에 필터 존재 여부 확인 후 삭제/변경 처리', '수동입력에는 있으나 필터별_상품수에는 없음');
      return;
    }

    if (resolved.changed) {
      if (item.hasMangoCount && diff !== 0 && diff !== '') {
        pushIssue_(2, '필터명 변경+수집수 차이', outputFilterName, brand, manualMango, apiMango, diff, '수동입력표를 최신 필터명/수량으로 정리', '수동필터 ' + item.filterName + ' → API필터 ' + outputFilterName);
      } else {
        pushIssue_(3, '필터명 변경반영', outputFilterName, brand, manualMango, apiMango, diff, '정상 이동 확인 후 수동입력표 최신화', '수동필터 ' + item.filterName + ' → API필터 ' + outputFilterName);
      }
      return;
    }

    if (item.hasMangoCount && diff !== 0 && diff !== '') {
      const priority = Math.abs(toNumber_(diff)) >= 100 ? 2 : 4;
      pushIssue_(priority, '수집수 차이', outputFilterName, brand, manualMango, apiMango, diff, '차이가 크면 더망고 팝업 기준으로 수동값 업데이트', '수동_더망고수집수와 API_totalCount 차이');
    }
  });

  Object.keys(filterSummaryMap || {}).sort().forEach(function(filterName) {
    if (!isValidLotteonFilterName_(filterName)) return;
    if (usedApiFilters[filterName]) return;
    summary.totalChecked++;

    const apiData = filterSummaryMap[filterName];
    const brand = (apiData && apiData['브랜드명']) || brandFromFilterName_(filterName) || '';
    const apiMango = toNumber_(apiData && apiData['API_totalCount']);
    pushIssue_(1, '수동입력 없음', filterName, brand, '', apiMango, '', '쿠팡전송수_수동입력에 행 추가', 'API에는 있으나 수동입력표에는 없음');
  });

  issues.sort(function(a, b) {
    if (a.priority !== b.priority) return a.priority - b.priority;
    const ad = Math.abs(toNumber_(a.diff));
    const bd = Math.abs(toNumber_(b.diff));
    if (ad !== bd) return bd - ad;
    return String(a.filterName || '').localeCompare(String(b.filterName || ''));
  });

  summary.issueCount = issues.length;
  return { issues: issues, summary: summary };
}

function formatDashboardAction_(sheet, rowCount, colCount) {
  if (!sheet) return;
  const rows = rowCount || Math.max(sheet.getLastRow(), 1);
  const cols = colCount || Math.max(sheet.getLastColumn(), 1);
  const dataRows = Math.max(rows - 1, 0);

  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, cols)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 42);
  } catch (e) {}

  if (dataRows > 0) {
    try {
      sheet.getRange(2, 1, dataRows, cols)
        .setWrap(false)
        .setVerticalAlignment('middle');
      sheet.getRange(2, 3, dataRows, 6).setNumberFormat('#,##0');
    } catch (e) {}
  }

  const widths = [90, 210, 140, 115, 115, 95, 95, 140, 120, 210, 340];
  widths.forEach(function(w, i) {
    try { sheet.setColumnWidth(i + 1, w); } catch (e) {}
  });
}

function getFilterDashboardSheet_() {
  return ensureSheet_(SpreadsheetApp.getActive(), '필터_대시보드');
}

function refreshFilterDashboardFastOnly() {
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);

  const filterSummaryMap = buildFilterSummaryMapFromFilterSheet_();
  buildDashboardFastFromFilterSummary_(filterSummaryMap);

  SpreadsheetApp.getUi().alert(
    '필터_대시보드 갱신 완료\n\n' +
    '필터/API 조치 요약은 이제 대시보드가 아니라 필터_대시보드에 표시됩니다.\n' +
    '매출/브랜드 판단 대시보드는 "대시보드 빠른 갱신"을 실행하세요.'
  );
}

function buildDashboardFastFromFilterSummary_(filterSummaryMap) {
  // v5.77:
  // 필터/API 조치 요약은 메인 "대시보드"를 덮어쓰지 않고 "필터_대시보드"에만 기록합니다.
  // 메인 대시보드는 핵심_브랜드요약 기준의 매출/브랜드 판단표로 고정합니다.
  const sheet = getFilterDashboardSheet_();

  const accountStat = {};
  const filterList = [];

  Object.keys(filterSummaryMap || {}).sort().forEach(function(filterName) {
    if (!isValidLotteonFilterName_(filterName)) return;

    const row = filterSummaryMap[filterName] || {};
    const accountId = String(row['쿠팡계정ID'] || '').trim();
    const accountNo = String(row['계정번호'] || '').trim();
    const totalCount = toNumber_(row['API_totalCount']) || 0;
    const brand = String(row['브랜드명'] || '').trim();

    const key = accountId || ('계정' + accountNo);
    if (!accountStat[key]) {
      accountStat[key] = { accountId: accountId, accountNo: accountNo, filterCount: 0, productCount: 0, zeroCount: 0 };
    }

    accountStat[key].filterCount += 1;
    accountStat[key].productCount += totalCount;
    if (totalCount === 0) accountStat[key].zeroCount += 1;
    filterList.push([filterName, brand, accountId || accountNo, totalCount]);
  });

  const totalFilters = filterList.length;
  const totalProducts = filterList.reduce(function(s, r) { return s + toNumber_(r[3]); }, 0);
  const issueResult = buildDashboardActionIssues_(filterSummaryMap);
  const issues = issueResult.issues;
  const issueSummary = issueResult.summary;

  const output = [];
  output.push(CONFIG.HEADERS.DASHBOARD);

  output.push(['요약', '대시보드 용도', '필터/API 조치 필요 요약판', '', '', '', '', '', '', '', '매출/브랜드 대시보드는 메인 대시보드에서 확인']);
  output.push(['요약', '갱신 기준', '필터별_상품수 + 쿠팡전송수_수동입력', '', '', '', '', '', '', '', '매출 분석과 분리된 필터 관리용']);
  output.push(['요약', '갱신 시각', now_(), '', '', '', '', '', '', '', 'Asia/Seoul']);
  output.push(['요약', '총 필터수', totalFilters, '', '', '', '', '', '', '', '롯백_01~04 유효 검색필터']);
  output.push(['요약', '총 API 상품수', totalProducts, '', '', '', '', '', '', '', '필터별_상품수 API_totalCount 합계']);
  output.push(['요약', '조치 필요 행수', issueSummary.issueCount, '', '', '', '', '', '', '', '아래 조치 필요 목록 기준']);
  output.push(['요약', 'API 미확인', issueSummary.apiMissing, '', '', '', '', '', '', '', '수동입력에는 있으나 API 필터 없음']);
  output.push(['요약', '수동입력 없음', issueSummary.manualMissing, '', '', '', '', '', '', '', 'API에는 있으나 수동입력표 없음']);
  output.push(['요약', '수집수 차이 100개 이상', issueSummary.collectDiffBig, '', '', '', '', '', '', '', '먼저 확인할 큰 차이']);
  output.push(['요약', '필터명 변경 관련', issueSummary.filterChanged, '', '', '', '', '', '', '', '02→04 등 이동 반영']);
  output.push(['', '', '', '', '', '', '', '', '', '', '']);

  output.push(['구분', '계정ID', '필터수', 'API 상품수', '0개 필터수', '', '', '', '', '', '메모']);
  Object.keys(accountStat).sort().forEach(function(key) {
    const a = accountStat[key];
    output.push(['계정요약', a.accountId || key, a.filterCount, a.productCount, a.zeroCount, '', '', '', '', '', '계정별 상품 쏠림/0개 필터 확인']);
  });
  output.push(['', '', '', '', '', '', '', '', '', '', '']);

  output.push(['우선순위', '판정', '검색필터명', '브랜드명', '수동수집수', 'API수집수', '차이', '해야 할 일', '', '', '메모']);

  issues.slice(0, 80).forEach(function(x) {
    output.push([
      x.priority,
      x.result,
      x.filterName,
      x.brand,
      x.manualMango,
      x.apiMango,
      x.diff,
      x.action,
      '',
      '',
      x.memo
    ]);
  });

  if (!issues.length) {
    output.push(['정상', '조치 필요 없음', '', '', '', '', '', '현재 문제행 없음', '', '', '수동입력_API검증도 확인 권장']);
  }

  const rows = output.length;
  const cols = output[0].length;
  const clearRows = Math.max(sheet.getLastRow(), rows);
  const clearCols = Math.max(sheet.getLastColumn(), cols);

  try { sheet.getRange(1, 1, clearRows, clearCols).clearContent(); } catch (e) {}

  sheet.getRange(1, 1, rows, cols).setValues(output);
  sheet.setFrozenRows(1);
  formatDashboardAction_(sheet, rows, cols);
}

function refreshDashboardFastOnly() {
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
  ensureSheet_(SpreadsheetApp.getActive(), CONFIG.SHEETS.DASHBOARD);

  buildDashboard_(null, null, null);

  SpreadsheetApp.getUi().alert(
    '대시보드 빠른 갱신 완료\n\n' +
    '대시보드는 핵심_브랜드요약 기준의 매출/브랜드 판단표입니다.\n' +
    '- 브랜드별 매출액: 핵심_브랜드요약과 동일\n' +
    '- 수량/비율/금액 단위: 표시 문자열로 통일\n' +
    '- 필터/API 조치 요약은 필터_대시보드에 분리됩니다.'
  );
}

/**
 * v5.49 신규: 빠른 갱신 후 대시보드 최소 서식 적용 (전체 스캔 금지)
 */
function formatDashboardFast_(sheet, rowCount, colCount) {
  if (!sheet) return;

  const lastRow = rowCount || Math.max(sheet.getLastRow(), 1);
  const lastCol = colCount || Math.max(sheet.getLastColumn(), 1);
  const dataRows = Math.max(lastRow - 1, 0);

  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 36);
  } catch (e) {}

  if (dataRows > 0) {
    try {
      sheet.getRange(2, 1, dataRows, lastCol)
        .setWrap(false)
        .setVerticalAlignment('middle');
    } catch (e) {}

    // v5.51: 셀별 조건부 서식 루프 제거. 주요 숫자 컬럼만 한 번에 서식 적용.
    try { sheet.getRange(2, 3, dataRows, 3).setNumberFormat('#,##0'); } catch (e) {}
  }

  // 고정폭으로 빠르게 적용. 전체 데이터 폭 스캔 금지.
  const colWidths = [80, 230, 140, 110, 110, 80, 80, 80, 80, 80, 280];
  colWidths.forEach(function(w, i) {
    try { sheet.setColumnWidth(i + 1, w); } catch (e) {}
  });
}


function syncCoupangWorkLogFromFastFilterSummary_(filterSummaryMap) {
  const sheet = ensureCoupangWorkLogSheet_();
  const headers = CONFIG.HEADERS.RETRANSMIT_LOG;
  const manual = readCoupangSentManualMap_();
  const existingRows = readTable_(sheet);
  const latestByBrand = buildLatestFilterSummaryByBrand_(filterSummaryMap, manual);
  const merged = {};

  existingRows.forEach(function(r) {
    const oldFilterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const oldBrand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim() || brandFromFilterName_(oldFilterName);
    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(oldBrand, oldFilterName);
    const latest = latestByBrand[brandKey];

    if (oldFilterName && isInactiveManualFilterName_(oldFilterName, manual) && !latest) return;

    const obj = {};
    headers.forEach(function(h) { obj[normalizeHeaderKey_(h)] = getObjectValueByHeader_(r, h); });

    if (latest) {
      obj['검색필터명'] = latest.filterName;
      obj['브랜드명'] = latest.brand;
      obj['쿠팡계정ID'] = latest.accountId || obj['쿠팡계정ID'] || '';
      if (latest.totalCount !== '') obj['최종_더망고수집수'] = latest.totalCount;
      const manualItem = manual.byFilter[latest.filterName] || manual.byBrand[brandKey] || null;
      if (manualItem && manualItem.hasTotal) obj['최종_쿠팡전송수'] = manualItem.total;
    }

    const key = latest ? ('BRAND|' + normalizeCoupangWorkLogBrandMergeKey_(latest.brand, latest.filterName)) :
      getCoupangWorkLogMergeKey_(r, {});
    if (!key) return;
    if (!merged[key]) merged[key] = obj;
    else merged[key] = mergeCoupangWorkLogObjects_(merged[key], obj);
  });

  Object.keys(latestByBrand).forEach(function(brandKey) {
    const latest = latestByBrand[brandKey];
    const key = 'BRAND|' + brandKey;
    if (merged[key]) return;
    const manualItem = manual.byFilter[latest.filterName] || manual.byBrand[brandKey] || null;
    const obj = {
      '검색필터명': latest.filterName, '브랜드명': latest.brand, '쿠팡계정ID': latest.accountId || '',
      '최초전송일': '', '누적매출액': '', '월평균매출액': '',
      '작업유형': manualItem && manualItem.isAdded ? '신규전송' : '',
      '1단계_추가수집일': '', '2단계_재전송일': '', '3단계_정리재수집일': '',
      '최종_더망고수집수': latest.totalCount,
      '최종_쿠팡전송수': manualItem && manualItem.hasTotal ? manualItem.total : '',
      '추가수집상품수': '', '작업사유': '', '비고': ''
    };
    merged[key] = obj;
  });

  const out = Object.keys(merged).sort(function(a, b) {
    return String(merged[a]['검색필터명'] || '').localeCompare(String(merged[b]['검색필터명'] || ''));
  }).map(function(key) {
    return headers.map(function(h) { return getCoupangWorkLogValueByAlias_(merged[key], h); });
  });

  replaceDataFastLimited_(sheet, headers, out);
  formatCoupangWorkLogSheetFastOnly_(sheet);
}

function formatCoupangWorkLogSheetFastOnly_(sheet) {
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const dataRows = Math.max(lastRow - 1, 0);
  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#e8f0fe').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
    sheet.setRowHeight(1, 54);
  } catch (e) {}
  if (dataRows > 0) {
    try {
      sheet.getRange(2, 1, dataRows, lastCol).setWrap(false).setVerticalAlignment('middle');
      sheet.getRange(2, 4, dataRows, 1).setNumberFormat('M/dd');
      sheet.getRange(2, 5, dataRows, 2).setNumberFormat('"₩"#,##0');
      sheet.getRange(2, 8, dataRows, 3).setNumberFormat('M/dd');
      sheet.getRange(2, 11, dataRows, 3).setNumberFormat('#,##0');
    } catch (e) {}
  }
  var widths = [170, 130, 95, 70, 95, 105, 95, 85, 85, 95, 95, 95, 95, 150, 230];
  widths.forEach(function(w, i) { try { sheet.setColumnWidth(i + 1, w); } catch (e) {} });
}

function buildLatestFilterSummaryByBrand_(filterSummaryMap, manual) {
  const latestByBrand = {};
  Object.keys(filterSummaryMap || {}).forEach(function(filterName) {
    if (!isValidLotteonFilterName_(filterName)) return;
    if (isInactiveManualFilterName_(filterName, manual)) return;
    const row = filterSummaryMap[filterName] || {};
    const brand = String(row['브랜드명'] || '').trim() || brandFromFilterName_(filterName);
    if (!brand) return;
    const account = accountFromFilterName_(filterName);
    const item = {
      filterName: filterName, brand: brand,
      accountNo: row['계정번호'] || account.accountNo || '',
      accountId: row['쿠팡계정ID'] || account.accountId || '',
      totalCount: String(row['API_totalCount'] || '').trim() !== '' ? toNumber_(row['API_totalCount']) : ''
    };
    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
    const prev = latestByBrand[brandKey];
    if (!prev || preferFastFilterSummary_(item, prev, manual)) latestByBrand[brandKey] = item;
  });
  return latestByBrand;
}

function preferFastFilterSummary_(candidate, current, manual) {
  const cManual = manual && manual.byFilter ? manual.byFilter[candidate.filterName] : null;
  const pManual = manual && manual.byFilter ? manual.byFilter[current.filterName] : null;
  const cPriority = cManual ? 3 : 1;
  const pPriority = pManual ? 3 : 1;
  if (cPriority !== pPriority) return cPriority > pPriority;
  const cCode = toNumber_(filterCodeFromFilterName_(candidate.filterName));
  const pCode = toNumber_(filterCodeFromFilterName_(current.filterName));
  if (cCode !== pCode) return cCode > pCode;
  const cLen = String(candidate.filterName || '').length;
  const pLen = String(current.filterName || '').length;
  if (cLen !== pLen) return cLen < pLen;
  return String(candidate.filterName || '') > String(current.filterName || '');
}

function getFastFilterCountTmpSheetName_() {
  return '빠른갱신_TMP';
}

function getFastFilterCountTmpHeaders_() {
  return [
    '검색필터명',
    '브랜드명',
    '계정번호',
    '쿠팡계정ID',
    'API_totalCount',
    'API_totalPage',
    '필터코드',
    'API_최근수집일자',
    'API_필터생성일',
    'API_최근수집일자_필드',
    'API_필터생성일_필드',
    '처리상태',
    '오류',
    '갱신시각'
  ];
}

function getFastFilterCountTmpSheet_() {
  return ensureSheet_(SpreadsheetApp.getActive(), getFastFilterCountTmpSheetName_());
}

function ensureFastFilterCountTmpSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheetWithHeader_(ss, getFastFilterCountTmpSheetName_(), getFastFilterCountTmpHeaders_());
  try { sheet.hideSheet(); } catch (e) {}
  return sheet;
}

function continueFastFilterCountSync() {
  // v5.98: 과거 이어실행 트리거가 남아 호출될 경우에도 1페이지 안전 실행만 수행합니다.
  return runFilterListResumeBatch_({ startFresh: false });
}

function processFastFilterCountBatch_(startedAt, isManualStart) {
  // v5.98: productList 배치 대신 filterList 1페이지 안전 실행을 사용합니다.
  const result = runFilterListResumeBatch_({ startFresh: false });
  return {
    done: result.done,
    doneCount: result.filterCount,
    pendingCount: result.done ? 0 : -1,
    totalCount: result.filterCount,
    elapsedSec: result.elapsedSec
  };
}

function buildFilterSummaryMapFromFastTmpValues_(values) {
  const map = {};
  (values || []).forEach(function(r) {
    const filterName = String(r[0] || '').trim();
    if (!isValidLotteonFilterName_(filterName)) return;
    map[filterName] = {
      '검색필터명': filterName,
      '브랜드명': String(r[1] || '').trim() || brandFromFilterName_(filterName),
      '계정번호': r[2],
      '쿠팡계정ID': r[3],
      'API_totalCount': toNumber_(r[4]),
      'API_totalPage': toNumber_(r[5]),
      '이번조회_행수': 0,
      '필터코드': r[6] || filterCodeFromFilterName_(filterName),
      'API_최근수집일자': r[7] || '',
      'API_필터생성일': r[8] || '',
      'API_최근수집일자_필드': r[9] || '',
      'API_필터생성일_필드': r[10] || '',
      '메모': '분할 빠른갱신: productList 1페이지 totalCount 기준 / 상품상세 미수집'
    };
  });
  return map;
}

function scheduleFastFilterCountContinuation_() {
  deleteFastFilterCountTriggers_();
  ScriptApp.newTrigger('continueFastFilterCountSync')
    .timeBased()
    .after(60 * 1000)
    .create();
}

function deleteFastFilterCountTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    try {
      if (t.getHandlerFunction && t.getHandlerFunction() === 'continueFastFilterCountSync') {
        ScriptApp.deleteTrigger(t);
      }
    } catch (e) {}
  });
}

function fetchLotteonFilterItemsFast_() {
  const all = [];
  let page = 1;
  while (page <= CONFIG.MAX_PAGE_PER_QUERY) {
    const payload = { page: String(page), searchQuery: { searchKeyword: '', siteId: '', filterGroup: 'all', sort: 'nameAsc' } };
    const resp = apiRequest_('filterList', 'post', payload);
    const data = extractData_(resp);
    const items = extractItems_(resp, false);
    if (items && items.length) items.forEach(function(item) { all.push(item); });
    const totalPage = toNumber_(data.totalPage);
    if (!items.length || (totalPage && page >= totalPage)) break;
    page++;
    Utilities.sleep(CONFIG.REQUEST_DELAY_MS);
  }
  return all;
}


function apiRequestFilterListSingle_(payload) {
  const cred = getCredentials_();
  const url = buildApiUrl_('filterList');
  const props = PropertiesService.getScriptProperties();

  props.setProperty('LAST_API_URL', maskSensitiveUrl_(url));
  props.setProperty('LAST_API_ROUTE', 'FILTERLIST_SINGLE');

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + cred.apiKey,
      'X-API-SENDER': cred.sender
    },
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('filterList 단일 호출 실패 HTTP ' + code + ': ' + text.slice(0, 500));
  }

  try {
    const parsed = JSON.parse(text);
    recordApiTokenUsage_('filterList');
    return parsed;
  } catch (e) {
    throw new Error('filterList 응답 JSON 파싱 실패: ' + text.slice(0, 500));
  }
}

function fetchFilterListPageForLotteon_(page) {
  // v5.98:
  // 시간초과 방지를 위해 apiRequest_의 다중 URL/5회 재시도 루틴을 거치지 않고
  // filterList 전용 단일 호출만 수행합니다.
  const payload = {
    page: String(page),
    searchQuery: {
      searchKeyword: '롯백',
      siteId: '',
      filterGroup: 'all',
      sort: 'nameAsc'
    }
  };

  const resp = apiRequestFilterListSingle_(payload);
  const data = extractData_(resp);

  let filters = [];
  if (data && Array.isArray(data.filters)) filters = data.filters;
  else if (data && data.data && Array.isArray(data.data.filters)) filters = data.data.filters;
  else if (Array.isArray(data)) filters = data;
  else if (resp && resp.data && resp.data.filters && Array.isArray(resp.data.filters)) filters = resp.data.filters;
  else if (resp && resp.filters && Array.isArray(resp.filters)) filters = resp.filters;
  else filters = extractItems_(resp, false);

  return {
    items: (filters || []).filter(isObject_),
    totalPage: toNumber_(data && data.totalPage),
    totalCount: toNumber_(data && data.totalCount)
  };
}

function runFilterListResumeBatch_(options) {
  options = options || {};
  const started = Date.now();

  // v5.98:
  // Apps Script 시간초과 방지 최우선.
  // 1회 실행당 filterList 1페이지만 처리하고 무조건 종료합니다.
  const maxPages = 1;
  const startFresh = !!options.startFresh;

  if (startFresh) resetFilterListResumeWork_();

  let state = getFilterListResumeState_();
  if (!state || startFresh) {
    state = {
      page: 1,
      done: false,
      totalRows: 0,
      startedAt: now_(),
      lastUpdatedAt: '',
      lastMessage: ''
    };
  }

  let page = Math.max(1, Number(state.page || 1));
  let processedPages = 0;
  let done = false;
  let lastCount = 0;
  let apiTotalPage = 0;

  writeSyncStatus_({
    phase: 'FILTERLIST_RESUME_RUNNING',
    status: 'filterList 1페이지 안전 실행 중',
    currentFilter: 'filterList page ' + page,
    lastUrl: buildApiUrl_('filterList'),
    tmpRows: state.totalRows || ''
  });

  const pageResult = fetchFilterListPageForLotteon_(page);
  const items = pageResult.items || [];
  lastCount = items.length;
  apiTotalPage = pageResult.totalPage || 0;

  if (!items.length) {
    done = true;
  } else {
    const rows = buildFilterSummaryRowsFromFilterListItems_(items);
    appendFilterListResumeRows_(rows);
    processedPages = 1;

    if (apiTotalPage && page >= apiTotalPage) {
      done = true;
      page++;
    } else {
      page++;
      if (lastCount < 20) done = true;
    }
  }

  const accumulated = dedupeFilterListRows_(readFilterListResumeRows_());

  state.page = page;
  state.done = done;
  state.totalRows = accumulated.length;
  state.lastUpdatedAt = now_();
  state.lastMessage = done ? '완료' : '이어실행 필요';
  setFilterListResumeState_(state);

  if (done) {
    ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
    replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.FILTERS), CONFIG.HEADERS.FILTERS, accumulated);
    clearFilterListResumeState_();

    writeSyncStatus_({
      phase: 'DONE',
      status: 'filterList 1페이지 안전 실행 완료 - 필터별_상품수 최종 반영',
      filterCount: accumulated.length,
      filterIndex: accumulated.length,
      currentFilter: '',
      lastUrl: buildApiUrl_('filterList'),
      tmpRows: accumulated.length,
      lastError: ''
    });

    return {
      done: true,
      filterCount: accumulated.length,
      processedPages: processedPages,
      nextPage: '',
      elapsedSec: Math.round((Date.now() - started) / 1000)
    };
  }

  writeSyncStatus_({
    phase: 'FILTERLIST_RESUME_WAIT',
    status: 'filterList 1페이지 처리 완료 - 같은 메뉴를 다시 실행하세요',
    filterCount: accumulated.length,
    filterIndex: accumulated.length,
    currentFilter: 'next page ' + page,
    lastUrl: buildApiUrl_('filterList'),
    tmpRows: accumulated.length,
    lastError: ''
  });

  return {
    done: false,
    filterCount: accumulated.length,
    processedPages: processedPages,
    nextPage: page,
    elapsedSec: Math.round((Date.now() - started) / 1000)
  };
}


function fetchProductListTotalCountFast_(filterNames) {
  const cred = getCredentials_();
  const url = buildApiUrl_('productList');
  const props = PropertiesService.getScriptProperties();
  const batchSize = Math.max(1, Math.min(toNumber_(CONFIG.FAST_COUNT_BATCH_SIZE) || 25, filterNames.length || 1));
  const result = {};

  for (var i = 0; i < filterNames.length; i += batchSize) {
    const chunk = filterNames.slice(i, i + batchSize);
    const requests = chunk.map(function(filterName) {
      return {
        url: url,
        method: 'post',
        contentType: 'application/json',
        headers: {
          Authorization: 'Bearer ' + cred.apiKey,
          'X-API-SENDER': cred.sender
        },
        payload: JSON.stringify({
          page: '1',
          searchQuery: {
            siteId: CONFIG.SITE_ID,
            searchType: 'filterName',
            searchKeyword: filterName,
            condition: '',
            sort: 'dateDesc'
          }
        }),
        muteHttpExceptions: true
      };
    });

    var responses = [];
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      // v5.60: fetchAll이 실패하면 이 배치는 0+오류로 기록하고 다음 실행을 계속합니다.
      chunk.forEach(function(filterName) {
        result[filterName] = {
          totalCount: 0,
          totalPage: 0,
          error: 'FETCH_ALL_ERROR: ' + String(e && e.message ? e.message : e)
        };
      });
      continue;
    }

    responses.forEach(function(resp, idx) {
      const filterName = chunk[idx];
      try {
        const code = resp.getResponseCode ? resp.getResponseCode() : 200;
        const text = resp.getContentText ? resp.getContentText() : String(resp || '');

        if (code < 200 || code >= 300) {
          result[filterName] = {
            totalCount: 0,
            totalPage: 0,
            error: 'HTTP_' + code
          };
          return;
        }

        const parsed = JSON.parse(text);
        recordApiTokenUsage_('productList');
        const data = extractData_(parsed);
        result[filterName] = {
          totalCount: toNumber_(data.totalCount),
          totalPage: toNumber_(data.totalPage),
          error: ''
        };
      } catch (e) {
        result[filterName] = {
          totalCount: 0,
          totalPage: 0,
          error: String(e && e.message ? e.message : e)
        };
      }
    });
  }

  props.setProperty('LAST_API_ROUTE', 'DIRECT_FETCH_ALL');
  props.setProperty('LAST_API_URL', maskSensitiveUrl_(url));
  props.setProperty('API_TOKEN_LAST_ENDPOINT', 'productList');
  props.setProperty('API_TOKEN_LAST_AT', now_());

  return result;
}

function fetchProductListTotalCountSingleResponse_(filterName) {
  const resp = apiRequest_('productList', 'post', { page: '1', searchQuery: { siteId: CONFIG.SITE_ID, searchType: 'filterName', searchKeyword: filterName, condition: '', sort: 'dateDesc' } });
  return { getResponseCode: function() { return 200; }, getContentText: function() { return JSON.stringify(resp); } };
}

function extractFilterListMeta_(item) {
  const recentCandidates = ['updateDate','modifiedAt','modifyDate','updatedAt','collectionDate','collectDate','collectedAt','crawlDate','scrapDate','lastCollectDate','lastCollectionDate','lastUpdateDate'];
  const createCandidates = ['createDate','createdAt','regDate','registerDate','insertDate','registeredAt','filterCreateDate'];
  const recent = findDateWithField_(item, recentCandidates);
  const created = findDateWithField_(item, createCandidates);
  return { recentDate: recent.value || '', recentField: recent.field || '', createDate: created.value || '', createField: created.field || '' };
}

function findDateWithField_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    const raw = getAny_(obj, [keys[i]]);
    const value = normalizeDateText_(raw);
    if (value) return { value: value, field: keys[i] };
  }
  for (const k in obj) {
    const lk = String(k || '').toLowerCase();
    if ((lk.indexOf('date') >= 0 || lk.indexOf('time') >= 0) &&
        (lk.indexOf('update') >= 0 || lk.indexOf('modify') >= 0 || lk.indexOf('collect') >= 0 || lk.indexOf('create') >= 0 || lk.indexOf('reg') >= 0)) {
      const value = normalizeDateText_(obj[k]);
      if (value) return { value: value, field: k };
    }
  }
  return { value: '', field: '' };
}



function auditReadTableFastByName_(sheetName, maxCols, maxRows) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = Math.min(Math.max(sheet.getLastRow(), 1), maxRows || 5000);
  if (lastRow < 2) return [];
  const lastCol = Math.min(Math.max(sheet.getLastColumn(), 1), maxCols || 80);
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(function(h) { return normalizeHeaderKey_(h); });
  const rows = [];
  for (var r = 1; r < values.length; r++) {
    const obj = {};
    let nonEmpty = false;
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = values[r][c];
      if (String(values[r][c] || '').trim() !== '') nonEmpty = true;
    }
    if (nonEmpty) rows.push(obj);
  }
  return rows;
}

function auditFastRows_(sheetName) {
  const h = CONFIG.HEADERS;
  const s = CONFIG.SHEETS;
  if (sheetName === s.SALES_CLEAN) return auditReadTableFastByName_(sheetName, h.SALES_CLEAN.length, 5000);
  if (sheetName === s.BRAND_SUMMARY) return auditReadTableFastByName_(sheetName, h.BRAND_SUMMARY.length, 1000);
  if (sheetName === s.BRAND) return auditReadTableFastByName_(sheetName, h.BRAND.length, 1000);
  if (sheetName === s.RETRANSMIT_LOG) return auditReadTableFastByName_(sheetName, h.RETRANSMIT_LOG.length, 1000);
  if (sheetName === s.FILTERS) return auditReadTableFastByName_(sheetName, h.FILTERS.length, 1000);
  if (sheetName === s.COUPANG_SENT_MANUAL) return auditReadTableFastByName_(sheetName, h.COUPANG_SENT_MANUAL.length, 1000);
  if (sheetName === s.MANUAL_API_CHECK) return auditReadTableFastByName_(sheetName, h.MANUAL_API_CHECK.length, 300);
  if (sheetName === s.MATCH_DIAG) return auditReadTableFastByName_(sheetName, h.MATCH_DIAG.length, 200);
  return auditReadTableFastByName_(sheetName, 80, 5000);
}

function auditFastBrandFromRaw_(row) {
  return String(
    getObjectValueByHeader_(row, '브랜드') ||
    getObjectValueByHeader_(row, '브랜드명') ||
    getObjectValueByHeader_(row, '상표') ||
    ''
  ).trim();
}

function auditFastRawStats_(rows) {
  const map = {};
  (rows || []).forEach(function(r) {
    if (!auditIsLotteonSalesRow_(r)) return;
    const brand = auditFastBrandFromRaw_(r);
    if (!brand || isGenericInvalidBrand_(brand)) return;
    const amount = getObjectValueByHeader_(r, '결제금액합계(원)') || getObjectValueByHeader_(r, '결제금액합계') || getObjectValueByHeader_(r, '결제금액') || 0;
    const orderKey = getObjectValueByHeader_(r, '마켓주문번호') || getObjectValueByHeader_(r, '더망고주문고유번호') || getObjectValueByHeader_(r, '주문번호');
    const productKey = getObjectValueByHeader_(r, '사이트상품번호') || getObjectValueByHeader_(r, '마켓상품번호') || getObjectValueByHeader_(r, '상품URL') || getObjectValueByHeader_(r, '상품명');
    auditAddStat_(map, brand, '', amount, orderKey, productKey);
  });
  return map;
}

function auditFastCleanStats_(rows) {
  const map = {};
  (rows || []).forEach(function(r) {
    const target = String(getObjectValueByHeader_(r, '분석대상여부') || '').trim();
    if (target && target !== 'Y' && target !== 'TRUE' && target !== 'true' && target !== '1') return;
    const brand = getObjectValueByHeader_(r, '브랜드명_매칭') || getObjectValueByHeader_(r, '브랜드명_원본');
    if (!brand || isGenericInvalidBrand_(brand)) return;
    const amount = getObjectValueByHeader_(r, '결제금액합계') || getObjectValueByHeader_(r, '결제금액합계(원)') || 0;
    const orderKey = getObjectValueByHeader_(r, '마켓주문번호') || getObjectValueByHeader_(r, '원본행');
    const productKey = getObjectValueByHeader_(r, '사이트상품번호') || getObjectValueByHeader_(r, '마켓상품번호') || getObjectValueByHeader_(r, '상품URL') || getObjectValueByHeader_(r, '원문상품명');
    auditAddStat_(map, brand, '', amount, orderKey, productKey);
  });
  return map;
}

function auditFastLatestDate_(rows, headerNames) {
  let latest = '';
  (rows || []).forEach(function(r) {
    for (var i = 0; i < headerNames.length; i++) {
      const raw = getObjectValueByHeader_(r, headerNames[i]);
      if (raw === '' || raw === null || raw === undefined) continue;
      let d = raw;
      if (Object.prototype.toString.call(raw) === '[object Date]') {
        d = Utilities.formatDate(raw, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      } else if (typeof raw === 'number') {
        // Google Sheets serial date to JS date
        const js = new Date(Math.round((raw - 25569) * 86400 * 1000));
        d = Utilities.formatDate(js, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      } else {
        d = normalizeDateText_(raw);
      }
      if (d && (!latest || String(d) > String(latest))) latest = String(d);
      break;
    }
  });
  return latest ? formatShortDate_(latest) : '';
}

function formatAuditReportSheetUltraLight_(sheet) {
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = CONFIG.HEADERS.AUDIT_REPORT.length;
  const dataRows = Math.max(lastRow - 1, 0);
  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 42);
  } catch (e) {}
  if (dataRows > 0) {
    try {
      sheet.getRange(2, 1, dataRows, lastCol).setWrap(false).setVerticalAlignment('middle');
      sheet.getRange(2, 6, dataRows, 3).setNumberFormat('#,##0');
    } catch (e) {}
  }
  const widths = [90, 175, 75, 125, 190, 105, 105, 95, 210, 300];
  widths.forEach(function(w, i) { try { sheet.setColumnWidth(i + 1, w); } catch (e) {} });
}

function auditMetricValue_(rows, itemName) {
  const target = String(itemName || '').trim();
  for (var i = 0; i < (rows || []).length; i++) {
    const item = String(getObjectValueByHeader_(rows[i], '항목') || '').trim();
    if (item === target) return getObjectValueByHeader_(rows[i], '값');
  }
  return '';
}

function auditMetricRatio_(rows, itemName) {
  const target = String(itemName || '').trim();
  for (var i = 0; i < (rows || []).length; i++) {
    const item = String(getObjectValueByHeader_(rows[i], '항목') || '').trim();
    if (item === target) return getObjectValueByHeader_(rows[i], '비중');
  }
  return '';
}

function auditSheetRowCount_(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return 0;
  return Math.max(0, sheet.getLastRow() - 1);
}

function replaceAuditReportQuick_(sheet, rows) {
  const headers = CONFIG.HEADERS.AUDIT_REPORT;
  const data = [headers].concat(rows || []);
  const cols = headers.length;
  const clearRows = Math.max(250, data.length + 20);

  // v5.69: 기존 상세행 잔여물을 제거하되, 시트 전체 clear/getLastRow 대량 처리를 피합니다.
  try { sheet.getRange(1, 1, clearRows, cols).clearContent(); } catch (e) {}
  sheet.getRange(1, 1, data.length, cols).setValues(data);
  try { sheet.setFrozenRows(1); } catch (e) {}
}

function formatAuditReportSheetInstant_(sheet) {
  if (!sheet) return;
  const lastCol = CONFIG.HEADERS.AUDIT_REPORT.length;
  try {
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 42);
  } catch (e) {}
  const widths = [85, 190, 70, 130, 190, 115, 115, 95, 210, 320];
  widths.forEach(function(w, i) { try { sheet.setColumnWidth(i + 1, w); } catch (e) {} });
}

function generateAuditReport() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheetWithHeader_(ss, CONFIG.SHEETS.AUDIT_REPORT, CONFIG.HEADERS.AUDIT_REPORT);
  const generatedAt = now_();

  // v5.69:
  // 검수리포트는 더 이상 매출데이터_붙여넣기/매출데이터_정리 전체를 다시 스캔하지 않습니다.
  // ⑧이 이미 만든 매칭진단 + 핵심_브랜드요약 + 쿠팡재전송_로그 + 수동입력_API검증 판정요약만 사용합니다.
  const matchDiagRows = auditFastRows_(CONFIG.SHEETS.MATCH_DIAG);
  const summaryRows = auditFastRows_(CONFIG.SHEETS.BRAND_SUMMARY);
  const logRows = auditFastRows_(CONFIG.SHEETS.RETRANSMIT_LOG);
  const filterRows = auditFastRows_(CONFIG.SHEETS.FILTERS);
  const manualApiRows = auditFastRows_(CONFIG.SHEETS.MANUAL_API_CHECK);

  const summaryStats = auditBuildSummaryStats_(summaryRows);
  const logStats = auditBuildRetransmitStats_(logRows);
  const summaryTotal = auditSumStatsAmount_(summaryStats);
  const logTotal = auditSumStatsAmount_(logStats);

  const sourceRows = toNumber_(auditMetricValue_(matchDiagRows, '원본 매출행수'));
  const lotteonRows = toNumber_(auditMetricValue_(matchDiagRows, 'LOTTEON 분석대상 행수'));
  const analyzedRows = toNumber_(auditMetricValue_(matchDiagRows, '분석 반영 행수')) || toNumber_(auditMetricValue_(matchDiagRows, '전체매출 행수'));
  const cancelRows = toNumber_(auditMetricValue_(matchDiagRows, '더망고 취소/반품 주문행수'));
  const cancelRevenue = toNumber_(auditMetricValue_(matchDiagRows, '더망고 취소/반품 매출액'));
  const unmatchedRows = toNumber_(auditMetricValue_(matchDiagRows, '미매칭 행수'));
  const unmatchedRevenue = toNumber_(auditMetricValue_(matchDiagRows, '미매칭 매출액'));
  const analyzedRevenue = toNumber_(auditMetricValue_(matchDiagRows, '분석 반영 매출액'));

  const rows = [];
  const push = function(section, item, result, brand, filterName, v1, v2, diff, action, memo) {
    rows.push([section || '', item || '', result || '', brand || '', filterName || '', v1, v2, diff, action || '', memo || '']);
  };

  push('요약', '검수 생성시각', 'INFO', '', '', generatedAt, '', '', '검수리포트 확인', 'Asia/Seoul');
  push('요약', '실행모드', 'INFO', '', '', '즉시요약', '', '', '정상', 'v5.98: 매출 원본/정리 전체 스캔 없음');
  push('요약', '매출데이터_붙여넣기 원본행', sourceRows ? 'INFO' : '주의', '', '', sourceRows || auditSheetRowCount_(CONFIG.SHEETS.SALES_IN), '', '', sourceRows ? '정상' : '⑧ 분석/매출 원본 확인', '매칭진단 기준');
  push('요약', 'LOTTEON 분석대상 행수', lotteonRows ? 'INFO' : '주의', '', '', lotteonRows, '', '', lotteonRows ? '정상' : '구매사이트명/⑧ 분석 확인', '매칭진단 기준');
  push('요약', '분석 반영 행수', analyzedRows ? 'INFO' : '주의', '', '', analyzedRows, '', '', analyzedRows ? '정상' : '⑧ 분석 실행 확인', '취소/반품 포함 전체 주문반응 기준');
  push('요약', '분석 반영 매출액', analyzedRevenue ? 'INFO' : '주의', '', '', formatWon_(analyzedRevenue), '', '', analyzedRevenue ? '정상' : '⑧ 분석 실행 확인', '매칭진단 기준');
  push('요약', '핵심요약 전체매출 합계', 'INFO', '', '', formatWon_(summaryTotal), '', '', '참고', '핵심_브랜드요약 기준');
  push('요약', '재전송로그 누적매출 합계', 'INFO', '', '', formatWon_(logTotal), '', '', '참고', '쿠팡재전송_로그 기준');
  push('취소정보', '더망고 취소/반품 주문행수', 'INFO', '', '', cancelRows, '', '', '참고', '더망고주문상태 기준, 주 분석 매출에는 포함');
  push('취소정보', '더망고 취소/반품 매출액', 'INFO', '', '', formatWon_(cancelRevenue), '', '', '참고', '정보성 지표');

  let severeCount = 0;
  let warnCount = 0;

  const totalDiff = analyzedRevenue - summaryTotal;
  if (analyzedRevenue && Math.abs(totalDiff) >= 50000) {
    warnCount++;
    push('매출차이', '분석반영 vs 핵심요약 차이', '주의', '', '', formatWon_(analyzedRevenue), formatWon_(summaryTotal), formatWon_(totalDiff), '⑧ 매출분석 재실행', '매칭진단과 핵심요약 합계 차이');
  }

  if (unmatchedRows > 0 || unmatchedRevenue > 0) {
    severeCount++;
    push('매칭진단', '미매칭 존재', '긴급', '', '', unmatchedRows, formatWon_(unmatchedRevenue), '', '브랜드명_매칭표/매출정리 확인', '미매칭은 핵심요약 누락 가능성 있음');
  } else {
    push('매칭진단', '미매칭', '정상', '', '', 0, formatWon_(0), '', '조치 없음', '미매칭 행수/매출액 0');
  }

  // 필터 0개는 상세행이 아니라 개수만 요약
  let zeroApi = 0;
  (filterRows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const count = toNumber_(getObjectValueByHeader_(r, 'API_totalCount'));
    if (filterName && count === 0) zeroApi++;
  });
  if (zeroApi) {
    push('필터관리', 'API 상품수 0 필터 수', '참고', '', '', zeroApi, '', '', '필터별_상품수 시트 확인', '의도한 0개 필터인지 확인');
  }

  // 수동/API검증도 판정별 카운트만 출력
  if (manualApiRows && manualApiRows.length) {
    const counts = {};
    manualApiRows.forEach(function(r) {
      const result = String(getObjectValueByHeader_(r, '판정') || '').trim() || '미분류';
      counts[result] = (counts[result] || 0) + 1;
    });
    Object.keys(counts).sort().forEach(function(k) {
      const level = /API 미확인|수동입력 없음|차이/.test(k) ? '주의' : '참고';
      if (level === '주의') warnCount += counts[k] ? 1 : 0;
      push('수동/API검증', k, level, '', '', counts[k], '', '', '수동입력_API검증 시트 상세 확인', '판정별 문제행 수');
    });
  } else {
    push('수동/API검증', '검증시트 행 없음', '참고', '', '', 0, '', '', '필요 시 수동입력/API 검증 시트 갱신', '문제행만 출력 방식이면 정상일 수 있음');
  }

  auditAppendMatchDiagLite_(push, matchDiagRows);

  push('요약', '긴급 이슈 수', severeCount ? '긴급' : '정상', '', '', severeCount, '', '', severeCount ? '미매칭부터 처리' : '긴급 이슈 없음', '매출 누락 가능성 중심');
  push('요약', '주의 이슈 수', warnCount ? '주의' : '정상', '', '', warnCount, '', '', warnCount ? '수동/API검증 요약 확인' : '주의 이슈 없음', '상세 문제행은 수동입력_API검증 시트에서 확인');

  replaceAuditReportQuick_(sheet, rows);
  formatAuditReportSheetInstant_(sheet);
  log_('audit_report_instant', '검수리포트 즉시요약 생성 완료 / rows=' + rows.length + ', severe=' + severeCount + ', warn=' + warnCount);

  SpreadsheetApp.getUi().alert(
    '검수리포트 생성 완료\n\n' +
    '실행모드: 즉시요약\n' +
    '긴급 이슈: ' + severeCount + '개\n' +
    '주의 요약: ' + warnCount + '개\n\n' +
    '매출 원본/정리 전체 스캔과 상세 문제행 출력은 하지 않습니다.'
  );
}

function auditAppendFilterIssuesLite_(push, filterRows, manualRows) {
  let active02 = 0, api02 = 0, noBrandManual = 0, noBrandApi = 0, zeroApi = 0;
  const statusInactive = /삭제|변경|비활성|중지/i;
  (manualRows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || brandFromFilterName_(filterName) || '').trim();
    const status = String(getObjectValueByHeader_(r, '확인일') || '').trim();
    if (!brand && filterName) noBrandManual++;
    if (!statusInactive.test(status) && /_02_/.test(filterName)) active02++;
  });
  (filterRows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || brandFromFilterName_(filterName) || '').trim();
    const count = toNumber_(getObjectValueByHeader_(r, 'API_totalCount'));
    if (!brand && filterName) noBrandApi++;
    if (/_02_/.test(filterName)) api02++;
    if (count === 0 && filterName) zeroApi++;
  });
  if (active02) push('필터관리', '활성 02 필터 잔존', '주의', '', '', active02, '', '', '04로 변경/삭제 상태 처리', '수동입력 기준');
  if (api02) push('필터관리', 'API 02 필터 존재', '주의', '', '', api02, '', '', '더망고 필터명 변경 여부 확인', '필터별_상품수 기준');
  if (noBrandManual) push('필터관리', '수동입력 브랜드명 빈칸', '주의', '', '', noBrandManual, '', '', '브랜드명 채우기', '쿠팡전송수_수동입력');
  if (noBrandApi) push('필터관리', 'API 브랜드 없는 필터명', '주의', '', '', noBrandApi, '', '', '검색필터명 수정/삭제', '필터별_상품수 기준');
  if (zeroApi) push('필터관리', 'API 상품수 0', '참고', '', '', zeroApi, '', '', '필터 조건/수집 상태 확인', '0개 필터 수');
}

function auditAppendManualApiIssuesLite_(push, manualApiRows) {
  if (!manualApiRows || !manualApiRows.length) {
    push('수동/API검증', '검증시트 행 없음', '참고', '', '', 0, '', '', '필요 시 수동입력/API 검증 시트 갱신 실행', '문제행만 출력 방식이면 정상일 수 있음');
    return;
  }
  const counts = {};
  manualApiRows.forEach(function(r) {
    const result = String(getObjectValueByHeader_(r, '판정') || '').trim() || '미분류';
    counts[result] = (counts[result] || 0) + 1;
  });
  Object.keys(counts).sort().forEach(function(k) {
    const level = /API 미확인|수동입력 없음|차이/.test(k) ? '주의' : '참고';
    push('수동/API검증', k, level, '', '', counts[k], '', '', '수동입력_API검증 시트 상세 확인', '판정별 문제행 수');
  });
}

function auditAppendMatchDiagLite_(push, matchDiagRows) {
  if (!matchDiagRows || !matchDiagRows.length) return;
  (matchDiagRows || []).forEach(function(r) {
    const item = String(getObjectValueByHeader_(r, '항목') || '').trim();
    if (!item) return;
    if (/미매칭|전체매출|분석 반영|LOTTEON/.test(item)) {
      push('매칭진단', item, /미매칭/.test(item) && toNumber_(getObjectValueByHeader_(r, '값')) > 0 ? '주의' : 'INFO', '', '', getObjectValueByHeader_(r, '값'), getObjectValueByHeader_(r, '비중'), '', /미매칭/.test(item) ? '브랜드명_매칭표/매출정리 확인' : '참고', getObjectValueByHeader_(r, '설명'));
    }
  });
}



function auditReadTableByName_(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return [];
  return readTable_(sheet);
}

function auditEmptyStat_() { return { brand: '', filterName: '', amount: 0, orders: 0, products: 0 }; }

function auditAddStat_(map, brand, filterName, amount, orderKey, productKey) {
  const b = String(brand || '').trim();
  if (!b) return;
  const key = normalizeBrandKey_(b);
  if (!key) return;
  if (!map[key]) map[key] = { brand: b, filterName: filterName || '', amount: 0, orders: 0, products: 0, orderSet: {}, productSet: {} };
  map[key].amount += toNumber_(amount);
  if (filterName && !map[key].filterName) map[key].filterName = filterName;
  const ok = String(orderKey || '').trim();
  if (ok && !map[key].orderSet[ok]) { map[key].orderSet[ok] = true; map[key].orders += 1; }
  const pk = String(productKey || '').trim();
  if (pk && !map[key].productSet[pk]) { map[key].productSet[pk] = true; map[key].products += 1; }
}

function auditBuildBrandCandidates_(filterRows, manualRows, summaryRows, brandRows, logRows, cleanRows) {
  const out = [];
  function add_(v) { const s = String(v || '').trim(); if (s && out.indexOf(s) < 0 && !/^롯백_\d{2}_?$/.test(s)) out.push(s); }
  [filterRows, manualRows, summaryRows, brandRows, logRows, cleanRows].forEach(function(rows) {
    (rows || []).forEach(function(r) {
      add_(getObjectValueByHeader_(r, '브랜드명'));
      add_(getObjectValueByHeader_(r, '브랜드명_매칭'));
      add_(getObjectValueByHeader_(r, '브랜드명_원본'));
      add_(brandFromFilterName_(getObjectValueByHeader_(r, '검색필터명')));
      add_(brandFromFilterName_(getObjectValueByHeader_(r, '대표검색필터명')));
      add_(brandFromFilterName_(getObjectValueByHeader_(r, '대표\n검색필터명')));
    });
  });
  out.sort(function(a, b) { return String(b).length - String(a).length; });
  return out;
}

function auditGuessBrandFromTexts_(row, candidates) {
  const direct = String(getObjectValueByHeader_(row, '브랜드') || getObjectValueByHeader_(row, '브랜드명') || getObjectValueByHeader_(row, '브랜드명_매칭') || '').trim();
  if (direct) return direct;
  const text = [
    getObjectValueByHeader_(row, '마켓상품명'),
    getObjectValueByHeader_(row, '상품명(원문)'),
    getObjectValueByHeader_(row, '원문상품명'),
    getObjectValueByHeader_(row, '상품명')
  ].join(' ');
  const normText = normalizeProductNameKey_(text);
  for (var i = 0; i < (candidates || []).length; i++) {
    const b = candidates[i];
    if (!b) continue;
    if (normText.indexOf(normalizeProductNameKey_(b)) >= 0) return b;
  }
  return '';
}

function auditIsLotteonSalesRow_(row) {
  const site = String(getObjectValueByHeader_(row, '구매사이트명') || '').trim();
  return CONFIG.LOTTEON_SITE_REGEX.test(site);
}

function auditIsValidSalesRow_(row) {
  // v5.65: 검수 리포트의 원본/정리 매출 비교도 취소/반품 포함 전체 매출 기준입니다.
  // 취소 여부는 더망고주문상태 기준 정보성 지표로만 별도 확인합니다.
  return true;
}

function auditIsCancelledByMangoStatus_(row) {
  return isCancelledByMangoStatus_(getObjectValueByHeader_(row, '더망고주문상태'));
}

function auditBuildRawSalesStats_(rows, candidates) {
  const map = {};
  (rows || []).forEach(function(r) {
    if (!auditIsLotteonSalesRow_(r)) return;
    if (!auditIsValidSalesRow_(r)) return;
    const brand = auditGuessBrandFromTexts_(r, candidates);
    const amount = getObjectValueByHeader_(r, '결제금액합계(원)') || getObjectValueByHeader_(r, '결제금액합계') || 0;
    const orderKey = getObjectValueByHeader_(r, '마켓주문번호') || getObjectValueByHeader_(r, '더망고주문고유번호');
    const productKey = getObjectValueByHeader_(r, '사이트상품번호') || getObjectValueByHeader_(r, '마켓상품번호') || getObjectValueByHeader_(r, '상품URL');
    auditAddStat_(map, brand, '', amount, orderKey, productKey);
  });
  return map;
}

function auditBuildCleanSalesStats_(rows) {
  const map = {};
  (rows || []).forEach(function(r) {
    const target = String(getObjectValueByHeader_(r, '분석대상여부') || '').trim();
    const valid = String(getObjectValueByHeader_(r, '유효여부') || '').trim();
    if (target && target !== 'Y' && target !== 'TRUE' && target !== 'true' && target !== '1') return;
    if (valid && valid !== 'Y' && valid !== 'TRUE' && valid !== 'true' && valid !== '1') return;
    const brand = getObjectValueByHeader_(r, '브랜드명_매칭') || getObjectValueByHeader_(r, '브랜드명_원본');
    const amount = getObjectValueByHeader_(r, '결제금액합계') || getObjectValueByHeader_(r, '결제금액합계(원)') || 0;
    const orderKey = getObjectValueByHeader_(r, '마켓주문번호') || getObjectValueByHeader_(r, '원본행');
    const productKey = getObjectValueByHeader_(r, '사이트상품번호') || getObjectValueByHeader_(r, '마켓상품번호') || getObjectValueByHeader_(r, '상품URL');
    auditAddStat_(map, brand, '', amount, orderKey, productKey);
  });
  return map;
}

function auditBuildSummaryStats_(rows) {
  const map = {};
  (rows || []).forEach(function(r) {
    const brand = getObjectValueByHeader_(r, '브랜드명');
    const filterName = getObjectValueByHeader_(r, '대표검색필터명') || getObjectValueByHeader_(r, '대표\n검색필터명');
    const amount =
      getObjectValueByHeader_(r, '전체\n매출') ||
      getObjectValueByHeader_(r, '전체매출') ||
      getObjectValueByHeader_(r, '유효\n매출') ||
      getObjectValueByHeader_(r, '유효매출') ||
      getObjectValueByHeader_(r, '매출액') ||
      getObjectValueByHeader_(r, '매출') ||
      0;
    const orders = getObjectValueByHeader_(r, '주문건수') || getObjectValueByHeader_(r, '주문\n건수') || 0;
    const products = getObjectValueByHeader_(r, '매출상품수') || getObjectValueByHeader_(r, '매출\n상품수') || 0;
    const key = normalizeBrandKey_(brand);
    if (!key) return;
    map[key] = { brand: brand, filterName: filterName, amount: toNumber_(amount), orders: toNumber_(orders), products: toNumber_(products) };
  });
  return map;
}

function auditBuildBrandDetailStats_(rows) {
  const map = {};
  (rows || []).forEach(function(r) {
    const brand = getObjectValueByHeader_(r, '브랜드명');
    const filterName = getObjectValueByHeader_(r, '대표검색필터명');
    const amount = getObjectValueByHeader_(r, '전체_매출액') || 0;
    const orders = getObjectValueByHeader_(r, '전체_주문건수') || 0;
    const products = getObjectValueByHeader_(r, '전체_매출상품수') || 0;
    const key = normalizeBrandKey_(brand);
    if (!key) return;
    map[key] = { brand: brand, filterName: filterName, amount: toNumber_(amount), orders: toNumber_(orders), products: toNumber_(products) };
  });
  return map;
}

function auditBuildRetransmitStats_(rows) {
  const map = {};
  (rows || []).forEach(function(r) {
    const brand = getObjectValueByHeader_(r, '브랜드명');
    const filterName = getObjectValueByHeader_(r, '검색필터명');
    const amount = getObjectValueByHeader_(r, '누적매출액') || getObjectValueByHeader_(r, '누적\n매출액') || 0;
    const key = normalizeBrandKey_(brand);
    if (!key) return;
    map[key] = { brand: brand, filterName: filterName, amount: toNumber_(amount), orders: 0, products: 0 };
  });
  return map;
}

function auditUnionKeys_() {
  const seen = {}, out = [];
  for (var a = 0; a < arguments.length; a++) {
    const map = arguments[a] || {};
    Object.keys(map).forEach(function(k) { if (k && !seen[k]) { seen[k] = true; out.push(k); } });
  }
  return out.sort();
}

function auditSumStatsAmount_(map) {
  let sum = 0;
  Object.keys(map || {}).forEach(function(k) { sum += toNumber_(map[k].amount); });
  return sum;
}

function auditLatestOrderDate_(rows) {
  let latest = '';
  (rows || []).forEach(function(r) {
    const d = normalizeDateText_(getObjectValueByHeader_(r, 'ㄹ') || getObjectValueByHeader_(r, '주문일시') || getObjectValueByHeader_(r, '일시'));
    if (d && (!latest || d > latest)) latest = d;
  });
  return latest ? formatShortDate_(latest) : '';
}

function auditLatestCleanOrderDate_(rows) {
  let latest = '';
  (rows || []).forEach(function(r) {
    const d = normalizeDateText_(getObjectValueByHeader_(r, '주문일시'));
    if (d && (!latest || d > latest)) latest = d;
  });
  return latest ? formatShortDate_(latest) : '';
}

function auditAppendFilterIssues_(push, filterRows, manualRows) {
  const activeManualByBrand = {};
  const statusInactive = /삭제|변경|비활성|중지/i;
  (manualRows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || brandFromFilterName_(filterName) || '').trim();
    const status = String(getObjectValueByHeader_(r, '확인일') || '').trim();
    if (!filterName && !brand) return;
    if (!brand) push('필터관리', '브랜드명 빈칸', '주의', brand, filterName, '', '', '', '수동입력 브랜드명 채우기', '쿠팡전송수_수동입력');
    if (/^롯백_\d{2}_?$/.test(filterName)) push('필터관리', '브랜드 없는 필터명', '주의', brand, filterName, '', '', '', '필터명/브랜드명 수정', '예: 롯백_04_ 처럼 브랜드가 없는 필터');
    if (!statusInactive.test(status)) {
      if (/_02_/.test(filterName)) push('필터관리', '활성 02 필터 잔존', '주의', brand, filterName, status, '', '', '04로 변경/삭제 상태 처리', '현재 02 필터를 쓰지 않는 운영 기준');
      const key = normalizeBrandKey_(brand);
      if (key) {
        if (!activeManualByBrand[key]) activeManualByBrand[key] = [];
        if (filterName) activeManualByBrand[key].push(filterName);
      }
    }
  });
  Object.keys(activeManualByBrand).forEach(function(k) {
    const arr = unique_(activeManualByBrand[k]);
    if (arr.length > 1) push('필터관리', '동일 브랜드 활성필터 중복', '주의', '', arr.join(', '), arr.length, '', '', '대표 운영 필터 1개만 남기는지 확인', '브랜드 기준 중복 필터');
  });
  (filterRows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || brandFromFilterName_(filterName) || '').trim();
    const count = toNumber_(getObjectValueByHeader_(r, 'API_totalCount'));
    if (/_02_/.test(filterName)) push('필터관리', 'API 02 필터 존재', '주의', brand, filterName, count, '', '', '더망고 필터명 변경 여부 확인', '필터별_상품수 기준');
    if (/^롯백_\d{2}_?$/.test(filterName)) push('필터관리', 'API 브랜드 없는 필터명', '주의', brand, filterName, count, '', '', '더망고 검색필터명 수정/삭제', '필터별_상품수 기준');
    if (count === 0 && filterName) push('필터관리', 'API 상품수 0', '참고', brand, filterName, count, '', '', '필터 조건/수집 상태 확인', '0개 필터가 의도한 상태인지 확인');
  });
}

function auditAppendManualApiIssues_(push, manualApiRows) {
  if (!manualApiRows || !manualApiRows.length) {
    push('수동/API검증', '검증시트 행 없음', '참고', '', '', 0, '', '', '필요 시 수동입력/API 검증 시트 갱신 실행', '문제행만 출력 방식이면 정상일 수 있음');
    return;
  }
  const counts = {};
  manualApiRows.forEach(function(r) {
    const result = String(getObjectValueByHeader_(r, '판정') || '').trim() || '미분류';
    counts[result] = (counts[result] || 0) + 1;
  });
  Object.keys(counts).sort().forEach(function(k) {
    const level = /API 미확인|수동입력 없음|차이/.test(k) ? '주의' : '참고';
    push('수동/API검증', k, level, '', '', counts[k], '', '', '수동입력_API검증 시트 상세 확인', '판정별 문제행 수');
  });
  manualApiRows.slice(0, 30).forEach(function(r) {
    const result = String(getObjectValueByHeader_(r, '판정') || '').trim();
    if (!result || result === '일치') return;
    push('수동/API검증', '상세 문제행', /API 미확인|수동입력 없음|차이/.test(result) ? '주의' : '참고', getObjectValueByHeader_(r, '브랜드명'), getObjectValueByHeader_(r, '검색필터명'), getObjectValueByHeader_(r, '수동_더망고수집수'), getObjectValueByHeader_(r, 'API_더망고수집수'), getObjectValueByHeader_(r, '수집수차이'), '수동입력 또는 필터명 확인', result + ' / ' + String(getObjectValueByHeader_(r, '메모') || '').slice(0, 160));
  });
}

function auditAppendMatchDiag_(push, matchDiagRows) {
  if (!matchDiagRows || !matchDiagRows.length) return;
  (matchDiagRows || []).forEach(function(r) {
    const item = String(getObjectValueByHeader_(r, '항목') || '').trim();
    if (!item) return;
    if (/미매칭|취소|매출액|LOTTEON|분석 반영/.test(item)) {
      push('매칭진단', item, /미매칭/.test(item) && toNumber_(getObjectValueByHeader_(r, '값')) > 0 ? '주의' : 'INFO', '', '', getObjectValueByHeader_(r, '값'), getObjectValueByHeader_(r, '비중'), '', /미매칭/.test(item) ? '브랜드명_매칭표/매출정리 확인' : '참고', getObjectValueByHeader_(r, '설명'));
    }
  });
}

function formatAuditReportSheet_(sheet) {
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = CONFIG.HEADERS.AUDIT_REPORT.length;
  const dataRows = Math.max(lastRow - 1, 0);
  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 42);
  } catch (e) {}
  if (dataRows > 0) {
    try {
      sheet.getRange(2, 1, dataRows, lastCol).setWrap(false).setVerticalAlignment('middle');
      sheet.getRange(2, 6, dataRows, 3).setNumberFormat('#,##0');
    } catch (e) {}
  }
  const widths = [95, 190, 80, 130, 220, 120, 120, 110, 260, 420];
  widths.forEach(function(w, i) { try { sheet.setColumnWidth(i + 1, w); } catch (e) {} });
}

function runSelfDiagnostic() {
  const ui = SpreadsheetApp.getUi();
  const results = [];
  const startedAt = now_();
  setupLotteonSheets(false);
  diagStep_(results, '시트 구조 정리', function() { return '필수 시트 생성/헤더 갱신 완료'; });
  diagStep_(results, 'API 인증값 저장 여부', function() {
    const props = PropertiesService.getScriptProperties();
    const baseUrl = props.getProperty('THE_MANGO_BASE_URL') || CONFIG.BASE_URL_DEFAULT;
    const apiKey = String(props.getProperty('THE_MANGO_API_KEY') || '').trim();
    const sender = String(props.getProperty('THE_MANGO_SENDER') || '').trim();
    if (!apiKey || !sender) throw new Error('API Key 또는 X-API-SENDER가 저장되어 있지 않습니다. ② API 인증값 저장 또는 ★ 원클릭 준비를 먼저 실행하세요.');
    return '저장됨 / URL=' + normalizeBaseUrl_(baseUrl) + ' / API Key, Sender 값은 미표시';
  });
  diagStep_(results, '필수 시트/헤더 점검', function() {
    const check = checkRequiredSheetsAndHeaders_();
    if (check.errors.length) throw new Error(check.errors.join(' | '));
    return 'OK / 점검 시트 ' + check.checked + '개';
  });
  let sampleFilterName = '';
  diagStep_(results, 'filterList API 실호출', function() {
    const resp = apiRequest_('filterList', 'post', { page: '1', searchQuery: { searchKeyword: CONFIG.FILTER_PREFIX, siteId: '', filterGroup: 'all', sort: 'nameAsc' } });
    const data = extractData_(resp);
    const items = extractItems_(resp, false);
    sampleFilterName = findFirstLotteonFilterName_(items);
    const msg = '성공 / 조회행수=' + items.length + ' / totalPage=' + (data.totalPage || '');
    return sampleFilterName ? (msg + ' / 샘플필터=' + sampleFilterName) : (msg + ' / 샘플필터 없음: productList 샘플은 건너뜀');
  });
  diagStep_(results, 'productList API 샘플 실호출', function() {
    if (!sampleFilterName) return 'SKIP / filterList 1페이지에서 롯백 샘플필터를 찾지 못함';
    const resp = apiRequest_('productList', 'post', { page: '1', searchQuery: { siteId: CONFIG.SITE_ID, searchType: 'filterName', searchKeyword: sampleFilterName, condition: '', sort: 'dateDesc' } });
    const data = extractData_(resp);
    const items = extractItems_(resp, true);
    return '성공 / 샘플필터=' + sampleFilterName + ' / 조회행수=' + items.length + ' / totalCount=' + (data.totalCount || '') + ' / totalPage=' + (data.totalPage || '');
  });
  diagStep_(results, '트리거 생성/삭제 테스트', function() {
    const testHandler = 'selfDiagnosticDummy_';
    deleteTriggersForHandler_(testHandler);
    const trigger = ScriptApp.newTrigger(testHandler).timeBased().after(60 * 1000).create();
    const triggerId = trigger.getUniqueId ? trigger.getUniqueId() : 'unknown';
    const triggers = ScriptApp.getProjectTriggers();
    const created = triggers.some(function(t) { return t.getHandlerFunction() === testHandler; });
    if (!created) throw new Error('트리거 생성 실패');
    const deleted = safeDeleteTrigger_(trigger);
    if (deleted) return 'PASS / 테스트 트리거 생성 후 즉시 삭제 성공';
    return 'PARTIAL / 생성됨(ID=' + triggerId + '), 삭제 실패 → Apps Script 왼쪽 "트리거" 메뉴에서 수동 삭제 필요';
  });
  diagStep_(results, '토큰 추정값 점검', function() {
    const t = getApiTokenUsageSnapshot_();
    return '오늘=' + formatCount_(t.todayUsed) + '/' + formatCount_(CONFIG.API_TOKEN_DAILY_LIMIT) + ', 이번달=' + formatCount_(t.monthUsed) + '/' + formatCount_(CONFIG.API_TOKEN_MONTHLY_LIMIT) + ', 최근서비스=' + (t.lastEndpoint || '');
  });
  writeSelfDiagnosticResults_(results, startedAt);
  const failed = results.filter(function(r) { return r[1] === 'FAIL'; }).length;
  const skipped = results.filter(function(r) { return r[1] === 'SKIP'; }).length;
  const passed = results.filter(function(r) { return r[1] === 'PASS'; }).length;
  const partial = results.filter(function(r) { return r[1] === 'PARTIAL'; }).length;
  ui.alert('자가진단 완료\n\nPASS: ' + passed + '\nPARTIAL: ' + partial + '\nSKIP: ' + skipped + '\nFAIL: ' + failed + '\n\n상세 내용은 "자가진단" 시트에서 확인하세요.');
}

function diagStep_(results, name, fn) {
  try {
    const detail = fn();
    let status = 'PASS';
    if (String(detail || '').indexOf('SKIP') === 0) status = 'SKIP';
    else if (String(detail || '').indexOf('PARTIAL') === 0) status = 'PARTIAL';
    results.push([name, status, detail || '', now_()]);
  } catch (e) { results.push([name, 'FAIL', String(e && e.message ? e.message : e), now_()]); }
}

function writeSelfDiagnosticResults_(rows, startedAt) {
  const sheet = getSheet_(CONFIG.SHEETS.SELF_DIAGNOSTIC);
  const output = [CONFIG.HEADERS.SELF_DIAGNOSTIC, ['진단시작', 'INFO', startedAt, now_()]].concat(rows);
  sheet.clearContents();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
  try { sheet.autoResizeColumns(1, 4); } catch (e) {}
  for (var r = 2; r <= output.length; r++) {
    const status = String(sheet.getRange(r, 2).getValue() || '');
    if (status === 'FAIL') sheet.getRange(r, 1, 1, 4).setBackground('#fce8e6');
    else if (status === 'SKIP') sheet.getRange(r, 1, 1, 4).setBackground('#fff4ce');
    else if (status === 'PARTIAL') sheet.getRange(r, 1, 1, 4).setBackground('#fff9e6');
    else if (status === 'PASS') sheet.getRange(r, 1, 1, 4).setBackground('#e6f4ea');
  }
  log_('self_diagnostic', '자가진단 완료 / FAIL=' + rows.filter(function(r) { return r[1] === 'FAIL'; }).length);
}

function checkRequiredSheetsAndHeaders_() {
  const ss = SpreadsheetApp.getActive();
  const targets = {};
  targets[CONFIG.SHEETS.DASHBOARD] = CONFIG.HEADERS.DASHBOARD;
  targets[CONFIG.SHEETS.PRODUCTS] = CONFIG.HEADERS.PRODUCTS;
  targets[CONFIG.SHEETS.FILTERS] = CONFIG.HEADERS.FILTERS;
  targets[CONFIG.SHEETS.SALES_CLEAN] = CONFIG.HEADERS.SALES_CLEAN;
  targets[CONFIG.SHEETS.MONTHLY_BRAND] = CONFIG.HEADERS.MONTHLY_BRAND;
  targets[CONFIG.SHEETS.BRAND] = CONFIG.HEADERS.BRAND;
  targets[CONFIG.SHEETS.COUPANG_SENT_MANUAL] = CONFIG.HEADERS.COUPANG_SENT_MANUAL;
  targets[CONFIG.SHEETS.BRAND_ALIAS] = CONFIG.HEADERS.BRAND_ALIAS;
  targets[CONFIG.SHEETS.MATCH_DIAG] = CONFIG.HEADERS.MATCH_DIAG;
  targets[CONFIG.SHEETS.SYNC_STATUS] = CONFIG.HEADERS.SYNC_STATUS;
  targets[CONFIG.SHEETS.LOG] = CONFIG.HEADERS.LOG;
  targets[CONFIG.SHEETS.SELF_DIAGNOSTIC] = CONFIG.HEADERS.SELF_DIAGNOSTIC;
  targets[CONFIG.SHEETS.AUDIT_REPORT] = CONFIG.HEADERS.AUDIT_REPORT;
  const errors = [];
  let checked = 0;
  Object.keys(targets).forEach(function(name) {
    checked++;
    const sheet = ss.getSheetByName(name);
    if (!sheet) { errors.push('시트 없음: ' + name); return; }
    const expected = targets[name].map(function(v) { return String(v || '').trim(); });
    const actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0].map(function(v) { return String(v || '').trim(); });
    for (var i = 0; i < expected.length; i++) {
      if (actual[i] !== expected[i]) { errors.push(name + ' 헤더 불일치: ' + (i + 1) + '열 기대=' + expected[i] + ', 실제=' + actual[i]); break; }
    }
  });
  return { checked: checked, errors: errors };
}

function findFirstLotteonFilterName_(items) {
  for (var i = 0; i < (items || []).length; i++) {
    const name = findFilterName_(items[i]);
    if (name && String(name).trim().startsWith(CONFIG.FILTER_PREFIX)) return String(name).trim();
  }
  return '';
}

function selfDiagnosticDummy_() {}

function setupLotteonSheets(showPopup, skipFormat) {
  const ss = SpreadsheetApp.getActive();
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.DASHBOARD, CONFIG.HEADERS.DASHBOARD);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.PRODUCTS, CONFIG.HEADERS.PRODUCTS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.PRODUCTS_TMP, CONFIG.HEADERS.TMP_PRODUCTS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.MARKETS, CONFIG.HEADERS.MARKETS);
  const salesIn = ensureSheet_(ss, CONFIG.SHEETS.SALES_IN);
  if (salesIn.getLastRow() === 0) salesIn.getRange(1, 1, 1, 2).setValues([['여기에 더망고 매출 엑셀을 헤더 포함 그대로 붙여넣기','필수 필드: 마켓주문일자/마켓명/마켓주문번호/마켓아이디/마켓상품번호/사이트상품번호/브랜드/마켓상품명/상품명(원문)/구매사이트명/결제수량/결제금액합계(원)/마켓주문상태/더망고주문상태']]);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SALES_CLEAN, CONFIG.HEADERS.SALES_CLEAN);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.MONTHLY_BRAND, CONFIG.HEADERS.MONTHLY_BRAND);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND, CONFIG.HEADERS.BRAND);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.PRODUCT_PERF, CONFIG.HEADERS.PRODUCT_PERF);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.DELETE, CONFIG.HEADERS.DELETE);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.ACCOUNT, CONFIG.HEADERS.ACCOUNT);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SENT_CHECK, CONFIG.HEADERS.SENT_CHECK);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SENT_MANUAL, CONFIG.HEADERS.SENT_MANUAL);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.MANUAL_API_CHECK, CONFIG.HEADERS.MANUAL_API_CHECK);
  ensureCoupangWorkLogSheet_();
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.FILTER_RAW, CONFIG.HEADERS.FILTER_RAW);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND_ALIAS, CONFIG.HEADERS.BRAND_ALIAS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.MATCH_DIAG, CONFIG.HEADERS.MATCH_DIAG);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SYNC_STATUS, CONFIG.HEADERS.SYNC_STATUS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.LOG, CONFIG.HEADERS.LOG);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SELF_DIAGNOSTIC, CONFIG.HEADERS.SELF_DIAGNOSTIC);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.AUDIT_REPORT, CONFIG.HEADERS.AUDIT_REPORT);
  writeSettings_();
  if (skipFormat !== true) formatAllSheets_();
  if (showPopup !== false) {
    log_('setup', '초기 시트 생성/정리 완료');
    SpreadsheetApp.getUi().alert('초기 시트 생성/정리 완료\n\nv5.49 구조로 주요 시트 헤더를 갱신했습니다.\n매출데이터_붙여넣기 시트의 원본 데이터는 보존합니다.');
  }
}

function saveTheMangoCredentials() {
  const ui = SpreadsheetApp.getUi();
  const baseUrlResp = ui.prompt('TheMangoClient URL', '기본값: https://tmg2007.cafe24.com', ui.ButtonSet.OK_CANCEL);
  if (baseUrlResp.getSelectedButton() !== ui.Button.OK) return;
  const apiKeyResp = ui.prompt('API Key 입력', 'API Key를 붙여넣으세요. 실제 값은 시트에 노출하지 않습니다.', ui.ButtonSet.OK_CANCEL);
  if (apiKeyResp.getSelectedButton() !== ui.Button.OK) return;
  const senderResp = ui.prompt('X-API-SENDER 입력', 'X-API-SENDER를 붙여넣으세요. 실제 값은 시트에 노출하지 않습니다.', ui.ButtonSet.OK_CANCEL);
  if (senderResp.getSelectedButton() !== ui.Button.OK) return;
  const baseUrl = normalizeBaseUrl_(baseUrlResp.getResponseText() || CONFIG.BASE_URL_DEFAULT);
  const apiKey = String(apiKeyResp.getResponseText() || '').trim();
  const sender = String(senderResp.getResponseText() || '').trim();
  if (!apiKey || !sender) { ui.alert('API Key와 X-API-SENDER는 필수입니다.'); return; }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('THE_MANGO_BASE_URL', baseUrl);
  props.setProperty('THE_MANGO_API_KEY', apiKey);
  props.setProperty('THE_MANGO_SENDER', sender);
  log_('credentials', '인증값 저장 완료. 실제 값은 미표시');
  ui.alert('저장 완료. 실제 인증값은 시트/로그에 표시하지 않습니다.');
}


function startLotteonProductSync() {
  setupLotteonSheets(false);
  replaceData_(getSheet_(CONFIG.SHEETS.PRODUCTS_TMP), CONFIG.HEADERS.TMP_PRODUCTS, []);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SYNC_ACTIVE', 'Y'); props.setProperty('SYNC_PHASE', CONFIG.PHASE.COLLECT_FILTERS);
  props.setProperty('SYNC_FILTER_NAMES', '[]'); props.setProperty('SYNC_FILTER_INDEX', '0');
  props.setProperty('SYNC_PAGE', '1'); props.setProperty('SYNC_FILTER_LIST_PAGE', '1');
  props.setProperty('SYNC_CURRENT_FILTER', ''); props.setProperty('SYNC_CURRENT_TOTAL_PAGE', '');
  props.setProperty('SYNC_LAST_ERROR', ''); props.setProperty('SYNC_LAST_MESSAGE', '초기화 완료. 트리거로 분할 실행 시작 대기.');
  props.setProperty('SYNC_STARTED_AT', now_()); props.setProperty('SYNC_FINISHED_AT', ''); props.setProperty('SYNC_TOKEN_USED_EST', '0');
  writeSyncStatus_({ phase: CONFIG.PHASE.COLLECT_FILTERS, status: '초기화 완료 (트리거 대기)', filterCount: 0, filterIndex: 0, page: 1, currentFilter: '', totalPage: '', tmpRows: 0, lastUrl: '', lastError: '' });
  deleteTriggersForHandler_('continueLotteonProductSync');
  deleteTriggersForHandler_('runAnalysisAfterSync');
  ScriptApp.newTrigger('continueLotteonProductSync').timeBased().after(CONFIG.TRIGGER_AFTER_MS).create();
  log_('product_sync', '분할 업데이트 시작 - 초기화만 수행. 30초 후 트리거로 이어집니다.');
  SpreadsheetApp.getUi().alert('상품목록 분할 업데이트를 시작했습니다.\n\n진행상황은 "동기화상태" 시트에서 확인하세요.');
}

function runProductSyncOneStepManual() {
  const props = PropertiesService.getScriptProperties();
  const phase = props.getProperty('SYNC_PHASE') || CONFIG.PHASE.IDLE;
  if (phase === CONFIG.PHASE.FINALIZE) { finalizeProductSyncManual(); return; }
  if (props.getProperty('SYNC_ACTIVE') !== 'Y') {
    SpreadsheetApp.getUi().alert('현재 진행 중인 상품목록 동기화가 없습니다.\n\n새로 시작하려면 "⑤ LOTTEON 상품목록 분할 업데이트 시작"을 실행하세요.\n이미 임시 수집 데이터가 있고 최종 반영만 필요하면 "상품목록 최종 정리만 실행"을 실행하세요.');
    return;
  }
  continueLotteonProductSync();
  SpreadsheetApp.getUi().alert('분할 업데이트 1회 수동 실행을 완료했습니다.\n\n"동기화상태" 시트에서 PHASE와 필터 진행 상황을 확인하세요.');
}

function finalizeProductSyncManual() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { SpreadsheetApp.getUi().alert('다른 실행이 진행 중이라 최종 정리를 실행하지 못했습니다. 잠시 후 다시 시도하세요.'); return; }
  try {
    setupLotteonSheets(false, true);
    const tmpCount = countDataRows_(getSheet_(CONFIG.SHEETS.PRODUCTS_TMP));
    if (!tmpCount) { SpreadsheetApp.getUi().alert('LOTTEON_상품목록_TMP에 반영할 임시 데이터가 없습니다.\n\n먼저 "⑤ LOTTEON 상품목록 분할 업데이트 시작"을 실행해 상품목록을 수집하세요.'); return; }
    finalizeProductSync_();
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SYNC_ACTIVE', 'N'); props.setProperty('SYNC_PHASE', CONFIG.PHASE.DONE);
    props.setProperty('SYNC_FINISHED_AT', now_()); props.setProperty('SYNC_LAST_MESSAGE', '수동 최종 정리/반영 완료. 다음 단계로 ⑧ 분석 갱신을 실행하세요.'); props.setProperty('SYNC_LAST_ERROR', '');
    deleteTriggersForHandler_('continueLotteonProductSync'); deleteTriggersForHandler_('runAnalysisAfterSync');
    writeSyncStatus_({ phase: CONFIG.PHASE.DONE, status: '수동 최종 정리/반영 완료 - ⑧ 분석 갱신 필요', tmpRows: tmpCount, lastUrl: props.getProperty('LAST_API_URL') || '', lastError: '' });
    SpreadsheetApp.getUi().alert('상품목록 최종 정리/반영이 완료되었습니다.\n\n반영 임시행수: ' + formatCount_(tmpCount) + '행\n\n다음 단계로 "⑧ 매출데이터 정리 + 핵심분석 갱신"을 실행하세요.');
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function isTransientUrlFetchError_(msg) {
  return /URL_FETCH_ERROR|사용할 수 없는 주소|Address unavailable|DNS|timed out|Timeout|Exception:/.test(String(msg || ''));
}

function handleTransientSyncError_(msg) {
  const props = PropertiesService.getScriptProperties();
  const count = toNumber_(props.getProperty('SYNC_TRANSIENT_ERROR_COUNT')) + 1;
  props.setProperty('SYNC_TRANSIENT_ERROR_COUNT', String(count));
  props.setProperty('SYNC_LAST_ERROR', msg);
  props.setProperty('SYNC_LAST_MESSAGE', '일시적 API 접속 오류 재시도 대기: ' + count + '/' + CONFIG.MAX_TRANSIENT_URLFETCH_ERRORS);
  writeSyncStatus_({ phase: props.getProperty('SYNC_PHASE') || CONFIG.PHASE.COLLECT_FILTERS, status: '일시적 API 접속 오류 - 자동 재시도 예약됨 (' + count + '/' + CONFIG.MAX_TRANSIENT_URLFETCH_ERRORS + ')', currentFilter: props.getProperty('SYNC_CURRENT_FILTER') || '', page: toNumber_(props.getProperty('SYNC_PAGE')), totalPage: props.getProperty('SYNC_CURRENT_TOTAL_PAGE') || '', lastUrl: props.getProperty('LAST_API_URL') || '', lastError: msg });
  deleteTriggersForHandler_('continueLotteonProductSync');
  if (count >= CONFIG.MAX_TRANSIENT_URLFETCH_ERRORS) {
    props.setProperty('SYNC_ACTIVE', 'ERROR'); props.setProperty('SYNC_PHASE', CONFIG.PHASE.ERROR); props.setProperty('SYNC_LAST_MESSAGE', '일시적 API 접속 오류 반복 초과: ' + msg);
    writeSyncStatus_({ phase: CONFIG.PHASE.ERROR, status: '오류 - 일시적 API 접속 오류 반복 초과', currentFilter: props.getProperty('SYNC_CURRENT_FILTER') || '', page: toNumber_(props.getProperty('SYNC_PAGE')), totalPage: props.getProperty('SYNC_CURRENT_TOTAL_PAGE') || '', lastUrl: props.getProperty('LAST_API_URL') || '', lastError: msg });
    log_('product_sync_transient_error_limit', msg); return;
  }
  ScriptApp.newTrigger('continueLotteonProductSync').timeBased().after(CONFIG.TRANSIENT_RETRY_AFTER_MS).create();
  log_('product_sync_transient_retry', '일시적 API 접속 오류. ' + (CONFIG.TRANSIENT_RETRY_AFTER_MS / 1000) + '초 후 재시도 예약 / ' + msg);
}

function continueLotteonProductSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { log_('product_sync', '이미 다른 실행이 진행 중이라 이번 실행은 스킵'); return; }
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('SYNC_ACTIVE') !== 'Y') { deleteTriggersForHandler_('continueLotteonProductSync'); return; }
    const phase = props.getProperty('SYNC_PHASE') || CONFIG.PHASE.COLLECT_FILTERS;
    const started = Date.now();
    let nextPhase = phase;
    try {
      if (phase === CONFIG.PHASE.COLLECT_FILTERS) nextPhase = runPhaseCollectFilters_(started);
      else if (phase === CONFIG.PHASE.COLLECT_PRODUCTS) nextPhase = runPhaseCollectProducts_(started);
      else if (phase === CONFIG.PHASE.FINALIZE) nextPhase = runPhaseFinalize_();
      else { props.setProperty('SYNC_ACTIVE', 'N'); deleteTriggersForHandler_('continueLotteonProductSync'); return; }
      props.setProperty('SYNC_PHASE', nextPhase);
      if (nextPhase === CONFIG.PHASE.DONE) {
        props.setProperty('SYNC_ACTIVE', 'N'); props.setProperty('SYNC_FINISHED_AT', now_()); props.setProperty('SYNC_LAST_MESSAGE', '상품목록 수집/확정 완료. 분석은 2분 뒤 자동 실행 예정.');
        deleteTriggersForHandler_('continueLotteonProductSync'); deleteTriggersForHandler_('runAnalysisAfterSync');
        ScriptApp.newTrigger('runAnalysisAfterSync').timeBased().after(CONFIG.ANALYSIS_AFTER_MS).create();
        writeSyncStatus_({ phase: CONFIG.PHASE.DONE, status: '상품목록 확정 완료, 약 2분 뒤 분석 자동 실행 예정' });
        log_('product_sync', '분할 업데이트 전체 완료. 분석은 별도 트리거로 약 2분 뒤 실행');
      } else {
        deleteTriggersForHandler_('continueLotteonProductSync');
        ScriptApp.newTrigger('continueLotteonProductSync').timeBased().after(CONFIG.TRIGGER_AFTER_MS).create();
      }
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      if (isTransientUrlFetchError_(msg)) { handleTransientSyncError_(msg); return; }
      props.setProperty('SYNC_ACTIVE', 'ERROR'); props.setProperty('SYNC_PHASE', CONFIG.PHASE.ERROR); props.setProperty('SYNC_LAST_ERROR', msg); props.setProperty('SYNC_LAST_MESSAGE', '오류: ' + msg);
      writeSyncStatus_({ phase: CONFIG.PHASE.ERROR, status: '오류 - 자동 트리거 해제됨', currentFilter: props.getProperty('SYNC_CURRENT_FILTER') || '', page: toNumber_(props.getProperty('SYNC_PAGE')), totalPage: props.getProperty('SYNC_CURRENT_TOTAL_PAGE') || '', lastUrl: props.getProperty('LAST_API_URL') || '', lastError: msg });
      deleteTriggersForHandler_('continueLotteonProductSync'); log_('product_sync_error', msg); throw e;
    }
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function runAnalysisAfterSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { log_('analysis_after_sync', '락 획득 실패. 다른 실행과 충돌'); return; }
  try {
    deleteTriggersForHandler_('runAnalysisAfterSync');
    try { refreshLotteonAnalysis(); log_('analysis_after_sync', '동기화 완료 후 분석 자동 실행 완료'); }
    catch (e) { log_('analysis_after_sync_error', String(e && e.message ? e.message : e)); }
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function runPhaseCollectFilters_(startedMs) {
  const props = PropertiesService.getScriptProperties();
  let page = toNumber_(props.getProperty('SYNC_FILTER_LIST_PAGE')) || 1;
  let collected = JSON.parse(props.getProperty('SYNC_FILTER_NAMES') || '[]');
  let processedPages = 0;
  while (Date.now() - startedMs < CONFIG.MAX_RUN_MS && processedPages < CONFIG.MAX_FILTER_PAGES_PER_RUN) {
    const payload = { page: String(page), searchQuery: { searchKeyword: '', siteId: '', filterGroup: 'all', sort: 'nameAsc' } };
    writeSyncStatus_({ phase: CONFIG.PHASE.COLLECT_FILTERS, status: '필터목록 수집 중', page: page, currentFilter: 'filterList page=' + page, lastUrl: buildApiUrl_('filterList'), filterCount: collected.length });
    const resp = apiRequest_('filterList', 'post', payload);
    props.setProperty('SYNC_TRANSIENT_ERROR_COUNT', '0');
    const data = extractData_(resp);
    const items = extractItems_(resp, false);
    items.forEach(function(item) { const name = findFilterName_(item); if (name && String(name).trim().startsWith(CONFIG.FILTER_PREFIX)) collected.push(String(name).trim()); });
    const totalPage = toNumber_(data.totalPage);
    processedPages++;
    if (!items.length || (totalPage && page >= totalPage) || page >= CONFIG.MAX_PAGE_PER_QUERY) {
      const uniqueSorted = unique_(collected).sort();
      props.setProperty('SYNC_FILTER_NAMES', JSON.stringify(uniqueSorted)); props.setProperty('SYNC_FILTER_INDEX', '0'); props.setProperty('SYNC_PAGE', '1'); props.setProperty('SYNC_FILTER_LIST_PAGE', String(page + 1)); props.setProperty('SYNC_LAST_MESSAGE', '필터목록 수집 완료: ' + uniqueSorted.length + '개');
      writeSyncStatus_({ phase: CONFIG.PHASE.COLLECT_PRODUCTS, status: '필터목록 수집 완료, 상품 수집 대기', filterCount: uniqueSorted.length, filterIndex: 0, page: 1, currentFilter: uniqueSorted[0] || '', lastUrl: '' });
      log_('product_sync', '필터목록 수집 완료 / 롯백 필터 ' + uniqueSorted.length + '개');
      if (!uniqueSorted.length) throw new Error('롯백으로 시작하는 검색필터를 찾지 못했습니다.');
      return CONFIG.PHASE.COLLECT_PRODUCTS;
    }
    page++; Utilities.sleep(CONFIG.REQUEST_DELAY_MS);
  }
  props.setProperty('SYNC_FILTER_LIST_PAGE', String(page)); props.setProperty('SYNC_FILTER_NAMES', JSON.stringify(collected)); props.setProperty('SYNC_LAST_MESSAGE', '필터목록 수집 진행중 page=' + page);
  writeSyncStatus_({ phase: CONFIG.PHASE.COLLECT_FILTERS, status: '필터목록 수집 진행중 (다음 실행 대기)', page: page, currentFilter: 'filterList page=' + page, filterCount: unique_(collected).length });
  return CONFIG.PHASE.COLLECT_FILTERS;
}

function runPhaseCollectProducts_(startedMs) {
  const props = PropertiesService.getScriptProperties();
  const filterNames = JSON.parse(props.getProperty('SYNC_FILTER_NAMES') || '[]');
  let filterIndex = toNumber_(props.getProperty('SYNC_FILTER_INDEX')) || 0;
  let page = toNumber_(props.getProperty('SYNC_PAGE')) || 1;
  if (!filterNames.length) throw new Error('수집된 필터목록이 없습니다.');
  const tmp = getSheet_(CONFIG.SHEETS.PRODUCTS_TMP);
  const filterRows = readTable_(getSheet_(CONFIG.SHEETS.FILTERS));
  const filterSummaryMap = {};
  filterRows.forEach(function(r) { filterSummaryMap[String(r['검색필터명'] || '')] = r; });
  let processedPages = 0, summaryDirty = false;
  const rowsBuffer = [];
  try {
    while (filterIndex < filterNames.length && Date.now() - startedMs < CONFIG.MAX_RUN_MS && processedPages < CONFIG.MAX_PAGES_PER_RUN) {
      const filterName = filterNames[filterIndex];
      const account = accountFromFilterName_(filterName);
      const brand = brandFromFilterName_(filterName);
      const payload = { page: String(page), searchQuery: { siteId: CONFIG.SITE_ID, searchType: 'filterName', searchKeyword: filterName, condition: '', sort: 'dateDesc' } };
      props.setProperty('SYNC_CURRENT_FILTER', filterName);
      if (processedPages % 5 === 0) writeSyncStatus_({ phase: CONFIG.PHASE.COLLECT_PRODUCTS, status: '상품 수집 중 (버퍼누적 ' + rowsBuffer.length + '행)', filterCount: filterNames.length, filterIndex: filterIndex, page: page, currentFilter: filterName, totalPage: props.getProperty('SYNC_CURRENT_TOTAL_PAGE') || '', tmpRows: countDataRows_(tmp), lastUrl: buildApiUrl_('productList') });
      const resp = apiRequest_('productList', 'post', payload);
      props.setProperty('SYNC_TRANSIENT_ERROR_COUNT', '0');
      const data = extractData_(resp);
      const items = extractItems_(resp, true);
      const totalPage = toNumber_(data.totalPage), totalCount = toNumber_(data.totalCount);
      props.setProperty('SYNC_CURRENT_TOTAL_PAGE', String(totalPage || ''));
      const updateAt = now_();
      if (page === 1) {
        filterSummaryMap[filterName] = { '검색필터명': filterName, '브랜드명': brand, '계정번호': account.accountNo, '쿠팡계정ID': account.accountId, 'API_totalCount': totalCount, 'API_totalPage': totalPage, '이번조회_행수': filterSummaryMap[filterName] ? toNumber_(filterSummaryMap[filterName]['이번조회_행수']) : 0, '필터코드': filterCodeFromFilterName_(filterName), '메모': '' };
        summaryDirty = true;
        log_('product_sync', filterName + ' 시작 / totalCount=' + totalCount + ', totalPage=' + totalPage);
      }
      if (items.length) {
        items.forEach(function(item) { rowsBuffer.push([updateAt, getAny_(item, ['productId','prdNo','id','goodsNo']), getAny_(item, ['siteId','siteID']) || CONFIG.SITE_ID, getAny_(item, ['itemName','productName','name','modelName','title']), getAny_(item, ['modelName','model']), getAny_(item, ['image','imageUrl','thumbnail','thumb']), getAny_(item, ['filterId','filterID']), filterName, brand, account.accountNo, account.accountId, findDateValue_(item, ['createDate','createdAt','regDate','registerDate','insertDate','registeredAt']), findDateValue_(item, ['collectionDate','collectDate','collectedAt','crawlDate','scrapDate']), findDateValue_(item, ['updateDate','modifiedAt','modifyDate','updatedAt']), findDateValue_(item, ['marketSendDate','sendDate','sentDate','transmitDate','marketRegDate','coupangSendDate']), 'siteId=' + CONFIG.SITE_ID + ', filterName=' + filterName]); });
        filterSummaryMap[filterName]['이번조회_행수'] = toNumber_(filterSummaryMap[filterName]['이번조회_행수']) + items.length;
        summaryDirty = true;
      }
      page++; processedPages++;
      if (!items.length || (totalPage && page > totalPage) || page > CONFIG.MAX_PAGE_PER_QUERY) { filterIndex++; page = 1; props.setProperty('SYNC_CURRENT_TOTAL_PAGE', ''); }
      Utilities.sleep(CONFIG.REQUEST_DELAY_MS);
    }
  } finally {
    if (rowsBuffer.length) { appendRows_(tmp, rowsBuffer); log_('product_sync_buffer_flush', '버퍼 flush: ' + rowsBuffer.length + '행'); }
    props.setProperty('SYNC_FILTER_INDEX', String(filterIndex)); props.setProperty('SYNC_PAGE', String(page));
    if (summaryDirty) { try { writeFilterSummary_(filterSummaryMap); } catch (e) { log_('filter_summary_write_error', String(e && e.message ? e.message : e)); } }
  }
  if (filterIndex >= filterNames.length) {
    props.setProperty('SYNC_LAST_MESSAGE', '상품 수집 완료, 최종화 대기');
    writeSyncStatus_({ phase: CONFIG.PHASE.FINALIZE, status: '상품 수집 완료, 최종 정리 대기', filterCount: filterNames.length, filterIndex: filterIndex, page: page, currentFilter: '', tmpRows: countDataRows_(tmp) });
    return CONFIG.PHASE.FINALIZE;
  }
  props.setProperty('SYNC_LAST_MESSAGE', '진행중: 필터 ' + (filterIndex + 1) + '/' + filterNames.length + ', page=' + page);
  writeSyncStatus_({ phase: CONFIG.PHASE.COLLECT_PRODUCTS, status: '상품 수집 진행중 (다음 실행 대기)', filterCount: filterNames.length, filterIndex: filterIndex, page: page, currentFilter: filterNames[filterIndex] || '', totalPage: props.getProperty('SYNC_CURRENT_TOTAL_PAGE') || '', tmpRows: countDataRows_(tmp) });
  log_('product_sync_progress', '필터 ' + (filterIndex + 1) + '/' + filterNames.length + ', page=' + page);
  return CONFIG.PHASE.COLLECT_PRODUCTS;
}

function runPhaseFinalize_() { finalizeProductSync_(); writeSyncStatus_({ phase: CONFIG.PHASE.DONE, status: '완료', lastUrl: PropertiesService.getScriptProperties().getProperty('LAST_API_URL') || '' }); return CONFIG.PHASE.DONE; }

function finalizeProductSync_() {
  const tmpRows = readTable_(getSheet_(CONFIG.SHEETS.PRODUCTS_TMP));
  const existingRows = readTable_(getSheet_(CONFIG.SHEETS.PRODUCTS));
  const sentManualRows = readTable_(getSheet_(CONFIG.SHEETS.SENT_MANUAL));
  const existingFirstSeen = {};
  existingRows.forEach(function(r) { const productId = String(r['상품ID'] || '').trim(); if (productId) existingFirstSeen[productId] = r['최초확인일']; });
  const manualSentMap = {};
  sentManualRows.forEach(function(r) { const productId = String(r['상품ID'] || '').trim(); if (productId && r['마켓전송일']) manualSentMap[productId] = r['마켓전송일']; });
  const seen = {};
  const rows = [];
  const nowText = now_();
  tmpRows.forEach(function(r) {
    const productId = String(r['상품ID'] || '').trim();
    const fallbackKey = productId || String(r['상품명'] || '') + '|' + String(r['검색필터명'] || '');
    if (!fallbackKey || seen[fallbackKey]) return;
    seen[fallbackKey] = true;
    const firstSeen = existingFirstSeen[productId] || nowText;
    const apiReg = r['API_등록일'], apiCollect = r['API_수집일'], apiSend = r['API_마켓전송일'];
    const manualSend = manualSentMap[productId] || '';
    const baseDate = normalizeDateText_(apiReg || apiCollect || firstSeen);
    const lotteonCode = extractLotteonProductCodeFromFields_(r['이미지URL'], r['상품명'], r['모델명']);
    const rowNo = rows.length + 2;
    const daysFormula = '=IF(S' + rowNo + '="","",MAX(0,TODAY()-DATEVALUE(LEFT(S' + rowNo + ',10))))';
    rows.push([r['업데이트일시'], firstSeen, productId, r['사이트ID'], r['상품명'], r['모델명'], r['이미지URL'], lotteonCode, r['필터ID'], r['검색필터명'], r['브랜드명'], r['계정번호'], r['쿠팡계정ID'], apiReg, apiCollect, r['API_수정일'], apiSend, manualSend, baseDate, daysFormula, r['원본_scope']]);
  });
  replaceData_(getSheet_(CONFIG.SHEETS.PRODUCTS), CONFIG.HEADERS.PRODUCTS, rows);
  log_('product_sync_finalize', '최종 상품목록 생성 완료 / 상품수=' + rows.length);
}

function resetProductSyncState() {
  const props = PropertiesService.getScriptProperties();
  ['SYNC_ACTIVE','SYNC_PHASE','SYNC_FILTER_NAMES','SYNC_FILTER_INDEX','SYNC_PAGE','SYNC_FILTER_LIST_PAGE','SYNC_LAST_ERROR','SYNC_CURRENT_FILTER','SYNC_CURRENT_TOTAL_PAGE','SYNC_LAST_MESSAGE','SYNC_STARTED_AT','SYNC_FINISHED_AT','SYNC_TOKEN_USED_EST','LAST_API_URL'].forEach(function(k) { props.deleteProperty(k); });
  deleteTriggersForHandler_('continueLotteonProductSync'); deleteTriggersForHandler_('runAnalysisAfterSync');
  writeSyncStatus_({ phase: CONFIG.PHASE.IDLE, status: '초기화', lastUrl: '' });
  log_('product_sync', '동기화 상태 초기화');
  SpreadsheetApp.getUi().alert('상품목록 동기화 상태를 초기화했습니다.');
}


function ensureAnalysisSheetsLight_() {
  // v5.67:
  // ⑧ 매출분석은 더 이상 LOTTEON_상품목록 전체를 읽지 않습니다.
  // 브랜드 매출/효율 분석에 필요한 최소 시트만 보장합니다.
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, CONFIG.SHEETS.SALES_IN);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SALES_CLEAN, CONFIG.HEADERS.SALES_CLEAN);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND, CONFIG.HEADERS.BRAND);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.HEADERS.BRAND_SUMMARY);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.MATCH_DIAG, CONFIG.HEADERS.MATCH_DIAG);
  ensureCoupangWorkLogSheet_();
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.SYNC_STATUS, CONFIG.HEADERS.SYNC_STATUS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.LOG, CONFIG.HEADERS.LOG);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.BRAND_ALIAS, CONFIG.HEADERS.BRAND_ALIAS);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.AUDIT_REPORT, CONFIG.HEADERS.AUDIT_REPORT);
}

function buildAnalysisContextWithoutProductList_() {
  // buildMatchingContext_는 빈 상품목록이어도 필터별_상품수/수동입력/브랜드명_매칭표의 브랜드를
  // 매칭 후보로 넣도록 되어 있습니다. 상품 단위 매칭은 줄지만 브랜드 매출 분석은 훨씬 빨라집니다.
  return buildMatchingContext_([]);
}

function buildBrandMetricsFastWithoutProductList_(sales) {
  return buildBrandMetrics_([], {}, sales || []);
}

function refreshLotteonAnalysis() {
  const started = Date.now();
  ensureAnalysisSheetsLight_();

  writeSyncStatus_({
    phase: 'ANALYSIS_RUNNING',
    status: '⑧ 매출데이터 정리 + 핵심분석 갱신 중 - 상품목록 미조회 초고속 모드',
    currentFilter: '',
    lastUrl: '',
    tmpRows: ''
  });

  // v5.67:
  // 시간초과 핵심 원인이던 LOTTEON_상품목록 3~4만행 readTable_를 제거합니다.
  // 상품수/전송수는 필터별_상품수 + 쿠팡전송수_수동입력 기준으로 가져오고,
  // 매출은 매출데이터_붙여넣기 브랜드 기준으로 집계합니다.
  const ctx = buildAnalysisContextWithoutProductList_();
  const salesResult = normalizeSalesData_([], ctx);
  const sales = salesResult.rows;

  const brandMetrics = buildBrandMetricsFastWithoutProductList_(sales);
  const brandRows = buildBrandRows_(brandMetrics);
  const summaryRows = buildBrandSummaryRowsFromBrandRows_(brandRows);

  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND), CONFIG.HEADERS.BRAND, brandRows);
  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND_SUMMARY), CONFIG.HEADERS.BRAND_SUMMARY, summaryRows);

  syncCoupangWorkLogMetrics_(brandRows);
  writeMatchDiagnostic_(salesResult.diag);

  formatCoreAnalysisSheetsLightOnly_();

  const elapsedSec = Math.round((Date.now() - started) / 1000);
  writeSyncStatus_({
    phase: 'ANALYSIS_DONE',
    status: '⑧ 매출데이터 정리 + 핵심분석 갱신 완료 - 상품목록 미조회 초고속 모드',
    filterCount: '',
    filterIndex: '',
    page: '',
    currentFilter: '',
    totalPage: '',
    tmpRows: '',
    lastUrl: '',
    lastError: ''
  });

  log_('analysis_fast_without_product_list', '완료 / 매출정리행=' + sales.length + ', 브랜드=' + brandRows.length + ', 소요초=' + elapsedSec);

  SpreadsheetApp.getUi().alert(
    '⑧ 매출데이터 정리 + 핵심분석 갱신 완료\n\n' +
    '실행모드: 상품목록 미조회 초고속 모드\n' +
    '매출정리 행수: ' + formatCount_(sales.length) + '행\n' +
    '브랜드 수: ' + formatCount_(brandRows.length) + '개\n' +
    '소요 시간: 약 ' + elapsedSec + '초\n\n' +
    '상품수/전송수 기준:\n' +
    '  ① 필터별_상품수\n' +
    '  ② 쿠팡전송수_수동입력\n\n' +
    '상품 단위 상세성과가 아니라 브랜드/필터별 매출 효율 분석에 최적화된 모드입니다.'
  );
}

function buildMatchingContext_(products) {
  const productIdToBrand = {}, productIdToName = {}, productIdToFilter = {}, siteProductToProductId = {}, dmBrandSet = {}, dmBrandCanonical = {}, brandProductIndex = {};

  function addBrandToContext_(brand) {
    const b = String(brand || '').trim();
    if (!b) return;
    const key = normalizeBrandKey_(b);
    if (!key) return;
    dmBrandSet[key] = true;
    if (!dmBrandCanonical[key]) dmBrandCanonical[key] = b;
  }

  products.forEach(function(p) {
    const pid = String(p['상품ID'] || '').trim();
    const brand = String(p['브랜드명'] || '').trim();
    const name = String(p['상품명'] || '').trim();
    const lotteonCode = String(p['롯데온상품번호_추출'] || '').trim().toUpperCase();

    if (pid) {
      productIdToBrand[pid] = brand;
      productIdToName[pid] = name;
      productIdToFilter[pid] = p['검색필터명'];
    }

    addBrandToContext_(brand);

    if (lotteonCode && pid && !siteProductToProductId[lotteonCode]) siteProductToProductId[lotteonCode] = pid;

    if (brand && name && pid) {
      const key = normalizeBrandKey_(brand) + '|' + normalizeProductNameKey_(name);
      if (key !== '|' && !brandProductIndex[key]) brandProductIndex[key] = pid;
    }
  });

  // v5.62:
  // 기존에는 LOTTEON_상품목록에 들어 있는 브랜드만 "분석 가능한 브랜드"로 봤습니다.
  // 02→04 필터 이동, 빠른갱신, 상품목록 미갱신 상황에서는 상품목록에 없는 브랜드 매출이 0으로 빠질 수 있어
  // 필터별_상품수와 쿠팡전송수_수동입력의 브랜드도 매칭 후보에 포함합니다.
  try {
    readTable_(getSheet_(CONFIG.SHEETS.FILTERS)).forEach(function(r) {
      addBrandToContext_(getObjectValueByHeader_(r, '브랜드명') || brandFromFilterName_(getObjectValueByHeader_(r, '검색필터명')));
    });
  } catch (e) {}

  try {
    readTable_(getSheet_(CONFIG.SHEETS.COUPANG_SENT_MANUAL)).forEach(function(r) {
      addBrandToContext_(getObjectValueByHeader_(r, '브랜드명') || brandFromFilterName_(getObjectValueByHeader_(r, '검색필터명')));
    });
  } catch (e) {}

  const aliasMap = {};
  try {
    readTable_(getSheet_(CONFIG.SHEETS.BRAND_ALIAS)).forEach(function(r) {
      const src = normalizeBrandKey_(r['매출브랜드명']);
      const dst = String(r['더망고브랜드명'] || '').trim();
      if (src && dst) {
        aliasMap[src] = dst;
        addBrandToContext_(dst);
      }
    });
  } catch (e) {}

  return {
    productIdToBrand: productIdToBrand,
    productIdToName: productIdToName,
    productIdToFilter: productIdToFilter,
    siteProductToProductId: siteProductToProductId,
    dmBrandSet: dmBrandSet,
    dmBrandCanonical: dmBrandCanonical,
    aliasMap: aliasMap,
    brandProductIndex: brandProductIndex
  };
}

function matchSalesRow_(siteProductCode, salesBrandRaw, salesProductName, ctx) {
  const code = String(siteProductCode || '').trim().toUpperCase();

  if (code && ctx.siteProductToProductId[code]) {
    const pid = ctx.siteProductToProductId[code];
    return { brand: ctx.productIdToBrand[pid] || '', productId: pid, method: 'siteProductId', isProductMatched: true };
  }

  const salesBrand = String(salesBrandRaw || '').trim();
  const salesBrandKey = normalizeBrandKey_(salesBrand);
  let mangoBrand = '', brandMethod = '';

  if (salesBrandKey) {
    if (ctx.dmBrandSet[salesBrandKey]) {
      mangoBrand = ctx.dmBrandCanonical[salesBrandKey];
      brandMethod = 'brandDirect';
    } else if (ctx.aliasMap[salesBrandKey]) {
      mangoBrand = ctx.aliasMap[salesBrandKey];
      brandMethod = 'brandAlias';
    } else if (salesBrand && !isGenericInvalidBrand_(salesBrand)) {
      // v5.62:
      // 더망고 매출데이터의 "브랜드" 컬럼은 실제 주문 기준값이므로,
      // LOTTEON_상품목록이 최신이 아니어서 ctx에 없는 브랜드라도 브랜드 매출로 반영합니다.
      // 예: 상품목록은 과거 02 필터 기준인데 매출데이터는 롯백_04_엄브로 주문으로 들어온 경우
      mangoBrand = salesBrand;
      brandMethod = 'brandDirect';
    }
  }

  if (mangoBrand && salesProductName) {
    const nameKey = normalizeProductNameKey_(salesProductName);
    const key = normalizeBrandKey_(mangoBrand) + '|' + nameKey;
    if (nameKey && ctx.brandProductIndex[key]) {
      return { brand: mangoBrand, productId: ctx.brandProductIndex[key], method: 'brandProductName', isProductMatched: true };
    }
  }

  if (brandMethod === 'brandDirect') return { brand: mangoBrand, productId: '', method: 'brandDirect', isProductMatched: false };
  if (brandMethod === 'brandAlias') return { brand: mangoBrand, productId: '', method: 'brandAlias', isProductMatched: false };

  // v5.62: 브랜드 컬럼이 비어 있거나 깨진 경우 상품명/마켓상품명 안의 브랜드명으로 최종 추정
  const inferredBrand = inferBrandFromSalesText_(salesProductName, ctx);
  if (inferredBrand) {
    return { brand: inferredBrand, productId: '', method: 'brandText', isProductMatched: false };
  }

  return { brand: salesBrand, productId: '', method: '미매칭', isProductMatched: false };
}

function isGenericInvalidBrand_(brand) {
  const s = normalizeBrandKey_(brand);
  if (!s) return true;
  if (s === '브랜드' || s === '브랜드명' || s === '없음' || s === '미확인' || s === '기타') return true;
  return false;
}

function inferBrandFromSalesText_(salesProductName, ctx) {
  const textKey = normalizeProductNameKey_(salesProductName);
  if (!textKey || !ctx || !ctx.dmBrandCanonical) return '';

  const candidates = [];
  Object.keys(ctx.dmBrandCanonical).forEach(function(brandKey) {
    const brand = String(ctx.dmBrandCanonical[brandKey] || '').trim();
    if (!brand) return;

    const normalizedBrand = normalizeProductNameKey_(brand);
    if (!normalizedBrand) return;

    // 한 글자 브랜드(예: 숲, 랩)는 상품명 텍스트 자동추정에서 오탐 가능성이 있어 제외합니다.
    if (normalizedBrand.length < 2) return;

    if (textKey.indexOf(normalizedBrand) >= 0) {
      candidates.push({ brand: brand, key: normalizedBrand, len: normalizedBrand.length });
    }
  });

  if (!candidates.length) return '';
  candidates.sort(function(a, b) {
    if (b.len !== a.len) return b.len - a.len;
    return String(a.brand).localeCompare(String(b.brand));
  });
  return candidates[0].brand;
}

function isCancelledByMangoStatus_(mangoStatus) {
  // v5.65: 매출 분석의 주 기준은 취소/반품 포함 전체 주문/매출입니다.
  // 취소/반품 여부는 정보성 지표로만 분리하며, 판정 기준은 더망고주문상태 열만 사용합니다.
  return CONFIG.CANCEL_PATTERN.test(String(mangoStatus || ''));
}

function normalizeSalesData_(products, ctx) {
  const src = getSheet_(CONFIG.SHEETS.SALES_IN), out = getSheet_(CONFIG.SHEETS.SALES_CLEAN);
  const diag = {
    total: 0,
    lotteonTarget: 0,
    excluded: 0,
    valid: 0,
    cancelled: 0,
    cancelledRevenue: 0,
    bySiteProductId: 0,
    byBrandProductName: 0,
    byBrandDirect: 0,
    byBrandAlias: 0,
    byBrandText: 0,
    unmatched: 0,
    revenueProductMatched: 0,
    revenueBrandOnly: 0,
    revenueUnmatched: 0,
    revenueAnalyzed: 0
  };

  if (src.getLastRow() < 2) {
    replaceDataFastLimited_(out, CONFIG.HEADERS.SALES_CLEAN, []);
    return { rows: [], diag: diag };
  }

  const values = getTrimmedSheetValues_(src, 3000, 80);
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const rows = [];

  for (var r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(function(v) { return String(v || '').trim() === ''; })) continue;
    diag.total++;

    const orderDate = pickByHeader_(headers, row, ['마켓주문일자','주문일시','주문일','결제일시','결제일','주문등록일시']);
    const marketName = pickByHeader_(headers, row, ['마켓명','판매처','채널','쇼핑몰','판매마켓']);
    const marketOrderNo = pickByHeader_(headers, row, ['마켓주문번호','주문번호','주문ID','ordNo']);
    const acct = normalizeAccountId_(pickByHeader_(headers, row, ['마켓아이디','판매자ID','계정ID','쿠팡계정ID','아이디']));
    const marketProductNo = pickByHeader_(headers, row, ['마켓상품번호','쿠팡상품번호']);
    const siteProductNo = pickByHeader_(headers, row, ['사이트상품번호','lotteon상품번호','롯데온상품번호','원상품번호']);
    const brandRaw = pickByHeader_(headers, row, ['브랜드','브랜드명','상표']);
    const marketProductName = pickByHeader_(headers, row, ['마켓상품명']);
    let originalProductName = pickByHeader_(headers, row, ['상품명(원문)','원문상품명','상품명원문','소싱상품명']);
    if (!originalProductName) originalProductName = pickByHeader_(headers, row, ['상품명']) || '';
    const productUrl = pickByHeader_(headers, row, ['상품URL','상품링크']);
    const productImage = pickByHeader_(headers, row, ['상품이미지','상품이미지URL','썸네일']);
    const purchaseSite = pickByHeader_(headers, row, ['구매사이트명','구매사이트','소싱사이트','소싱사이트명']);
    const qty = toNumber_(pickByHeader_(headers, row, ['결제수량','수량','주문수량','판매수량'])) || 1;
    const paymentAmount = toNumber_(pickByHeader_(headers, row, ['결제금액합계(원)','결제금액합계','결제금액','판매금액','매출금액','상품금액','주문금액']));
    const shippingFee = toNumber_(pickByHeader_(headers, row, ['결제배송비(원)','결제배송비','배송비']));
    const settlementAmount = toNumber_(pickByHeader_(headers, row, ['정산예정금액(원)','정산예정금액','정산금액']));
    let marketStatus = pickByHeader_(headers, row, ['마켓주문상태']);
    const mangoStatus = pickByHeader_(headers, row, ['더망고주문상태']);
    if (!marketStatus) marketStatus = pickByHeader_(headers, row, ['주문상태','상태','배송상태']);

    const revenue = paymentAmount || settlementAmount;
    const isAnalysisTarget = CONFIG.LOTTEON_SITE_REGEX.test(String(purchaseSite || ''));
    const isCancelledInfo = isCancelledByMangoStatus_(mangoStatus);
    // v5.65: 취소/반품과 관계없이 수집/전송 후 발생한 주문은 모두 주 분석값에 반영합니다.
    const isValid = true;

    if (isAnalysisTarget) {
      diag.lotteonTarget++;
      diag.valid++;
      if (isCancelledInfo) {
        diag.cancelled++;
        diag.cancelledRevenue += revenue;
      }
    } else {
      diag.excluded++;
    }

    const matchText = [originalProductName, marketProductName].filter(Boolean).join(' ');
    const match = matchSalesRow_(siteProductNo, brandRaw, matchText, ctx);

    if (isAnalysisTarget) {
      if (match.method === 'siteProductId') diag.bySiteProductId++;
      else if (match.method === 'brandProductName') diag.byBrandProductName++;
      else if (match.method === 'brandDirect') diag.byBrandDirect++;
      else if (match.method === 'brandAlias') diag.byBrandAlias++;
      else if (match.method === 'brandText') diag.byBrandText++;
      else diag.unmatched++;

      if (match.isProductMatched) {
        diag.revenueProductMatched += revenue;
        diag.revenueAnalyzed += revenue;
      } else if (match.method === 'brandDirect' || match.method === 'brandAlias' || match.method === 'brandText') {
        diag.revenueBrandOnly += revenue;
        diag.revenueAnalyzed += revenue;
      } else {
        diag.revenueUnmatched += revenue;
      }
    }

    rows.push([
      monthKey_(orderDate),
      normalizeDateText_(orderDate),
      marketName,
      marketOrderNo,
      acct.original,
      acct.normalized,
      acct.accountNo,
      marketProductNo,
      siteProductNo,
      brandRaw,
      match.brand,
      marketProductName,
      originalProductName,
      productUrl,
      productImage,
      purchaseSite,
      qty,
      revenue,
      shippingFee,
      settlementAmount,
      marketStatus,
      mangoStatus,
      isValid ? 'Y' : 'N',
      isAnalysisTarget ? 'Y' : 'N',
      match.method,
      match.productId,
      r + 1,
      isCancelledInfo ? 'Y' : 'N',
      '더망고주문상태'
    ]);
  }

  replaceDataFastLimited_(out, CONFIG.HEADERS.SALES_CLEAN, rows);
  return { rows: readTable_(out), diag: diag };
}

function writeMatchDiagnostic_(diag) {
  const total = diag.total || 0;
  const lotteon = diag.lotteonTarget || 0;
  const included = diag.valid || 0;
  const totalRevenueAttempted = diag.revenueAnalyzed + diag.revenueUnmatched;
  const pct = function(n, d) { return d ? (n / d) : 0; };

  const rows = [
    ['원본 매출행수', total, 1, '매출데이터_붙여넣기 전체 행'],
    ['LOTTEON 분석대상 행수', lotteon, pct(lotteon, total), '구매사이트명에 lotteon 포함'],
    ['분석대상 제외 행수', diag.excluded, pct(diag.excluded, total), 'LOTTEON 외 구매사이트'],
    ['분석 반영 행수', included, pct(included, lotteon), '취소/반품 포함 전체 LOTTEON 주문행'],
    ['더망고 취소/반품 주문행수', diag.cancelled, pct(diag.cancelled, lotteon), '정보성: 더망고주문상태 기준 취소/반품/교환'],
    ['더망고 취소/반품 매출액', diag.cancelledRevenue || 0, pct(diag.cancelledRevenue || 0, totalRevenueAttempted), '정보성: 주 분석 매출에는 포함'],
    ['사이트상품번호 매칭 행수', diag.bySiteProductId, pct(diag.bySiteProductId, included), '1순위 매칭'],
    ['브랜드+상품명 매칭 행수', diag.byBrandProductName, pct(diag.byBrandProductName, included), '2순위 매칭'],
    ['브랜드명 직접 매칭 행수', diag.byBrandDirect, pct(diag.byBrandDirect, included), '3순위 매칭'],
    ['브랜드명_매칭표 매칭 행수', diag.byBrandAlias, pct(diag.byBrandAlias, included), '4순위 매칭'],
    ['상품명 브랜드추정 매칭 행수', diag.byBrandText || 0, pct(diag.byBrandText || 0, included), '5순위 매칭: 상품명/마켓상품명 안의 브랜드명 추정'],
    ['미매칭 행수', diag.unmatched, pct(diag.unmatched, included), '브랜드명_매칭표에 추가 필요'],
    ['상품매칭 매출액', diag.revenueProductMatched, pct(diag.revenueProductMatched, totalRevenueAttempted), '1+2순위 매출 합계'],
    ['브랜드만 매칭 매출액', diag.revenueBrandOnly, pct(diag.revenueBrandOnly, totalRevenueAttempted), '3+4+5순위 매출 합계'],
    ['미매칭 매출액', diag.revenueUnmatched, pct(diag.revenueUnmatched, totalRevenueAttempted), '미매칭 매출 합계'],
    ['분석 반영 매출액', diag.revenueAnalyzed, pct(diag.revenueAnalyzed, totalRevenueAttempted), '취소/반품 포함 전체 분석 반영 매출 합계']
  ];

  replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.MATCH_DIAG), CONFIG.HEADERS.MATCH_DIAG, rows);
}


function applyRawBrandSalesFallbackToBrandMetrics_(brandMap, sales) {
  const rawMap = {};

  (sales || []).forEach(function(s) {
    if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y') return;

    const rawBrand = String(s['브랜드명_원본'] || s['브랜드명_매칭'] || '').trim();
    if (!rawBrand || isGenericInvalidBrand_(rawBrand)) return;

    const brand = canonicalBrandNameFromKnownSources_(rawBrand);
    if (!brand) return;

    if (!rawMap[brand]) {
      rawMap[brand] = {
        brand: brand,
        orderCount: 0,
        revenue: 0,
        soldKeys: {},
        latestMonth: '',
        latestMonthOrders: 0,
        latestMonthRevenue: 0,
        latestMonthSoldKeys: {},
        cancelledOrders: 0,
        cancelledRevenue: 0
      };
    }

    const qty = toNumber_(s['결제수량']) || 1;
    const revenue = toNumber_(s['결제금액합계']);
    const month = String(s['주문월'] || '').trim();
    const productKey = String(s['사이트상품번호'] || '').trim() ||
      String(s['마켓상품번호'] || '').trim() ||
      normalizeProductNameKey_(String(s['원문상품명'] || s['마켓상품명'] || ''));

    const x = rawMap[brand];
    x.orderCount += qty;
    x.revenue += revenue;
    if (String(s['취소여부'] || '') === 'Y') {
      x.cancelledOrders += qty;
      x.cancelledRevenue += revenue;
    }
    if (productKey) x.soldKeys[productKey] = true;

    if (month && (!x.latestMonth || month > x.latestMonth)) {
      x.latestMonth = month;
      x.latestMonthOrders = 0;
      x.latestMonthRevenue = 0;
      x.latestMonthSoldKeys = {};
    }

    if (month && month === x.latestMonth) {
      x.latestMonthOrders += qty;
      x.latestMonthRevenue += revenue;
      if (productKey) x.latestMonthSoldKeys[productKey] = true;
    }
  });

  Object.keys(rawMap).forEach(function(brand) {
    const raw = rawMap[brand];
    if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, '', '', '');
    const b = brandMap[brand];

    // 원본 브랜드 기준 매출액이 기존 매칭 매출보다 크면 원본 기준으로 보정합니다.
    // 이미 상품매칭으로 잡힌 매출은 유지하고, 부족분만 상품미매칭 매출로 추가 반영합니다.
    if (raw.revenue > toNumber_(b.revenue)) {
      const prevRevenue = toNumber_(b.revenue);
      const diffRevenue = raw.revenue - prevRevenue;
      b.revenue = raw.revenue;
      b.revenueProductUnmatched = Math.max(toNumber_(b.revenueProductUnmatched), toNumber_(b.revenueProductUnmatched) + diffRevenue);
    }

    if (raw.orderCount > toNumber_(b.orderCount)) {
      const prevOrders = toNumber_(b.orderCount);
      const diffOrders = raw.orderCount - prevOrders;
      b.orderCount = raw.orderCount;
      b.ordersProductUnmatched = Math.max(toNumber_(b.ordersProductUnmatched), toNumber_(b.ordersProductUnmatched) + diffOrders);
    }

    const rawSoldCount = Object.keys(raw.soldKeys || {}).length;
    if (rawSoldCount > toNumber_(b.soldProductCount)) b.soldProductCount = rawSoldCount;
    b.cancelledOrderCount = Math.max(toNumber_(b.cancelledOrderCount), toNumber_(raw.cancelledOrders));
    b.cancelledRevenue = Math.max(toNumber_(b.cancelledRevenue), toNumber_(raw.cancelledRevenue));

    if (raw.latestMonth && (!b.latestMonth || raw.latestMonth >= b.latestMonth)) {
      b.latestMonth = raw.latestMonth;
      b.latestMonthOrders = Math.max(toNumber_(b.latestMonthOrders), raw.latestMonthOrders);
      b.latestMonthRevenue = Math.max(toNumber_(b.latestMonthRevenue), raw.latestMonthRevenue);
      b.latestMonthSoldProducts = Math.max(toNumber_(b.latestMonthSoldProducts), Object.keys(raw.latestMonthSoldKeys || {}).length);
    }

    b.rawBrandSalesFallbackApplied = true;
  });
}

function canonicalBrandNameFromKnownSources_(brandRaw) {
  const raw = String(brandRaw || '').trim();
  if (!raw || isGenericInvalidBrand_(raw)) return '';

  const key = normalizeBrandKey_(raw);

  // 브랜드명_매칭표 우선 적용
  try {
    const aliasRows = readTable_(getSheet_(CONFIG.SHEETS.BRAND_ALIAS));
    for (var i = 0; i < aliasRows.length; i++) {
      const src = normalizeBrandKey_(aliasRows[i]['매출브랜드명']);
      const dst = String(aliasRows[i]['더망고브랜드명'] || '').trim();
      if (src && dst && src === key) return dst;
    }
  } catch (e) {}

  // 필터별_상품수 / 수동입력에 동일 정규화 브랜드가 있으면 그 표기를 사용
  const sources = [CONFIG.SHEETS.FILTERS, CONFIG.SHEETS.COUPANG_SENT_MANUAL];
  for (var s = 0; s < sources.length; s++) {
    try {
      const rows = readTable_(getSheet_(sources[s]));
      for (var r = 0; r < rows.length; r++) {
        const f = String(getObjectValueByHeader_(rows[r], '검색필터명') || '').trim();
        const b = String(getObjectValueByHeader_(rows[r], '브랜드명') || '').trim() || brandFromFilterName_(f);
        if (b && normalizeBrandKey_(b) === key) return b;
      }
    } catch (e) {}
  }

  return raw;
}

function buildProductMetrics_(products, sales) {
  const map = {};
  products.forEach(function(p) { const productId = String(p['상품ID'] || '').trim(); const key = productId || String(p['상품명'] || '') + '|' + String(p['검색필터명'] || ''); map[key] = { productId: productId, productName: p['상품명'], brand: p['브랜드명'], filterName: p['검색필터명'], accountNo: p['계정번호'], accountId: p['쿠팡계정ID'], baseDate: p['기준등록일'], lotteonCode: p['롯데온상품번호_추출'], sellDays: toNumber_(p['판매기간일수']), orderCount: 0, revenue: 0, sold: false, latestMonth: '' }; });
  sales.forEach(function(s) { if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y') return; const method = String(s['매칭방식'] || ''); if (method !== 'siteProductId' && method !== 'brandProductName') return; const pid = String(s['매칭상품ID'] || '').trim(); if (!pid || !map[pid]) return; const qty = toNumber_(s['결제수량']) || 1, revenue = toNumber_(s['결제금액합계']); map[pid].orderCount += qty; map[pid].revenue += revenue; map[pid].sold = true; const m = String(s['주문월'] || ''); if (m && (!map[pid].latestMonth || m > map[pid].latestMonth)) map[pid].latestMonth = m; });
  return map;
}

function buildBrandMetrics_(products, productMetrics, sales) {
  const brandMap = {};
  products.forEach(function(p) { const brand = String(p['브랜드명'] || '').trim() || '브랜드미확인'; if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, p['검색필터명'], p['계정번호'], p['쿠팡계정ID']); const b = brandMap[brand]; b.productCount++; if (p['기준등록일']) b.baseDates.push(p['기준등록일']); if (p['API_마켓전송일']) { b.sendDates.push(p['API_마켓전송일']); b.apiSendCount++; } if (p['마켓전송일_수동']) { b.sendDates.push(p['마켓전송일_수동']); b.manualSendCount++; } });
  Object.keys(productMetrics).forEach(function(k) { const pm = productMetrics[k], brand = pm.brand || '브랜드미확인'; if (brandMap[brand] && pm.sold) brandMap[brand].soldProductCount++; });
  sales.forEach(function(s) { if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y') return; const method = String(s['매칭방식'] || ''); if (method === '미매칭') return; const brand = String(s['브랜드명_매칭'] || '').trim(); if (!brand) return; if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, '', '', ''); const b = brandMap[brand]; const qty = toNumber_(s['결제수량']) || 1, revenue = toNumber_(s['결제금액합계']), isProductMatched = (method === 'siteProductId' || method === 'brandProductName'); b.orderCount += qty; b.revenue += revenue; if (String(s['취소여부'] || '') === 'Y') { b.cancelledOrderCount += qty; b.cancelledRevenue += revenue; } if (isProductMatched) { b.revenueProductMatched += revenue; b.ordersProductMatched += qty; } else { b.revenueProductUnmatched += revenue; b.ordersProductUnmatched += qty; } const m = String(s['주문월'] || ''); if (m && (!b.latestMonth || m > b.latestMonth)) b.latestMonth = m; });
  sales.forEach(function(s) { if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y' || s['매칭방식'] === '미매칭') return; const b = brandMap[String(s['브랜드명_매칭'] || '').trim()]; if (!b || !b.latestMonth || s['주문월'] !== b.latestMonth) return; b.latestMonthOrders += toNumber_(s['결제수량']) || 1; b.latestMonthRevenue += toNumber_(s['결제금액합계']); });
  Object.keys(productMetrics).forEach(function(k) { const pm = productMetrics[k], b = brandMap[pm.brand || '브랜드미확인']; if (b && b.latestMonth && pm.latestMonth === b.latestMonth && pm.sold) b.latestMonthSoldProducts++; });

  // v5.64:
  // 매출데이터_정리의 브랜드명_원본 기준 매출액을 브랜드 성과에 강제 보정합니다.
  // 상품목록/매칭컨텍스트가 오래되어도 매출데이터에 브랜드가 있으면 핵심요약 0원으로 빠지지 않게 하기 위함입니다.
  applyRawBrandSalesFallbackToBrandMetrics_(brandMap, sales);

  applyFastFilterSummaryToBrandMetrics_(brandMap);
  applyCoupangSentManualToBrandMetrics_(brandMap);
  applyFilterRawDatesToBrandMetrics_(brandMap);
  applyCleanupRefreshLogToBrandMetrics_(brandMap);
  applyRetransmitLogToBrandMetrics_(brandMap);
  applyAdditionalCollectionLogToBrandMetrics_(brandMap);
  applyRetransmitSalesToBrandMetrics_(brandMap, sales);
  applyAdditionalCollectionSalesToBrandMetrics_(brandMap, sales);
  return brandMap;
}

function createBrandMetric_(brand, filterName, accountNo, accountId) {
  return { brand: brand, filterName: filterName, accountNo: accountNo, accountId: accountId, productCount: 0, soldProductCount: 0, orderCount: 0, revenue: 0, cancelledOrderCount: 0, cancelledRevenue: 0, revenueProductMatched: 0, revenueProductUnmatched: 0, ordersProductMatched: 0, ordersProductUnmatched: 0, baseDates: [], sendDates: [], latestMonth: '', latestMonthSoldProducts: 0, latestMonthOrders: 0, latestMonthRevenue: 0, apiSendCount: 0, manualSendCount: 0, coupangSentCount: 0, manualCoupangSentCount: 0, representativeCoupangSendDate: '', coupangSentStatus: '전송정보 미확인', manualCoupangAccountCounts: {}, filterRecentCollectDate: '', filterCreateDate: '', filterRecentCollectField: '', filterCreateField: '', retransmitDate: '', retransmitOperationDate: '', retransmitType: '', retransmitBeforeSentCount: 0, retransmitAfterSentCount: 0, retransmitReason: '', retransmitMemo: '', retransmitBeforeRevenue: 0, retransmitAfterRevenue: 0, retransmitBeforeOrders: 0, retransmitAfterOrders: 0, retransmitBeforeDailyRevenue: 0, retransmitAfterDailyRevenue: 0, retransmitEffect: '', additionalCollectionDate: '', additionalCollectionOperationDate: '', additionalCollectionType: '', additionalBeforeMangoCount: 0, additionalAfterMangoCount: 0, additionalBeforeSentCount: 0, additionalAfterSentCount: 0, additionalAddedCount: 0, additionalCollectionReason: '', additionalCollectionMemo: '', additionalBeforeRevenue: 0, additionalAfterRevenue: 0, additionalBeforeOrders: 0, additionalAfterOrders: 0, additionalBeforeDailyRevenue: 0, additionalAfterDailyRevenue: 0, additionalCollectionEffect: '', cleanupRefreshDate: '', cleanupRefreshType: '', cleanupRefreshFinalMangoCount: 0, cleanupRefreshFinalSentCount: 0, cleanupRefreshReason: '', cleanupRefreshMemo: '' };
}

function readCoupangSentManualMap_() {
  const rows = readTable_(getSheet_(CONFIG.SHEETS.COUPANG_SENT_MANUAL));
  const byBrand = {}, byFilter = {}, byBrandAll = {}, byFilterAll = {}, allRows = [];
  rows.forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim() || brandFromFilterName_(filterName);
    if (!filterName && !brand) return;
    const statusRaw = String(getObjectValueByHeader_(r, '확인일') || '').trim();
    const statusType = classifyCoupangManualStatus_(statusRaw);
    const c1021 = toNumber_(getObjectValueByHeader_(r, 'beliun1021'));
    const c1023 = toNumber_(getObjectValueByHeader_(r, 'beliun1023'));
    const c1024 = toNumber_(getObjectValueByHeader_(r, 'beliun1024'));
    let total = toNumber_(getObjectValueByHeader_(r, '쿠팡전체전송수'));
    const sumAccounts = c1021 + c1023 + c1024;
    if (!total && sumAccounts) total = sumAccounts;
    const mangoRaw = getObjectValueByHeader_(r, '더망고 수집수') || getObjectValueByHeader_(r, '더망고수집수') || getObjectValueByHeader_(r, '더망고등록상품수');
    const mangoCount = toNumber_(mangoRaw);
    const accountFromFilter = accountFromFilterName_(filterName);
    const inferred = inferAccountFromManualCounts_({ beliun1021: c1021, beliun1023: c1023, beliun1024: c1024 });
    const item = { confirmDate: statusType.isDate ? normalizeDateText_(statusRaw) : '', statusRaw: statusRaw, statusType: statusType.type, isAdded: statusType.type === 'ADD', isDeleted: statusType.type === 'DELETE', isMoved: statusType.type === 'MOVED', inactive: statusType.inactive, filterName: filterName, brand: brand, total: total, hasTotal: String(getObjectValueByHeader_(r, '쿠팡전체전송수') || '').trim() !== '' || sumAccounts > 0, mangoCount: mangoCount, hasMangoCount: String(mangoRaw || '').trim() !== '', accountCounts: { beliun1021: c1021, beliun1023: c1023, beliun1024: c1024 }, accountId: inferred.accountId || accountFromFilter.accountId || '', accountNo: inferred.accountNo || accountFromFilter.accountNo || '', sendDate: normalizeDateText_(getObjectValueByHeader_(r, '대표쿠팡전송일')), memo: String(getObjectValueByHeader_(r, '확인메모') || '').trim() };
    allRows.push(item);
    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
    if (brandKey) byBrandAll[brandKey] = item;
    if (filterName) byFilterAll[String(filterName).trim()] = item;
    if (item.inactive) return;
    if (brandKey) byBrand[brandKey] = item;
    if (filterName) byFilter[String(filterName).trim()] = item;
  });
  return { byBrand: byBrand, byFilter: byFilter, byBrandAll: byBrandAll, byFilterAll: byFilterAll, rows: allRows };
}

function classifyCoupangManualStatus_(raw) {
  const s = String(raw || '').trim();
  if (!s) return { type: 'NORMAL', inactive: false, isDate: false };
  if (s.indexOf('삭제') >= 0) return { type: 'DELETE', inactive: true, isDate: false };
  if (/\d+\s*로\s*변경/.test(s) || s.indexOf('변경') >= 0) return { type: 'MOVED', inactive: true, isDate: false };
  if (s.indexOf('추가') >= 0) return { type: 'ADD', inactive: false, isDate: false };
  return { type: 'NORMAL', inactive: false, isDate: !!normalizeDateText_(s) };
}

function applyFastFilterSummaryToBrandMetrics_(brandMap) {
  const filterRows = readTable_(getSheet_(CONFIG.SHEETS.FILTERS));
  if (!filterRows || !filterRows.length) return;
  filterRows.forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    if (!isValidLotteonFilterName_(filterName)) return;
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim() || brandFromFilterName_(filterName) || '브랜드미확인';
    const account = accountFromFilterName_(filterName);
    const accountNo = getObjectValueByHeader_(r, '계정번호') || account.accountNo || '';
    const accountId = getObjectValueByHeader_(r, '쿠팡계정ID') || account.accountId || '';
    const totalRaw = getObjectValueByHeader_(r, 'API_totalCount');
    const totalCount = String(totalRaw || '').trim() !== '' ? toNumber_(totalRaw) : 0;
    const recentDate = normalizeDateText_(getObjectValueByHeader_(r, 'API_최근수집일자'));
    const createDate = normalizeDateText_(getObjectValueByHeader_(r, 'API_필터생성일'));
    if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, filterName, accountNo, accountId);
    const b = brandMap[brand];
    // 필터별_상품수를 단일 원천으로: 최신 필터명/계정번호를 대표값으로 갱신
    b.filterName = filterName;
    b.accountNo = accountNo;
    b.accountId = accountId;
    if (String(totalRaw || '').trim() !== '') b.productCount = totalCount;
    if (recentDate) b.filterRecentCollectDate = recentDate;
    if (createDate) b.filterCreateDate = createDate;
    if (getObjectValueByHeader_(r, 'API_최근수집일자_필드')) b.filterRecentCollectField = getObjectValueByHeader_(r, 'API_최근수집일자_필드');
    if (getObjectValueByHeader_(r, 'API_필터생성일_필드')) b.filterCreateField = getObjectValueByHeader_(r, 'API_필터생성일_필드');
  });
}

function applyCoupangSentManualToBrandMetrics_(brandMap) {
  const manual = readCoupangSentManualMap_();
  (manual.rows || []).forEach(function(item) {
    if (!item || item.inactive) return;
    if (!item.filterName && !item.brand) return;
    const brand = item.brand || brandFromFilterName_(item.filterName) || '브랜드미확인';
    if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, item.filterName, item.accountNo || '', item.accountId || '');
  });
  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand];
    const item = manual.byFilter[String(b.filterName || '').trim()] || manual.byBrand[normalizeCoupangWorkLogBrandMergeKey_(brand, b.filterName)] || manual.byBrand[normalizeBrandKey_(brand)];
    if (!item || item.inactive) return;
    if (item.hasTotal) { b.manualCoupangSentCount = item.total || 0; b.coupangSentCount = item.total || 0; } else { b.manualCoupangSentCount = item.total || 0; b.coupangSentCount = item.total || b.coupangSentCount || 0; }
    if (item.hasMangoCount && item.mangoCount > 0) b.productCount = item.mangoCount;
    b.representativeCoupangSendDate = item.sendDate || b.representativeCoupangSendDate || '';
    b.manualCoupangAccountCounts = item.accountCounts || {};
    if (item.filterName) b.filterName = item.filterName;
    if (item.accountId) { b.accountId = item.accountId; b.accountNo = item.accountNo || b.accountNo; }
    b.coupangSentStatus = b.coupangSentCount > 0 ? (b.representativeCoupangSendDate ? '수동입력 전송수+전송일 확인' : '수동입력 전송수 확인/전송일 미확인') : '수동입력 전송수 0';
  });
}

function inferAccountFromManualCounts_(counts) {
  const arr = [{ id: 'beliun1021', no: 1, count: toNumber_(counts.beliun1021) }, { id: 'beliun1023', no: 3, count: toNumber_(counts.beliun1023) }, { id: 'beliun1024', no: 2, count: toNumber_(counts.beliun1024) }].filter(function(x) { return x.count > 0; });
  arr.sort(function(a, b) { return b.count - a.count; });
  if (!arr.length) return { accountId: '', accountNo: '' };
  if (arr.length === 1) return { accountId: arr[0].id, accountNo: arr[0].no };
  return { accountId: arr.map(function(x) { return x.id; }).join(','), accountNo: arr.map(function(x) { return x.no; }).join(',') };
}

function getCoupangBasis_(b) {
  const baseDate = minDateText_(b.baseDates);
  const mangoDays = daysSince_(baseDate);
  const cleanupRefreshDate = normalizeDateText_(b.cleanupRefreshDate);
  const retransmitDate = normalizeDateText_(b.retransmitDate);
  const manualOrProductSendDate = normalizeDateText_(b.representativeCoupangSendDate) || minDateText_(b.sendDates);
  const firstSendDate = minDateText_([normalizeDateText_(b.representativeCoupangSendDate)].concat((b.sendDates || []).map(normalizeDateText_).filter(Boolean)));
  const recentCollectDate = normalizeDateText_(b.filterRecentCollectDate);
  const filterCreateDate = normalizeDateText_(b.filterCreateDate);
  let representativeSendDate = '', dateBasis = '';
  if (cleanupRefreshDate) { representativeSendDate = cleanupRefreshDate; dateBasis = '정리재수집_로그'; }
  else if (retransmitDate) { representativeSendDate = retransmitDate; dateBasis = '쿠팡재전송_로그'; }
  else if (manualOrProductSendDate) { representativeSendDate = manualOrProductSendDate; dateBasis = b.coupangSentStatus && b.coupangSentStatus.indexOf('수동입력') >= 0 ? '수동입력_대표쿠팡전송일' : '상품/API_마켓전송일'; }
  else if (recentCollectDate) { representativeSendDate = recentCollectDate; dateBasis = 'API_최근수집일자_추정'; }
  else if (filterCreateDate) { representativeSendDate = filterCreateDate; dateBasis = 'API_필터생성일_추정'; }
  else { dateBasis = '전송일/필터일자 미확인'; }
  const coupangDays = representativeSendDate ? daysSince_(representativeSendDate) : 0;
  const sentCount = toNumber_(b.coupangSentCount) || (toNumber_(b.apiSendCount) + toNumber_(b.manualSendCount));
  let sentStatus = '전송정보 미확인';
  if (sentCount > 0) {
    if (cleanupRefreshDate) sentStatus = '정리재수집 기준 판매기간';
    else if (retransmitDate) sentStatus = '재전송로그 기준 판매기간';
    else if (manualOrProductSendDate) sentStatus = b.coupangSentStatus || '전송수+전송일 확인';
    else if (recentCollectDate) sentStatus = '전송수 확인/API_최근수집일자 기준 판매기간 추정';
    else if (filterCreateDate) sentStatus = '전송수 확인/API_필터생성일 기준 판매기간 추정';
    else sentStatus = '전송수 확인/전송일 미확인';
  } else { if (representativeSendDate) sentStatus = dateBasis + '/전송수 미확인'; }
  const sellRate = sentCount ? b.soldProductCount / sentCount : 0;
  const revenuePerProduct = sentCount ? b.revenue / sentCount : 0;
  const dailyRevenue = coupangDays ? b.revenue / coupangDays : 0;
  const revenue30d = coupangDays ? dailyRevenue * 30 : 0;
  return { baseDate: baseDate, mangoDays: mangoDays, representativeSendDate: representativeSendDate, firstSendDate: firstSendDate, coupangDays: coupangDays, sentCount: sentCount, sentStatus: sentStatus, sellRate: sellRate, revenuePerProduct: revenuePerProduct, dailyRevenue: dailyRevenue, revenue30d: revenue30d, recentCollectDate: recentCollectDate, filterCreateDate: filterCreateDate, dateBasis: dateBasis, retransmitDate: retransmitDate };
}

function judgeBrandByCoupang_(b, basis) {
  if (!basis.sentCount || !basis.representativeSendDate) return { label: '전송정보 확인필요', reason: '쿠팡 전송수 또는 대표쿠팡전송일이 없어 판매기간/전송수 기준 판정 보류' };
  return judgeBrand_(b, basis.coupangDays, basis.sellRate, basis.revenue30d);
}

function buildBrandRows_(brandMetrics) {
  const rows = [];
  Object.keys(brandMetrics).sort().forEach(function(brand) {
    const b = brandMetrics[brand];
    const basis = getCoupangBasis_(b);
    const decision = judgeBrandByCoupang_(b, basis);
    const totalRevenueAll = b.revenueProductMatched + b.revenueProductUnmatched;
    const totalOrdersAll = b.ordersProductMatched + b.ordersProductUnmatched;
    const productMatchRate = totalOrdersAll ? b.ordersProductMatched / totalOrdersAll : 0;
    const revenueMatchRate = totalRevenueAll ? b.revenueProductMatched / totalRevenueAll : 0;
    rows.push([b.brand, b.filterName, b.accountNo, b.accountId, basis.baseDate, basis.representativeSendDate, basis.firstSendDate, basis.mangoDays, basis.coupangDays, b.productCount, basis.sentCount, basis.sentStatus, b.soldProductCount, basis.sellRate, b.orderCount, b.revenue, b.latestMonth, b.latestMonthSoldProducts, b.latestMonthOrders, b.latestMonthRevenue, basis.revenuePerProduct, basis.dailyRevenue, basis.revenue30d, decision.label, decision.reason, b.revenueProductMatched, b.revenueProductUnmatched, b.ordersProductMatched, b.ordersProductUnmatched, productMatchRate, revenueMatchRate, b.cleanupRefreshDate ? 'Y' : 'N', b.cleanupRefreshDate, b.cleanupRefreshDate ? daysSince_(b.cleanupRefreshDate) : '', makeCleanupRefreshMemo_(b), b.retransmitDate ? 'Y' : 'N', b.retransmitDate, b.retransmitDate ? daysSince_(b.retransmitDate) : '', b.retransmitBeforeDailyRevenue, b.retransmitAfterDailyRevenue, b.retransmitEffect, makeRetransmitMemo_(b), b.additionalCollectionDate ? 'Y' : 'N', b.additionalCollectionDate, b.additionalCollectionDate ? daysSince_(b.additionalCollectionDate) : '', b.additionalAddedCount, b.additionalBeforeDailyRevenue, b.additionalAfterDailyRevenue, b.additionalCollectionEffect, makeAdditionalCollectionMemo_(b), b.cancelledOrderCount || 0, b.cancelledRevenue || 0]);
  });
  rows.sort(function(a, b) { return toNumber_(b[21]) - toNumber_(a[21]); });
  return rows;
}


function buildBrandSummaryRowsFromBrandRows_(brandRows) {
  const headers = CONFIG.HEADERS.BRAND;
  const logMap = readRetransmitLogDateMap_();

  const objects = (brandRows || []).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return applyRetransmitLogDatesToBrandObj_(obj, logMap);
  }).filter(function(obj) { return String(obj['브랜드명'] || '').trim() !== ''; });

  const rows = objects.map(function(b) {
    const classification = classifyBrandSummary_(b);
    const priority = brandSummaryPriority_(classification, b);
    const action = recommendBrandSummaryAction_(classification, b);
    const memo = makeBrandSummaryMemo_(classification, b);
    const salesBasis = summarizeSalesBasis_(b['전송일확인상태']);
    const rotation = buildProductRotationPlan_(b, classification);

    return [
      priority,
      classification,
      b['브랜드명'],
      b['대표검색필터명'],
      b['쿠팡계정ID'],
      toNumber_(b['쿠팡전송확인상품수']),
      toNumber_(b['전체_매출상품수']),
      toNumber_(b['쿠팡전송수대비_매출상품률']),
      toNumber_(b['전체_주문건수']),
      toNumber_(b['전체_매출액']),
      toNumber_(b['쿠팡전송후판매기간일수']),
      toNumber_(b['30일환산매출_쿠팡전송기준']),
      salesBasis,
      action,
      memo,
      rotation.nextAction,
      rotation.stage1,
      rotation.stage2,
      rotation.stage3,
      rotation.memo,
      b['정리재수집여부'] || 'N',
      formatShortDate_(b['최근정리재수집일']),
      b['정리재수집메모'] || '',
      b['재전송여부'] || 'N',
      formatShortDate_(b['최근재전송일']),
      summarizeEffect_(b['재전송효과']),
      b['재전송메모'] || '',
      b['추가수집여부'] || 'N',
      formatShortDate_(b['최근추가수집일']),
      summarizeEffect_(b['추가수집효과']),
      b['추가수집메모'] || '',
      toNumber_(b['취소수량']),
      toNumber_(b['취소매출액'])
    ];
  });

  rows.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] - b[0];
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]));
    if (toNumber_(b[11]) !== toNumber_(a[11])) return toNumber_(b[11]) - toNumber_(a[11]);
    if (toNumber_(b[8]) !== toNumber_(a[8])) return toNumber_(b[8]) - toNumber_(a[8]);
    return toNumber_(b[9]) - toNumber_(a[9]);
  });
  return rows;
}

function summarizeSalesBasis_(status) {
  const s = String(status || '');
  if (!s) return '';
  if (s.indexOf('재전송') >= 0 || s.indexOf('쿠팡재전송') >= 0) return '재전송';
  if (s.indexOf('수동입력') >= 0) return '수동';
  if (s.indexOf('API_마켓전송일') >= 0 || s.indexOf('상품/API') >= 0) return 'API전송';
  if (s.indexOf('최근수집') >= 0) return '수집일추정';
  if (s.indexOf('필터생성') >= 0) return '필터일추정';
  if (s.indexOf('미확인') >= 0) return '미확인';
  return s.length > 10 ? s.slice(0, 10) : s;
}

function formatShortDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'M/dd');
  const s = normalizeDateText_(v);
  const m = String(s || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return Number(m[2]) + '/' + pad2_(Number(m[3]));
  const md = String(v || '').trim().match(/^(\d{1,2})[\/\.\-\s]+(\d{1,2})$/);
  if (md) return Number(md[1]) + '/' + pad2_(Number(md[2]));
  return String(v || '').trim();
}

function pad2_(n) { n = Number(n); return n < 10 ? '0' + n : String(n); }

function summarizeEffect_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  if (s.indexOf('7일미만') >= 0) return '7일미만';
  if (s.indexOf('판단보류') >= 0) return '보류';
  if (s.indexOf('개선') >= 0) return '개선';
  if (s.indexOf('악화') >= 0) return '악화';
  if (s.indexOf('변화없음') >= 0) return '변화없음';
  return s.length > 8 ? s.slice(0, 8) : s;
}

function buildProductRotationPlan_(b, classification) {
  const sentCount = toNumber_(b['쿠팡전송확인상품수']);
  const soldProducts = toNumber_(b['전체_매출상품수']);
  const unsoldCount = Math.max(0, sentCount - soldProducts);

  // v5.98:
  // 상품갈이 기준일 우선순위
  // 1) 최근 정리재수집일
  // 2) 최근 재전송일
  // 3) 최근 추가수집일
  // 4) 대표/최초 전송일 또는 필터 생성일
  const baseDate = normalizeDateForComparison_(
    b['최근정리재수집일'] ||
    b['최근재전송일'] ||
    b['최근추가수집일'] ||
    b['대표쿠팡전송일'] ||
    b['최초쿠팡전송일'] ||
    b['더망고기준등록일']
  );

  const stage1Base = normalizeDateForComparison_(b['최근추가수집일'] || b['최근정리재수집일'] || baseDate);
  const stage2Base = normalizeDateForComparison_(b['최근재전송일'] || b['최근정리재수집일'] || baseDate);
  const stage3Base = normalizeDateForComparison_(b['최근정리재수집일'] || baseDate);

  const stage1 = calcRotationStatus_(stage1Base, CONFIG.ROTATION_STAGE1_DAYS, true, false, sentCount);
  const stage2 = calcRotationStatus_(stage2Base, CONFIG.ROTATION_STAGE2_DAYS, unsoldCount > 0, true, sentCount);
  const stage3 = calcRotationStatus_(stage3Base, CONFIG.ROTATION_STAGE3_DAYS, unsoldCount > 0, true, sentCount);

  let nextAction = '대기';
  if (sentCount <= 0) nextAction = '전송확인';
  else if (stage3 === '필요') nextAction = '3단계';
  else if (stage2 === '필요') nextAction = '2단계';
  else if (stage1 === '필요') nextAction = '1단계';

  return {
    nextAction: nextAction,
    stage1: stage1,
    stage2: stage2,
    stage3: stage3,
    memo: makeRotationMemo_(nextAction, unsoldCount, baseDate)
  };
}

function calcRotationStatus_(baseDate, cycleDays, hasTarget, needsUnsold, sentCount) {
  if (sentCount <= 0) return '전송확인';
  if (needsUnsold && !hasTarget) return '대상없음';
  const d = normalizeDateForComparison_(baseDate);
  if (!d) return '기준없음';
  const elapsed = daysSince_(d);
  if (elapsed >= cycleDays) return '필요';
  return 'D-' + (cycleDays - elapsed);
}

function makeRotationMemo_(nextAction, unsoldCount, baseDate) {
  if (nextAction === '전송확인') return '전송수 확인';
  const parts = [];
  if (baseDate) parts.push('기준 ' + formatShortDate_(baseDate));
  if (unsoldCount > 0) parts.push('미판매 ' + formatCount_(unsoldCount));
  if (nextAction !== '대기') parts.push(nextAction + ' 필요');
  else parts.push('대기');
  return parts.join(' / ');
}

function classifyBrandSummary_(b) {
  const sentCount = toNumber_(b['쿠팡전송확인상품수']);
  const soldProducts = toNumber_(b['전체_매출상품수']);
  const sellRate = toNumber_(b['쿠팡전송수대비_매출상품률']);
  const orders = toNumber_(b['전체_주문건수']);
  const revenue = toNumber_(b['전체_매출액']);
  const days = toNumber_(b['쿠팡전송후판매기간일수']);
  const revenue30d = toNumber_(b['30일환산매출_쿠팡전송기준']);
  const status = String(b['전송일확인상태'] || '');
  if (!sentCount) return '전송정보 확인';
  if (orders > 0 || revenue > 0 || soldProducts > 0) {
    if (orders >= 10 || sellRate >= 0.01 || revenue >= 800000 || revenue30d >= 1000000) return '확대 우선';
    return '유지/관찰';
  }
  if (days >= 30 && sentCount >= 300) return '삭제/압축 검토';
  if (days >= 20 && sentCount >= 300) return '30일 도달 대기';
  if (status.indexOf('추정') >= 0) return '무매출 관찰';
  return '무매출 관찰';
}

function brandSummaryPriority_(classification) { if (classification === '확대 우선') return 10; if (classification === '삭제/압축 검토') return 20; if (classification === '30일 도달 대기') return 30; if (classification === '유지/관찰') return 40; if (classification === '전송정보 확인') return 50; if (classification === '무매출 관찰') return 60; return 90; }
function recommendBrandSummaryAction_(classification) { if (classification === '확대 우선') return '확대'; if (classification === '유지/관찰') return '관찰'; if (classification === '삭제/압축 검토') return '압축검토'; if (classification === '30일 도달 대기') return '30일대기'; if (classification === '전송정보 확인') return '전송확인'; if (classification === '무매출 관찰') return '관찰'; return '보류'; }

function makeBrandSummaryMemo_(classification, b) {
  const sentCount = toNumber_(b['쿠팡전송확인상품수']);
  const soldProducts = toNumber_(b['전체_매출상품수']);
  const orders = toNumber_(b['전체_주문건수']);
  const revenue = toNumber_(b['전체_매출액']);
  const days = toNumber_(b['쿠팡전송후판매기간일수']);
  if (classification === '전송정보 확인') return '전송수/계정 확인';
  if (orders > 0 || revenue > 0) return '전송 ' + formatCount_(sentCount) + ' / 매출상품 ' + formatCount_(soldProducts) + ' / 주문 ' + formatCount_(orders) + ' / 매출 ' + formatWonCompact_(revenue);
  if (classification === '삭제/압축 검토') return '전송 ' + formatCount_(sentCount) + ' / ' + formatCount_(days) + '일 / 주문0 / 압축';
  if (classification === '30일 도달 대기') return '전송 ' + formatCount_(sentCount) + ' / ' + formatCount_(days) + '일 / 주문0 / 대기';
  return '전송 ' + formatCount_(sentCount) + ' / ' + formatCount_(days) + '일 / 주문0';
}

function formatWonCompact_(n) { const v = Math.round(toNumber_(n)); if (Math.abs(v) >= 100000000) return '₩' + (v / 100000000).toFixed(1).replace(/\.0$/, '') + '억'; if (Math.abs(v) >= 10000) return '₩' + (v / 10000).toFixed(0) + '만'; return '₩' + formatCount_(v); }

function dashboardGet_(row, names) {
  for (var i = 0; i < (names || []).length; i++) {
    const v = getObjectValueByHeader_(row, names[i]);
    if (v !== '' && v !== null && v !== undefined) return v;
  }
  return '';
}

function dashboardPercentText_(v) {
  if (v === '' || v === null || v === undefined) return '0.00%';
  if (typeof v === 'string' && v.indexOf('%') >= 0) {
    const n = Number(v.replace('%', '').replace(/,/g, '').trim());
    return isNaN(n) ? '0.00%' : n.toFixed(2) + '%';
  }
  const n = toNumber_(v);
  return (n * 100).toFixed(2) + '%';
}

function dashboardCountText_(v) {
  return formatCount_(toNumber_(v));
}

function dashboardMoneyText_(v) {
  return formatWon_(toNumber_(v));
}

function dashboardBrandObjFromSummaryRow_(r) {
  const sent = toNumber_(dashboardGet_(r, ['전송수']));
  const sold = toNumber_(dashboardGet_(r, ['매출\n상품수', '매출상품수']));
  const orders = toNumber_(dashboardGet_(r, ['주문\n건수', '주문건수']));
  const revenue = toNumber_(dashboardGet_(r, ['전체\n매출', '유효\n매출', '전체매출', '유효매출', '매출액', '매출']));
  let rate = dashboardGet_(r, ['매출\n상품률', '매출상품률']);
  if (rate === '' || rate === null || rate === undefined) rate = sent ? sold / sent : 0;

  return {
    priority: toNumber_(dashboardGet_(r, ['우선\n순위', '우선순위'])),
    cls: String(dashboardGet_(r, ['운영\n분류', '운영분류']) || ''),
    brand: String(dashboardGet_(r, ['브랜드명']) || ''),
    filterName: String(dashboardGet_(r, ['대표\n검색필터명', '대표검색필터명']) || ''),
    accountId: String(dashboardGet_(r, ['쿠팡\n계정ID', '쿠팡계정ID']) || ''),
    sent: sent,
    sold: sold,
    rate: rate,
    orders: orders,
    revenue: revenue,
    days: toNumber_(dashboardGet_(r, ['판매\n일수', '판매일수'])),
    action: String(dashboardGet_(r, ['액션']) || ''),
    summaryMemo: String(dashboardGet_(r, ['요약\n메모', '요약메모']) || ''),
    nextAction: String(dashboardGet_(r, ['다음\n상품갈이', '다음상품갈이']) || ''),
    stage1: String(dashboardGet_(r, ['1단계\n추가수집', '1단계추가수집']) || ''),
    stage2: String(dashboardGet_(r, ['2단계\n재전송', '2단계재전송']) || ''),
    stage3: String(dashboardGet_(r, ['3단계\n정리재수집', '3단계정리재수집']) || ''),
    rotationMemo: String(dashboardGet_(r, ['상품갈이\n메모', '상품갈이메모']) || ''),
    cancelOrders: toNumber_(dashboardGet_(r, ['취소\n수량', '취소\n주문', '취소수량', '취소주문'])),
    cancelRevenue: toNumber_(dashboardGet_(r, ['취소\n매출', '취소매출']))
  };
}

function dashboardBrandDisplayRow_(section, x, action, memo) {
  return [
    section,
    x.brand,
    x.filterName,
    x.accountId,
    dashboardCountText_(x.sent),
    dashboardCountText_(x.sold),
    dashboardPercentText_(x.rate),
    dashboardCountText_(x.orders),
    dashboardMoneyText_(x.revenue),
    action || '',
    memo || ''
  ];
}

function dashboardBuildStatsFromCore_(items) {
  const stat = {
    brandCount: 0, sentCount: 0, soldProducts: 0, orders: 0, revenue: 0, sellRate: 0,
    needCheck: 0, expand: 0, observe: 0, compress: 0, wait30: 0,
    stage1Due: 0, stage2Due: 0, stage3Due: 0,
    cancelOrders: 0, cancelRevenue: 0
  };

  (items || []).forEach(function(x) {
    stat.brandCount += 1;
    stat.sentCount += toNumber_(x.sent);
    stat.soldProducts += toNumber_(x.sold);
    stat.orders += toNumber_(x.orders);
    stat.revenue += toNumber_(x.revenue);
    stat.cancelOrders += toNumber_(x.cancelOrders);
    stat.cancelRevenue += toNumber_(x.cancelRevenue);

    if (x.stage1 === '필요') stat.stage1Due += 1;
    if (x.stage2 === '필요') stat.stage2Due += 1;
    if (x.stage3 === '필요') stat.stage3Due += 1;

    if (x.cls === '확대 우선') stat.expand += 1;
    else if (x.cls === '유지/관찰') stat.observe += 1;
    else if (x.cls === '전송정보 확인') stat.needCheck += 1;
    else if (x.cls === '삭제/압축 검토') stat.compress += 1;
    else if (x.cls === '30일 도달 대기') stat.wait30 += 1;
  });

  stat.sellRate = stat.sentCount ? stat.soldProducts / stat.sentCount : 0;
  return stat;
}

function dashboardMatchDiagValue_(itemName) {
  try {
    const rows = readTable_(getSheet_(CONFIG.SHEETS.MATCH_DIAG));
    return auditMetricValue_(rows, itemName);
  } catch (e) {
    return '';
  }
}

function dashboardWrite_(sheet, output) {
  const rows = output.length;
  const cols = output[0].length;
  const clearRows = Math.max(sheet.getLastRow(), rows, 250);
  const clearCols = Math.max(sheet.getLastColumn(), cols, 11);
  try { sheet.getRange(1, 1, clearRows, clearCols).clearContent(); } catch (e) {}
  sheet.getRange(1, 1, rows, cols).setValues(output);
  sheet.setFrozenRows(1);
  formatDashboardDisplayOnly_(sheet, rows, cols);
}

function formatDashboardDisplayOnly_(sheet, rowCount, colCount) {
  if (!sheet) return;
  const rows = rowCount || Math.max(sheet.getLastRow(), 1);
  const cols = colCount || Math.max(sheet.getLastColumn(), 1);

  try {
    sheet.getRange(1, 1, rows, cols)
      .setNumberFormat('@')
      .setVerticalAlignment('middle')
      .setWrap(false);
  } catch (e) {}

  try {
    sheet.getRange(1, 1, 1, cols)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 42);
  } catch (e) {}

  try {
    const vals = sheet.getRange(1, 1, rows, 1).getValues();
    for (var r = 1; r <= rows; r++) {
      const section = String(vals[r - 1][0] || '');
      if (section === '구분') {
        sheet.getRange(r, 1, 1, cols)
          .setFontWeight('bold')
          .setBackground('#e8f0fe')
          .setWrap(true)
          .setHorizontalAlignment('center');
        sheet.setRowHeight(r, 42);
      }
    }
  } catch (e) {}

  const widths = [90, 150, 210, 115, 90, 90, 90, 80, 115, 115, 360];
  widths.forEach(function(w, i) { try { sheet.setColumnWidth(i + 1, w); } catch (e) {} });
}

function buildDashboard_(monthlyBrandRows, brandRows, deleteRows) {
  // v5.98:
  // 대시보드는 핵심_브랜드요약을 단일 원천으로 사용합니다.
  // 기존처럼 brandRows로 다시 요약을 재생성하지 않으므로 핵심요약/검수리포트와 매출액이 어긋나지 않습니다.
  // 또한 모든 수량/비율/금액을 표시 문자열로 만들어 단위 깨짐을 방지합니다.
  const sheet = getSheet_(CONFIG.SHEETS.DASHBOARD);
  try {
    const charts = sheet.getCharts();
    charts.forEach(function(chart) { sheet.removeChart(chart); });
    SpreadsheetApp.flush();
  } catch (e) {}

  const coreRows = readTable_(getSheet_(CONFIG.SHEETS.BRAND_SUMMARY));
  const logMap = readRetransmitLogDateMap_();
  const items = (coreRows || [])
    .map(dashboardBrandObjFromSummaryRow_)
    .map(function(x) { return buildRotationPlanFromDashboardItemWithLog_(x, logMap); })
    .filter(function(x) { return x.brand; });

  const stat = dashboardBuildStatsFromCore_(items);
  const analyzedRevenue = toNumber_(dashboardMatchDiagValue_('분석 반영 매출액'));
  const revenueDiff = analyzedRevenue ? analyzedRevenue - stat.revenue : 0;

  const output = [];
  output.push(['구분','항목','값1','값2','값3','값4','값5','값6','값7','값8','메모']);

  output.push(['요약','갱신기준','핵심_브랜드요약 + 쿠팡재전송_로그 날짜 기준', '', '', '', '', '', '', '', 'v5.98: 상품갈이 날짜 로그연동']);
  output.push(['요약','갱신시각', now_(), '', '', '', '', '', '', '', 'Asia/Seoul']);
  output.push(['요약','분석브랜드수', dashboardCountText_(stat.brandCount), '', '', '', '', '', '', '', '핵심_브랜드요약 기준']);
  output.push(['요약','쿠팡전송상품수', dashboardCountText_(stat.sentCount), '', '', '', '', '', '', '', '수동입력/전송확인 기준']);
  output.push(['요약','매출상품수', dashboardCountText_(stat.soldProducts), '', '', '', '', '', '', '', '주문 발생 상품 수']);
  output.push(['요약','주문건수', dashboardCountText_(stat.orders), '', '', '', '', '', '', '', '취소/반품 포함 전체']);
  output.push(['요약','매출액', dashboardMoneyText_(stat.revenue), '', '', '', '', '', '', '', '핵심_브랜드요약 합계']);
  output.push(['요약','매칭진단 매출액', analyzedRevenue ? dashboardMoneyText_(analyzedRevenue) : '', '', '', '', '', '', '', '', '분석 반영 매출액']);
  output.push(['요약','대시보드-매칭진단 차이', analyzedRevenue ? dashboardMoneyText_(revenueDiff) : '', '', '', '', '', '', '', '', revenueDiff === 0 ? '정상' : '차이 발생 시 ⑧ 재실행 필요']);
  output.push(['요약','전체 매출상품률', dashboardPercentText_(stat.sellRate), '', '', '', '', '', '', '', '매출상품수 ÷ 쿠팡전송상품수']);
  output.push(['요약','취소/반품 수량', dashboardCountText_(stat.cancelOrders), '', '', '', '', '', '', '', '정보성: 더망고주문상태 기준 수량 합계']);
  output.push(['요약','취소/반품 매출', dashboardMoneyText_(stat.cancelRevenue), '', '', '', '', '', '', '', '정보성: 주 분석 매출에는 포함']);
  output.push(['요약','전송정보 확인 필요', dashboardCountText_(stat.needCheck), '', '', '', '', '', '', '', '쿠팡전송수 0 또는 계정 미확인']);
  output.push(['요약','확대 우선', dashboardCountText_(stat.expand), '', '', '', '', '', '', '', '신규계정/주력계정 후보']);
  output.push(['요약','유지/관찰', dashboardCountText_(stat.observe), '', '', '', '', '', '', '', '매출은 있으나 추가 관찰']);
  output.push(['요약','삭제/압축 검토', dashboardCountText_(stat.compress), '', '', '', '', '', '', '', '판매기간 충분 + 무매출']);
  output.push(['요약','30일 도달 대기', dashboardCountText_(stat.wait30), '', '', '', '', '', '', '', '30일 도달 후 재판정']);
  output.push(['요약','1단계 추가수집 필요', dashboardCountText_(stat.stage1Due), '', '', '', '', '', '', '', '주1회 상품 추가수집 대상']);
  output.push(['요약','2단계 재전송 필요', dashboardCountText_(stat.stage2Due), '', '', '', '', '', '', '', '월1회 미판매상품 재전송 대상']);
  output.push(['요약','3단계 정리재수집 필요', dashboardCountText_(stat.stage3Due), '', '', '', '', '', '', '', '2개월1회 미판매상품 완전정리 대상']);

  output.push(['','','','','','','','','','','']);
  output.push(['구분','브랜드명','대표검색필터명','계정ID','전송수','매출상품','매출률','주문','매출','액션','메모']);

  items.filter(function(x) { return x.cls === '확대 우선'; })
    .sort(function(a, b) { return b.orders !== a.orders ? b.orders - a.orders : b.revenue - a.revenue; })
    .slice(0, 15)
    .forEach(function(x) {
      output.push(dashboardBrandDisplayRow_(
        '확대TOP',
        x,
        x.action,
        x.summaryMemo || ('전송 ' + dashboardCountText_(x.sent) + ' / 매출상품 ' + dashboardCountText_(x.sold) + ' / 주문 ' + dashboardCountText_(x.orders) + ' / 매출 ' + dashboardMoneyText_(x.revenue))
      ));
    });

  output.push(['','','','','','','','','','','']);
  output.push(['구분','브랜드명','대표검색필터명','계정ID','전송수','매출상품','매출률','주문','매출','다음작업','상품갈이메모']);

  items.filter(function(x) { return x.stage1 === '필요' || x.stage2 === '필요' || x.stage3 === '필요'; })
    .sort(function(a, b) {
      const aw = rotationWeight_(a.nextAction), bw = rotationWeight_(b.nextAction);
      return aw !== bw ? bw - aw : b.sent - a.sent;
    })
    .slice(0, 25)
    .forEach(function(x) {
      output.push(dashboardBrandDisplayRow_('상품갈이', x, x.nextAction, x.rotationMemo));
    });

  output.push(['','','','','','','','','','','']);
  output.push(['구분','브랜드명','대표검색필터명','계정ID','전송수','매출상품','매출률','주문','매출','액션','메모']);

  items.filter(function(x) {
      return x.cls === '삭제/압축 검토' || x.cls === '30일 도달 대기' || x.cls === '전송정보 확인';
    })
    .sort(function(a, b) { return a.priority !== b.priority ? a.priority - b.priority : b.sent - a.sent; })
    .slice(0, 20)
    .forEach(function(x) {
      output.push(dashboardBrandDisplayRow_('조치필요', x, x.action, x.summaryMemo));
    });

  dashboardWrite_(sheet, output);
}

function refreshCoreSummaryAndDashboardWithRetransmitLogDates() {
  const ui = SpreadsheetApp.getUi();
  try {
    const brandRows = getValuesWithoutHeader_(CONFIG.SHEETS.BRAND);
    const summaryRows = buildBrandSummaryRowsFromBrandRows_(brandRows);
    replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND_SUMMARY), CONFIG.HEADERS.BRAND_SUMMARY, summaryRows);
    buildDashboard_(null, null, null);
    ui.alert(
      '핵심_브랜드요약 + 대시보드 날짜 연동 갱신 완료\n\n' +
      '반영 기준:\n' +
      '1. 쿠팡재전송_로그 3단계_정리재수집일\n' +
      '2. 쿠팡재전송_로그 2단계_재전송일\n' +
      '3. 쿠팡재전송_로그 1단계_추가수집일\n' +
      '4. 쿠팡재전송_로그 최초전송일 / 필터생성일\n\n' +
      '상품갈이 다음작업과 상품갈이메모를 다시 계산했습니다.'
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    ui.alert('핵심_브랜드요약 + 대시보드 날짜 연동 갱신 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function rotationWeight_(nextAction) { const s = String(nextAction || ''); if (s.indexOf('3') >= 0) return 3; if (s.indexOf('2') >= 0) return 2; if (s.indexOf('1') >= 0) return 1; return 0; }

function buildDashboardStatsFromSummary_(summaryRows) {
  const stat = { brandCount: 0, sentCount: 0, soldProducts: 0, orders: 0, revenue: 0, sellRate: 0, expand: 0, observe: 0, needCheck: 0, compress: 0, wait30: 0, retransmit: 0, additionalCollection: 0, stage1Due: 0, stage2Due: 0, stage3Due: 0 };
  (summaryRows || []).forEach(function(r) {
    const cls = String(r[1] || '');
    stat.brandCount += 1; stat.sentCount += toNumber_(r[5]); stat.soldProducts += toNumber_(r[6]); stat.orders += toNumber_(r[8]); stat.revenue += toNumber_(r[9]);
    if (String(r[16] || '') === '필요') stat.stage1Due += 1;
    if (String(r[17] || '') === '필요') stat.stage2Due += 1;
    if (String(r[18] || '') === '필요') stat.stage3Due += 1;
    if (String(r[23] || '') === 'Y') stat.retransmit += 1;
    if (String(r[27] || '') === 'Y') stat.additionalCollection += 1;
    if (cls === '확대 우선') stat.expand += 1;
    else if (cls === '유지/관찰') stat.observe += 1;
    else if (cls === '전송정보 확인') stat.needCheck += 1;
    else if (cls === '삭제/압축 검토') stat.compress += 1;
    else if (cls === '30일 도달 대기') stat.wait30 += 1;
  });
  stat.sellRate = stat.sentCount ? stat.soldProducts / stat.sentCount : 0;
  return stat;
}


function syncCoupangWorkLogMetrics_(brandRows) {
  const sheet = ensureCoupangWorkLogSheet_();
  const headers = CONFIG.HEADERS.RETRANSMIT_LOG;
  const metricMap = buildCoupangWorkLogMetricMap_(brandRows);
  const manual = readCoupangSentManualMap_();
  const existingRows = readTable_(sheet);
  const merged = {};
  const metricKeysUsed = {};

  existingRows.forEach(function(r) {
    if (shouldSkipInactiveCoupangWorkLogRow_(r, manual)) return;
    const metric = getCoupangWorkLogMetricForRow_(r, metricMap);
    const key = getCoupangWorkLogMergeKey_(r, metric);
    if (!key) return;
    const obj = buildCoupangWorkLogOutputObj_(r, metric);
    if (!merged[key]) merged[key] = obj;
    else merged[key] = mergeCoupangWorkLogObjects_(merged[key], obj);
    if (metric && metric.filterName) metricKeysUsed[getCoupangWorkLogMetricKey_(metric)] = true;
  });

  Object.keys(metricMap.byFilter).sort().forEach(function(filterName) {
    const metric = metricMap.byFilter[filterName];
    if (isInactiveManualFilterName_(filterName, manual)) return;
    const key = getCoupangWorkLogMetricKey_(metric);
    if (metricKeysUsed[key] || merged[key]) return;
    merged[key] = buildCoupangWorkLogOutputObj_({}, metric);
  });

  const out = Object.keys(merged).sort(function(a, b) {
    return String(merged[a]['검색필터명'] || '').localeCompare(String(merged[b]['검색필터명'] || ''));
  }).map(function(key) {
    return headers.map(function(h) { return getCoupangWorkLogValueByAlias_(merged[key], h); });
  });

  replaceDataFastLimited_(sheet, headers, out);
  formatCoupangWorkLogSheetFastOnly_(sheet);
}

function patchRetransmitLogDatesFromFilterSummary_() {
  const sheet = ensureCoupangWorkLogSheet_();
  const headers = CONFIG.HEADERS.RETRANSMIT_LOG;
  const rows = readTable_(sheet);
  const filterSummaryMap = buildFilterSummaryMapFromFilterSheet_();
  const out = [];
  let patchedFirst = 0;
  let patchedStage1 = 0;
  let clearedStage1 = 0;
  let missing = 0;

  (rows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim();
    const info = getFilterSummaryDateInfoForLog_(filterName, brand, filterSummaryMap);

    const obj = {};
    headers.forEach(function(h) {
      const key = normalizeHeaderKey_(h);
      obj[key] = getObjectValueByHeader_(r, key);
    });

    if (info && info.hasApiRow) {
      if (info.createDisplay) {
        if (obj['최초전송일'] !== info.createDisplay) patchedFirst++;
        obj['최초전송일'] = info.createDisplay;
      }

      // v5.98:
      // 1단계_추가수집일은 기존 값과 병합하지 않고 API 조건값으로 직접 덮어씁니다.
      // K2처럼 API_최근수집일자 6/01, API_필터생성일 4/24이면 6/01이 들어갑니다.
      const nextStage1 = info.additionalDisplay || '';
      if (nextStage1) {
        if (obj['1단계_추가수집일'] !== nextStage1) patchedStage1++;
      } else if (obj['1단계_추가수집일']) {
        clearedStage1++;
      }
      obj['1단계_추가수집일'] = nextStage1;

      const memo = '필터일자 기준: 생성 ' + (info.createDisplay || '-') + ' / 최근수집 ' + (info.recentDisplay || '-');
      obj['비고'] = mergeShortText_(obj['비고'] || '', memo);
    } else {
      missing++;
    }

    out.push(headers.map(function(h) { return obj[normalizeHeaderKey_(h)] || ''; }));
  });

  replaceDataFastLimited_(sheet, headers, out);
  formatCoupangWorkLogSheetFastOnly_(sheet);

  log_(
    'patch_retransmit_log_dates_v592',
    'rows=' + out.length +
    ', first=' + patchedFirst +
    ', stage1=' + patchedStage1 +
    ', clearedStage1=' + clearedStage1 +
    ', missing=' + missing
  );

  return {
    rows: out.length,
    patchedFirst: patchedFirst,
    patchedStage1: patchedStage1,
    clearedStage1: clearedStage1,
    missing: missing
  };
}

function patchRetransmitLogDatesFromFilterSummary() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = patchRetransmitLogDatesFromFilterSummary_();
    ui.alert(
      '쿠팡재전송_로그 날짜 직접패치 완료\n\n' +
      '처리 행수: ' + result.rows + '\n' +
      '최초전송일 패치: ' + result.patchedFirst + '\n' +
      '1단계_추가수집일 패치: ' + result.patchedStage1 + '\n' +
      '1단계_추가수집일 삭제: ' + result.clearedStage1 + '\n' +
      '필터정보 미매칭: ' + result.missing
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    ui.alert('쿠팡재전송_로그 날짜 직접패치 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}


function findSheetHeaderCol_(sheet, headerName) {
  const target = normalizeHeaderKey_(headerName);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeaderKey_(headers[i]) === target) return i + 1;
  }
  return 0;
}

function findOrAppendRetransmitLogRowByFilterName_(sheet, filterName) {
  const colFilter = findSheetHeaderCol_(sheet, '검색필터명');
  if (!colFilter) throw new Error('쿠팡재전송_로그에서 검색필터명 컬럼을 찾지 못했습니다.');

  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow >= 2) {
    const values = sheet.getRange(2, colFilter, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === String(filterName || '').trim()) return i + 2;
    }
  }

  return lastRow + 1;
}

function setRetransmitLogCell_(sheet, row, headerName, value) {
  const col = findSheetHeaderCol_(sheet, headerName);
  if (!col) throw new Error('쿠팡재전송_로그에서 ' + headerName + ' 컬럼을 찾지 못했습니다.');
  sheet.getRange(row, col).setValue(value);
}

function getRetransmitLogCell_(sheet, row, headerName) {
  const col = findSheetHeaderCol_(sheet, headerName);
  if (!col) return '';
  return sheet.getRange(row, col).getValue();
}

function patchOneRetransmitLogDateFromApiExact_(filterName) {
  filterName = String(filterName || '').trim();
  if (!filterName) throw new Error('검색필터명이 비어 있습니다.');

  const item = findExactFilterListItemByFilterName_(filterName);
  if (!item) {
    const items = fetchFilterListItemsByKeyword_(filterName);
    const names = (items || []).map(function(x) { return String(findFilterName_(x) || '').trim(); }).filter(Boolean).slice(0, 20);
    throw new Error(
      'filterList에서 정확히 일치하는 필터를 찾지 못했습니다.\n' +
      '검색어: ' + filterName + '\n' +
      'API 반환 필터명: ' + (names.length ? names.join(', ') : '(없음)')
    );
  }

  const builtRows = buildFilterSummaryRowsFromFilterListItems_([item]);
  if (!builtRows || !builtRows.length) {
    throw new Error(filterName + ' 응답을 필터별_상품수 행으로 변환하지 못했습니다.');
  }

  // 필터별_상품수 캐시도 같이 최신화
  mergeFilterRowsIntoFilterSummarySheet_(builtRows);

  const row = builtRows[0];
  const brand = row[1] || brandFromFilterName_(filterName);
  const accountId = row[3] || accountFromFilterName_(filterName).accountId || '';
  const recentDate = row[8] || '';
  const createDate = row[9] || '';

  const singleMap = {};
  singleMap[filterName] = {
    '검색필터명': filterName,
    '브랜드명': brand,
    '쿠팡계정ID': accountId,
    'API_최근수집일자': recentDate,
    'API_필터생성일': createDate
  };

  const info = getFilterSummaryDateInfoForLog_(filterName, brand, singleMap);

  const sheet = ensureCoupangWorkLogSheet_();
  const targetRow = findOrAppendRetransmitLogRowByFilterName_(sheet, filterName);

  // v5.98:
  // replaceData/병합 로직을 거치지 않고, 사용자가 보고 있는 쿠팡재전송_로그의 실제 셀에 직접 씁니다.
  setRetransmitLogCell_(sheet, targetRow, '검색필터명', filterName);
  setRetransmitLogCell_(sheet, targetRow, '브랜드명', brand);
  setRetransmitLogCell_(sheet, targetRow, '쿠팡계정ID', accountId);
  setRetransmitLogCell_(sheet, targetRow, '최초전송일', info.createDisplay || '');
  setRetransmitLogCell_(sheet, targetRow, '1단계_추가수집일', info.additionalDisplay || '');

  const existingMemo = getRetransmitLogCell_(sheet, targetRow, '비고');
  const memo = '필터일자 기준: 생성 ' + (info.createDisplay || '-') + ' / 최근수집 ' + (info.recentDisplay || '-');
  setRetransmitLogCell_(sheet, targetRow, '비고', mergeShortText_(existingMemo || '', memo));

  try { formatCoupangWorkLogSheetFastOnly_(sheet); } catch (e) {}

  log_(
    'direct_patch_retransmit_log_date_v594',
    filterName +
    ' / row=' + targetRow +
    ' / create=' + (info.createDisplay || '') +
    ' / recent=' + (info.recentDisplay || '') +
    ' / stage1=' + (info.additionalDisplay || '')
  );

  return {
    filterName: filterName,
    brand: brand,
    accountId: accountId,
    row: targetRow,
    createDisplay: info.createDisplay || '',
    recentDisplay: info.recentDisplay || '',
    additionalDisplay: info.additionalDisplay || ''
  };
}


function patchK2RetransmitLogDateFromApi() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = patchOneRetransmitLogDateFromApiExact_('롯백_04_K2');

    ui.alert(
      'K2 날짜 직접반영 완료\n\n' +
      '검색필터명: ' + result.filterName + '\n' +
      '반영 행: ' + result.row + '행\n' +
      '최초전송일: ' + (result.createDisplay || '-') + '\n' +
      'API_최근수집일자: ' + (result.recentDisplay || '-') + '\n' +
      '1단계_추가수집일: ' + (result.additionalDisplay || '-') + '\n\n' +
      '쿠팡재전송_로그의 실제 셀에 직접 반영했습니다.\n' +
      '이후 핵심요약+대시보드 날짜연동 갱신을 실행하세요.'
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    ui.alert('K2 날짜 직접반영 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function shouldSkipInactiveCoupangWorkLogRow_(r, manual) {
  const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
  if (filterName && isInactiveManualFilterName_(filterName, manual)) return true;
  if (!filterName) {
    const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim();
    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, '');
    const item = brandKey && manual && manual.byBrandAll ? manual.byBrandAll[brandKey] : null;
    if (item && item.inactive) return true;
  }
  return false;
}

function isInactiveManualFilterName_(filterName, manual) {
  const key = String(filterName || '').trim();
  if (!key || !manual || !manual.byFilterAll) return false;
  const item = manual.byFilterAll[key];
  return !!(item && item.inactive);
}

function buildCoupangWorkLogMetricMap_(brandRows) {
  const brandHeaders = CONFIG.HEADERS.BRAND;
  const byFilter = {}, byBrand = {}, byBrandAccount = {}, byMergeKey = {};
  const manual = readCoupangSentManualMap_();
  const filterSummaryMap = buildFilterSummaryMapFromFilterSheet_();

  (brandRows || []).forEach(function(row) {
    const b = {};
    brandHeaders.forEach(function(h, i) { b[h] = row[i]; });
    const filterName = String(b['대표검색필터명'] || '').trim();
    const brand = String(b['브랜드명'] || '').trim();
    const accountId = String(b['쿠팡계정ID'] || '').trim();
    if (!filterName && !brand) return;

    const manualForFilter = manual.byFilterAll[filterName] ||
      manual.byBrandAll[normalizeCoupangWorkLogBrandMergeKey_(brand, filterName)] || null;
    if (manualForFilter && manualForFilter.inactive) return;

    const filterDate = getFilterSummaryDateInfoForLog_(filterName, brand, filterSummaryMap);

    // v5.98:
    // 쿠팡재전송_로그의 최초전송일은 상품/수동 전송일이 아니라 검색필터 API의 필터생성일을 우선 사용합니다.
    const fallbackFirstSendRaw = normalizeDateText_(b['최초쿠팡전송일'] || b['대표쿠팡전송일'] || b['더망고기준등록일']);
    const firstSendRaw = filterDate.createDate || fallbackFirstSendRaw;

    const revenue = toNumber_(b['전체_매출액']);

    const metric = {
      filterName: filterDate.filterName || filterName,
      brand: filterDate.brand || brand,
      accountId: accountId || filterDate.accountId,
      firstSendRaw: firstSendRaw,
      firstSendDisplay: formatShortDate_(firstSendRaw),
      filterCreateDate: filterDate.createDate || '',
      filterRecentCollectDate: filterDate.recentDate || '',
      additionalCollectionRaw: filterDate.additionalDate || '',
      hasFilterDateInfo: !!(filterDate.hasApiRow || filterDate.hasDateInfo),
      hasFilterDateInfo: !!(filterDate.hasApiRow || filterDate.hasDateInfo),
      additionalCollectionDisplay: filterDate.additionalDisplay || '',
      cumulativeRevenue: revenue,
      monthlyAverageRevenue: calcMonthlyAverageRevenue_(revenue, firstSendRaw),
      productCount: toNumber_(b['LOTTEON_더망고등록상품수']),
      sentCount: toNumber_(b['쿠팡전송확인상품수']),
      manualStatusType: manualForFilter ? manualForFilter.statusType : ''
    };

    if (manualForFilter && !manualForFilter.inactive) {
      if (manualForFilter.hasMangoCount && manualForFilter.mangoCount > 0) metric.productCount = manualForFilter.mangoCount;
      if (manualForFilter.hasTotal) metric.sentCount = manualForFilter.total || 0;
      if (manualForFilter.accountId) metric.accountId = manualForFilter.accountId;
      if (manualForFilter.filterName) {
        metric.filterName = manualForFilter.filterName;
        const manualDate = getFilterSummaryDateInfoForLog_(manualForFilter.filterName, manualForFilter.brand || metric.brand, filterSummaryMap);
        if (manualDate.createDate) {
          metric.filterCreateDate = manualDate.createDate;
          metric.filterRecentCollectDate = manualDate.recentDate || metric.filterRecentCollectDate || '';
          metric.additionalCollectionRaw = manualDate.additionalDate || metric.additionalCollectionRaw || '';
          metric.additionalCollectionDisplay = formatShortDate_(metric.additionalCollectionRaw);
          metric.firstSendRaw = manualDate.createDate;
          metric.firstSendDisplay = formatShortDate_(manualDate.createDate);
          metric.monthlyAverageRevenue = calcMonthlyAverageRevenue_(revenue, manualDate.createDate);
        }
      }
      if (manualForFilter.brand) metric.brand = manualForFilter.brand;
    }

    const mergeKey = getCoupangWorkLogMetricKey_(metric);
    if (!mergeKey) return;
    if (!byMergeKey[mergeKey]) byMergeKey[mergeKey] = cloneCoupangWorkLogMetric_(metric);
    else byMergeKey[mergeKey] = mergeCoupangWorkLogMetric_(byMergeKey[mergeKey], metric);
  });

  (manual.rows || []).forEach(function(item) {
    if (!item || item.inactive) return;
    if (!item.filterName && !item.brand) return;

    const filterDate = getFilterSummaryDateInfoForLog_(item.filterName, item.brand || brandFromFilterName_(item.filterName), filterSummaryMap);
    const firstSendRaw = filterDate.createDate || item.sendDate || '';

    const metric = {
      filterName: filterDate.filterName || item.filterName || '',
      brand: filterDate.brand || item.brand || brandFromFilterName_(item.filterName) || '',
      accountId: item.accountId || filterDate.accountId || '',
      firstSendRaw: firstSendRaw,
      firstSendDisplay: formatShortDate_(firstSendRaw),
      filterCreateDate: filterDate.createDate || '',
      filterRecentCollectDate: filterDate.recentDate || '',
      additionalCollectionRaw: filterDate.additionalDate || '',
      additionalCollectionDisplay: filterDate.additionalDisplay || '',
      cumulativeRevenue: 0,
      monthlyAverageRevenue: 0,
      productCount: item.mangoCount || 0,
      sentCount: item.total || 0,
      manualStatusType: item.statusType || ''
    };

    const mergeKey = getCoupangWorkLogMetricKey_(metric);
    if (!mergeKey) return;
    if (!byMergeKey[mergeKey]) byMergeKey[mergeKey] = cloneCoupangWorkLogMetric_(metric);
    else byMergeKey[mergeKey] = mergeCoupangWorkLogMetric_(byMergeKey[mergeKey], metric);
  });

  Object.keys(byMergeKey).forEach(function(mergeKey) {
    const metric = byMergeKey[mergeKey];
    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(metric.brand, metric.filterName);
    if (metric.filterName) byFilter[metric.filterName] = metric;
    if (brandKey) byBrand[brandKey] = metric;
    const baKey = makeBrandAccountMergeKey_(metric.brand, metric.accountId);
    if (baKey) byBrandAccount[baKey] = metric;
  });

  return { byFilter: byFilter, byBrand: byBrand, byBrandAccount: byBrandAccount, byMergeKey: byMergeKey };
}

function cloneCoupangWorkLogMetric_(metric) {
  return {
    filterName: metric.filterName || '',
    brand: metric.brand || '',
    accountId: metric.accountId || '',
    firstSendRaw: metric.firstSendRaw || '',
    firstSendDisplay: metric.firstSendDisplay || '',
    filterCreateDate: metric.filterCreateDate || '',
    filterRecentCollectDate: metric.filterRecentCollectDate || '',
    additionalCollectionRaw: metric.additionalCollectionRaw || '',
    additionalCollectionDisplay: metric.additionalCollectionDisplay || '',
    cumulativeRevenue: toNumber_(metric.cumulativeRevenue),
    monthlyAverageRevenue: toNumber_(metric.monthlyAverageRevenue),
    productCount: toNumber_(metric.productCount),
    sentCount: toNumber_(metric.sentCount),
    manualStatusType: metric.manualStatusType || '',
    hasFilterDateInfo: !!metric.hasFilterDateInfo
  };
}

function mergeCoupangWorkLogMetric_(base, incoming) {
  const preferred = choosePreferredCoupangWorkLogMetric_(base, incoming);

  // v5.98: 최초전송일은 filterList 필터생성일 기준이므로 여러 값이 있으면 더 이른 필터생성일을 사용합니다.
  const firstSendRaw = pickEarlierDateText_(base.firstSendRaw, incoming.firstSendRaw);
  const filterCreateDate = pickEarlierDateText_(base.filterCreateDate, incoming.filterCreateDate);
  const filterRecentCollectDate = pickLaterDateText_(base.filterRecentCollectDate, incoming.filterRecentCollectDate);
  const additionalCollectionRaw = pickLaterDateText_(base.additionalCollectionRaw, incoming.additionalCollectionRaw);

  const revenue = toNumber_(base.cumulativeRevenue) + toNumber_(incoming.cumulativeRevenue);
  const productCount = toNumber_(base.productCount) + toNumber_(incoming.productCount);
  const sentCount = toNumber_(base.sentCount) + toNumber_(incoming.sentCount);

  return {
    filterName: preferred.filterName || base.filterName || incoming.filterName || '',
    brand: preferred.brand || base.brand || incoming.brand || '',
    accountId: preferred.accountId || base.accountId || incoming.accountId || '',
    firstSendRaw: firstSendRaw,
    firstSendDisplay: formatShortDate_(firstSendRaw),
    filterCreateDate: filterCreateDate,
    filterRecentCollectDate: filterRecentCollectDate,
    additionalCollectionRaw: additionalCollectionRaw,
    additionalCollectionDisplay: formatShortDate_(additionalCollectionRaw),
    cumulativeRevenue: revenue,
    monthlyAverageRevenue: calcMonthlyAverageRevenue_(revenue, firstSendRaw),
    productCount: productCount,
    sentCount: sentCount,
    manualStatusType: preferred.manualStatusType || base.manualStatusType || incoming.manualStatusType || '',
    hasFilterDateInfo: !!(base.hasFilterDateInfo || incoming.hasFilterDateInfo)
  };
}

function choosePreferredCoupangWorkLogMetric_(a, b) {
  const ap = manualStatusPriority_(a.manualStatusType), bp = manualStatusPriority_(b.manualStatusType);
  if (ap !== bp) return bp > ap ? b : a;
  const aGeneric = isGenericCoupangWorkLogBrand_(a.brand, a.filterName);
  const bGeneric = isGenericCoupangWorkLogBrand_(b.brand, b.filterName);
  if (aGeneric !== bGeneric) return bGeneric ? b : a;
  const aLen = String(a.filterName || '').length || 999, bLen = String(b.filterName || '').length || 999;
  if (aLen !== bLen) return bLen < aLen ? b : a;
  const ad = normalizeDateText_(a.firstSendRaw), bd = normalizeDateText_(b.firstSendRaw);
  if (ad && bd && bd > ad) return b;
  return b.filterName ? b : a;
}

function manualStatusPriority_(statusType) { const s = String(statusType || ''); if (s === 'ADD') return 3; if (s === 'NORMAL') return 2; return 1; }
function isGenericCoupangWorkLogBrand_(brand, filterName) { const raw = normalizeBrandKey_(brand || brandFromFilterName_(filterName)); if (raw === 'lee') return true; return false; }
function pickEarlierDateText_(a, b) { const ad = normalizeDateText_(a), bd = normalizeDateText_(b); if (!ad) return bd || ''; if (!bd) return ad || ''; return bd <= ad ? bd : ad; }

function buildCoupangWorkLogOutputObj_(r, metric) {
  r = r || {};
  metric = metric || {};

  const existingStage1 = getObjectValueByHeader_(r, '1단계_추가수집일') || getObjectValueByHeader_(r, '추가수집일');

  let stage1Display = '';
  if (metric.hasFilterDateInfo) {
    // v5.98:
    // API 기준 날짜가 있으면 기존 로그값을 유지하지 않고 API_최근수집일자 조건 결과로 덮어씁니다.
    // 이 처리로 이전 실행에서 잘못 들어간 6/13 같은 미래 날짜를 제거합니다.
    stage1Display = metric.additionalCollectionRaw ? formatShortDate_(metric.additionalCollectionRaw) : '';
  } else {
    // API 날짜 정보가 없는 경우에만 기존 수동 입력값을 보존합니다.
    stage1Display = isFutureDateText_(existingStage1) ? '' : formatShortDate_(existingStage1);
  }

  return {
    '검색필터명': metric.filterName || getObjectValueByHeader_(r, '검색필터명') || '',
    '브랜드명': metric.brand || getObjectValueByHeader_(r, '브랜드명') || '',
    '쿠팡계정ID': metric.accountId || getObjectValueByHeader_(r, '쿠팡계정ID') || '',

    // v5.98: 최초전송일은 검색필터 API의 필터생성일을 우선 반영합니다.
    '최초전송일': metric.firstSendDisplay || formatShortDate_(getObjectValueByHeader_(r, '최초전송일')) || '',

    '누적매출액': metric.cumulativeRevenue || toNumber_(getObjectValueByHeader_(r, '누적매출액')),
    '월평균매출액': metric.monthlyAverageRevenue || toNumber_(getObjectValueByHeader_(r, '월평균매출액')),
    '작업유형': getObjectValueByHeader_(r, '작업유형') || '',

    // v5.98: 최근수집일자가 필터생성일보다 최근이고, 오늘 이후 미래 날짜가 아닐 때만 반영합니다.
    '1단계_추가수집일': stage1Display,

    '2단계_재전송일': formatShortDate_(getObjectValueByHeader_(r, '2단계_재전송일') || getObjectValueByHeader_(r, '재전송일') || getObjectValueByHeader_(r, '기준일')),
    '3단계_정리재수집일': formatShortDate_(getObjectValueByHeader_(r, '3단계_정리재수집일') || getObjectValueByHeader_(r, '정리재수집일')),
    '최종_더망고수집수': metric.productCount || toNumber_(getObjectValueByHeader_(r, '최종_더망고수집수')) || '',
    '최종_쿠팡전송수': metric.sentCount || toNumber_(getObjectValueByHeader_(r, '최종_쿠팡전송수')) || '',
    '추가수집상품수': getObjectValueByHeader_(r, '추가수집상품수') || '',
    '작업사유': getObjectValueByHeader_(r, '작업사유') || '',
    '비고': mergeShortText_(
      getObjectValueByHeader_(r, '비고') || '',
      metric.filterCreateDate || metric.filterRecentCollectDate
        ? ('필터일자 기준: 생성 ' + (formatShortDate_(metric.filterCreateDate) || '-') + ' / 최근수집 ' + (formatShortDate_(metric.filterRecentCollectDate) || '-'))
        : ''
    )
  };
}

function getCoupangWorkLogMetricForRow_(r, metricMap) {
  const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
  const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim() || brandFromFilterName_(filterName);
  const accountId = String(getObjectValueByHeader_(r, '쿠팡계정ID') || '').trim();
  if (filterName && metricMap.byFilter[filterName]) return metricMap.byFilter[filterName];
  const baKey = makeBrandAccountMergeKey_(brand, accountId);
  if (baKey && metricMap.byBrandAccount[baKey]) return metricMap.byBrandAccount[baKey];
  const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
  if (brandKey && metricMap.byBrand[brandKey]) return metricMap.byBrand[brandKey];
  return {};
}

function getCoupangWorkLogMergeKey_(r, metric) {
  metric = metric || {};
  if (metric.filterName || metric.brand) return getCoupangWorkLogMetricKey_(metric);
  const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
  const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim() || brandFromFilterName_(filterName);
  const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
  if (brandKey) return 'BRAND|' + brandKey;
  return filterName ? 'FILTER|' + filterName : '';
}

function getCoupangWorkLogMetricKey_(metric) {
  const brandKey = normalizeCoupangWorkLogBrandMergeKey_(metric.brand, metric.filterName);
  if (brandKey) return 'BRAND|' + brandKey;
  return metric.filterName ? 'FILTER|' + metric.filterName : '';
}

function makeBrandAccountMergeKey_(brand, accountId) {
  const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, '');
  const account = String(accountId || '').trim().toLowerCase();
  if (!brandKey && !account) return '';
  return brandKey + '|' + account;
}

function normalizeCoupangWorkLogBrandMergeKey_(brand, filterName) {
  var s = normalizeBrandKey_(brand || brandFromFilterName_(filterName));
  if (s === 'lee키즈' || s === 'leekids' || s === 'lee') return 'lee';
  return s;
}

function mergeCoupangWorkLogObjects_(base, incoming) {
  const out = {};
  CONFIG.HEADERS.RETRANSMIT_LOG.forEach(function(h) {
    const key = normalizeHeaderKey_(h);
    const a = getObjectValueByHeader_(base, key);
    const b = getObjectValueByHeader_(incoming, key);
    if (key === '검색필터명' || key === '브랜드명' || key === '쿠팡계정ID') out[key] = b || a;
    else if (key === '최초전송일') out[key] = pickEarlierShortDate_(a, b);
    else if (key === '1단계_추가수집일' || key === '2단계_재전송일' || key === '3단계_정리재수집일') out[key] = pickLaterShortDate_(a, b);
    else if (key === '누적매출액' || key === '월평균매출액' || key === '최종_더망고수집수' || key === '최종_쿠팡전송수' || key === '추가수집상품수') { const bn = toNumber_(b), an = toNumber_(a); out[key] = bn || an || ''; }
    else if (key === '작업유형') out[key] = mergeWorkType_(a, b);
    else if (key === '작업사유' || key === '비고') out[key] = mergeShortText_(a, b);
    else out[key] = b || a || '';
  });
  return out;
}

function mergeWorkType_(a, b) { const av = String(a || '').trim(), bv = String(b || '').trim(); if (!av) return bv; if (!bv) return av; if (av === bv) return av; const combined = av + '+' + bv; if (combined.indexOf('정리재수집') >= 0) return '정리재수집'; if (combined.indexOf('삭제후재전송') >= 0) return '삭제후재전송'; if (combined.indexOf('추가수집') >= 0) return '추가수집'; if (combined.indexOf('신규전송') >= 0) return '신규전송'; return bv || av; }
function mergeShortText_(a, b) { const av = String(a || '').trim(), bv = String(b || '').trim(); if (!av) return bv; if (!bv || av === bv) return av; if (av.indexOf(bv) >= 0) return av; if (bv.indexOf(av) >= 0) return bv; return av + ' / ' + bv; }
function pickLaterShortDate_(a, b) { const ad = normalizeDateText_(a), bd = normalizeDateText_(b); if (!ad) return formatShortDate_(bd || b); if (!bd) return formatShortDate_(ad || a); return formatShortDate_(bd >= ad ? bd : ad); }
function pickEarlierShortDate_(a, b) { const ad = normalizeDateText_(a), bd = normalizeDateText_(b); if (!ad) return formatShortDate_(bd || b); if (!bd) return formatShortDate_(ad || a); return formatShortDate_(bd <= ad ? bd : ad); }

function buildMonthlyBrandRows_(products, sales) {
  const productByBrand = {};
  products.forEach(function(p) { const brand = p['브랜드명'] || '브랜드미확인'; if (!productByBrand[brand]) productByBrand[brand] = []; productByBrand[brand].push(p); });
  const manual = readCoupangSentManualMap_();
  const map = {};
  sales.forEach(function(s) {
    if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y' || s['매칭방식'] === '미매칭') return;
    const month = String(s['주문월'] || ''), brand = String(s['브랜드명_매칭'] || '').trim();
    if (!month || !brand) return;
    const key = month + '|' + brand;
    if (!map[key]) {
      const p0 = (productByBrand[brand] || [])[0] || {};
      const manualItem = manual.byFilter[String(p0['검색필터명'] || '').trim()] || manual.byBrand[normalizeBrandKey_(brand)] || {};
      map[key] = { month: month, brand: brand, accountNo: p0['계정번호'] || '', accountId: p0['쿠팡계정ID'] || '', mangoProductCount: countProductsByMonthEnd_(productByBrand[brand] || [], month), coupangSentCount: toNumber_(manualItem.total), productIds: {}, orders: 0, revenue: 0, cancelledOrders: 0, cancelledRevenue: 0 };
    }
    const method = String(s['매칭방식'] || '');
    if (method === 'siteProductId' || method === 'brandProductName') { const pid = String(s['매칭상품ID'] || '').trim(); if (pid) map[key].productIds[pid] = true; }
    map[key].orders += toNumber_(s['결제수량']) || 1;
    map[key].revenue += toNumber_(s['결제금액합계']);
    if (String(s['취소여부'] || '') === 'Y') {
      map[key].cancelledOrders += toNumber_(s['결제수량']) || 1;
      map[key].cancelledRevenue += toNumber_(s['결제금액합계']);
    }
  });
  const rows = Object.keys(map).sort().map(function(key) {
    const m = map[key]; const soldProductCount = Object.keys(m.productIds).length; const denom = m.coupangSentCount || 0;
    const sellRate = denom ? soldProductCount / denom : 0, revenuePerProduct = denom ? m.revenue / denom : 0, orderPerProduct = denom ? m.orders / denom : 0;
    return [m.month, m.brand, m.accountNo, m.accountId, m.mangoProductCount, denom, soldProductCount, sellRate, m.orders, m.revenue, revenuePerProduct, orderPerProduct, judgeMonthly_(m.revenue, m.orders, sellRate), m.cancelledOrders || 0, m.cancelledRevenue || 0];
  });
  rows.sort(function(a, b) { return a[0] !== b[0] ? String(b[0]).localeCompare(String(a[0])) : b[9] - a[9]; });
  return rows;
}

function buildProductPerfRows_(products, productMetrics) {
  const rows = [];
  products.forEach(function(p) { const productId = String(p['상품ID'] || '').trim(); const key = productId || String(p['상품명'] || '') + '|' + String(p['검색필터명'] || ''); const pm = productMetrics[key] || {}; rows.push([productId, p['롯데온상품번호_추출'], p['상품명'], p['브랜드명'], p['검색필터명'], p['계정번호'], p['쿠팡계정ID'], p['기준등록일'], toNumber_(p['판매기간일수']), pm.sold ? 'Y' : 'N', pm.orderCount || 0, pm.revenue || 0, pm.latestMonth || '', pm.sold ? '매출 발생 상품 (사이트상품번호 또는 브랜드+상품명 매칭)' : '매출 없음 또는 상품단위 미매칭']); });
  rows.sort(function(a, b) { return b[11] - a[11]; });
  return rows;
}

function buildDeleteRows_(brandMetrics) {
  const rows = [];
  Object.keys(brandMetrics).forEach(function(brand) {
    const b = brandMetrics[brand]; if (!b.productCount) return;
    const basis = getCoupangBasis_(b); const decision = judgeBrandByCoupang_(b, basis);
    rows.push([deletePriority_(decision.label), b.brand, b.accountNo, b.accountId, basis.mangoDays, basis.coupangDays, b.productCount, basis.sentCount, b.soldProductCount, basis.sellRate, b.orderCount, b.revenue, basis.revenue30d, decision.label, decision.reason, b.filterName, basis.sentStatus]);
  });
  rows.sort(function(a, b) { if (a[0] !== b[0]) return a[0] - b[0]; if (a[11] !== b[11]) return a[11] - b[11]; return b[6] - a[6]; });
  return rows;
}

function buildAccountRows_(brandMetrics) {
  const rows = [];
  Object.keys(brandMetrics).forEach(function(brand) {
    const b = brandMetrics[brand]; if (!b.productCount) return;
    const basis = getCoupangBasis_(b); const decision = judgeBrandByCoupang_(b, basis);
    rows.push([b.accountNo, b.accountId, b.brand, b.filterName, b.productCount, basis.sentCount, b.soldProductCount, b.orderCount, b.revenue, basis.coupangDays, basis.revenue30d, recommendAccountAction_(decision.label), accountMemo_(decision, b)]);
  });
  rows.sort(function(a, b) { return String(a[0]) !== String(b[0]) ? String(a[0]).localeCompare(String(b[0])) : b[10] - a[10]; });
  return rows;
}

function buildSentRows_(products, productMetrics) {
  const map = {};
  products.forEach(function(p) {
    const brand = p['브랜드명'] || '브랜드미확인';
    if (!map[brand]) map[brand] = { brand: brand, filterName: p['검색필터명'], accountNo: p['계정번호'], accountId: p['쿠팡계정ID'], productCount: 0, apiSendCount: 0, manualSendCount: 0, soldProductIds: {}, manualTotal: 0, manualSendDate: '', manualStatus: '' };
    const b = map[brand]; b.productCount++;
    if (p['API_마켓전송일']) b.apiSendCount++;
    if (p['마켓전송일_수동']) b.manualSendCount++;
    const pid = String(p['상품ID'] || '').trim(); const pm = productMetrics[pid];
    if (pm && pm.sold && pid) b.soldProductIds[pid] = true;
  });
  const manual = readCoupangSentManualMap_();
  Object.keys(map).forEach(function(brand) {
    const b = map[brand]; const item = manual.byFilter[String(b.filterName || '').trim()] || manual.byBrand[normalizeBrandKey_(brand)];
    if (!item) return;
    b.manualTotal = toNumber_(item.total); b.manualSendDate = normalizeDateText_(item.sendDate);
    b.manualStatus = b.manualTotal ? (b.manualSendDate ? '수동입력 전송수+전송일 확인' : '수동입력 전송수 확인/전송일 미확인') : '';
    if (!b.accountId) { const inferred = inferAccountFromManualCounts_(item.accountCounts || {}); b.accountId = inferred.accountId || b.accountId; b.accountNo = inferred.accountNo || b.accountNo; }
  });
  return Object.keys(map).sort().map(function(brand) {
    const b = map[brand]; const soldCount = Object.keys(b.soldProductIds).length;
    const coupangSentCount = b.manualTotal || b.apiSendCount || b.manualSendCount || 0;
    const representativeDate = b.manualSendDate || '';
    let sentStatus = '전송정보 미확인', status = '전송 여부 확인 불가';
    if (b.manualTotal > 0) { status = '수동입력 전송수 확인'; sentStatus = b.manualStatus || '수동입력 전송수 확인'; }
    else if (b.apiSendCount > 0) { status = 'API 응답 전송일로 일부 확인'; sentStatus = '상품별 API 전송일 일부 확인'; }
    else if (b.manualSendCount > 0) { status = '상품별 수동전송일로 일부 확인'; sentStatus = '상품별 수동전송일 일부 확인'; }
    else if (soldCount > 0) { status = '주문발생으로 전송 확인'; sentStatus = '주문발생상품만 전송 확인'; }
    return [b.brand, b.filterName, b.accountNo, b.accountId, b.productCount, coupangSentCount, b.manualTotal, b.apiSendCount, b.manualSendCount, soldCount, representativeDate, sentStatus, status, status === '전송 여부 확인 불가' ? '쿠팡전송수_수동입력 시트에 팝업 기준 전송수를 입력하세요.' : '전송/주문 근거 있음'];
  });
}


function refreshManualInputApiValidation() {
  // v5.55:
  // 기본 검증은 시간초과 방지를 위해 "문제행만" 출력합니다.
  // LOTTEON_상품목록 전체를 읽지 않고, 필터별_상품수 + 쿠팡전송수_수동입력만 비교합니다.
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.MANUAL_API_CHECK, CONFIG.HEADERS.MANUAL_API_CHECK);

  const filterSummaryMap = buildFilterSummaryMapFromFilterSheet_();
  const result = writeManualInputApiValidationFast_(filterSummaryMap, { issueOnly: true, ultraLight: true });

  SpreadsheetApp.getUi().alert(
    '수동입력/API 검증 시트 갱신 완료\n\n' +
    'v5.55 초경량 검증 방식으로 실행했습니다.\n' +
    '- 출력 방식: 문제행만 표시\n' +
    '- 기준 1: 필터별_상품수\n' +
    '- 기준 2: 쿠팡전송수_수동입력\n' +
    '- LOTTEON_상품목록 전체 읽기 없음\n\n' +
    '출력 행수: ' + result.writtenRows + '행\n' +
    '전체 검토 행수: ' + result.totalRows + '행\n\n' +
    '일치 행까지 모두 보고 싶으면 refreshManualInputApiValidationAllRows_()를 직접 실행하세요.'
  );
}

function refreshManualInputApiValidationAllRows_() {
  // 필요할 때만 직접 실행하는 전체행 경량 검증입니다.
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);
  ensureSheetWithHeader_(SpreadsheetApp.getActive(), CONFIG.SHEETS.MANUAL_API_CHECK, CONFIG.HEADERS.MANUAL_API_CHECK);

  const filterSummaryMap = buildFilterSummaryMapFromFilterSheet_();
  const result = writeManualInputApiValidationFast_(filterSummaryMap, { issueOnly: false, ultraLight: true });

  SpreadsheetApp.getUi().alert(
    '수동입력/API 검증 전체행 갱신 완료\n\n' +
    '출력 행수: ' + result.writtenRows + '행\n' +
    '주의: 평소에는 메뉴의 문제행 검증만 사용하세요.'
  );
}


function refreshManualInputApiValidationFullProducts_() {
  setupLotteonSheets(false, true);
  const products = readTable_(getSheet_(CONFIG.SHEETS.PRODUCTS));
  writeManualInputApiValidation_(products);
  SpreadsheetApp.getUi().alert('전체상품목록 기준 수동입력/API 검증 완료\n\n주의: LOTTEON_상품목록 전체를 읽는 무거운 검증입니다. 평소에는 메뉴의 경량 검증을 사용하세요.');
}

function writeManualInputApiValidation_(products) {
  const manual = readCoupangSentManualMap_();
  const apiMap = buildApiFilterCountMapFromProducts_(products || []);
  const rows = [];
  (manual.rows || []).forEach(function(item) {
    const api = apiMap[item.filterName] || {};
    const manualSent = item.hasTotal ? item.total : '';
    const apiSent = api.apiSentCount || 0;
    const sentDiff = item.hasTotal ? (toNumber_(manualSent) - apiSent) : '';
    const manualMango = item.hasMangoCount ? item.mangoCount : '';
    const apiMango = api.productCount || 0;
    const mangoDiff = item.hasMangoCount ? (toNumber_(manualMango) - apiMango) : '';
    let result = '확인';
    const memos = [];
    if (item.inactive) { result = '비활성행'; memos.push(item.statusRaw + ' 표시 - 분석/대표로그 제외'); }
    else if (!api.exists) { result = 'API 미확인'; memos.push('현재 LOTTEON_상품목록에 필터 없음'); }
    else if ((item.hasTotal && sentDiff !== 0) || (item.hasMangoCount && mangoDiff !== 0)) {
      result = '차이';
      if (item.hasTotal && sentDiff !== 0) memos.push('전송수 차이 ' + formatCount_(sentDiff));
      if (item.hasMangoCount && mangoDiff !== 0) memos.push('수집수 차이 ' + formatCount_(mangoDiff));
    } else { result = '일치'; memos.push('수동값과 API 집계값 일치'); }
    rows.push([item.statusRaw || item.confirmDate || '', item.filterName, item.brand, manualSent, apiSent, sentDiff, manualMango, apiMango, mangoDiff, result, memos.join(' / ')]);
  });
  replaceData_(getSheet_(CONFIG.SHEETS.MANUAL_API_CHECK), CONFIG.HEADERS.MANUAL_API_CHECK, rows);
  formatManualInputApiValidationSheetFast_(getSheet_(CONFIG.SHEETS.MANUAL_API_CHECK));
}

function formatManualInputApiValidationSheetFast_(sheet) {
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1), lastCol = Math.max(sheet.getLastColumn(), 1), dataRows = Math.max(lastRow - 1, 0);
  try { sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#e8f0fe').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center'); sheet.setRowHeight(1, 42); } catch (e) {}
  if (dataRows > 0) { try { sheet.getRange(2, 1, dataRows, lastCol).setNumberFormat('General').setWrap(false).setVerticalAlignment('middle'); sheet.getRange(2, 4, dataRows, 6).setNumberFormat('#,##0'); } catch (e) {} }
  var widths = [92, 180, 130, 92, 110, 85, 105, 115, 85, 80, 260];
  widths.forEach(function(w, i) { try { sheet.setColumnWidth(i + 1, w); } catch (e) {} });
}

function buildApiFilterCountMapFromProducts_(products) {
  const map = {};
  (products || []).forEach(function(p) {
    const filterName = String(getObjectValueByHeader_(p, '검색필터명') || '').trim();
    if (!filterName) return;
    if (!map[filterName]) map[filterName] = { exists: true, filterName: filterName, brand: String(getObjectValueByHeader_(p, '브랜드명') || '').trim(), productCount: 0, apiSentCount: 0, productCountSource: 'LOTTEON_상품목록' };
    map[filterName].productCount += 1;
    if (getObjectValueByHeader_(p, 'API_마켓전송일') || getObjectValueByHeader_(p, '마켓전송일_수동')) map[filterName].apiSentCount += 1;
  });
  try {
    const filterRows = readTable_(getSheet_(CONFIG.SHEETS.FILTERS));
    filterRows.forEach(function(r) {
      const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
      if (!filterName) return;
      const totalRaw = getObjectValueByHeader_(r, 'API_totalCount');
      const hasTotal = String(totalRaw || '').trim() !== '';
      const totalCount = toNumber_(totalRaw);
      if (!map[filterName]) { map[filterName] = { exists: true, filterName: filterName, brand: String(getObjectValueByHeader_(r, '브랜드명') || '').trim(), productCount: hasTotal ? totalCount : 0, apiSentCount: 0, productCountSource: '필터별_상품수' }; }
      else if (hasTotal) { map[filterName].exists = true; map[filterName].productCount = totalCount; map[filterName].brand = map[filterName].brand || String(getObjectValueByHeader_(r, '브랜드명') || '').trim(); map[filterName].productCountSource = '필터별_상품수'; }
    });
  } catch (e) {}
  return map;
}

function calcMonthlyAverageRevenue_(revenue, firstSendDate) {
  const r = toNumber_(revenue), d = normalizeDateText_(firstSendDate);
  if (!r) return 0;
  if (!d) return r;
  const days = Math.max(1, daysSince_(d)), months = Math.max(1, days / 30);
  return r / months;
}

function readCleanupRefreshLogMap_() {
  var rows = [];
  try { rows = readTable_(getSheet_(CONFIG.SHEETS.RETRANSMIT_LOG)); } catch (e) { return { byBrand: {}, byFilter: {} }; }
  const byBrand = {}, byFilter = {};
  rows.forEach(function(r) {
    const typeText = String(r['작업유형'] || '').trim();
    const hasCleanupDate = r['3단계_정리재수집일'] || r['정리재수집일'];
    if (typeText.indexOf('정리재수집') < 0 && typeText.indexOf('더망고삭제') < 0 && typeText.indexOf('완전정리') < 0 && !hasCleanupDate) return;
    const filterName = String(r['검색필터명'] || '').trim();
    const brand = String(r['브랜드명'] || '').trim() || brandFromFilterName_(filterName);
    if (!filterName && !brand) return;
    const basisDate = normalizeDateText_(r['3단계_정리재수집일'] || r['정리재수집일'] || r['2단계_재전송일'] || r['재전송일']);
    if (!basisDate) return;
    const item = { filterName: filterName, brand: brand, accountId: String(r['쿠팡계정ID'] || '').trim(), type: typeText || '정리재수집', basisDate: basisDate, finalSentCount: toNumber_(r['최종_쿠팡전송수']), finalMangoCount: toNumber_(r['최종_더망고수집수']), reason: String(r['작업사유'] || '').trim(), memo: String(r['비고'] || '').trim() };
    const brandKey = normalizeBrandKey_(brand);
    if (filterName) { const prev = byFilter[filterName]; if (!prev || String(item.basisDate) > String(prev.basisDate)) byFilter[filterName] = item; }
    if (brandKey) { const prev = byBrand[brandKey]; if (!prev || String(item.basisDate) > String(prev.basisDate)) byBrand[brandKey] = item; }
  });
  return { byBrand: byBrand, byFilter: byFilter };
}

function applyCleanupRefreshLogToBrandMetrics_(brandMap) {
  const logMap = readCleanupRefreshLogMap_();
  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand]; const item = logMap.byFilter[String(b.filterName || '').trim()] || logMap.byBrand[normalizeBrandKey_(brand)];
    if (!item) return;
    b.cleanupRefreshDate = item.basisDate || ''; b.cleanupRefreshType = item.type || '정리재수집';
    b.cleanupRefreshFinalMangoCount = item.finalMangoCount || 0; b.cleanupRefreshFinalSentCount = item.finalSentCount || 0;
    b.cleanupRefreshReason = item.reason || ''; b.cleanupRefreshMemo = item.memo || '';
    if (item.finalSentCount > 0) { b.coupangSentCount = item.finalSentCount; b.manualCoupangSentCount = item.finalSentCount; }
    if (item.finalMangoCount > 0 && item.finalMangoCount > b.productCount) b.productCount = item.finalMangoCount;
    if (item.accountId && !b.accountId) b.accountId = item.accountId;
    if (b.cleanupRefreshDate) b.coupangSentStatus = '정리재수집_로그 반영';
  });
}

function makeCleanupRefreshMemo_(b) { if (!b.cleanupRefreshDate) return ''; const parts = [formatShortDate_(b.cleanupRefreshDate)]; if (b.cleanupRefreshFinalMangoCount) parts.push('수집 ' + formatCount_(b.cleanupRefreshFinalMangoCount)); if (b.cleanupRefreshFinalSentCount) parts.push('전송 ' + formatCount_(b.cleanupRefreshFinalSentCount)); return parts.join(' / '); }

function readAdditionalCollectionLogMap_() {
  var rows = [];
  try { rows = readTable_(getSheet_(CONFIG.SHEETS.RETRANSMIT_LOG)); } catch (e) { return { byBrand: {}, byFilter: {} }; }
  const byBrand = {}, byFilter = {};
  rows.forEach(function(r) {
    const typeText = String(r['작업유형'] || '').trim();
    if (typeText.indexOf('추가수집') < 0) return;
    const filterName = String(r['검색필터명'] || '').trim();
    const brand = String(r['브랜드명'] || '').trim() || brandFromFilterName_(filterName);
    if (!filterName && !brand) return;
    const basisDate = normalizeDateText_(r['1단계_추가수집일'] || r['추가수집일'] || r['기준일']);
    if (!basisDate) return;
    const item = { operationDate: '', filterName: filterName, brand: brand, accountId: String(r['쿠팡계정ID'] || '').trim(), type: typeText || '추가수집', beforeMangoCount: 0, afterMangoCount: toNumber_(r['최종_더망고수집수']), beforeSentCount: 0, afterSentCount: toNumber_(r['최종_쿠팡전송수']), addedCount: toNumber_(r['추가수집상품수']), basisDate: basisDate, reason: String(r['작업사유'] || '').trim(), memo: String(r['비고'] || '').trim() };
    const brandKey = normalizeBrandKey_(brand);
    if (filterName) { const prev = byFilter[filterName]; if (!prev || String(item.basisDate) > String(prev.basisDate)) byFilter[filterName] = item; }
    if (brandKey) { const prev = byBrand[brandKey]; if (!prev || String(item.basisDate) > String(prev.basisDate)) byBrand[brandKey] = item; }
  });
  return { byBrand: byBrand, byFilter: byFilter };
}

function applyAdditionalCollectionLogToBrandMetrics_(brandMap) {
  const logMap = readAdditionalCollectionLogMap_();
  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand]; const item = logMap.byFilter[String(b.filterName || '').trim()] || logMap.byBrand[normalizeBrandKey_(brand)];
    if (!item) return;
    b.additionalCollectionDate = item.basisDate || ''; b.additionalCollectionOperationDate = item.operationDate || ''; b.additionalCollectionType = item.type || '추가수집';
    b.additionalBeforeMangoCount = item.beforeMangoCount || 0; b.additionalAfterMangoCount = item.afterMangoCount || 0;
    b.additionalBeforeSentCount = item.beforeSentCount || 0; b.additionalAfterSentCount = item.afterSentCount || 0;
    b.additionalAddedCount = item.addedCount || 0; b.additionalCollectionReason = item.reason || ''; b.additionalCollectionMemo = item.memo || '';
    if (item.afterSentCount > 0) { b.coupangSentCount = item.afterSentCount; b.manualCoupangSentCount = item.afterSentCount; }
    if (item.afterMangoCount > 0 && item.afterMangoCount > b.productCount) b.productCount = item.afterMangoCount;
    if (item.accountId && !b.accountId) b.accountId = item.accountId;
  });
}

function applyAdditionalCollectionSalesToBrandMetrics_(brandMap, sales) {
  Object.keys(brandMap).forEach(function(brand) { const b = brandMap[brand]; if (!b.additionalCollectionDate) return; b.additionalBeforeRevenue = 0; b.additionalAfterRevenue = 0; b.additionalBeforeOrders = 0; b.additionalAfterOrders = 0; });
  (sales || []).forEach(function(s) {
    if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y') return;
    const method = String(s['매칭방식'] || ''); if (method === '미매칭') return;
    const brand = String(s['브랜드명_매칭'] || '').trim(); const b = brandMap[brand];
    if (!b || !b.additionalCollectionDate) return;
    const orderDate = normalizeDateText_(s['주문일시']); if (!orderDate) return;
    const qty = toNumber_(s['결제수량']) || 1, revenue = toNumber_(s['결제금액합계']);
    if (String(orderDate).slice(0, 10) >= String(b.additionalCollectionDate).slice(0, 10)) { b.additionalAfterOrders += qty; b.additionalAfterRevenue += revenue; }
    else { b.additionalBeforeOrders += qty; b.additionalBeforeRevenue += revenue; }
  });
  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand]; if (!b.additionalCollectionDate) return;
    const beforeBase = normalizeDateText_(b.retransmitDate) || minDateText_(b.sendDates) || normalizeDateText_(b.filterRecentCollectDate) || normalizeDateText_(b.filterCreateDate) || minDateText_(b.baseDates);
    const beforeDays = beforeBase ? daysBetweenDates_(beforeBase, b.additionalCollectionDate) : 0;
    const afterDays = daysSince_(b.additionalCollectionDate);
    b.additionalBeforeDailyRevenue = beforeDays ? b.additionalBeforeRevenue / beforeDays : 0;
    b.additionalAfterDailyRevenue = afterDays ? b.additionalAfterRevenue / afterDays : 0;
    b.additionalCollectionEffect = judgeAdditionalCollectionEffect_(b, beforeDays, afterDays);
  });
}

function judgeAdditionalCollectionEffect_(b, beforeDays, afterDays) { if (!b.additionalCollectionDate) return ''; if (afterDays < 7) return '판단보류(7일미만)'; const beforeDaily = toNumber_(b.additionalBeforeDailyRevenue), afterDaily = toNumber_(b.additionalAfterDailyRevenue); if (!beforeDaily && afterDaily > 0) return '개선'; if (!beforeDaily && !afterDaily) return '변화없음'; if (afterDaily >= beforeDaily * 1.2) return '개선'; if (afterDaily <= beforeDaily * 0.8) return '악화'; return '변화없음'; }
function makeAdditionalCollectionMemo_(b) { if (!b.additionalCollectionDate) return ''; const parts = [formatShortDate_(b.additionalCollectionDate)]; if (b.additionalAddedCount) parts.push('추가 ' + formatCount_(b.additionalAddedCount)); if (b.additionalAfterSentCount) parts.push('전송 ' + formatCount_(b.additionalAfterSentCount)); const effect = summarizeEffect_(b.additionalCollectionEffect); if (effect) parts.push(effect); return parts.join(' / '); }

function readRetransmitLogMap_() {
  var rows = [];
  try { rows = readTable_(getSheet_(CONFIG.SHEETS.RETRANSMIT_LOG)); } catch (e) { return { byBrand: {}, byFilter: {} }; }
  const byBrand = {}, byFilter = {};
  rows.forEach(function(r) {
    const typeText = String(r['작업유형'] || '').trim();
    if (typeText.indexOf('추가수집') >= 0 && typeText.indexOf('재전송') < 0 && typeText.indexOf('신규전송') < 0) return;
    if (typeText.indexOf('정리재수집') >= 0 || typeText.indexOf('더망고삭제') >= 0 || typeText.indexOf('완전정리') >= 0 || r['3단계_정리재수집일'] || r['정리재수집일']) return;
    const filterName = String(r['검색필터명'] || '').trim(); const brand = String(r['브랜드명'] || '').trim() || brandFromFilterName_(filterName);
    if (!filterName && !brand) return;
    const basisDate = normalizeDateText_(r['2단계_재전송일'] || r['재전송일'] || r['기준일']); if (!basisDate) return;
    const item = { operationDate: '', filterName: filterName, brand: brand, accountId: String(r['쿠팡계정ID'] || '').trim(), type: typeText || '삭제후재전송', beforeCount: 0, afterCount: toNumber_(r['최종_쿠팡전송수'] || r['재전송후_쿠팡전송수']), finalMangoCount: toNumber_(r['최종_더망고수집수']), addedCount: toNumber_(r['추가수집상품수']), basisDate: basisDate, reason: String(r['작업사유'] || '').trim(), memo: String(r['비고'] || '').trim() };
    const brandKey = normalizeBrandKey_(brand);
    if (filterName) { const prev = byFilter[filterName]; if (!prev || String(item.basisDate) > String(prev.basisDate)) byFilter[filterName] = item; }
    if (brandKey) { const prev = byBrand[brandKey]; if (!prev || String(item.basisDate) > String(prev.basisDate)) byBrand[brandKey] = item; }
  });
  return { byBrand: byBrand, byFilter: byFilter };
}

function applyRetransmitLogToBrandMetrics_(brandMap) {
  const logMap = readRetransmitLogMap_();
  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand]; const item = logMap.byFilter[String(b.filterName || '').trim()] || logMap.byBrand[normalizeBrandKey_(brand)];
    if (!item) return;
    b.retransmitDate = item.basisDate || ''; b.retransmitOperationDate = item.operationDate || ''; b.retransmitType = item.type || '삭제후재전송';
    b.retransmitBeforeSentCount = 0; b.retransmitAfterSentCount = item.afterCount || 0;
    b.retransmitReason = item.reason || ''; b.retransmitMemo = item.memo || '';
    if (item.afterCount > 0) { b.coupangSentCount = item.afterCount; b.manualCoupangSentCount = item.afterCount; }
    if (item.finalMangoCount > 0 && item.finalMangoCount > b.productCount) b.productCount = item.finalMangoCount;
    if (item.accountId && !b.accountId) b.accountId = item.accountId;
    if (b.retransmitDate) b.coupangSentStatus = '쿠팡재전송_로그 반영';
  });
}

function applyRetransmitSalesToBrandMetrics_(brandMap, sales) {
  Object.keys(brandMap).forEach(function(brand) { const b = brandMap[brand]; if (!b.retransmitDate) return; b.retransmitBeforeRevenue = 0; b.retransmitAfterRevenue = 0; b.retransmitBeforeOrders = 0; b.retransmitAfterOrders = 0; });
  (sales || []).forEach(function(s) {
    if (s['유효여부'] !== 'Y' || s['분석대상여부'] !== 'Y') return;
    const method = String(s['매칭방식'] || ''); if (method === '미매칭') return;
    const brand = String(s['브랜드명_매칭'] || '').trim(); const b = brandMap[brand];
    if (!b || !b.retransmitDate) return;
    const orderDate = normalizeDateText_(s['주문일시']); if (!orderDate) return;
    const qty = toNumber_(s['결제수량']) || 1, revenue = toNumber_(s['결제금액합계']);
    if (String(orderDate).slice(0, 10) >= String(b.retransmitDate).slice(0, 10)) { b.retransmitAfterOrders += qty; b.retransmitAfterRevenue += revenue; }
    else { b.retransmitBeforeOrders += qty; b.retransmitBeforeRevenue += revenue; }
  });
  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand]; if (!b.retransmitDate) return;
    const beforeBase = minDateText_(b.sendDates) || normalizeDateText_(b.filterRecentCollectDate) || normalizeDateText_(b.filterCreateDate) || minDateText_(b.baseDates);
    const beforeDays = beforeBase ? daysBetweenDates_(beforeBase, b.retransmitDate) : 0;
    const afterDays = daysSince_(b.retransmitDate);
    b.retransmitBeforeDailyRevenue = beforeDays ? b.retransmitBeforeRevenue / beforeDays : 0;
    b.retransmitAfterDailyRevenue = afterDays ? b.retransmitAfterRevenue / afterDays : 0;
    b.retransmitEffect = judgeRetransmitEffect_(b, beforeDays, afterDays);
  });
}

function judgeRetransmitEffect_(b, beforeDays, afterDays) { if (!b.retransmitDate) return ''; if (afterDays < 7) return '판단보류(7일미만)'; const beforeDaily = toNumber_(b.retransmitBeforeDailyRevenue), afterDaily = toNumber_(b.retransmitAfterDailyRevenue); if (!beforeDaily && afterDaily > 0) return '개선'; if (!beforeDaily && !afterDaily) return '변화없음'; if (afterDaily >= beforeDaily * 1.2) return '개선'; if (afterDaily <= beforeDaily * 0.8) return '악화'; return '변화없음'; }
function makeRetransmitMemo_(b) { if (!b.retransmitDate) return ''; const parts = [formatShortDate_(b.retransmitDate)]; if (b.retransmitAfterSentCount) parts.push('전송 ' + formatCount_(b.retransmitAfterSentCount)); const effect = summarizeEffect_(b.retransmitEffect); if (effect) parts.push(effect); return parts.join(' / '); }

function createRetransmitLogSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureCoupangWorkLogSheet_();

  let brandRows = [];
  try {
    brandRows = getValuesWithoutHeader_(CONFIG.SHEETS.BRAND);
  } catch (e) {
    brandRows = [];
  }

  let filterRefresh = null;
  try {
    // v5.98:
    // 쿠팡재전송_로그 갱신 전, 로그/브랜드분석 대상 검색필터를 filterList로 개별 최신화합니다.
    // K2처럼 더망고 화면의 최근수집일자가 바뀌었는데 필터별_상품수 캐시가 오래된 경우를 방지합니다.
    filterRefresh = refreshFilterSummaryForRetransmitLogTargetsNoAlert_(brandRows);
  } catch (e) {
    filterRefresh = { error: String(e && e.message ? e.message : e) };
    log_('filterlist_target_refresh_error_v591', filterRefresh.error);
  }

  let patchResult = null;
  if (brandRows && brandRows.length) {
    syncCoupangWorkLogMetrics_(brandRows);
    patchResult = patchRetransmitLogDatesFromFilterSummary_();
  }

  try {
    sheet.showSheet();
    ss.setActiveSheet(sheet);
    formatCoupangWorkLogSheetFastOnly_(sheet);
  } catch (e) {}

  const refreshMsg = filterRefresh && !filterRefresh.error
    ? ('대상필터=' + filterRefresh.targetCount + ' / 최신화=' + filterRefresh.updatedCount + ' / 실패=' + (filterRefresh.errors ? filterRefresh.errors.length : 0) + ' / 소요초=' + filterRefresh.elapsedSec)
    : ('필터정보 최신화 실패: ' + (filterRefresh && filterRefresh.error ? filterRefresh.error : ''));

  ui.alert(
    '쿠팡재전송_로그 시트 생성/갱신 완료\n\n' +
    'v5.98 반영:\n' +
    '- 갱신 전 로그/브랜드분석 대상 필터를 정확한 검색필터명으로 filterList 개별 최신화\n' +
    '- 최초전송일 = 필터별_상품수의 API_필터생성일 기준\n' +
    '- API_최근수집일자가 API_필터생성일보다 최근이면 1단계_추가수집일에 반영\n' +
    '- 오늘 이후 미래 날짜는 무효 처리\n\n' +
    '필터정보 최신화 결과:\n' + refreshMsg + '\n\n' +
    '로그 날짜 직접패치 결과:\n' +
    (patchResult ? ('최초전송일=' + patchResult.patchedFirst + ' / 1단계=' + patchResult.patchedStage1 + ' / 1단계삭제=' + patchResult.clearedStage1 + ' / 미매칭=' + patchResult.missing) : '미실행') + '\n\n' +
    '다음 단계:\n' +
    'LOTTEON 자동화 → 핵심요약+대시보드 날짜연동 갱신'
  );
}

function diagnoseRetransmitLogFilterDateK2() {
  const ui = SpreadsheetApp.getUi();
  const target = '롯백_04_K2';
  try {
    const items = fetchFilterListItemsByKeyword_(target);
    const item = pickExactFilterListItem_(items, target);

    if (!item) {
      const names = (items || []).map(function(x) { return String(findFilterName_(x) || '').trim(); }).filter(Boolean).slice(0, 20);
      ui.alert(
        'K2 filterList 진단 결과(조회만)\n\n' +
        '검색어: ' + target + '\n' +
        '정확히 일치하는 필터를 찾지 못했습니다.\n\n' +
        'API가 반환한 필터명:\n' +
        (names.length ? names.join('\n') : '(없음)') + '\n\n' +
        '이 메뉴는 조회만 하며 쿠팡재전송_로그를 수정하지 않습니다.'
      );
      return;
    }

    const row = buildFilterSummaryRowsFromFilterListItems_([item])[0];
    const createDate = row ? row[9] : '';
    const recentDate = row ? row[8] : '';
    const additional = getFilterSummaryDateInfoForLog_(target, 'K2', {
      '롯백_04_K2': {
        '검색필터명': target,
        '브랜드명': 'K2',
        'API_최근수집일자': recentDate,
        'API_필터생성일': createDate,
        '쿠팡계정ID': 'beliun1024'
      }
    });

    ui.alert(
      'K2 filterList 진단 결과(조회만)\n\n' +
      '검색어: ' + target + '\n' +
      '매칭필터명: ' + (row ? row[0] : '') + '\n' +
      'API_필터생성일: ' + (formatShortDate_(createDate) || createDate || '-') + '\n' +
      'API_최근수집일자: ' + (formatShortDate_(recentDate) || recentDate || '-') + '\n' +
      '1단계_추가수집일 예정값: ' + (additional.additionalDisplay || '-') + '\n\n' +
      '중요: 이 메뉴는 진단용이라 시트를 수정하지 않습니다.\n' +
      '실제 반영은 LOTTEON 자동화 → K2 날짜 직접반영(실제수정)을 실행하세요.'
    );
  } catch (e) {
    ui.alert('K2 filterList 진단 중 오류가 발생했습니다.\n\n' + String(e && e.message ? e.message : e));
    throw e;
  }
}

function cleanFutureDatesInCoupangWorkLog() {
  const sheet = ensureCoupangWorkLogSheet_();
  const rows = readTable_(sheet);
  if (!rows || !rows.length) {
    SpreadsheetApp.getUi().alert('쿠팡재전송_로그에 정리할 데이터가 없습니다.');
    return;
  }

  const headers = CONFIG.HEADERS.RETRANSMIT_LOG;
  const cleaned = [];
  let changed = 0;

  rows.forEach(function(r) {
    const obj = {};
    headers.forEach(function(h) {
      const key = normalizeHeaderKey_(h);
      let v = getObjectValueByHeader_(r, key);
      if ((key === '최초전송일' || key === '1단계_추가수집일' || key === '2단계_재전송일' || key === '3단계_정리재수집일') && isFutureDateText_(v)) {
        v = '';
        changed++;
      }
      obj[key] = v;
    });
    cleaned.push(headers.map(function(h) { return obj[normalizeHeaderKey_(h)] || ''; }));
  });

  replaceDataFastLimited_(sheet, headers, cleaned);
  formatCoupangWorkLogSheetFastOnly_(sheet);
  SpreadsheetApp.getUi().alert('쿠팡재전송_로그 미래 날짜 정리 완료\n\n정리된 셀 수: ' + changed + '개');
}

function createAdditionalCollectionLogSheet() { createRetransmitLogSheet(); }

function createCoupangSentManualSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheetWithHeader_(ss, CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.HEADERS.COUPANG_SENT_MANUAL);

  // v5.56:
  // autoResizeColumns는 시트 데이터가 쌓이면 매우 느려져 시간초과 원인이 됩니다.
  // 기존 입력값은 절대 지우지 않고, 헤더/고정폭/기본 서식만 가볍게 적용합니다.
  try {
    sheet.showSheet();
    ss.setActiveSheet(sheet);
    formatCoupangSentManualSheetFast_(sheet);
  } catch (e) {}

  log_('manual_coupang_sent_sheet', '쿠팡전송수_수동입력 시트 경량 생성/갱신 완료');
  ui.alert(
    '쿠팡전송수_수동입력 시트 준비 완료\n\n' +
    '기존 입력값은 보존했습니다.\n' +
    'v5.56부터는 시간초과 방지를 위해 자동 열너비 조정(autoResize)을 사용하지 않습니다.\n\n' +
    '확인일에는 날짜 외에 "추가", "삭제", "01로 변경" 같은 상태값도 입력할 수 있습니다.\n' +
    '더망고 수집수와 쿠팡전체전송수는 분석에 수동 기준값으로 반영됩니다.'
  );
}

function formatCoupangSentManualSheetFast_(sheet) {
  if (!sheet) return;

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = CONFIG.HEADERS.COUPANG_SENT_MANUAL.length;
  const dataRows = Math.max(lastRow - 1, 0);

  try {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 42);
  } catch (e) {}

  if (dataRows > 0) {
    try {
      sheet.getRange(2, 1, dataRows, lastCol)
        .setWrap(false)
        .setVerticalAlignment('middle');
      sheet.getRange(2, 4, dataRows, 4).setNumberFormat('#,##0');
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat('M/dd');
    } catch (e) {}
  }

  const widths = [80, 190, 130, 105, 105, 95, 95, 95, 260];
  widths.forEach(function(w, i) {
    try { sheet.setColumnWidth(i + 1, w); } catch (e) {}
  });
}


function updateFilterRawFieldCheck() {
  const ui = SpreadsheetApp.getUi(); const ss = SpreadsheetApp.getActive();
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.FILTER_RAW, CONFIG.HEADERS.FILTER_RAW);
  ensureSheetWithHeader_(ss, CONFIG.SHEETS.FILTERS, CONFIG.HEADERS.FILTERS);
  const rows = []; let page = 1, totalPage = 1, apiRows = 0; const checkedAt = now_();
  try {
    do {
      const payload = { page: String(page), searchQuery: { searchKeyword: CONFIG.FILTER_PREFIX, siteId: '', filterGroup: 'all', sort: 'nameAsc' } };
      const resp = apiRequest_('filterList', 'post', payload); const data = extractData_(resp); const items = extractItems_(resp, false);
      totalPage = toNumber_(data.totalPage) || totalPage || 1;
      items.forEach(function(item) {
        const filterName = findFilterName_(item);
        if (!filterName || !String(filterName).trim().startsWith(CONFIG.FILTER_PREFIX)) return;
        const recent = getFilterRecentDateCandidate_(item); const created = getFilterCreateDateCandidate_(item);
        const account = accountFromFilterName_(filterName); const brand = brandFromFilterName_(filterName);
        rows.push([checkedAt, page, filterName, getAny_(item, ['filterId','filterID','id','seq','filterNo']), brand, account.accountNo, account.accountId, recent.value, created.value, recent.field, created.field, safeJsonStringify_(item)]);
      });
      apiRows += items.length; page += 1; Utilities.sleep(CONFIG.REQUEST_DELAY_MS);
    } while (page <= totalPage && page <= CONFIG.MAX_PAGE_PER_QUERY);

    replaceData_(getSheet_(CONFIG.SHEETS.FILTER_RAW), CONFIG.HEADERS.FILTER_RAW, rows);
    updateFilterSummaryDatesFromRaw_();
    const sheet = getSheet_(CONFIG.SHEETS.FILTER_RAW);
    try { sheet.showSheet(); ss.setActiveSheet(sheet); sheet.autoResizeColumns(1, Math.min(CONFIG.HEADERS.FILTER_RAW.length, 10)); } catch (e) {}
    log_('filter_raw_check', 'filterList 원본필드 확인 완료 / API행=' + apiRows + ', 롯백필터=' + rows.length);
    ui.alert('API 필터 원본필드 확인 완료\n\nAPI 조회 행수: ' + apiRows + '\n롯백 필터 수: ' + rows.length + '\ntotalPage: ' + totalPage + '\n\nAPI_필터원본확인 시트와 필터별_상품수 시트에 날짜 후보를 반영했습니다.');
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    log_('filter_raw_check_error', msg); ui.alert('API 필터 원본필드 확인 실패\n\n' + msg); throw e;
  }
}

function updateFilterSummaryDatesFromRaw_() {
  const rawMap = readFilterRawDateMap_(); const sheet = getSheet_(CONFIG.SHEETS.FILTERS); const currentRows = readTable_(sheet); const map = {};
  currentRows.forEach(function(r) { const filterName = String(r['검색필터명'] || '').trim(); if (!filterName) return; map[filterName] = r; });
  Object.keys(rawMap.byFilter).forEach(function(filterName) {
    const d = rawMap.byFilter[filterName];
    if (!map[filterName]) { const account = accountFromFilterName_(filterName); map[filterName] = { '검색필터명': filterName, '브랜드명': brandFromFilterName_(filterName), '계정번호': account.accountNo, '쿠팡계정ID': account.accountId, 'API_totalCount': '', 'API_totalPage': '', '이번조회_행수': '', '필터코드': filterCodeFromFilterName_(filterName), '메모': '' }; }
    map[filterName]['API_최근수집일자'] = d.recentDate || ''; map[filterName]['API_필터생성일'] = d.createDate || ''; map[filterName]['API_최근수집일자_필드'] = d.recentField || ''; map[filterName]['API_필터생성일_필드'] = d.createField || '';
  });
  writeFilterSummary_(map);
}

function readFilterRawDateMap_() {
  const map = { byFilter: {}, byBrand: {} };
  var rows = [];
  try { rows = readTable_(getSheet_(CONFIG.SHEETS.FILTER_RAW)); } catch (e) { return map; }
  rows.forEach(function(r) {
    const filterName = String(r['검색필터명'] || '').trim(); const brand = String(r['브랜드명'] || '').trim() || brandFromFilterName_(filterName);
    if (!filterName && !brand) return;
    const item = { filterName: filterName, brand: brand, recentDate: normalizeDateText_(r['최근수집일자_후보']), createDate: normalizeDateText_(r['필터생성일_후보']), recentField: String(r['최근수집일자_필드'] || '').trim(), createField: String(r['필터생성일_필드'] || '').trim() };
    if (filterName) map.byFilter[filterName] = item;
    if (brand) map.byBrand[normalizeBrandKey_(brand)] = item;
  });
  return map;
}

function applyFilterRawDatesToBrandMetrics_(brandMap) {
  const rawMap = readFilterRawDateMap_();
  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand]; const item = rawMap.byFilter[String(b.filterName || '').trim()] || rawMap.byBrand[normalizeBrandKey_(brand)];
    if (!item) return;
    b.filterRecentCollectDate = item.recentDate || ''; b.filterCreateDate = item.createDate || ''; b.filterRecentCollectField = item.recentField || ''; b.filterCreateField = item.createField || '';
    if (!b.representativeCoupangSendDate && b.coupangSentCount > 0) {
      if (b.filterRecentCollectDate) b.coupangSentStatus = '수동입력 전송수 확인/API_최근수집일자 기준 추정';
      else if (b.filterCreateDate) b.coupangSentStatus = '수동입력 전송수 확인/API_필터생성일 기준 추정';
    }
  });
}

function getFilterRecentDateCandidate_(obj) {
  const exact = findFirstDateCandidateWithField_(obj, ['recentCollectDate','recentCollectionDate','lastCollectDate','lastCollectionDate','collectDate','collectionDate','collectedAt','lastCrawlDate','crawlDate','scrapDate','scrapedAt','updateDate','updatedAt','modifyDate','modifiedAt','lastUpdateDate','lastUpdatedAt','recentDate','최근수집일자','최근수집일','최근수집일시','최근갱신일자','최근갱신일']);
  if (exact.value) return exact;
  return findDateByKeyHeuristic_(obj, ['collect','collection','crawl','scrap','update','updated','modify','modified','recent','last']);
}

function getFilterCreateDateCandidate_(obj) {
  const exact = findFirstDateCandidateWithField_(obj, ['filterCreateDate','filterCreatedAt','filterRegDate','createDate','createdAt','regDate','registerDate','registeredAt','insertDate','insertedAt','makeDate','madeAt','필터생성일','생성일','등록일자','등록일']);
  if (exact.value) return exact;
  return findDateByKeyHeuristic_(obj, ['create','created','reg','register','insert','made']);
}

function findFirstDateCandidateWithField_(obj, keys) {
  for (var i = 0; i < keys.length; i++) { const value = getNestedValueByKey_(obj, keys[i]); const normalized = normalizeDateText_(value); if (looksLikeDateText_(normalized)) return { value: normalized, field: keys[i] }; }
  return { value: '', field: '' };
}

function findDateByKeyHeuristic_(obj, keywords) {
  const found = []; scanObjectForDateFields_(obj, '', keywords, found);
  found.sort(function(a, b) { const sa = heuristicDateFieldScore_(a.field), sb = heuristicDateFieldScore_(b.field); if (sa !== sb) return sb - sa; return String(b.value).localeCompare(String(a.value)); });
  return found.length ? found[0] : { value: '', field: '' };
}

function scanObjectForDateFields_(obj, path, keywords, out) {
  if (!obj || typeof obj !== 'object') return;
  if (Array.isArray(obj)) { obj.forEach(function(v, i) { scanObjectForDateFields_(v, path + '[' + i + ']', keywords, out); }); return; }
  Object.keys(obj).forEach(function(k) {
    const value = obj[k]; const fullPath = path ? path + '.' + k : k; const lk = String(k).toLowerCase();
    const keywordHit = (keywords || []).some(function(keyword) { return lk.indexOf(String(keyword).toLowerCase()) >= 0; });
    if (keywordHit) { const normalized = normalizeDateText_(value); if (looksLikeDateText_(normalized)) out.push({ value: normalized, field: fullPath }); }
    if (value && typeof value === 'object') scanObjectForDateFields_(value, fullPath, keywords, out);
  });
}

function heuristicDateFieldScore_(field) { const f = String(field || '').toLowerCase(); let score = 0; if (f.indexOf('recent') >= 0 || f.indexOf('last') >= 0) score += 5; if (f.indexOf('collect') >= 0 || f.indexOf('collection') >= 0) score += 4; if (f.indexOf('update') >= 0 || f.indexOf('modify') >= 0) score += 4; if (f.indexOf('create') >= 0 || f.indexOf('reg') >= 0 || f.indexOf('insert') >= 0) score += 3; if (f.indexOf('date') >= 0 || f.indexOf('at') >= 0) score += 2; return score; }

function getNestedValueByKey_(obj, key) {
  if (!obj || typeof obj !== 'object') return '';
  if (obj[key] !== undefined && obj[key] !== null && String(obj[key]).trim() !== '') return obj[key];
  const lowerKey = String(key).toLowerCase();
  for (const k in obj) { if (String(k).toLowerCase() === lowerKey && obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== '') return obj[k]; }
  for (const k in obj) { const v = obj[k]; if (v && typeof v === 'object') { const found = getNestedValueByKey_(v, key); if (found !== undefined && found !== null && String(found).trim() !== '') return found; } }
  return '';
}

function looksLikeDateText_(v) { return /(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/.test(String(v || '').trim()); }
function safeJsonStringify_(obj) { try { return JSON.stringify(obj); } catch (e) { return String(obj || ''); } }


function ensureCoupangWorkLogSheet_() {
  const ss = SpreadsheetApp.getActive(); const sheet = ensureSheet_(ss, CONFIG.SHEETS.RETRANSMIT_LOG); const headers = CONFIG.HEADERS.RETRANSMIT_LOG;
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) { sheet.getRange(1, 1, 1, headers.length).setValues([headers]); sheet.setFrozenRows(1); return sheet; }
  const oldLastRow = sheet.getLastRow(), oldLastCol = Math.max(sheet.getLastColumn(), 1);
  const oldHeaderDisplay = sheet.getRange(1, 1, 1, oldLastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  const oldHeaders = oldHeaderDisplay.map(function(h) { return normalizeHeaderKey_(h); });
  const targetHeaders = headers.map(function(h) { return normalizeHeaderKey_(h); });
  const same = oldHeaders.length === targetHeaders.length && oldHeaders.every(function(h, i) { return h === targetHeaders[i]; });
  if (same) { sheet.getRange(1, 1, 1, headers.length).setValues([headers]); sheet.setFrozenRows(1); return sheet; }
  const oldValues = oldLastRow >= 2 ? sheet.getRange(2, 1, oldLastRow - 1, oldLastCol).getValues() : [];
  const migrated = oldValues.map(function(row) {
    const obj = {}; oldHeaders.forEach(function(h, i) { obj[h] = row[i]; });
    return targetHeaders.map(function(h) { return getCoupangWorkLogValueByAlias_(obj, h); });
  }).filter(function(row) { return row.some(function(v) { return String(v || '').trim() !== ''; }); });
  sheet.clearContents(); sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (migrated.length) sheet.getRange(2, 1, migrated.length, headers.length).setValues(migrated);
  sheet.setFrozenRows(1); return sheet;
}

function getCoupangWorkLogValueByAlias_(obj, header) {
  const cleanHeader = normalizeHeaderKey_(header);
  const aliases = {
    '검색필터명': ['검색필터명'], '브랜드명': ['브랜드명'], '쿠팡계정ID': ['쿠팡계정ID'],
    '최초전송일': ['최초전송일','최초쿠팡전송일','대표쿠팡전송일'],
    '누적매출액': ['누적매출액','누적매출','매출액'], '월평균매출액': ['월평균매출액','월평균매출'],
    '작업유형': ['작업유형'],
    '1단계_추가수집일': ['1단계_추가수집일','추가수집일'],
    '2단계_재전송일': ['2단계_재전송일','재전송일','기준일'],
    '3단계_정리재수집일': ['3단계_정리재수집일','정리재수집일'],
    '최종_더망고수집수': ['최종_더망고수집수'], '최종_쿠팡전송수': ['최종_쿠팡전송수'],
    '추가수집상품수': ['추가수집상품수'], '작업사유': ['작업사유'], '비고': ['비고']
  };
  const keys = aliases[cleanHeader] || [cleanHeader, header];
  for (var i = 0; i < keys.length; i++) { const v = getObjectValueByHeader_(obj, keys[i]); if (String(v || '').trim() !== '') return v; }
  return '';
}

function replaceDataFastLimited_(sheet, headers, rows) {
  const data = [headers].concat(rows || []);
  const cols = headers.length;
  const newRows = data.length;

  // v5.98:
  // 과거 실행/서식 때문에 getLastRow가 크게 잡히는 경우 전체 삭제가 시간초과 원인이 됩니다.
  // 현재 출력 데이터 주변만 정리하고, 실제 데이터는 필요한 범위에만 씁니다.
  const oldRows = Math.min(Math.max(sheet.getLastRow(), newRows), Math.max(newRows + 200, 1000));

  try {
    sheet.getRange(1, 1, oldRows, cols).clearContent();
  } catch (e) {}

  if (newRows && cols) {
    sheet.getRange(1, 1, newRows, cols).setValues(data);
  }

  try { sheet.setFrozenRows(1); } catch (e) {}
}


function getFilterListResumeState_() {
  const raw = PropertiesService.getScriptProperties().getProperty('FILTERLIST_RESUME_STATE') || '';
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function setFilterListResumeState_(state) {
  PropertiesService.getScriptProperties().setProperty('FILTERLIST_RESUME_STATE', JSON.stringify(state || {}));
}

function clearFilterListResumeState_() {
  PropertiesService.getScriptProperties().deleteProperty('FILTERLIST_RESUME_STATE');
}

function getFilterListResumeSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheet_(ss, '필터별_상품수_TMP');
  const headers = CONFIG.HEADERS.FILTERS;
  if (sheet.getLastRow() < 1) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function readFilterListResumeRows_() {
  const sheet = getFilterListResumeSheet_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.HEADERS.FILTERS.length).getValues();
}

function appendFilterListResumeRows_(rows) {
  if (!rows || !rows.length) return;
  const sheet = getFilterListResumeSheet_();
  const start = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(start, 1, rows.length, CONFIG.HEADERS.FILTERS.length).setValues(rows);
}

function dedupeFilterListRows_(rows) {
  const map = {};
  (rows || []).forEach(function(row) {
    const name = String(row[0] || '').trim();
    if (!name) return;
    map[name] = row;
  });
  return Object.keys(map).sort().map(function(k) { return map[k]; });
}

function resetFilterListResumeWork_() {
  clearFilterListResumeState_();
  const sheet = getFilterListResumeSheet_();
  try { sheet.clearContents(); } catch (e) {}
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.FILTERS.length).setValues([CONFIG.HEADERS.FILTERS]);
}


function writeFilterSummaryFast_(filterSummaryMap) {
  // v5.58:
  // 빠른 갱신에서는 API_필터원본확인/RAW JSON 시트를 읽지 않습니다.
  // 날짜 후보는 filterList에서 방금 가져온 meta만 사용합니다.
  const rows = Object.keys(filterSummaryMap || {}).sort().map(function(k) {
    const r = filterSummaryMap[k];
    return [
      r['검색필터명'],
      r['브랜드명'],
      r['계정번호'],
      r['쿠팡계정ID'],
      r['API_totalCount'],
      r['API_totalPage'],
      r['이번조회_행수'],
      r['필터코드'],
      r['API_최근수집일자'] || '',
      r['API_필터생성일'] || '',
      r['API_최근수집일자_필드'] || '',
      r['API_필터생성일_필드'] || '',
      r['메모'] || ''
    ];
  });

  const sheet = getSheet_(CONFIG.SHEETS.FILTERS);
  replaceDataFastLimited_(sheet, CONFIG.HEADERS.FILTERS, rows);
  formatFilterSummarySheetFast_(sheet);
}

function formatFilterSummarySheetFast_(sheet) {
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = CONFIG.HEADERS.FILTERS.length;
  const dataRows = Math.max(lastRow - 1, 0);

  try {
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#e8f0fe')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 42);
  } catch (e) {}

  if (dataRows > 0) {
    try {
      sheet.getRange(2, 1, dataRows, lastCol).setWrap(false).setVerticalAlignment('middle');
      sheet.getRange(2, 3, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 5, dataRows, 3).setNumberFormat('#,##0');
      sheet.getRange(2, 8, dataRows, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 9, dataRows, 2).setNumberFormat('M/dd');
    } catch (e) {}
  }

  const widths = [190, 130, 80, 105, 105, 90, 90, 80, 100, 100, 120, 120, 280];
  widths.forEach(function(w, i) {
    try { sheet.setColumnWidth(i + 1, w); } catch (e) {}
  });
}

function writeFilterSummary_(filterSummaryMap) {
  const filterDateMap = readFilterRawDateMap_();
  const rows = Object.keys(filterSummaryMap).sort().map(function(k) {
    const r = filterSummaryMap[k]; const d = filterDateMap.byFilter[String(r['검색필터명'] || '').trim()] || {};
    return [r['검색필터명'], r['브랜드명'], r['계정번호'], r['쿠팡계정ID'], r['API_totalCount'], r['API_totalPage'], r['이번조회_행수'], r['필터코드'], d.recentDate || r['API_최근수집일자'] || '', d.createDate || r['API_필터생성일'] || '', d.recentField || r['API_최근수집일자_필드'] || '', d.createField || r['API_필터생성일_필드'] || '', r['메모'] || ''];
  });
  replaceData_(getSheet_(CONFIG.SHEETS.FILTERS), CONFIG.HEADERS.FILTERS, rows);
}

function writeSyncStatus_(info) {
  const props = PropertiesService.getScriptProperties(); info = info || {};
  const token = getApiTokenUsageSnapshot_();
  const phase = info.phase !== undefined ? info.phase : (props.getProperty('SYNC_PHASE') || CONFIG.PHASE.IDLE);
  const status = info.status !== undefined ? info.status : (props.getProperty('SYNC_LAST_MESSAGE') || '');
  const filterCount = info.filterCount !== undefined ? toNumber_(info.filterCount) : JSON.parse(props.getProperty('SYNC_FILTER_NAMES') || '[]').length;
  const filterIndex = info.filterIndex !== undefined ? toNumber_(info.filterIndex) : toNumber_(props.getProperty('SYNC_FILTER_INDEX'));
  const page = info.page !== undefined ? toNumber_(info.page) : toNumber_(props.getProperty('SYNC_PAGE'));
  const currentFilter = info.currentFilter !== undefined ? info.currentFilter : (props.getProperty('SYNC_CURRENT_FILTER') || '');
  const totalPage = info.totalPage !== undefined ? info.totalPage : (props.getProperty('SYNC_CURRENT_TOTAL_PAGE') || '');
  const tmpRows = info.tmpRows !== undefined ? info.tmpRows : '';
  const lastUrl = info.lastUrl !== undefined ? info.lastUrl : (props.getProperty('LAST_API_URL') || '');
  const lastError = info.lastError !== undefined ? info.lastError : (props.getProperty('SYNC_LAST_ERROR') || '');
  const currentNo = filterCount ? Math.min(filterIndex + 1, filterCount) : '';
  const progress = filterCount ? filterIndex / filterCount : '';
  const rows = [
    ['단계(PHASE)', phase, 'IDLE/COLLECT_FILTERS/COLLECT_PRODUCTS/FINALIZE/DONE/ERROR'],
    ['상태', status, '진행중/완료/오류'],
    ['전체 진행률', progress, filterCount ? ('필터 ' + filterIndex + '/' + filterCount + ' 완료 기준') : ''],
    ['필터 진행', filterCount ? (currentNo + ' / ' + filterCount) : '', '현재 처리 중인 필터 번호'],
    ['현재필터명', currentFilter, '현재 API 조회 중인 검색필터명'],
    ['현재페이지', page || '', 'productList page'],
    ['현재필터 전체페이지', totalPage, 'API totalPage가 확인된 경우 표시'],
    ['누적 임시행수', tmpRows, 'LOTTEON_상품목록_TMP 누적 행수'],
    ['최근 호출 URL', lastUrl, '인증값은 표시하지 않음'],
    ['최근 호출 경로', props.getProperty('LAST_API_ROUTE') || '', 'DIRECT 또는 GATEWAY'],
    ['최근 오류', lastError, '오류 발생 시 원문 표시'],
    ['예상 API 토큰(이번 동기화)', formatCount_(token.syncUsed), '이 구스에서 이번 상품목록 동기화 중 성공 호출 기준'],
    ['예상 API 토큰(오늘/구스)', formatCount_(token.todayUsed) + ' / ' + formatCount_(CONFIG.API_TOKEN_DAILY_LIMIT), '2026-05-18 기준값 3 포함. 더망고 전체와 차이 가능'],
    ['예상 잔여 토큰(오늘/구스)', formatCount_(token.todayRemain), '오늘 한도에서 위 예상 사용량을 차감한 값'],
    ['예상 API 토큰(이번달/구스)', formatCount_(token.monthUsed) + ' / ' + formatCount_(CONFIG.API_TOKEN_MONTHLY_LIMIT), '2026-05 기준값 4,504 포함. 2026-06부터 기준값 0'],
    ['예상 잔여 토큰(이번달/구스)', formatCount_(token.monthRemain), '월 한도에서 위 예상 사용량을 차감한 값'],
    ['최근 API 서비스', token.lastEndpoint, '마지막으로 성공한 API endpoint'],
    ['시작시각', props.getProperty('SYNC_STARTED_AT') || '', 'Asia/Seoul'],
    ['최근갱신', now_(), 'Asia/Seoul'],
    ['완료시각', props.getProperty('SYNC_FINISHED_AT') || '', '완료 시각']
  ];
  const sheet = getSheet_(CONFIG.SHEETS.SYNC_STATUS);
  replaceData_(sheet, CONFIG.HEADERS.SYNC_STATUS, rows);
  sheet.getRange(1, 2, sheet.getLastRow(), 1).setNumberFormat('@');
  sheet.getRange(3, 2).setNumberFormat('0.00%');
}

function writeSettings_() {
  const values = [
    ['분석대상', CONFIG.TARGET_SOURCE, '구매사이트명에 lotteon 포함된 매출만'],
    ['검색필터 prefix', CONFIG.FILTER_PREFIX, '롯백으로 시작하는 검색필터만 수집'], ['v5.51 변경', '빠른갱신 대시보드 시간초과 방지', '필터/상품수만 빠른 갱신에서 대시보드 차트삭제/셀별서식루프/sheet.clear를 제거하고 경량 갱신'], ['v5.52 변경', 'API URL fallback', '저장된 THE_MANGO_BASE_URL의 filterList 호출 실패 시 기본 API URL(tmg2007.cafe24.com)을 자동 재시도'], ['v5.53 변경', '필터명 변경 자동해결', '수동입력 필터명이 과거명이어도 동일 브랜드 최신 API 필터로 자동 매칭하고 롯백_00은 제외'], ['v5.54 변경', '빠른갱신 작업분리', '필터/상품수만 빠른 갱신은 필터별_상품수+쿠팡재전송_로그만 처리하고 검증/대시보드는 별도 메뉴로 분리'], ['v5.55 변경', '검증 문제행만 출력', '수동입력/API 검증 기본 메뉴는 일치/비활성행을 제외한 문제행만 출력하고 sheet.clearContents 없이 초경량 갱신'], ['v5.56 변경', '수동입력 시트 경량 갱신', '쿠팡전송수_수동입력 시트 생성/갱신에서 autoResizeColumns를 제거하고 고정폭 서식만 적용'], ['v5.57 변경', '대시보드 조치요약판', '대시보드를 단순 필터 목록이 아니라 API미확인/수동입력없음/수집수차이/필터변경 우선순위 요약판으로 재구성'], ['v5.58 변경', '빠른갱신 초경량화', '필터/상품수만 빠른 갱신에서 상태시트 반복갱신/TMP행수조회/RAW필드조회/대시보드·검증 보장을 제거'], ['v5.59 변경', '빠른갱신 분할실행', '필터별 상품수 조회를 8개씩 나누어 자동 트리거로 이어 실행하고 완료 시 필터별_상품수/쿠팡재전송_로그 반영'],
    ['v5.49 핵심수정', '빠른 갱신 4개 시트 동기화 보장', '필터별_상품수 → 쿠팡재전송_로그 → 수동입력_API검증 → 대시보드'],
    ['v5.49 신규', 'writeManualInputApiValidationFast_', 'LOTTEON_상품목록 읽지 않는 경량 검증'],
    ['v5.50 수정', '수동입력/API 검증 메뉴 경량화', '메뉴 실행 시 필터별_상품수 + 쿠팡전송수_수동입력만 비교'],
    ['v5.49 신규', 'buildDashboardFastFromFilterSummary_', '필터/상품수 기반 경량 대시보드'],
    ['v5.49 빠른갱신 금지', 'LOTTEON_상품목록 전체 읽기 / 매출분석 / 전체 서식', '6분 이내 완료 보장'],
    ['v5.49 단일 원천', '필터별_상품수 시트', '검색필터명/브랜드명/계정번호/API_totalCount 최신 기준'],
    ['v5.48 추가', '필터/상품수 빠른 갱신', 'filterList + productList 1페이지 totalCount 병렬 조회'],
    ['API 토큰 일 한도', CONFIG.API_TOKEN_DAILY_LIMIT, '더망고 관리자 화면 실제 기준은 관리자 화면에서 확인'],
    ['API 토큰 월 한도', CONFIG.API_TOKEN_MONTHLY_LIMIT, '더망고 관리자 화면 실제 기준은 관리자 화면에서 확인'],
    ['계정매핑_롯백_01', CONFIG.ACCOUNT_MAP['01'].accountId, '1번 계정'],
    ['계정매핑_롯백_02', CONFIG.ACCOUNT_MAP['02'].accountId, '2번 계정'],
    ['계정매핑_롯백_03', CONFIG.ACCOUNT_MAP['03'].accountId, '3번 계정'],
    ['계정매핑_롯백_04', CONFIG.ACCOUNT_MAP['04'].accountId, '4번 계정'],
    ['매출 매칭 우선순위', '①사이트상품번호 ②브랜드+상품명 ③브랜드직접 ④매칭표 ⑤상품명브랜드추정', '취소/반품 포함 전체 주문/매출을 브랜드성과에 반영'],
    ['브랜드 매칭표', CONFIG.SHEETS.BRAND_ALIAS, '매출브랜드명 → 더망고브랜드명 수동 매핑'],
    ['주문상태 검사', '더망고주문상태', '취소/반품/교환 포함 여부는 정보성 취소지표로만 분리. 주 분석 매출에는 포함'],
    ['개인정보 컬럼', '매출데이터_정리에 포함하지 않음', CONFIG.PERSONAL_INFO_BLOCKLIST.join(', ')]
  ];
  replaceData_(getSheet_(CONFIG.SHEETS.SETTINGS), CONFIG.HEADERS.SETTINGS, values);
}

function installDailyAnalysisTrigger() { removeLotteonTriggers(); ScriptApp.newTrigger('refreshLotteonAnalysis').timeBased().everyDays(1).atHour(7).create(); log_('trigger','매일 오전 7시 매출/분석 자동 업데이트 설치'); SpreadsheetApp.getUi().alert('매일 오전 7시 매출/분석 자동 업데이트를 설치했습니다.'); }

function removeLotteonTriggers() { deleteTriggersForHandler_('refreshLotteonAnalysis'); deleteTriggersForHandler_('continueLotteonProductSync'); deleteTriggersForHandler_('runAnalysisAfterSync'); log_('trigger','LOTTEON 관련 트리거 해제'); }

function judgeMonthly_(revenue, orders, sellRate) { if (orders >= CONFIG.GOOD_ORDER_30D || revenue >= CONFIG.GOOD_REVENUE_30D) return '월성과_좋음'; if (orders > 0 || revenue > 0) return '월성과_관찰'; return '월성과_없음'; }
function judgeBrand_(b, sellDays, sellRate, revenue30d) { if (b.orderCount > 0 || b.revenue > 0) { if (revenue30d >= CONFIG.GOOD_REVENUE_30D || b.orderCount >= CONFIG.GOOD_ORDER_30D || sellRate >= CONFIG.MIN_SELL_RATE_KEEP) return { label: '유지/확대', reason: '매출 발생 + 30일환산매출/주문/판매상품률 기준 양호' }; return { label: '유지/관찰', reason: '매출은 있으나 규모가 작아 1~2개월 추가 관찰' }; } if (sellDays < CONFIG.MIN_OBSERVE_DAYS) return { label: '신규 관찰', reason: '판매기간이 짧아 무성과 판단 보류' }; if (sellDays >= CONFIG.DELETE_REVIEW_DAYS && b.productCount >= CONFIG.VERY_LARGE_PRODUCT_COUNT) return { label: '삭제 1순위', reason: '판매기간 충분 + 대량 상품 + 매출 없음' }; if (sellDays >= CONFIG.DELETE_REVIEW_DAYS && b.productCount >= CONFIG.LARGE_PRODUCT_COUNT) return { label: '삭제/압축 검토', reason: '판매기간 충분 + 상품수 대비 매출 없음' }; return { label: '보류/관찰', reason: '매출은 없으나 상품수/판매기간 기준 즉시삭제 우선순위는 낮음' }; }
function deletePriority_(label) { if (label === '삭제 1순위') return 1; if (label === '삭제/압축 검토') return 2; if (label === '전송정보 확인필요') return 60; if (label === '유지/관찰') return 70; if (label === '신규 관찰') return 80; if (label === '유지/확대') return 90; return 50; }
function recommendAccountAction_(label) { if (label === '유지/확대') return '계속 유지/확대 후보'; if (label === '유지/관찰') return '1~2개월 경고메일/매출 관찰'; if (label === '신규 관찰') return '판매기간 확보 후 판단'; if (label === '전송정보 확인필요') return '쿠팡전송수/전송일 입력 후 판단'; if (label.indexOf('삭제') >= 0) return '신규계정 투입 중단/삭제 검토'; return '보류'; }
function accountMemo_(decision, b) { return decision.reason + ' / 현재 매핑 계정: ' + (b.accountId || '미확인'); }


function getCredentials_() { const props = PropertiesService.getScriptProperties(); const baseUrl = normalizeBaseUrl_(props.getProperty('THE_MANGO_BASE_URL') || CONFIG.BASE_URL_DEFAULT); const apiKey = String(props.getProperty('THE_MANGO_API_KEY') || '').trim(); const sender = String(props.getProperty('THE_MANGO_SENDER') || '').trim(); if (!apiKey || !sender) throw new Error('API 인증값이 없습니다. ② API 인증값 저장을 먼저 실행하세요.'); return { baseUrl: baseUrl, apiKey: apiKey, sender: sender }; }

function apiRequest_(endpoint, method, payload) {
  const cred = getCredentials_();
  const urls = buildApiUrlCandidates_(endpoint);
  const props = PropertiesService.getScriptProperties();
  const options = {
    method: method || 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + cred.apiKey, 'X-API-SENDER': cred.sender },
    muteHttpExceptions: true
  };
  if (payload && String(method).toLowerCase() !== 'get') options.payload = JSON.stringify(payload);

  const primaryBase = normalizeBaseUrl_(cred.baseUrl);
  const defaultBase = normalizeBaseUrl_(CONFIG.BASE_URL_DEFAULT);
  const errors = [];

  for (var u = 0; u < urls.length; u++) {
    const url = urls[u];
    const routeLabel = getApiRouteLabel_(url, primaryBase, defaultBase, u);
    props.setProperty('LAST_API_URL', maskSensitiveUrl_(url));

    let last = '';
    for (var attempt = 1; attempt <= 5; attempt++) {
      let resp;
      try {
        resp = UrlFetchApp.fetch(url, options);
      } catch (fetchError) {
        const msg = String(fetchError && fetchError.message ? fetchError.message : fetchError);
        errors.push(routeLabel + ' URL_FETCH_ERROR: ' + msg + ' / 호출URL=' + maskSensitiveUrl_(url));
        break;
      }

      const code = resp.getResponseCode();
      const text = resp.getContentText();
      last = text;

      if (code === 429) {
        Utilities.sleep(3000 * attempt);
        continue;
      }

      if (code < 200 || code >= 300) {
        const msg = routeLabel + ' ' + endpoint + ' HTTP ' + code + ': ' + text.slice(0, 800);
        errors.push(msg);
        // 401/403은 인증 문제일 가능성이 높지만, 저장 URL 문제일 수도 있어서 다음 후보 URL까지는 확인합니다.
        break;
      }

      try {
        const parsed = JSON.parse(text);
        recordApiTokenUsage_(endpoint);
        props.setProperty('LAST_API_ROUTE', routeLabel);
        props.setProperty('LAST_API_URL', maskSensitiveUrl_(url));

        // v5.52: 기본 API URL로 성공했으면 이후 실행도 안정적으로 기본 URL을 사용하게 보정
        if (routeLabel === 'DEFAULT_FALLBACK') {
          props.setProperty('THE_MANGO_LAST_WORKING_BASE_URL', defaultBase);
        }

        return parsed;
      } catch (e) {
        throw new Error(routeLabel + ' ' + endpoint + ' 응답 JSON 파싱 실패: ' + text.slice(0, 800));
      }
    }

    if (last) errors.push(routeLabel + ' 마지막 응답: ' + String(last).slice(0, 300));
  }

  throw new Error(endpoint + ' API 호출 실패. 시도한 경로: ' + errors.join(' || '));
}

function getApiRouteLabel_(url, primaryBase, defaultBase, index) {
  const s = String(url || '');
  const p = String(primaryBase || '').replace(/\/+$/, '');
  const d = String(defaultBase || '').replace(/\/+$/, '');

  if (p && s.indexOf(p + '/api/') === 0) return 'DIRECT';
  if (d && s.indexOf(d + '/api/') === 0) return 'DEFAULT_FALLBACK';
  if (s.indexOf(CONFIG.API_GATEWAY_BASE_URL) === 0) return 'GATEWAY';
  return index === 0 ? 'DIRECT' : 'FALLBACK_' + (index + 1);
}

function extractData_(resp) { if (resp && typeof resp === 'object' && resp.data && typeof resp.data === 'object') return resp.data; return resp || {}; }
function extractItems_(resp, preferItems) { const data = extractData_(resp); if (preferItems && data.items && Array.isArray(data.items)) return data.items.filter(isObject_); const arrays = []; collectArraysOfObjects_(data, arrays); if (!arrays.length) collectArraysOfObjects_(resp, arrays); if (!arrays.length) return []; arrays.sort(function(a, b) { return b.length - a.length; }); return arrays[0]; }
function collectArraysOfObjects_(obj, arrays) { if (Array.isArray(obj)) { if (obj.length && obj.every(isObject_)) arrays.push(obj); obj.forEach(function(v) { collectArraysOfObjects_(v, arrays); }); } else if (isObject_(obj)) Object.keys(obj).forEach(function(k) { collectArraysOfObjects_(obj[k], arrays); }); }
function isObject_(v) { return v && typeof v === 'object' && !Array.isArray(v); }

function ensureSheet_(ss, name) { let sheet = ss.getSheetByName(name); if (!sheet) sheet = ss.insertSheet(name); return sheet; }
function ensureSheetWithHeader_(ss, name, headers) {
  const sheet = ensureSheet_(ss, name); let existing = [];
  if (sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) existing = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), headers.length)).getValues()[0].map(function(v) { return String(v || '').trim(); });
  const target = headers.map(function(v) { return String(v || '').trim(); });
  const same = existing.length === target.length && existing.every(function(v, i) { return v === target[i]; });
  if (!same) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1); return sheet;
}
function getSheet_(name) { return ensureSheet_(SpreadsheetApp.getActive(), name); }
function replaceData_(sheet, headers, rows) { sheet.clearContents(); const data = [headers].concat(rows || []); if (data.length && headers.length) sheet.getRange(1, 1, data.length, headers.length).setValues(data); sheet.setFrozenRows(1); }
function appendRows_(sheet, rows) { if (!rows || !rows.length) return; const start = Math.max(sheet.getLastRow() + 1, 2); sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows); }

function getTrimmedSheetValues_(sheet, maxRows, maxCols) {
  if (!sheet) return [];
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const rowLimit = Math.max(1, Math.min(lastRow, maxRows || 3000));
  const colProbeLimit = Math.max(1, Math.min(lastCol, maxCols || 80));

  const headerRow = sheet.getRange(1, 1, 1, colProbeLimit).getValues()[0];
  let usedCols = 0;
  for (var c = 0; c < headerRow.length; c++) {
    if (String(headerRow[c] || '').trim() !== '') usedCols = c + 1;
  }
  if (usedCols <= 0) usedCols = colProbeLimit;

  return sheet.getRange(1, 1, rowLimit, usedCols).getValues();
}
function normalizeHeaderKey_(h) { return String(h || '').replace(/\s+/g, '').trim(); }
function getObjectValueByHeader_(obj, header) {
  const target = normalizeHeaderKey_(header);
  if (obj[header] !== undefined) return obj[header];
  if (obj[target] !== undefined) return obj[target];
  const keys = Object.keys(obj || {});
  for (var i = 0; i < keys.length; i++) { if (normalizeHeaderKey_(keys[i]) === target) return obj[keys[i]]; }
  return '';
}
function readTable_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = getTrimmedSheetValues_(sheet, 3000, 80);
  if (!values || values.length < 2) return [];

  const headers = values[0].map(function(h) { return normalizeHeaderKey_(h); });
  const rows = [];

  for (var r = 1; r < values.length; r++) {
    const row = values[r];
    const obj = {};
    let nonEmpty = false;

    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = row[c];
      if (String(row[c] || '').trim() !== '') nonEmpty = true;
    }

    if (nonEmpty) rows.push(obj);
  }

  return rows;
}
function getValuesWithoutHeader_(sheetName) { const sheet = getSheet_(sheetName); if (sheet.getLastRow() < 2) return []; return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues(); }

function formatCoreAnalysisSheetsLightOnly_() {
  // v5.66:
  // ⑧ 분석용 초경량 서식입니다.
  // 전체 자동 열너비 계산, 대시보드 서식, getValue 반복 호출을 하지 않습니다.
  const targetNames = [
    CONFIG.SHEETS.SALES_CLEAN,
    CONFIG.SHEETS.BRAND,
    CONFIG.SHEETS.BRAND_SUMMARY,
    CONFIG.SHEETS.MATCH_DIAG,
    CONFIG.SHEETS.RETRANSMIT_LOG
  ];

  targetNames.forEach(function(name) {
    const sheet = getSheet_(name);
    if (!sheet) return;

    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const dataRows = Math.max(lastRow - 1, 0);

    try {
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, lastCol)
        .setFontWeight('bold')
        .setBackground('#e8f0fe')
        .setWrap(true)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');
      sheet.setRowHeight(1, 42);
    } catch (e) {}

    if (dataRows <= 0) return;

    try {
      sheet.getRange(2, 1, dataRows, lastCol)
        .setWrap(false)
        .setVerticalAlignment('middle');
    } catch (e) {}

    try {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (var c = 0; c < headers.length; c++) {
        const h = String(headers[c] || '').replace(/\s+/g, '');
        const range = sheet.getRange(2, c + 1, dataRows, 1);
        if (isPercentHeader_(h)) range.setNumberFormat('0.00%');
        else if (isMoneyHeader_(h)) range.setNumberFormat('"₩"#,##0');
        else if (isCountHeader_(h)) range.setNumberFormat('#,##0');
        else if (isDateHeader_(h)) range.setNumberFormat('M/dd');
      }
    } catch (e) {}
  });

  setCoreSheetFixedWidthsLight_();
}

function setCoreSheetFixedWidthsLight_() {
  const ss = SpreadsheetApp.getActive();
  const widths = {};
  widths[CONFIG.SHEETS.BRAND_SUMMARY] = [60, 95, 130, 190, 105, 80, 80, 75, 70, 95, 80, 105, 85, 95, 220, 80, 80, 80, 80, 220, 80, 75, 210, 80, 75, 90, 210, 80, 75, 90, 210, 70, 95];
  widths[CONFIG.SHEETS.MATCH_DIAG] = [180, 120, 90, 360];
  widths[CONFIG.SHEETS.RETRANSMIT_LOG] = [190, 130, 105, 75, 105, 105, 90, 85, 85, 90, 105, 105, 100, 150, 230];

  Object.keys(widths).forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    widths[name].forEach(function(w, i) {
      try { sheet.setColumnWidth(i + 1, w); } catch (e) {}
    });
  });
}

function formatCoreAnalysisSheetsOnly_() {
  var targetNames = [CONFIG.SHEETS.SALES_CLEAN, CONFIG.SHEETS.BRAND, CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.SHEETS.MATCH_DIAG, CONFIG.SHEETS.DASHBOARD];
  targetNames.forEach(function(name) { formatOneSheetFast_(getSheet_(name)); });
}

function formatAnalysisSheetsOnly_() {
  var targetNames = [CONFIG.SHEETS.SALES_CLEAN, CONFIG.SHEETS.MONTHLY_BRAND, CONFIG.SHEETS.BRAND, CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.SHEETS.DELETE, CONFIG.SHEETS.ACCOUNT, CONFIG.SHEETS.SENT_CHECK, CONFIG.SHEETS.MATCH_DIAG, CONFIG.SHEETS.DASHBOARD];
  targetNames.forEach(function(name) { formatOneSheetFast_(getSheet_(name)); });
}

function formatAllSheets_() { SpreadsheetApp.getActive().getSheets().forEach(function(sheet) { formatOneSheetFast_(sheet); }); }

function formatOneSheetFast_(sheet) {
  if (!sheet) return;
  const name = sheet.getName(); const lastRow = Math.max(sheet.getLastRow(), 1); const lastCol = Math.max(sheet.getLastColumn(), 1);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#e8f0fe').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  try { sheet.setRowHeight(1, 42); } catch (e) {}
  if (name === CONFIG.SHEETS.DASHBOARD) { formatDashboard_(sheet); return; }
  if (lastRow >= 2) {
    try { sheet.getRange(2, 1, lastRow - 1, lastCol).setNumberFormat('General'); } catch (e) {}
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < headers.length; c++) {
      const h = String(headers[c] || '').replace(/\s+/g, '');
      const range = sheet.getRange(2, c + 1, Math.max(lastRow - 1, 1), 1);
      if (isPercentHeader_(h)) range.setNumberFormat('0.00%');
      else if (isCountHeader_(h)) range.setNumberFormat('#,##0');
      else if (isMoneyHeader_(h)) range.setNumberFormat('"₩"#,##0');
      else if (isDateHeader_(h)) range.setNumberFormat('M/dd');
    }
    if (name === CONFIG.SHEETS.BRAND_SUMMARY) formatBrandSummarySheet_(sheet);
    else if (name === CONFIG.SHEETS.RETRANSMIT_LOG) formatCoupangWorkLogSheet_(sheet);
  }
  fitColumnsByHeaderAndData_(sheet, getStandardColumnFitOptions_());
}

function formatCoupangWorkLogSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1); const lastCol = Math.max(sheet.getLastColumn(), 1); const dataRows = Math.max(lastRow - 1, 0);
  try { sheet.getRange(1, 1, 1, lastCol).setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center'); sheet.setRowHeight(1, 54); } catch (e) {}
  if (dataRows > 0) {
    safeSetNumberFormat_(sheet, 4, dataRows, 'M/dd'); safeSetNumberFormat_(sheet, 5, dataRows, '"₩"#,##0'); safeSetNumberFormat_(sheet, 6, dataRows, '"₩"#,##0');
    safeSetNumberFormat_(sheet, 8, dataRows, 'M/dd'); safeSetNumberFormat_(sheet, 9, dataRows, 'M/dd'); safeSetNumberFormat_(sheet, 10, dataRows, 'M/dd');
    safeSetNumberFormat_(sheet, 11, dataRows, '#,##0'); safeSetNumberFormat_(sheet, 12, dataRows, '#,##0'); safeSetNumberFormat_(sheet, 13, dataRows, '#,##0');
    try { sheet.getRange(2, 1, dataRows, lastCol).setWrap(false).setVerticalAlignment('middle'); } catch (e) {}
  }
  fitColumnsByHeaderAndData_(sheet, getStandardColumnFitOptions_());
}

function formatBrandSummarySheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1); const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return;
  const dataRows = lastRow - 1;
  safeSetNumberFormat_(sheet, 1, dataRows, '#,##0'); safeSetNumberFormat_(sheet, 6, dataRows, '#,##0'); safeSetNumberFormat_(sheet, 7, dataRows, '#,##0'); safeSetNumberFormat_(sheet, 8, dataRows, '0.00%'); safeSetNumberFormat_(sheet, 9, dataRows, '#,##0'); safeSetNumberFormat_(sheet, 10, dataRows, '"₩"#,##0'); safeSetNumberFormat_(sheet, 11, dataRows, '#,##0'); safeSetNumberFormat_(sheet, 12, dataRows, '"₩"#,##0');
  try { sheet.getRange(2, 22, dataRows, 1).setNumberFormat('@'); sheet.getRange(2, 25, dataRows, 1).setNumberFormat('@'); sheet.getRange(2, 29, dataRows, 1).setNumberFormat('@'); } catch (e) {}
  try { sheet.getRange(1, 1, lastRow, lastCol).setVerticalAlignment('middle'); sheet.getRange(2, 1, dataRows, lastCol).setWrap(false); sheet.getRange(1, 1, 1, lastCol).setWrap(true); } catch (e) {}
}

function safeSetNumberFormat_(sheet, col, rows, format) { try { if (rows > 0) sheet.getRange(2, col, rows, 1).setNumberFormat(format); } catch (e) {} }

function formatDashboard_(sheet) {
  formatDashboardDisplayOnly_(sheet, sheet ? sheet.getLastRow() : 0, sheet ? Math.max(sheet.getLastColumn(), 1) : 0);
}

function getStandardColumnFitOptions_() { return { minWidth: 46, maxWidth: 260, headerPadding: 18, dataPadding: 18, sampleRows: 350 }; }

function fitColumnsByHeaderAndData_(sheet, options) {
  options = options || {}; const lastRow = Math.max(sheet.getLastRow(), 1); const lastCol = Math.max(sheet.getLastColumn(), 1);
  const sampleRows = Math.min(Math.max(lastRow, 1), options.sampleRows || 300);
  const values = sheet.getRange(1, 1, sampleRows, lastCol).getDisplayValues();
  for (var c = 1; c <= lastCol; c++) {
    const headerText = String(values[0][c - 1] || '');
    const headerWidth = estimateHeaderWidth_(headerText) + (options.headerPadding || 16);
    let dataWidth = 0;
    for (var r = 1; r < values.length; r++) { const text = String(values[r][c - 1] || ''); if (!text) continue; dataWidth = Math.max(dataWidth, estimateTextWidth_(text)); }
    dataWidth += (options.dataPadding || 16);
    const minWidth = options.minWidth || 40, maxWidth = options.maxWidth || 260;
    const width = Math.min(maxWidth, Math.max(minWidth, headerWidth, dataWidth));
    try { sheet.setColumnWidth(c, Math.round(width)); } catch (e) {}
  }
}

function estimateHeaderWidth_(text) { const lines = String(text || '').split(/\n/); let max = 0; lines.forEach(function(line) { max = Math.max(max, estimateTextWidth_(line)); }); return max; }

function estimateTextWidth_(text) {
  const s = String(text || ''); let width = 0;
  for (var i = 0; i < s.length; i++) { const ch = s.charAt(i); if (/[가-힣ㄱ-ㅎㅏ-ㅣ]/.test(ch)) width += 12; else if (/[A-Z]/.test(ch)) width += 8; else if (/[0-9]/.test(ch)) width += 7; else if (ch === '_' || ch === '-' || ch === '/') width += 5; else if (ch === ' ') width += 4; else width += 7; }
  return width;
}

function isPercentHeader_(h) { return /률|비율|비중|percent|rate/i.test(String(h || '')); }
function isCountHeader_(h) { const s = String(h || '').replace(/\s+/g, ''); if (/매출액|매출액|결제금액|정산|배송비|구매가격|환산매출|일평균매출|상품당매출|브랜드매출/i.test(s)) return false; return /상품수|건수|주문수|수량|브랜드수|행수|페이지|필터수|토큰|값|기간|일수|순위|번호|전송수/i.test(s); }
function isMoneyHeader_(h) { const s = String(h || ''); if (isCountHeader_(s) || isPercentHeader_(s)) return false; return /매출액|매출액|결제금액|정산|배송비|구매가격|환산매출|일평균매출|상품당매출|브랜드매출/i.test(s); }
function isDateHeader_(h) { const s = String(h || '').replace(/\s+/g, ''); return /일시|일자|날짜|등록일|수집일|수정일|전송일|확인일|기준등록일|재전송일|추가수집일|정리재수집일|1단계_추가수집일|2단계_재전송일|3단계_정리재수집일/i.test(s); }

function hideNonEssentialSheets() {
  const ui = SpreadsheetApp.getUi(); const ss = SpreadsheetApp.getActive();
  const visibleNames = getEssentialVisibleSheetNames_(); const visibleMap = {};
  visibleNames.forEach(function(name) { visibleMap[name] = true; });
  let shown = 0, hidden = 0, failed = 0; const failedNames = [];
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (dashboard) { try { dashboard.showSheet(); ss.setActiveSheet(dashboard); } catch (e) {} }
  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    try { if (visibleMap[name]) { sheet.showSheet(); shown += 1; } else { sheet.hideSheet(); hidden += 1; } }
    catch (e) { failed += 1; failedNames.push(name + ': ' + String(e && e.message ? e.message : e)); }
  });
  try { log_('sheet_visibility_fast', 'v5.49 시트 정리 완료 / 표시=' + shown + ', 숨김=' + hidden + ', 실패=' + failed); } catch (e) {}
  ui.alert('시트 정리 완료\n\n표시 시트: ' + shown + '개\n숨김 시트: ' + hidden + '개\n실패: ' + failed + '개' + (failedNames.length ? '\n\n실패 시트:\n' + failedNames.slice(0, 5).join('\n') : '') + '\n\n숨긴 시트는 삭제된 것이 아닙니다.');
}

function showAllLotteonSheets() {
  const ui = SpreadsheetApp.getUi(); let shown = 0, failed = 0; const failedNames = [];
  SpreadsheetApp.getActive().getSheets().forEach(function(sheet) {
    try { sheet.showSheet(); shown += 1; } catch (e) { failed += 1; failedNames.push(sheet.getName() + ': ' + String(e && e.message ? e.message : e)); }
  });
  ui.alert('시트 전체 표시 완료\n\n표시 시트: ' + shown + '개\n실패: ' + failed + '개' + (failedNames.length ? '\n\n실패 시트:\n' + failedNames.slice(0, 5).join('\n') : ''));
}

function getEssentialVisibleSheetNames_() {
  return [CONFIG.SHEETS.DASHBOARD, CONFIG.SHEETS.BRAND_SUMMARY, CONFIG.SHEETS.MATCH_DIAG, CONFIG.SHEETS.SALES_IN, CONFIG.SHEETS.COUPANG_SENT_MANUAL, CONFIG.SHEETS.MANUAL_API_CHECK, CONFIG.SHEETS.RETRANSMIT_LOG];
}

function forceCleanupLotteonTriggers() {
  const ui = SpreadsheetApp.getUi();
  const targetHandlers = ['selfDiagnosticDummy_', 'continueLotteonProductSync', 'runAnalysisAfterSync', 'refreshLotteonAnalysis'];
  const result = { totalProjectTriggers: 0, totalTarget: 0, deleted: 0, failed: 0, details: [] };
  try {
    const triggers = ScriptApp.getProjectTriggers(); result.totalProjectTriggers = triggers.length;
    triggers.forEach(function(trigger) {
      let handler = ''; let triggerId = '';
      try { handler = typeof trigger.getHandlerFunction === 'function' ? trigger.getHandlerFunction() : ''; } catch (e) { handler = 'HANDLER_READ_FAILED'; }
      try { triggerId = typeof trigger.getUniqueId === 'function' ? trigger.getUniqueId() : ''; } catch (e) { triggerId = 'ID_READ_FAILED'; }
      if (targetHandlers.indexOf(handler) < 0) return;
      result.totalTarget += 1;
      const ok = safeDeleteTrigger_(trigger);
      if (ok) { result.deleted += 1; result.details.push([now_(), 'DELETE_SUCCESS', handler, triggerId, '']); }
      else { result.failed += 1; result.details.push([now_(), 'DELETE_FAILED', handler, triggerId, 'ScriptApp.deleteTrigger 실패. 수동 삭제 필요']); }
    });
    if (result.totalTarget === 0) result.details.push([now_(), 'NO_TARGET_TRIGGER', '', '', '삭제 대상 LOTTEON 트리거가 없습니다.']);
    writeTriggerCleanupLog_(result);
    log_('force_trigger_cleanup', '전체트리거=' + result.totalProjectTriggers + ', 대상=' + result.totalTarget + ', 삭제성공=' + result.deleted + ', 삭제실패=' + result.failed);
  } catch (e) { const msg = String(e && e.message ? e.message : e); result.failed += 1; result.details.push([now_(), 'LIST_TRIGGER_FAILED', '', '', msg]); writeTriggerCleanupLog_(result); log_('force_trigger_cleanup_error', msg); }
  const msg = '트리거 정리 결과\n\n전체: ' + result.totalProjectTriggers + '\n대상: ' + result.totalTarget + '\n삭제 성공: ' + result.deleted + '\n삭제 실패: ' + result.failed + '\n\n' + (result.failed > 0 ? '삭제 실패가 남아 있습니다.\nApps Script 왼쪽 "트리거" 메뉴에서 수동 삭제가 필요할 수 있습니다.' : '대상 트리거가 모두 정리되었습니다.');
  ui.alert(msg);
}

function writeTriggerCleanupLog_(result) {
  const ss = SpreadsheetApp.getActive(); let sheet = ss.getSheetByName('트리거정리로그');
  if (!sheet) sheet = ss.insertSheet('트리거정리로그');
  const header = ['일시', '결과', '핸들러', '트리거ID', '오류내용'];
  if (sheet.getLastRow() === 0) { sheet.getRange(1, 1, 1, header.length).setValues([header]); }
  else { const currentHeader = sheet.getRange(1, 1, 1, header.length).getValues()[0].map(function(v) { return String(v || '').trim(); }); if (!header.every(function(v, i) { return currentHeader[i] === v; })) sheet.getRange(1, 1, 1, header.length).setValues([header]); }
  if (result && result.details && result.details.length) sheet.getRange(sheet.getLastRow() + 1, 1, result.details.length, header.length).setValues(result.details);
  sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#e8f0fe');
  try { sheet.autoResizeColumns(1, header.length); } catch (e) {}
}

function log_(type, message) { const sheet = getSheet_(CONFIG.SHEETS.LOG); if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 3).setValues([CONFIG.HEADERS.LOG]); const first = sheet.getRange(1, 1).getValue(); if (sheet.getLastRow() === 1 && first !== '일시') replaceData_(sheet, CONFIG.HEADERS.LOG, []); sheet.appendRow([now_(), type, message]); }

function recordApiTokenUsage_(endpoint) {
  const props = PropertiesService.getScriptProperties(); const now = new Date();
  const today = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd'); const month = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM');
  const costMap = CONFIG.API_TOKEN_COST_MAP || {}; const cost = toNumber_(costMap[endpoint]) || 1;
  const savedDate = props.getProperty('API_TOKEN_DATE'); const savedMonth = props.getProperty('API_TOKEN_MONTH');
  let todayUsed = savedDate === today ? toNumber_(props.getProperty('API_TOKEN_TODAY_EST')) : 0;
  let monthUsed = savedMonth === month ? toNumber_(props.getProperty('API_TOKEN_MONTH_EST')) : 0;
  let syncUsed = toNumber_(props.getProperty('SYNC_TOKEN_USED_EST'));
  todayUsed += cost; monthUsed += cost; syncUsed += cost;
  props.setProperty('API_TOKEN_DATE', today); props.setProperty('API_TOKEN_MONTH', month);
  props.setProperty('API_TOKEN_TODAY_EST', String(todayUsed)); props.setProperty('API_TOKEN_MONTH_EST', String(monthUsed));
  props.setProperty('SYNC_TOKEN_USED_EST', String(syncUsed)); props.setProperty('API_TOKEN_LAST_ENDPOINT', String(endpoint || '')); props.setProperty('API_TOKEN_LAST_AT', now_());
}

function getApiTokenUsageSnapshot_() {
  const props = PropertiesService.getScriptProperties(); const now = new Date();
  const today = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd'); const month = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM');
  const scriptTodayUsed = props.getProperty('API_TOKEN_DATE') === today ? toNumber_(props.getProperty('API_TOKEN_TODAY_EST')) : 0;
  const scriptMonthUsed = props.getProperty('API_TOKEN_MONTH') === month ? toNumber_(props.getProperty('API_TOKEN_MONTH_EST')) : 0;
  const syncUsed = toNumber_(props.getProperty('SYNC_TOKEN_USED_EST'));
  const baselineToday = today === CONFIG.API_TOKEN_BASELINE_DAILY_DATE ? toNumber_(CONFIG.API_TOKEN_BASELINE_DAILY_USED) : 0;
  const baselineMonth = month === CONFIG.API_TOKEN_BASELINE_MONTH ? toNumber_(CONFIG.API_TOKEN_BASELINE_MONTH_USED) : 0;
  const todayUsed = baselineToday + scriptTodayUsed, monthUsed = baselineMonth + scriptMonthUsed;
  return { todayUsed: todayUsed, monthUsed: monthUsed, syncUsed: syncUsed, todayRemain: Math.max(0, CONFIG.API_TOKEN_DAILY_LIMIT - todayUsed), monthRemain: Math.max(0, CONFIG.API_TOKEN_MONTHLY_LIMIT - monthUsed), lastEndpoint: props.getProperty('API_TOKEN_LAST_ENDPOINT') || '', lastAt: props.getProperty('API_TOKEN_LAST_AT') || '', scriptTodayUsed: scriptTodayUsed, scriptMonthUsed: scriptMonthUsed, baselineToday: baselineToday, baselineMonth: baselineMonth };
}

function formatCount_(v) { const n = toNumber_(v); return Utilities.formatString('%s', n.toLocaleString('en-US')); }
function buildApiUrl_(endpoint) {
  return buildApiUrlCandidates_(endpoint)[0];
}

function buildApiUrlCandidates_(endpoint) {
  const cred = getCredentials_();
  const ep = String(endpoint).replace(/^\/+/, '');
  const bases = [];
  const urls = [];

  function addBase_(base) {
    const b = normalizeBaseUrl_(base);
    if (!b) return;
    if (bases.indexOf(b) < 0) bases.push(b);
  }

  // 1순위: 사용자가 저장한 URL
  addBase_(cred.baseUrl);

  // 2순위: 공식/기본 API URL
  // 저장된 URL이 tmg4688.mycafe24.com처럼 관리자/개별몰 도메인이고 filterList가 막히는 경우,
  // tmg2007.cafe24.com 기본 API URL로 자동 재시도합니다.
  addBase_(CONFIG.BASE_URL_DEFAULT);

  bases.forEach(function(base) {
    const direct = base + '/api/' + ep;
    if (urls.indexOf(direct) < 0) urls.push(direct);
  });

  if (CONFIG.USE_API_GATEWAY_FALLBACK) {
    const gatewayBase = String(CONFIG.API_GATEWAY_BASE_URL || '').replace(/\/+$/, '');
    if (gatewayBase) {
      urls.slice().forEach(function(direct) {
        const gateway = gatewayBase + '/' + direct;
        if (urls.indexOf(gateway) < 0) urls.push(gateway);
      });
    }
  }

  return urls;
}
function buildApiUrlCandidates_(endpoint) {
  const cred = getCredentials_(); const base = normalizeBaseUrl_(cred.baseUrl); const ep = String(endpoint).replace(/^\/+/, ''); const direct = base + '/api/' + ep; const urls = [direct];
  if (CONFIG.USE_API_GATEWAY_FALLBACK) { const gatewayBase = String(CONFIG.API_GATEWAY_BASE_URL || '').replace(/\/+$/, ''); if (gatewayBase) { const gateway = gatewayBase + '/' + direct; if (gateway !== direct) urls.push(gateway); } }
  return urls;
}
function maskSensitiveUrl_(url) { return String(url || '').replace(/([?&](apiKey|key|token|sender)=)[^&]+/ig, '$1***'); }
function countDataRows_(sheet) { if (!sheet) return 0; return Math.max(0, sheet.getLastRow() - 1); }
function accountFromFilterName_(filterName) { const code = filterCodeFromFilterName_(filterName); if (CONFIG.ACCOUNT_MAP[code]) return CONFIG.ACCOUNT_MAP[code]; return code ? { accountNo: toNumber_(code) || code, accountId: '' } : { accountNo: '', accountId: '' }; }
function filterCodeFromFilterName_(filterName) { const m = String(filterName || '').trim().match(/^롯백_(\d{2})/); return m ? m[1] : ''; }
function brandFromFilterName_(filterName) { const s = String(filterName || '').trim(); const m = s.match(/^롯백_\d{2}_?(.*)$/); if (m) return String(m[1] || '').trim(); return s; }
function isValidLotteonFilterName_(filterName) {
  const s = String(filterName || '').trim();
  if (!s || !s.startsWith(CONFIG.FILTER_PREFIX)) return false;

  const code = filterCodeFromFilterName_(s);
  // v5.53: 롯백_00 같은 과거/미지정 필터는 현재 운영필터에서 제외합니다.
  // 운영 필터는 CONFIG.ACCOUNT_MAP에 정의된 01~04만 유효 처리합니다.
  if (!code || !CONFIG.ACCOUNT_MAP[code]) return false;

  const brand = brandFromFilterName_(s);
  if (!brand) return false;
  if (/^롯백_\d{2}_?$/.test(s)) return false;
  return true;
}
function findFilterName_(row) { const keys = ['filterName', 'filtername', 'filter_name', 'name', 'title', 'searchFilterName', 'keyword']; for (var i = 0; i < keys.length; i++) { const v = String(row[keys[i]] || '').trim(); if (v) return v; } for (const k in row) { const lk = String(k).toLowerCase(); if (lk.indexOf('filter') >= 0 && lk.indexOf('name') >= 0) { const v = String(row[k] || '').trim(); if (v) return v; } } return ''; }
function getAny_(obj, keys) { for (var i = 0; i < keys.length; i++) { if (obj[keys[i]] !== undefined && obj[keys[i]] !== null && String(obj[keys[i]]).trim() !== '') return obj[keys[i]]; } return ''; }
function findDateValue_(obj, keys) { return normalizeDateText_(getAny_(obj, keys)); }
function pickByHeader_(headers, row, candidates) { for (var i = 0; i < candidates.length; i++) { const target = String(candidates[i]).toLowerCase().trim(); for (var c = 0; c < headers.length; c++) { const h = String(headers[c] || '').toLowerCase().trim(); if (h === target) return row[c]; } } for (var i = 0; i < candidates.length; i++) { const target = String(candidates[i]).toLowerCase().trim(); for (var c = 0; c < headers.length; c++) { const h = String(headers[c] || '').toLowerCase().trim(); if (!h) continue; if (h.indexOf(target) >= 0 || target.indexOf(h) >= 0) return row[c]; } } return ''; }
function normalizeBaseUrl_(url) { let s = String(url || CONFIG.BASE_URL_DEFAULT).trim().replace(/\/+$/, ''); const apiIdx = s.indexOf('/api/'); if (apiIdx > 0) s = s.slice(0, apiIdx); return s; }
function now_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function normalizeDateText_(v) { if (!v) return ''; if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); const s = String(v).trim(); if (!s) return ''; const m = s.match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})(?:[ T](\d{1,2}):?(\d{0,2}):?(\d{0,2}))?/); if (m) { const yy = m[1], mm = ('0' + m[2]).slice(-2), dd = ('0' + m[3]).slice(-2), hh = ('0' + (m[4] || '00')).slice(-2), mi = ('0' + (m[5] || '00')).slice(-2), ss = ('0' + (m[6] || '00')).slice(-2); return yy + '-' + mm + '-' + dd + ' ' + hh + ':' + mi + ':' + ss; } return s; }
function monthKey_(v) { const s = normalizeDateText_(v); const m = s.match(/(\d{4})-(\d{2})/); return m ? m[1] + '-' + m[2] : ''; }
function daysSince_(dateText) { const s = normalizeDateText_(dateText); const m = s.match(/(\d{4})-(\d{2})-(\d{2})/); if (!m) return 0; const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])), today = new Date(); return Math.max(0, Math.floor((today.getTime() - d.getTime()) / (1000 * 60 * 60 * 24))); }
function minDateText_(arr) { const dates = (arr || []).map(normalizeDateText_).filter(Boolean).sort(); return dates.length ? dates[0] : ''; }
function daysBetweenDates_(startDateText, endDateText) { const s = normalizeDateText_(startDateText), e = normalizeDateText_(endDateText); const sm = s.match(/(\d{4})-(\d{2})-(\d{2})/), em = e.match(/(\d{4})-(\d{2})-(\d{2})/); if (!sm || !em) return 0; const sd = new Date(Number(sm[1]), Number(sm[2]) - 1, Number(sm[3])), ed = new Date(Number(em[1]), Number(em[2]) - 1, Number(em[3])); return Math.max(0, Math.floor((ed.getTime() - sd.getTime()) / (1000 * 60 * 60 * 24))); }
function countProductsByMonthEnd_(products, month) { if (!month) return products.length; const end = month + '-31'; let count = 0; products.forEach(function(p) { const d = normalizeDateText_(p['기준등록일'] || p['최초확인일']); if (!d || d.slice(0, 10) <= end) count++; }); return count; }
function toNumber_(v) { if (v === null || v === undefined || v === '') return 0; if (typeof v === 'number') return v; const s = String(v).replace(/[₩$,]/g, '').replace(/\s/g, '').trim(); const n = Number(s); return isNaN(n) ? 0 : n; }
function unique_(arr) { const seen = {}, out = []; arr.forEach(function(v) { const key = String(v); if (!seen[key]) { seen[key] = true; out.push(v); } }); return out; }

function deleteTriggersForHandler_(handlerName) {
  try {
    const triggers = ScriptApp.getProjectTriggers(); let deletedCount = 0, failedCount = 0;
    triggers.forEach(function(t) {
      try { if (t.getHandlerFunction() === handlerName) { if (safeDeleteTrigger_(t)) deletedCount += 1; else failedCount += 1; } }
      catch (e) { failedCount += 1; log_('trigger_handler_error', 'handler=' + handlerName + ', error=' + String(e && e.message ? e.message : e)); }
    });
    if (deletedCount > 0) log_('trigger_delete_success', 'handler=' + handlerName + ', deleted=' + deletedCount + ', failed=' + failedCount);
    if (failedCount > 0) log_('trigger_delete_partial_fail', 'handler=' + handlerName + ', 수동 삭제 필요 - Apps Script 왼쪽 트리거 메뉴 확인');
  } catch (e) { const msg = String(e && e.message ? e.message : e); log_('trigger_delete_error', 'handler=' + handlerName + ', error=' + msg); }
}

function normalizeBrandKey_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ''); }
function normalizeProductNameKey_(s) { return String(s || '').trim().toLowerCase().replace(/[\[\]\(\)\{\}\-_\/\\,.|·•~!@#$%^&*+=:;'"<>?]/g, '').replace(/\s+/g, ''); }
function extractLotteonProductCodeFromFields_(image, name, model) { const candidates = [image, name, model]; for (var i = 0; i < candidates.length; i++) { const v = candidates[i]; if (!v) continue; const m = String(v).match(CONFIG.LOTTEON_PRODUCT_CODE_REGEX); if (m) return m[1].toUpperCase(); } return ''; }
function normalizeAccountId_(raw) { const s = String(raw || '').trim(); if (!s) return { original: '', normalized: '', accountNo: '' }; const lower = s.toLowerCase(); const m = lower.match(/(beliun\d+)/); const normalized = m ? m[1] : lower; let accountNo = ''; Object.keys(CONFIG.ACCOUNT_ID_PREFIX_MAP).forEach(function(prefix) { if (normalized.indexOf(prefix) >= 0) accountNo = CONFIG.ACCOUNT_ID_PREFIX_MAP[prefix]; }); return { original: s, normalized: normalized, accountNo: accountNo }; }
function buildBrandSummaryOnly() { const brandRows = getValuesWithoutHeader_(CONFIG.SHEETS.BRAND); const summaryRows = buildBrandSummaryRowsFromBrandRows_(brandRows); replaceDataFastLimited_(getSheet_(CONFIG.SHEETS.BRAND_SUMMARY), CONFIG.HEADERS.BRAND_SUMMARY, summaryRows); formatCoreAnalysisSheetsLightOnly_(); SpreadsheetApp.getUi().alert('핵심_브랜드요약 갱신 완료: ' + summaryRows.length + '개 브랜드'); }
function formatWon_(n) { return '₩' + formatCount_(Math.round(toNumber_(n))); }
function writeSheetVisibilityLog_(rows) { try { log_('sheet_visibility_log_skipped', 'v5.49에서는 시간초과 방지를 위해 상세 시트정리로그 기록을 생략합니다. rows=' + ((rows && rows.length) || 0)); } catch (e) {} }
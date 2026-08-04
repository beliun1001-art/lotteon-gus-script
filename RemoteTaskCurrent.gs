/**
 * Permanent remote task slot + unattended autopilot.
 * Current task: PR15 VAT future-account raw scan diagnostic.
 */
const LOTTEON_REMOTE_TASK = {
  id: 'PR15-v1.2-20260804',
  title: 'PR15 부가세 이상 계정·미래월 원본 raw scan 진단',
  sourceUrls: [
    'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-15-vat-future-account-diagnostic/PR15_Future_Account_Diagnostic.gs',
    'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-15-vat-future-account-diagnostic/PR15_Future_Account_Diagnostic_Hotfix_v1_2.gs'
  ],
  startEntry: 'runPr15FutureAccountDiagnosticV12',
  continueEntry: 'runPr15FutureAccountDiagnosticContinueV12',
  statusSheet: 'PR15_진단상태',
  terminalStatuses: ['PASS', 'ERROR']
};

const LOTTEON_REMOTE_NOTICE_EMAIL = 'beliun1001@gmail.com';
const LOTTEON_REMOTE_AUTOPILOT_HANDLER = 'runLotteonRemoteTaskStart';
const LOTTEON_REMOTE_ACTIVE_KEY = 'LOTTEON_REMOTE_ACTIVE_TASK_ID';
const LOTTEON_REMOTE_LAST_DONE_KEY = 'LOTTEON_REMOTE_LAST_DONE_TASK_ID';
const LOTTEON_REMOTE_LAST_NOTICE_KEY = 'LOTTEON_REMOTE_LAST_NOTICE';
const LOTTEON_REMOTE_SPREADSHEET_KEY = 'LOTTEON_REMOTE_SPREADSHEET_ID';
const LOTTEON_REMOTE_NOTICE_ERROR_KEY = 'LOTTEON_REMOTE_NOTICE_ERROR';
const LOTTEON_REMOTE_MAIN_BASE = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/';

function runLotteonRemoteTaskStartRemote_() {
  lotteonRememberSpreadsheet_();
  lotteonInstallRemoteAutopilot_();

  const props = PropertiesService.getScriptProperties();
  const info = lotteonReadRemoteTaskStatus_();
  const lastDone = props.getProperty(LOTTEON_REMOTE_LAST_DONE_KEY) || '';
  const active = props.getProperty(LOTTEON_REMOTE_ACTIVE_KEY) || '';

  if (lastDone === LOTTEON_REMOTE_TASK.id && lotteonIsTerminalStatus_(info.status)) {
    lotteonSendRemoteTaskNotice_(info);
    return {ok:true, skipped:true, reason:'ALREADY_COMPLETED', status:info.status};
  }
  if (active === LOTTEON_REMOTE_TASK.id && info.status && !lotteonIsTerminalStatus_(info.status)) {
    return {ok:true, skipped:true, reason:'ALREADY_RUNNING', status:info.status};
  }

  props.setProperty(LOTTEON_REMOTE_ACTIVE_KEY, LOTTEON_REMOTE_TASK.id);
  try {
    return lotteonFinalizeRemoteInvocation_(lotteonRunRemoteTaskEntry_(LOTTEON_REMOTE_TASK.startEntry));
  } catch (error) {
    lotteonHandleRemoteWrapperError_(error);
    throw error;
  }
}

function runLotteonRemoteTaskContinueRemote_() {
  PropertiesService.getScriptProperties().setProperty(LOTTEON_REMOTE_ACTIVE_KEY, LOTTEON_REMOTE_TASK.id);
  try {
    return lotteonFinalizeRemoteInvocation_(lotteonRunRemoteTaskEntry_(LOTTEON_REMOTE_TASK.continueEntry));
  } catch (error) {
    lotteonHandleRemoteWrapperError_(error);
    throw error;
  }
}

function lotteonRunRemoteTaskEntry_(entryName) {
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(entryName)) {
    throw new Error('잘못된 원격 작업 진입점: ' + entryName);
  }

  const mainBundle = typeof loadLotteonRemoteBundle_ === 'function'
    ? loadLotteonRemoteBundle_()
    : lotteonFetchMainBundle_();
  const taskUrls = LOTTEON_REMOTE_TASK.sourceUrls || [LOTTEON_REMOTE_TASK.sourceUrl];
  const taskCode = taskUrls.map(function(url, index) {
    return lotteonFetchText_(url, '원격 작업 코드 ' + (index + 1));
  }).join('\n\n;\n\n');
  const combined = mainBundle + '\n\n;\n\n' + taskCode;

  return eval(
    combined +
    '\n\n; if (typeof ' + entryName + ' !== "function") ' +
    'throw new Error("원격 작업 진입점 없음: ' + entryName + '"); ' + entryName + '();'
  );
}

function lotteonFetchMainBundle_() {
  const code = lotteonFetchText_(LOTTEON_REMOTE_MAIN_BASE + 'Code.gs', 'Code.gs');
  const patch = lotteonFetchText_(LOTTEON_REMOTE_MAIN_BASE + 'Patch_v6_24_bootstrap_auto_continue.gs', 'Patch bootstrap');
  return code + '\n\n;\n\n' + patch;
}

function lotteonFetchText_(url, label) {
  const response = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), {
    method:'get', muteHttpExceptions:true, followRedirects:true
  });
  const status = response.getResponseCode();
  const text = response.getContentText('UTF-8');
  if (status < 200 || status >= 300) {
    throw new Error(label + ' 로드 실패 HTTP ' + status + '\n' + url + '\n' + text.slice(0,500));
  }
  return text;
}

function lotteonFinalizeRemoteInvocation_(result) {
  const info = lotteonReadRemoteTaskStatus_();
  if (lotteonIsTerminalStatus_(info.status)) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(LOTTEON_REMOTE_LAST_DONE_KEY, LOTTEON_REMOTE_TASK.id);
    props.deleteProperty(LOTTEON_REMOTE_ACTIVE_KEY);
    lotteonSendRemoteTaskNotice_(info);
  }
  return result;
}

function lotteonRememberSpreadsheet_() {
  try {
    const ss = SpreadsheetApp.getActive();
    if (ss) PropertiesService.getScriptProperties().setProperty(LOTTEON_REMOTE_SPREADSHEET_KEY, ss.getId());
  } catch (ignore) {}
}

function lotteonReadRemoteTaskStatus_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (ignore) {}
  if (!ss) {
    const id = props.getProperty(LOTTEON_REMOTE_SPREADSHEET_KEY);
    if (id) try { ss = SpreadsheetApp.openById(id); } catch (ignore2) {}
  }
  if (!ss) return {status:'', rows:[], map:{}, spreadsheetName:''};

  props.setProperty(LOTTEON_REMOTE_SPREADSHEET_KEY, ss.getId());
  const sheet = ss.getSheetByName(LOTTEON_REMOTE_TASK.statusSheet);
  if (!sheet || sheet.getLastRow() < 2) {
    return {status:'', rows:[], map:{}, spreadsheetName:ss.getName()};
  }

  const values = sheet.getDataRange().getDisplayValues();
  const map = {};
  for (let r=1; r<values.length; r++) {
    map[String(values[r][0] || '').trim()] = String(values[r][1] || '').trim();
  }
  return {status:map['상태'] || '', rows:values, map:map, spreadsheetName:ss.getName()};
}

function lotteonIsTerminalStatus_(status) {
  return LOTTEON_REMOTE_TASK.terminalStatuses.indexOf(String(status || '').trim()) >= 0;
}

function lotteonInstallRemoteAutopilot_() {
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === LOTTEON_REMOTE_AUTOPILOT_HANDLER &&
      trigger.getEventType() === ScriptApp.EventType.CLOCK;
  });
  if (!exists) {
    ScriptApp.newTrigger(LOTTEON_REMOTE_AUTOPILOT_HANDLER).timeBased().everyHours(1).create();
  }
}

function lotteonSendRemoteTaskNotice_(info) {
  const key = LOTTEON_REMOTE_TASK.id + '|' + info.status;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY) === key) return true;

  const lines = [];
  (info.rows || []).forEach(function(row,index) {
    if (index === 0) return;
    const name = String(row[0] || '').trim();
    const value = String(row[1] || '').trim();
    if (name) lines.push(name + ': ' + value);
  });

  const subject = '[LOTTEON 자동작업 결과][' + info.status + '] ' + LOTTEON_REMOTE_TASK.title;
  const body = [
    'LOTTEON 원격 자동작업이 종료되었습니다.','',
    '작업: ' + LOTTEON_REMOTE_TASK.title,
    '작업 ID: ' + LOTTEON_REMOTE_TASK.id,
    '스프레드시트: ' + (info.spreadsheetName || ''),'',
    lines.join('\n')
  ].join('\n');

  try {
    MailApp.sendEmail(LOTTEON_REMOTE_NOTICE_EMAIL, subject, body);
    props.setProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY, key);
    props.deleteProperty(LOTTEON_REMOTE_NOTICE_ERROR_KEY);
    lotteonWriteNoticeStatus_('SENT','');
    return true;
  } catch (error) {
    const message = String(error && error.message ? error.message : error);
    props.setProperty(LOTTEON_REMOTE_NOTICE_ERROR_KEY, message);
    lotteonWriteNoticeStatus_('FAILED',message);
    return false;
  }
}

function lotteonHandleRemoteWrapperError_(error) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(LOTTEON_REMOTE_ACTIVE_KEY);
  const message = String(error && error.message ? error.message : error);
  const key = LOTTEON_REMOTE_TASK.id + '|WRAPPER_ERROR|' + message;
  if (props.getProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY) === key) return;

  try {
    MailApp.sendEmail(
      LOTTEON_REMOTE_NOTICE_EMAIL,
      '[LOTTEON 자동작업 결과][ERROR] ' + LOTTEON_REMOTE_TASK.title,
      '원격 자동작업 진입 또는 코드 로드 단계에서 오류가 발생했습니다.\n\n' +
      '작업: ' + LOTTEON_REMOTE_TASK.title + '\n' +
      '작업 ID: ' + LOTTEON_REMOTE_TASK.id + '\n' +
      '오류: ' + message
    );
    props.setProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY,key);
    props.deleteProperty(LOTTEON_REMOTE_NOTICE_ERROR_KEY);
    lotteonWriteNoticeStatus_('SENT','');
  } catch (mailError) {
    const mailMessage = String(mailError && mailError.message ? mailError.message : mailError);
    props.setProperty(LOTTEON_REMOTE_NOTICE_ERROR_KEY,mailMessage);
    lotteonWriteNoticeStatus_('FAILED',mailMessage);
  }
}

function lotteonWriteNoticeStatus_(status,error) {
  try {
    const id = PropertiesService.getScriptProperties().getProperty(LOTTEON_REMOTE_SPREADSHEET_KEY);
    if (!id) return;
    const ss = SpreadsheetApp.openById(id);
    const sheet = ss.getSheetByName(LOTTEON_REMOTE_TASK.statusSheet);
    if (!sheet) return;
    const values = sheet.getRange(1,1,Math.max(1,sheet.getLastRow()),2).getDisplayValues();
    const index = {};
    for (let r=0; r<values.length; r++) index[String(values[r][0] || '').trim()] = r+1;
    const write = function(name,value) {
      const row = index[name] || sheet.getLastRow()+1;
      sheet.getRange(row,1,1,2).setValues([[name,value]]);
      index[name] = row;
    };
    write('완료알림',status);
    write('완료알림오류',error || '');
    write('완료알림시각',new Date().toISOString());
  } catch (ignore) {}
}

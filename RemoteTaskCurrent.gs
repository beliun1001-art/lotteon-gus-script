/**
 * Permanent remote task slot + unattended autopilot.
 * Current task: PR38 v6.70 production apply.
 */
const LOTTEON_REMOTE_TASK = {
  id: 'PR38-v1.0-20260803',
  title: 'PR38 v6.70 운영 부가세 반영',
  sourceUrl: 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-38-pr33-production-apply/PR38_Remote_Production_Apply.gs',
  startEntry: 'runPr38ProductionApplyStart',
  continueEntry: 'runPr38ProductionApplyContinue',
  statusSheet: 'PR38_운영반영상태',
  terminalStatuses: ['PASS','ROLLED_BACK','ERROR','ROLLBACK_ERROR']
};

const LOTTEON_REMOTE_AUTOPILOT_HANDLER = 'runLotteonRemoteTaskStart';
const LOTTEON_REMOTE_ACTIVE_KEY = 'LOTTEON_REMOTE_ACTIVE_TASK_ID';
const LOTTEON_REMOTE_LAST_DONE_KEY = 'LOTTEON_REMOTE_LAST_DONE_TASK_ID';
const LOTTEON_REMOTE_LAST_NOTICE_KEY = 'LOTTEON_REMOTE_LAST_NOTICE';

function runLotteonRemoteTaskStartRemote_() {
  lotteonInstallRemoteAutopilot_();
  const props = PropertiesService.getScriptProperties();
  const currentStatus = lotteonReadRemoteTaskStatus_();
  const lastDone = props.getProperty(LOTTEON_REMOTE_LAST_DONE_KEY) || '';
  const active = props.getProperty(LOTTEON_REMOTE_ACTIVE_KEY) || '';

  if (lastDone === LOTTEON_REMOTE_TASK.id && lotteonIsTerminalStatus_(currentStatus.status)) {
    return {ok:true, skipped:true, reason:'ALREADY_COMPLETED', taskId:LOTTEON_REMOTE_TASK.id, status:currentStatus.status};
  }
  if (active === LOTTEON_REMOTE_TASK.id && currentStatus.status && !lotteonIsTerminalStatus_(currentStatus.status)) {
    return {ok:true, skipped:true, reason:'ALREADY_RUNNING', taskId:LOTTEON_REMOTE_TASK.id, status:currentStatus.status};
  }

  props.setProperty(LOTTEON_REMOTE_ACTIVE_KEY, LOTTEON_REMOTE_TASK.id);
  try {
    const result = lotteonRunRemoteTaskEntry_(LOTTEON_REMOTE_TASK.startEntry);
    return lotteonFinalizeRemoteInvocation_(result);
  } catch (error) {
    lotteonHandleRemoteWrapperError_(error);
    throw error;
  }
}

function runLotteonRemoteTaskContinueRemote_() {
  PropertiesService.getScriptProperties().setProperty(LOTTEON_REMOTE_ACTIVE_KEY, LOTTEON_REMOTE_TASK.id);
  try {
    const result = lotteonRunRemoteTaskEntry_(LOTTEON_REMOTE_TASK.continueEntry);
    return lotteonFinalizeRemoteInvocation_(result);
  } catch (error) {
    lotteonHandleRemoteWrapperError_(error);
    throw error;
  }
}

function lotteonRunRemoteTaskEntry_(entryName) {
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(entryName)) throw new Error('잘못된 원격 작업 진입점: ' + entryName);
  const response = UrlFetchApp.fetch(LOTTEON_REMOTE_TASK.sourceUrl + '?ts=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const status = response.getResponseCode();
  const code = response.getContentText('UTF-8');
  if (status < 200 || status >= 300) {
    throw new Error('원격 작업 코드 로드 실패 HTTP ' + status + '\n' + code.slice(0, 500));
  }
  return eval(code + '\n\n; if (typeof ' + entryName + ' !== "function") throw new Error("원격 작업 진입점 없음: ' + entryName + '"); ' + entryName + '();');
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

function lotteonReadRemoteTaskStatus_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (ignore) {}
  if (!ss) {
    try {
      const stateKeys = ['PR38_REMOTE_APPLY_STATE','PR33_PREVIEW_SPREADSHEET_ID','PR30_V12_STATE'];
      const props = PropertiesService.getScriptProperties();
      for (let i=0;i<stateKeys.length&&!ss;i++) {
        const raw = props.getProperty(stateKeys[i]);
        if (!raw) continue;
        try {
          const parsed = JSON.parse(raw);
          if (parsed && parsed.spreadsheetId) ss = SpreadsheetApp.openById(parsed.spreadsheetId);
        } catch (e) {
          try { ss = SpreadsheetApp.openById(raw); } catch (ignore2) {}
        }
      }
    } catch (ignore3) {}
  }
  if (!ss) return {status:'', rows:[], spreadsheetName:''};
  const sheet = ss.getSheetByName(LOTTEON_REMOTE_TASK.statusSheet);
  if (!sheet || sheet.getLastRow() < 2) return {status:'', rows:[], spreadsheetName:ss.getName()};
  const values = sheet.getDataRange().getDisplayValues();
  const map = {};
  for (let r=1;r<values.length;r++) map[String(values[r][0]||'').trim()] = String(values[r][1]||'').trim();
  return {status:map['상태']||'', rows:values, map:map, spreadsheetName:ss.getName()};
}

function lotteonIsTerminalStatus_(status) {
  return LOTTEON_REMOTE_TASK.terminalStatuses.indexOf(String(status||'').trim()) >= 0;
}

function lotteonInstallRemoteAutopilot_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() === LOTTEON_REMOTE_AUTOPILOT_HANDLER && trigger.getEventType() === ScriptApp.EventType.CLOCK;
  });
  if (!exists) ScriptApp.newTrigger(LOTTEON_REMOTE_AUTOPILOT_HANDLER).timeBased().everyHours(1).create();
}

function lotteonSendRemoteTaskNotice_(info) {
  const key = LOTTEON_REMOTE_TASK.id + '|' + info.status;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY) === key) return;

  const recipient = Session.getEffectiveUser().getEmail();
  if (!recipient) return;
  const lines = [];
  (info.rows || []).forEach(function(row, index) {
    if (index === 0) return;
    const k = String(row[0]||'').trim();
    const v = String(row[1]||'').trim();
    if (k) lines.push(k + ': ' + v);
  });
  const subject = '[LOTTEON 자동작업 결과][' + info.status + '] ' + LOTTEON_REMOTE_TASK.title;
  const body = [
    'LOTTEON 원격 자동작업이 종료되었습니다.',
    '',
    '작업: ' + LOTTEON_REMOTE_TASK.title,
    '작업 ID: ' + LOTTEON_REMOTE_TASK.id,
    '스프레드시트: ' + (info.spreadsheetName || ''),
    '',
    lines.join('\n')
  ].join('\n');
  MailApp.sendEmail(recipient, subject, body);
  props.setProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY, key);
}

function lotteonHandleRemoteWrapperError_(error) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(LOTTEON_REMOTE_ACTIVE_KEY);
  const recipient = Session.getEffectiveUser().getEmail();
  if (!recipient) return;
  const message = String(error && error.message ? error.message : error);
  const key = LOTTEON_REMOTE_TASK.id + '|WRAPPER_ERROR|' + message;
  if (props.getProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY) === key) return;
  MailApp.sendEmail(
    recipient,
    '[LOTTEON 자동작업 결과][ERROR] ' + LOTTEON_REMOTE_TASK.title,
    '원격 자동작업 진입 또는 코드 로드 단계에서 오류가 발생했습니다.\n\n작업: ' + LOTTEON_REMOTE_TASK.title + '\n작업 ID: ' + LOTTEON_REMOTE_TASK.id + '\n오류: ' + message
  );
  props.setProperty(LOTTEON_REMOTE_LAST_NOTICE_KEY, key);
}

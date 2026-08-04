/** Issue #15 diagnostic notice enrichment v1.1. */
var PR15_NOTICE_VERSION = 'v1.1-ROW-EVIDENCE-NOTICE';

function pr15CountStatuses_(rows) {
  var out = {
    rows:rows.length,
    sourceValid:0,
    yearRestoreBug:0,
    accountMapping:0,
    sourceDataError:0,
    unresolved:0,
    evidence:[]
  };
  rows.forEach(function(row, index) {
    var status = row[18];
    if (status === 'SOURCE_CONFIRMED_VALID') out.sourceValid++;
    else if (status === 'YEAR_RESTORE_BUG_CONFIRMED') out.yearRestoreBug++;
    else if (status === 'ACCOUNT_MAPPING_REQUIRED') out.accountMapping++;
    else if (status === 'SOURCE_DATA_ERROR') out.sourceDataError++;
    else out.unresolved++;

    if (index < 20) {
      out.evidence.push([
        '원본행=' + String(row[0] || ''),
        '신고행=' + String(row[1] || ''),
        '주문=' + String(row[2] || ''),
        '원본일=' + String(row[3] || ''),
        '신고월=' + String(row[7] || ''),
        '계정=' + String(row[8] || ''),
        '사업자=' + String(row[9] || ''),
        '매입=' + String(row[12] || 0),
        'queue=' + String(row[16] || 0),
        '원인=' + String(row[18] || ''),
        '조치=' + String(row[19] || '')
      ].join(' | '));
    }
  });
  return out;
}

function pr15WriteStatus_(ss, status, message, stats, error) {
  var sheet = ss.getSheetByName(PR15_STATUS_SHEET) || ss.insertSheet(PR15_STATUS_SHEET);
  stats = stats || {};
  var rows = [
    ['항목','값'],
    ['버전',PR15_DIAG_VERSION + ' / ' + PR15_NOTICE_VERSION],
    ['상태',status],
    ['메시지',message || ''],
    ['진단행',stats.rows || 0],
    ['SOURCE_CONFIRMED_VALID',stats.sourceValid || 0],
    ['YEAR_RESTORE_BUG_CONFIRMED',stats.yearRestoreBug || 0],
    ['ACCOUNT_MAPPING_REQUIRED',stats.accountMapping || 0],
    ['SOURCE_DATA_ERROR',stats.sourceDataError || 0],
    ['UNRESOLVED',stats.unresolved || 0],
    ['운영 원본 변경','0건'],
    ['오류',error || ''],
    ['갱신시각',new Date().toISOString()]
  ];
  (stats.evidence || []).forEach(function(value, index) {
    rows.push(['증거' + (index + 1), value]);
  });
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setColumnWidth(1,260);
  sheet.setColumnWidth(2,1200);
}

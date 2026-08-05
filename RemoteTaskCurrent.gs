/**
 * Issue #42 operating diagnostic.
 * Extracts H1 NO_MATCH/AMBIGUOUS orders and classifies exact-amount evidence windows.
 * Diagnostic-only: no production VAT sheets are modified.
 */
const LOTTEON_REMOTE_TASK = {
  id: 'ISSUE42-v1.0-20260805',
  title: '상반기 VAT 잔여 51건 원인 분류 진단',
  enabled: true,
  statusSheet: 'ISSUE42_진단상태',
  outputSheet: 'ISSUE42_잔여매칭진단'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');

  var statusSheet = issue42EnsureSheet_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue42WriteStatus_(statusSheet, [
    ['항목', '값'],
    ['버전', 'v1.0-ISSUE42-H1-UNMATCHED-DIAGNOSTIC'],
    ['상태', 'RUNNING'],
    ['단계', 'LOAD'],
    ['메시지', '잔여 카드매칭 진단 시작'],
    ['운영시트 변경', '0'],
    ['갱신시각', new Date().toISOString()]
  ]);

  try {
    if (typeof loadVatCardHistory_v660_ !== 'function') throw new Error('필수 함수 미로딩: loadVatCardHistory_v660_');
    if (typeof loadVatCardMaster_v660_ !== 'function') throw new Error('필수 함수 미로딩: loadVatCardMaster_v660_');
    if (typeof canonicalizeVatHistory_v664_ !== 'function') throw new Error('필수 함수 미로딩: canonicalizeVatHistory_v664_');

    var diagSheet = ss.getSheetByName('부가세_카드매칭검증');
    if (!diagSheet || diagSheet.getLastRow() < 2) {
      throw new Error('부가세_카드매칭검증 시트를 찾지 못했거나 데이터가 없습니다.');
    }

    var values = diagSheet.getDataRange().getValues();
    var headers = values[0].map(issue42Text_);
    var ix = issue42BuildHeaderIndex_(headers);
    var required = ['신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','주문매입금액','카드매칭상태','카드매칭근거'];
    var missingHeaders = required.filter(function(name) { return ix[name] == null; });
    if (missingHeaders.length) throw new Error('필수 헤더 누락: ' + missingHeaders.join(', '));

    var unresolved = [];
    var usedApprovals = {};
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var year = issue42Text_(row[ix['신고연도']]);
      var half = issue42Text_(row[ix['반기']]);
      var status = issue42Text_(row[ix['카드매칭상태']]).toUpperCase();
      var approvalNo = ix['승인번호'] == null ? '' : issue42Text_(row[ix['승인번호']]);
      var company = ix['구매카드사'] == null ? '' : issue42Text_(row[ix['구매카드사']]);
      if ((status === 'MATCHED' || status === 'MASTER_MATCHED' || status === 'NON_CARD') && approvalNo) {
        usedApprovals[issue42ApprovalKey_(company, approvalNo)] = true;
      }
      if (year === '2026' && half === '상반기' && (status === 'NO_MATCH' || status === 'AMBIGUOUS')) {
        unresolved.push({
          sourceRow: r + 1,
          year: year,
          half: half,
          orderDate: issue42DateText_(row[ix['주문일']]),
          business: issue42Text_(row[ix['사업자등록번호']]),
          account: issue42Text_(row[ix['쿠팡계정ID']]),
          orderNo: issue42Text_(row[ix['주문번호']]),
          payment: issue42Text_(row[ix['롯데결제수단']]),
          purchase: issue42Number_(row[ix['주문매입금액']]),
          currentStatus: status,
          currentReason: issue42Text_(row[ix['카드매칭근거']]),
          currentCandidateCount: ix['후보수'] == null ? 0 : issue42Number_(row[ix['후보수']])
        });
      }
    }

    if (unresolved.length !== 51) {
      throw new Error('대상 건수 불일치: 기대 51건, 실제 ' + unresolved.length + '건');
    }

    var rawHistory = loadVatCardHistory_v660_(ss) || [];
    var master = loadVatCardMaster_v660_(ss) || [];
    var canonical = canonicalizeVatHistory_v664_(rawHistory, master) || [];
    var usable = canonical.filter(function(h) {
      return h && !h.nonCard && !h.cancelRow && !h.v664FullyCanceled && !!h.lotteEvidence &&
        issue42Number_(h.v664EffectiveAmount || h.amount) > 0 && issue42DateText_(h.date);
    });

    var outputHeaders = [
      '원본검증행','신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호',
      '롯데결제수단','주문매입금액','현재상태','현재근거','현재후보수',
      '원인분류','추가검토등급',
      '0~+7일_exact후보','0~+7일_미사용후보','0~+7일_cardIdentity수',
      '+8~+14일_exact후보','+8~+14일_미사용후보','+8~+14일_cardIdentity수',
      '-7~-1일_exact후보','-7~-1일_미사용후보','-7~-1일_cardIdentity수',
      '±30일_exact후보','±30일_미사용후보','기배정후보수',
      '최적후보승인일','최적후보카드사','최적후보카드명','최적후보끝4',
      '최적후보승인번호','최적후보승인금액','최적후보일차','최적후보원본파일',
      '후보요약','진단메모'
    ];

    var categoryCounts = {};
    var reviewCounts = {};
    var outputRows = unresolved.map(function(order) {
      var exact = usable.filter(function(h) {
        return issue42Number_(h.v664EffectiveAmount || h.amount) === issue42Number_(order.purchase);
      }).map(function(h) {
        var lag = issue42DaysBetween_(order.orderDate, issue42DateText_(h.date));
        var approvalKey = issue42ApprovalKey_(h.company, h.approvalNo);
        var identity = issue42Identity_(h, master);
        return {
          h: h,
          lag: lag,
          used: !!(h.approvalNo && usedApprovals[approvalKey]),
          identity: identity
        };
      }).filter(function(c) { return c.lag >= -30 && c.lag <= 30; });

      var compatible = issue42FilterPaymentCompatible_(exact, order.payment);
      var pool = compatible.length ? compatible : exact;
      var w07 = issue42WindowStats_(pool, 0, 7);
      var w814 = issue42WindowStats_(pool, 8, 14);
      var wNeg = issue42WindowStats_(pool, -7, -1);
      var w30 = issue42WindowStats_(pool, -30, 30);
      var classification = issue42Classify_(order, w07, w814, wNeg, w30);
      var best = issue42PickBest_(w07, w814, wNeg, w30);
      var category = classification.category;
      var review = classification.review;

      categoryCounts[category] = (categoryCounts[category] || 0) + 1;
      reviewCounts[review] = (reviewCounts[review] || 0) + 1;

      return [
        order.sourceRow, order.year, order.half, order.orderDate, order.business, order.account, order.orderNo,
        order.payment, order.purchase, order.currentStatus, order.currentReason, order.currentCandidateCount,
        category, review,
        w07.total, w07.unused.length, w07.identityCount,
        w814.total, w814.unused.length, w814.identityCount,
        wNeg.total, wNeg.unused.length, wNeg.identityCount,
        w30.total, w30.unused.length, w30.used.length,
        best ? issue42DateText_(best.h.date) : '',
        best ? issue42Text_(best.h.company) : '',
        best ? issue42Text_(best.h.cardName) : '',
        best ? issue42Text_(best.h.cardEnd4) : '',
        best ? issue42Text_(best.h.approvalNo) : '',
        best ? issue42Number_(best.h.v664EffectiveAmount || best.h.amount) : 0,
        best ? best.lag : '',
        best ? issue42Text_(best.h.sourceFile) : '',
        issue42CandidateSummary_(pool),
        classification.memo
      ];
    });

    var outputSheet = issue42EnsureSheet_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    outputSheet.clear();
    outputSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
    outputSheet.getRange(2, 1, outputRows.length, outputHeaders.length).setValues(outputRows);
    outputSheet.setFrozenRows(1);
    outputSheet.getRange(1, 1, 1, outputHeaders.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    outputSheet.getRange(2, 9, outputRows.length, 1).setNumberFormat('#,##0');
    outputSheet.getRange(2, 32, outputRows.length, 1).setNumberFormat('#,##0');
    if (outputSheet.getFilter()) outputSheet.getFilter().remove();
    outputSheet.getRange(1, 1, outputRows.length + 1, outputHeaders.length).createFilter();
    [90,90,80,100,130,120,170,150,110,100,360,80,260,120].forEach(function(width, i) {
      outputSheet.setColumnWidth(i + 1, width);
    });
    for (var c = 15; c <= outputHeaders.length; c++) {
      outputSheet.setColumnWidth(c, c === 35 || c === 36 ? 600 : 110);
    }

    var categoryKeys = Object.keys(categoryCounts).sort();
    var reviewKeys = Object.keys(reviewCounts).sort();
    var statusRows = [
      ['항목', '값'],
      ['버전', 'v1.0-ISSUE42-H1-UNMATCHED-DIAGNOSTIC'],
      ['상태', 'PASS'],
      ['단계', 'DONE'],
      ['메시지', '상반기 잔여 카드매칭 원인 분류 진단 완료'],
      ['대상 주문', unresolved.length],
      ['NO_MATCH', unresolved.filter(function(o){ return o.currentStatus === 'NO_MATCH'; }).length],
      ['AMBIGUOUS', unresolved.filter(function(o){ return o.currentStatus === 'AMBIGUOUS'; }).length],
      ['canonical 증빙', canonical.length],
      ['유효 롯데 카드증빙', usable.length],
      ['운영시트 변경', '0'],
      ['출력시트', LOTTEON_REMOTE_TASK.outputSheet],
      ['추가검토등급 집계', reviewKeys.map(function(k){ return k + '=' + reviewCounts[k]; }).join(' / ')],
      ['원인분류 합계', categoryKeys.reduce(function(sum,k){ return sum + categoryCounts[k]; }, 0)],
      ['완료시각', new Date().toISOString()]
    ];
    categoryKeys.forEach(function(k) {
      statusRows.push(['분류_' + k, categoryCounts[k]]);
    });
    issue42WriteStatus_(statusSheet, statusRows);
    SpreadsheetApp.flush();
    outputSheet.showSheet();
    statusSheet.showSheet();
    statusSheet.activate();

    try {
      SpreadsheetApp.getUi().alert(
        'Issue42 진단 완료',
        '대상: ' + unresolved.length + '건\n' +
        'NO_MATCH: ' + unresolved.filter(function(o){ return o.currentStatus === 'NO_MATCH'; }).length + '건\n' +
        'AMBIGUOUS: ' + unresolved.filter(function(o){ return o.currentStatus === 'AMBIGUOUS'; }).length + '건\n' +
        '운영시트 변경: 0건\n\nISSUE42_진단상태 표 전체를 복사해 전달하세요.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (ignore) {}

    return {
      ok: true,
      taskId: LOTTEON_REMOTE_TASK.id,
      status: 'PASS',
      targetOrders: unresolved.length,
      categories: categoryCounts,
      reviews: reviewCounts,
      productionWrites: 0
    };
  } catch (error) {
    var message = String(error && error.message ? error.message : error);
    issue42WriteStatus_(statusSheet, [
      ['항목', '값'],
      ['버전', 'v1.0-ISSUE42-H1-UNMATCHED-DIAGNOSTIC'],
      ['상태', 'ERROR'],
      ['단계', 'FAILED'],
      ['메시지', '잔여 카드매칭 진단 실패'],
      ['오류', message],
      ['운영시트 변경', '0'],
      ['갱신시각', new Date().toISOString()]
    ]);
    throw error;
  }
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemoteTaskStartRemote_();
}

function issue42Classify_(order, w07, w814, wNeg, w30) {
  if (!order.orderDate) return { category:'주문일공란', review:'유지', memo:'주문일이 없어 일자구간 진단 불가' };
  if (!issue42Number_(order.purchase)) return { category:'금액0원', review:'유지', memo:'주문매입금액 0원' };

  if (w07.unused.length === 1 && w07.identityCount === 1) {
    return { category:'0~+7일_단일미사용후보', review:'최우선재검토', memo:'현행 규칙에서 미매칭으로 남은 이유 확인 필요' };
  }
  if (w07.unused.length > 1 && w07.identityCount > 1) {
    return { category:'0~+7일_다중카드', review:'유지', memo:'동일금액 서로 다른 카드 후보' };
  }
  if (w07.unused.length > 1 && w07.identityCount === 1) {
    return { category:'0~+7일_동일카드다중건', review:'수동증빙검토', memo:'카드는 같지만 승인건 1:1 식별 불가' };
  }
  if (w07.total > 0 && w07.unused.length === 0) {
    return { category:'0~+7일_exact후보_전부기배정', review:'수동증빙검토', memo:'동일금액 후보가 모두 다른 주문에 사용됨' };
  }

  if (w814.unused.length === 1 && w814.identityCount === 1) {
    return { category:'+8~+14일_단일미사용후보', review:'확장규칙검토', memo:'현재 +7일 범위 밖 단일후보' };
  }
  if (w814.unused.length > 1 && w814.identityCount > 1) {
    return { category:'+8~+14일_다중카드', review:'유지', memo:'+8~+14일 구간 서로 다른 카드 후보' };
  }
  if (w814.unused.length > 1 && w814.identityCount === 1) {
    return { category:'+8~+14일_동일카드다중건', review:'수동증빙검토', memo:'+8~+14일 카드 동일, 승인건 다중' };
  }

  if (wNeg.unused.length === 1 && wNeg.identityCount === 1) {
    return { category:'주문일이전_1~7일_단일미사용후보', review:'역방향검토', memo:'주문일 이전 승인 단일후보' };
  }
  if (wNeg.unused.length > 1 && wNeg.identityCount > 1) {
    return { category:'주문일이전_1~7일_다중카드', review:'유지', memo:'주문일 이전 서로 다른 카드 후보' };
  }
  if (wNeg.unused.length > 1 && wNeg.identityCount === 1) {
    return { category:'주문일이전_1~7일_동일카드다중건', review:'수동증빙검토', memo:'주문일 이전 카드 동일, 승인건 다중' };
  }

  if (w30.unused.length > 0) {
    return { category:'±30일_exact후보_허용범위외', review:'수동증빙검토', memo:'exact 금액 증빙은 있으나 현재 검토구간 밖' };
  }
  if (w30.total > 0 && w30.unused.length === 0) {
    return { category:'±30일_exact후보_전부기배정', review:'유지', memo:'±30일 exact 후보가 모두 다른 주문에 사용됨' };
  }
  if (!issue42Text_(order.payment)) {
    return { category:order.currentStatus === 'AMBIGUOUS' ? '결제수단공란_AMBIGUOUS유지' : '결제수단공란_exact증빙없음', review:'유지', memo:'결제수단 공란이며 exact 금액 증빙 없음' };
  }
  return { category:'exact금액증빙없음', review:'유지', memo:'±30일 내 동일금액 유효 롯데 카드증빙 없음' };
}

function issue42WindowStats_(candidates, minLag, maxLag) {
  var rows = (candidates || []).filter(function(c){ return c.lag >= minLag && c.lag <= maxLag; });
  var unused = rows.filter(function(c){ return !c.used; });
  var used = rows.filter(function(c){ return c.used; });
  var identities = {};
  unused.forEach(function(c){ identities[c.identity || 'UNKNOWN'] = true; });
  return { rows:rows, unused:unused, used:used, total:rows.length, identityCount:Object.keys(identities).length };
}

function issue42PickBest_(w07, w814, wNeg, w30) {
  var groups = [w07.unused, w814.unused, wNeg.unused, w30.unused, w07.rows, w814.rows, wNeg.rows, w30.rows];
  for (var i = 0; i < groups.length; i++) {
    if (groups[i] && groups[i].length) {
      return groups[i].slice().sort(function(a,b) {
        return Math.abs(a.lag) - Math.abs(b.lag) ||
          issue42Text_(a.h.date).localeCompare(issue42Text_(b.h.date)) ||
          issue42Text_(a.h.approvalNo).localeCompare(issue42Text_(b.h.approvalNo));
      })[0];
    }
  }
  return null;
}

function issue42FilterPaymentCompatible_(candidates, payment) {
  var raw = issue42Text_(payment);
  if (!raw) return [];
  var normalizedIssuer = typeof normalizeCardCompany_v660_ === 'function' ? normalizeCardCompany_v660_(raw) : '';
  if (normalizedIssuer) {
    var issuerRows = candidates.filter(function(c) {
      return normalizeCardCompany_v660_(c.h.company) === normalizedIssuer;
    });
    if (issuerRows.length) return issuerRows;
  }
  if (typeof filterEvidenceByLottePayment_v660_ === 'function') {
    var evidence = candidates.map(function(c){ return c.h; });
    var filtered = filterEvidenceByLottePayment_v660_(evidence, raw) || [];
    var keys = {};
    filtered.forEach(function(h){ keys[issue42EvidenceKey_(h)] = true; });
    var rows = candidates.filter(function(c){ return !!keys[issue42EvidenceKey_(c.h)]; });
    if (rows.length) return rows;
  }
  return [];
}

function issue42Identity_(h, master) {
  if (typeof cardIdentityKey_v662_ === 'function') {
    var key = cardIdentityKey_v662_(h, master || []);
    if (key) return key;
  }
  return [
    typeof normalizeCardCompany_v660_ === 'function' ? normalizeCardCompany_v660_(h.company) : issue42Text_(h.company),
    issue42Text_(h.cardEnd4),
    issue42Text_(h.cardName)
  ].join('|');
}

function issue42CandidateSummary_(candidates) {
  return (candidates || []).slice().sort(function(a,b){
    return a.lag - b.lag || issue42Text_(a.h.date).localeCompare(issue42Text_(b.h.date));
  }).slice(0, 12).map(function(c) {
    return [
      issue42DateText_(c.h.date),
      c.lag >= 0 ? '+' + c.lag : String(c.lag),
      issue42Text_(c.h.company),
      issue42Text_(c.h.cardEnd4),
      issue42Text_(c.h.approvalNo),
      issue42Number_(c.h.v664EffectiveAmount || c.h.amount),
      c.used ? 'USED' : 'UNUSED'
    ].join('/');
  }).join(' | ');
}

function issue42BuildHeaderIndex_(headers) {
  var map = {};
  (headers || []).forEach(function(h, i){ map[issue42Text_(h)] = i; });
  return map;
}

function issue42ApprovalKey_(company, approvalNo) {
  var issuer = typeof normalizeCardCompany_v660_ === 'function'
    ? normalizeCardCompany_v660_(company)
    : issue42Text_(company).toUpperCase();
  return issuer + '|' + issue42Text_(approvalNo);
}

function issue42EvidenceKey_(h) {
  return issue42Text_(h.v664CanonicalKey) || [issue42Text_(h.company), issue42Text_(h.approvalNo), issue42Text_(h.date), issue42Number_(h.amount)].join('|');
}

function issue42DaysBetween_(fromText, toText) {
  if (typeof daysBetween_v664_ === 'function') return daysBetween_v664_(fromText, toText);
  var a = issue42DateParts_(fromText), b = issue42DateParts_(toText);
  if (!a || !b) return 99999;
  return Math.round((Date.UTC(b[0],b[1]-1,b[2]) - Date.UTC(a[0],a[1]-1,a[2])) / 86400000);
}

function issue42DateParts_(value) {
  var m = issue42DateText_(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? [Number(m[1]),Number(m[2]),Number(m[3])] : null;
}

function issue42DateText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var text = issue42Text_(value);
  var m = text.match(/^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})/);
  if (!m) return text;
  return [m[1], ('0' + m[2]).slice(-2), ('0' + m[3]).slice(-2)].join('-');
}

function issue42Text_(value) {
  return value == null ? '' : String(value).trim();
}

function issue42Number_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  var text = issue42Text_(value).replace(/,/g, '').replace(/[^\d.\-]/g, '');
  var number = Number(text);
  return isFinite(number) ? number : 0;
}

function issue42EnsureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function issue42WriteStatus_(sheet, rows) {
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 720);
  sheet.showSheet();
  SpreadsheetApp.flush();
}

// ================================================================
// 윌로그 인사발령 자동 동기화 — SyncAppointments.gs
// v1.0 (2026-07-28) : CLAUDE.md 6-1 + 6-2 구현
//
//  syncAppointments()      … 트리거 진입점. 아래 3종을 한 배치로 처리 후 append
//    ① 공고 원천 시트 → [인사카드_발령]  (6-1)
//    ② [조직도]/[퇴직자] 입사일 → '입사' 발령 자동 생성 (6-2)
//    ③ [퇴직자] 퇴직일 → '퇴사' 발령 자동 생성 (6-2)
//  previewAppointments()   … ★ 쓰기 없이 파싱/추가예정 건만 로그 (첫 실행 전 검증용)
//  syncBalryeong()         … 공고 동기화(①)만 단독 실행
//
//  중복 방지 키: 사번 | 발령일 | 발령유형 | 변경후_소속  (공백·기호·대소문자 무시)
//    + '입사'는 사번당 1건, '퇴사'도 사번당 1건으로 추가 제한 (수기 입력분 보호)
//  안전장치:
//    - [조직도] 또는 [퇴직자]에 존재하는 사번만 반영 (공고 양식의 더미 사번 200000000 차단)
//    - LockService 로 트리거·수동 실행 동시성 차단
//    - 기존 행은 절대 수정·삭제하지 않음 (append 전용)
//
//  ※ 웹앱 라우트와 무관한 내부 배치 → clasp push 만으로 반영 (새 버전 배포 불필요)
// ================================================================

// ── 상수 ────────────────────────────────────────────────────────
var SA_GONGGO_SS_ID = '1-pZ0qKTddrINBjWKfrs9kRytcMunr1xOq3vnJu_aPy4'; // 인사발령 공고 원천 시트
var SA_APPT_SHEET   = '인사카드_발령';

// 퇴직자에게도 '입사' 발령을 소급 생성할지 여부.
//   true  = [조직도] + [퇴직자] 전원 (퇴사 발령만 있고 입사 발령이 없는 불일치 방지)
//   false = [조직도] 재직·퇴직예정 인원만 (CLAUDE.md 6-2 문구 그대로)
var SA_INCLUDE_LEAVER_HIRE = true;

// 파싱에서 제외할 탭 이름 패턴 (양식·데모·작업중 탭)
//   ※ listGonggoTabs() 로 탭 목록을 확인한 뒤 필요하면 여기에 추가
var SA_SKIP_TAB_RE = /양식|서식|템플릿|템플렛|샘플|예시|견본|작성중|작업중|백업|template|sample|form|draft|copy|사본/i;

// 직책 서열 — '인사 발령'처럼 유형이 불명확한 행의 승격/변경 판별용
var SA_POS_RANK = {
  '팀원': 1, '파트장': 2, '팀장': 3, '조직장': 4, '부서장': 5, '본부장': 6, 'COO': 7, 'CEO': 8
};
var SA_POS_TOKENS = '팀장|팀원|부서장|본부장|조직장|파트장|CEO|COO|CTO|CFO|CHO';


// ================================================================
//  진입점
// ================================================================
function syncAppointments() {
  return sa_run({ dryRun: false, gonggo: true, hire: true, leave: true });
}

/** 쓰기 없이 결과만 로그 — 최초 실행 전 반드시 이걸로 먼저 확인 */
function previewAppointments() {
  return sa_run({ dryRun: true, gonggo: true, hire: true, leave: true });
}

/** 공고 시트 → [인사카드_발령] 동기화만 단독 실행 (CLAUDE.md 6-1) */
function syncBalryeong() {
  return sa_run({ dryRun: false, gonggo: true, hire: false, leave: false });
}

/** 공고 원천 시트의 탭 목록 + 각 탭의 공고번호·개인발령 행수를 로그 (제외 탭 판별용) */
function listGonggoTabs() {
  var src = SpreadsheetApp.openById(SA_GONGGO_SS_ID);
  var out = [];
  src.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    var data = sheet.getDataRange().getValues();
    var no = '', cnt = 0;
    data.forEach(function(row) {
      row.forEach(function(c) {
        var s = String(c === null || c === undefined ? '' : c).trim();
        if (!no) {
          var m = s.match(/(?:인사|징계)\s*공고\s*[0-9][0-9\-]*\s*호/);
          if (m) no = m[0].replace(/\s+/g, ' ');
        }
        if (/^\d{9}$/.test(s)) cnt++;
      });
    });
    out.push((SA_SKIP_TAB_RE.test(name) ? '⛔ ' : '   ') + name + '  |  ' + (no || '(공고번호 없음)') + '  |  사번행 ' + cnt);
  });
  Logger.log('공고 원천 시트 탭 목록 (' + out.length + '개)\n' + out.join('\n') +
             '\n\n⛔ = SA_SKIP_TAB_RE 로 제외되는 탭');
  return out;
}


// ================================================================
//  본체
// ================================================================
function sa_run(opts) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('⛔ 다른 실행이 진행 중입니다. 중단.');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var apptSheet = ss.getSheetByName(SA_APPT_SHEET);
    if (!apptSheet) { Logger.log('[' + SA_APPT_SHEET + '] 시트 없음. setupHrCardSheets() 먼저 실행하세요.'); return; }

    var warn = [];
    var roster   = sa_loadRoster(ss, warn);          // 사번 → 인적사항 (조직도 + 퇴직자)
    var existing = sa_readExisting(apptSheet, warn); // 기존 발령 행 인덱스
    if (!existing.hdr.length) { Logger.log(warn.join('\n')); return; }

    var cand = [];  // 후보 레코드 {sabun,nick,type,date,beforeOrg,beforePos,afterOrg,afterPos,note,src}
    if (opts.gonggo) cand = cand.concat(sa_parseGonggo(roster, warn));
    if (opts.hire)   cand = cand.concat(sa_buildHireRecords(roster, existing, warn));
    if (opts.leave)  cand = cand.concat(sa_buildLeaveRecords(roster, existing, warn));

    // ── 필터링: 기존 중복 + 배치 내 중복 제거 ──────────────────
    var batchKeys = {};
    var fresh = [], skipped = 0;
    cand.forEach(function(r) {
      var key = sa_apptKey(r.sabun, r.date, r.type, r.afterOrg);
      if (existing.keys[key] || batchKeys[key]) { skipped++; return; }

      // 사번당 1건 제한 유형 (수기 입력분과의 표기 차이로 인한 중복 방지)
      if (r.type === '입사' && existing.hasHire[r.sabun]) { skipped++; return; }
      if (r.type === '퇴사' && existing.hasLeave[r.sabun]) {
        if (existing.hasLeave[r.sabun] !== sa_toDateStr(r.date)) {
          warn.push('퇴사 발령 이미 존재하나 날짜 상이 → 건너뜀: ' + r.sabun +
                    ' (기존 ' + existing.hasLeave[r.sabun] + ' / 신규 ' + sa_toDateStr(r.date) + ')');
        }
        skipped++; return;
      }

      // 유사 중복 경고 (같은 사번·날짜·유형인데 소속 표기만 다름 → 사람이 확인 필요)
      var loose = sa_looseKey(r.sabun, r.date, r.type);
      if (existing.loose[loose]) {
        warn.push('유사 중복 의심 (소속 표기 차이): ' + r.sabun + ' ' + sa_toDateStr(r.date) + ' ' + r.type +
                  ' | 기존 "' + existing.loose[loose] + '" vs 신규 "' + r.afterOrg + '"');
      }

      batchKeys[key] = true;
      if (r.type === '입사') existing.hasHire[r.sabun] = true;
      if (r.type === '퇴사') existing.hasLeave[r.sabun] = sa_toDateStr(r.date);
      fresh.push(r);
    });

    // 발령일 오름차순 (시트가 시간순으로 쌓이도록)
    fresh.sort(function(a, b) {
      var x = sa_toDateStr(a.date), y = sa_toDateStr(b.date);
      return x < y ? -1 : (x > y ? 1 : 0);
    });

    // ── 쓰기 ────────────────────────────────────────────────
    if (fresh.length && !opts.dryRun) {
      var rows = fresh.map(function(r) { return sa_toSheetRow(r, existing.hdr, existing.col); });
      var start = Math.max(apptSheet.getLastRow() + 1, existing.hdrIdx + 2);
      var range = apptSheet.getRange(start, 1, rows.length, existing.hdr.length);
      range.setValues(rows);
      range.setFontFamily('Arial').setFontSize(9);
    }

    // ── 로그 ────────────────────────────────────────────────
    var byType = {};
    fresh.forEach(function(r) { byType[r.type] = (byType[r.type] || 0) + 1; });
    var log = [];
    log.push(opts.dryRun ? '🔍 [DRY-RUN] 실제 쓰기 없음' : '✍️ [인사카드_발령] 반영 완료');
    log.push('후보 ' + cand.length + '건 / 신규 ' + fresh.length + '건 / 중복 스킵 ' + skipped + '건 (기존 ' + existing.count + '행)');
    if (fresh.length) {
      log.push('유형별: ' + Object.keys(byType).map(function(k) { return k + ' ' + byType[k]; }).join(', '));
      fresh.forEach(function(r) {
        log.push('  + ' + sa_toDateStr(r.date) + ' | ' + r.sabun + ' ' + (r.nick || '') + ' | ' + r.type +
                 ' | ' + (r.beforeOrg || '-') + (r.beforePos ? ' (' + r.beforePos + ')' : '') +
                 ' → ' + (r.afterOrg || '-') + (r.afterPos ? ' (' + r.afterPos + ')' : '') +
                 (r.note ? ' | ' + r.note : ''));
      });
    }
    if (warn.length) {
      log.push('');
      log.push('⚠️ 확인 필요 ' + warn.length + '건');
      warn.forEach(function(w) { log.push('  - ' + w); });
    }
    Logger.log('인사발령 자동 동기화\n' + log.join('\n'));
    return { added: fresh.length, skipped: skipped, warnings: warn.length };

  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  ① 공고 원천 시트 파싱 (CLAUDE.md 5장 규칙)
// ================================================================
function sa_parseGonggo(roster, warn) {
  var src;
  try {
    src = SpreadsheetApp.openById(SA_GONGGO_SS_ID);
  } catch (e) {
    warn.push('공고 원천 시트를 열 수 없음 (' + SA_GONGGO_SS_ID + '): ' + e.message);
    return [];
  }

  var out = [];
  var unknown = {};
  var seenNo = {};

  src.getSheets().forEach(function(sheet) {
    var tab = sheet.getName();
    if (SA_SKIP_TAB_RE.test(tab)) { warn.push('탭 제외 (이름 규칙): [' + tab + ']'); return; }

    // rows = 교차검증 통과분 / allRows = 교차검증 이전 전체 (양식 탭 판정용)
    var rows = [], allRows = [], nos = {};
    sa_parseGonggoSheet(sheet, roster, rows, allRows, unknown, warn, nos, tab);

    // 양식·데모 탭 자동 격리 — 한 사람이 같은 날 여러 발령유형을 받는 구성은 실제 공고에 없음
    // ※ 교차검증에서 이미 걸러진 행까지 포함해 판정해야 양식 탭의 특징이 드러난다
    var demo = sa_looksLikeTemplate(allRows);
    if (demo) {
      warn.push('양식·데모 탭으로 판단해 ' + allRows.length + '건 전체 제외: [' + tab + '] — ' + demo +
                ' (실제 공고라면 SA_SKIP_TAB_RE 조정 필요)');
      return;
    }

    Object.keys(nos).forEach(function(no) {
      if (seenNo[no] && seenNo[no] !== tab) {
        warn.push('공고번호 중복: ' + no + ' → 탭 [' + seenNo[no] + '] / [' + tab + '] 내용 확인 필요');
      } else {
        seenNo[no] = tab;
      }
    });

    rows.forEach(function(r) { out.push(r); });
  });

  var us = Object.keys(unknown);
  if (us.length) {
    warn.push('명부에 없는 사번 → 반영 제외 (' + us.length + '건): ' +
              us.map(function(k) { return k + '(' + unknown[k] + ')'; }).join(', '));
  }
  return out;
}

/** 같은 사번·같은 발령일에 발령유형이 3종 이상 → 양식/데모 탭으로 간주 */
function sa_looksLikeTemplate(rows) {
  var g = {};
  rows.forEach(function(r) {
    var k = r.sabun + ' ' + sa_toDateStr(r.date);
    if (!g[k]) g[k] = {};
    g[k][r.type] = true;
  });
  var keys = Object.keys(g);
  for (var i = 0; i < keys.length; i++) {
    var types = Object.keys(g[keys[i]]);
    if (types.length >= 3) {
      return '동일 사번·동일 발령일에 발령유형 ' + types.length + '종 (' + keys[i] + ': ' + types.join(', ') + ')';
    }
  }
  return '';
}

function sa_parseGonggoSheet(sheet, roster, out, allRows, unknown, warn, nos, tab) {
  var data;
  try { data = sheet.getDataRange().getValues(); } catch (e) { return; }
  if (!data.length) return;

  var ctx = { no: '', date: '' };
  var map = null;

  for (var i = 0; i < data.length; i++) {
    var cells = data[i].map(function(c) {
      return (c instanceof Date) ? sa_toDateStr(c) : String(c === null || c === undefined ? '' : c).trim();
    });

    // 공고 번호 ("인사 공고 26-06호" / "징계 공고 25-1호")
    for (var c = 0; c < cells.length; c++) {
      var m = cells[c].match(/(?:인사|징계)\s*공고\s*[0-9][0-9\-]*\s*호/);
      // ※ sa_norm 은 하이픈을 공백으로 바꾸므로 공고번호에는 쓰지 않는다 ("26-03호" 보존)
      if (m) { ctx.no = m[0].replace(/\s+/g, ' ').trim(); ctx.date = ''; map = null; nos[ctx.no] = true; break; }
    }

    // 공고일자
    var di = cells.indexOf('공고일자');
    if (di >= 0) {
      for (var c2 = di + 1; c2 < cells.length; c2++) {
        var dv = sa_extractDate(cells[c2]);
        if (dv) { ctx.date = dv; break; }
      }
    }

    // 개인발령 테이블 헤더
    var hm = sa_gonggoHeaderMap(cells);
    if (hm) { map = hm; continue; }
    if (!map) continue;

    // 데이터 행: 사번 칸이 9자리 숫자여야 함
    var sabun = map.사번 >= 0 ? (cells[map.사번] || '') : '';
    if (!/^\d{9}$/.test(sabun)) continue;

    var g = function(idx) { return idx >= 0 && idx < cells.length ? cells[idx] : ''; };
    var person = sa_parseName(g(map.성명));
    var emp = roster[sabun];
    if (!emp) {
      // 공고 양식의 더미 행(사번 200000000 / 성명 OOO) 및 오기입 차단
      unknown[sabun] = person.name || person.nick || tab;
      continue;
    }

    var dateStr = sa_extractDate(g(map.발령일)) || ctx.date;
    if (!dateStr) {
      warn.push('발령일자 파싱 실패 → 제외: [' + tab + '] ' + sabun + ' ' + person.name);
      continue;
    }

    var rec;
    if (map.발령사항 >= 0) {
      rec = sa_buildGonggoRec(sabun, emp, person, g(map.발령사항), g(map.변경전), g(map.변경후), dateStr, g(map.비고));
    } else if (map.징계 >= 0) {
      // 징계 공고: 컬럼 구성이 다름 (징계 종류 및 내용 / 징계 사유) → 소속 비우고 비고에 요약
      var kind = sa_summary(g(map.징계), 120);
      var why  = sa_summary(g(map.사유), 120);
      rec = {
        sabun: sabun, nick: emp.nick || person.nick, type: /해고|해임/.test(kind) ? '퇴사' : '징계',
        date: dateStr, beforeOrg: '', beforePos: '', afterOrg: '', afterPos: '',
        note: sa_join([kind, why], ' / ')
      };
    } else {
      continue;
    }

    rec.note = sa_join([ctx.no, rec.note], ' · ');
    allRows.push(rec);

    // ── 교차검증: 명부와 모순되는 입·퇴사 발령은 반영하지 않음 (양식 행·오기입 차단)
    if (rec.type === '퇴사' && !emp.left && String(emp.status || '').indexOf('퇴') < 0) {
      warn.push('공고에는 퇴사 발령이나 명부상 재직 중 → 제외: [' + tab + '] ' + sabun + ' ' +
                (emp.nick || person.name) + ' ' + dateStr);
      continue;
    }
    if (rec.type === '입사' && emp.join && sa_toDateStr(dateStr) !== emp.join) {
      warn.push('공고 입사일이 명부 입사일과 불일치 → 제외: [' + tab + '] ' + sabun + ' ' +
                (emp.nick || person.name) + ' 공고 ' + dateStr + ' vs 명부 ' + emp.join);
      continue;
    }

    out.push(rec);
  }
}

/** 개인발령/징계 테이블 헤더 행이면 컬럼 매핑 반환, 아니면 null */
function sa_gonggoHeaderMap(cells) {
  var iS = cells.indexOf('사번');
  var iN = cells.indexOf('성명');
  if (iS < 0 || iN < 0) return null;

  var pick = function(names) {
    for (var k = 0; k < names.length; k++) {
      var j = cells.indexOf(names[k]);
      if (j >= 0) return j;
    }
    // 공백 차이 허용 (예: "변경  전")
    for (var k2 = 0; k2 < names.length; k2++) {
      var flat = sa_flat(names[k2]);
      for (var i = 0; i < cells.length; i++) if (sa_flat(cells[i]) === flat) return i;
    }
    return -1;
  };

  return {
    사번:     iS,
    성명:     iN,
    발령사항: pick(['발령사항', '발령 사항']),
    징계:     pick(['징계 종류 및 내용', '징계종류 및 내용', '징계 종류']),
    사유:     pick(['징계 사유', '징계사유']),
    변경전:   pick(['변경 전', '변경전']),
    변경후:   pick(['변경 후', '변경후']),
    발령일:   pick(['발령일자', '시행일자', '발령일']),
    비고:     pick(['비고'])
  };
}

/** 개인발령 행 1건 → 레코드 */
function sa_buildGonggoRec(sabun, emp, person, action, beforeRaw, afterRaw, dateStr, memo) {
  var bp = sa_splitOrgPos(beforeRaw);
  var ap = sa_splitOrgPos(afterRaw);
  var type = sa_deriveType(action, beforeRaw, afterRaw);
  var note = sa_norm(memo);

  if (type === '레벨업') {
    // 공고에 전/후 레벨 수치가 없음 → 소속·직책 칸을 비워 오해 방지
    bp = { org: '', pos: '' };
    ap = { org: '', pos: '' };
    note = sa_join(['Job-Level UP', note], ' · ');
  } else if (type === '정규직 전환') {
    bp = { org: '', pos: '' };
    ap = { org: '', pos: '' };
    note = sa_join(['수습 → 정규직', note], ' · ');
  } else if (type === '입사') {
    bp = { org: '', pos: '' };
  } else if (type === '퇴사') {
    if (!bp.org && !bp.pos) bp = ap;
    ap = { org: '', pos: '' };
  }

  return {
    sabun: sabun,
    nick: emp.nick || person.nick,
    type: type,
    date: dateStr,
    beforeOrg: bp.org, beforePos: bp.pos,
    afterOrg: ap.org,  afterPos: ap.pos,
    note: note
  };
}


// ── 소속 / 직책 분리 ─────────────────────────────────────────────
function sa_splitOrgPos(raw) {
  var t = sa_norm(raw);
  if (!t) return { org: '', pos: '' };

  // 1) 괄호 표기: "Synapse 부서 PO/PM 팀 (팀장)" / "(팀장/겸)"
  var m = t.match(new RegExp('\\(([^()]*(?:' + SA_POS_TOKENS + ')[^()]*)\\)\\s*$'));
  if (m) return { org: sa_norm(t.slice(0, m.index)), pos: sa_norm(m[1]) };

  // 2) 말미 직책 토큰: "Streamline Works 부서 Backend 팀장"
  var m2 = t.match(new RegExp('(' + SA_POS_TOKENS + ')\\s*(\\(겸\\))?\\s*$'));
  if (m2) {
    var pos = m2[1] + (m2[2] ? '(겸)' : '');
    var org = sa_norm(t.slice(0, m2.index));
    // "Backend 팀장"처럼 조직 단위 명사가 생략된 경우 보정
    // (앞부분에 다른 직책 토큰이 남아 있으면 구조가 복잡한 케이스이므로 손대지 않음)
    if (org && !new RegExp('(' + SA_POS_TOKENS + '|Head of)', 'i').test(org)) {
      if (/^(팀장|팀원|파트장)$/.test(m2[1]) && !/팀$/.test(org))   org += ' 팀';
      else if (m2[1] === '부서장' && !/부서$/.test(org))            org += ' 부서';
      else if (m2[1] === '본부장' && !/본부$/.test(org))            org += ' 본부';
    }
    return { org: org, pos: pos };
  }

  // 3) 해석 불가 (겸직 담당자 나열 등) → 원문을 소속에 보존
  return { org: t, pos: '' };
}


// ── 발령유형 도출 ────────────────────────────────────────────────
function sa_deriveType(action, before, after) {
  var a = sa_norm(action);

  if (/job\s*-?\s*level\s*up/i.test(a) || /레벨\s*업/.test(a)) return '레벨업';
  if (a.indexOf('정규직') >= 0) return '정규직 전환';
  if (a.indexOf('퇴사') >= 0 || a.indexOf('해고') >= 0 || a.indexOf('사직') >= 0) return '퇴사';
  if (a.indexOf('입사') >= 0 || a.indexOf('채용') >= 0) return '입사';
  if (a.indexOf('복직') >= 0) return '복직';
  if (a.indexOf('휴직') >= 0) return '휴직';

  var parts = [];
  if (/소속\s*변경|부서\s*이동|팀\s*이동|소속\s*이동|전보/.test(a)) parts.push('부서이동');
  if (/직무\s*변경/.test(a)) parts.push('직무변경');
  if (/직책\s*변경|보직\s*변경|승격|승진/.test(a)) parts.push(sa_isPromotion(before, after) ? '직책승격' : '직책변경');
  if (/겸임|겸직/.test(a)) parts.push('겸직');
  if (parts.length) return parts.join('/');

  // '인사 발령' 같은 포괄 표현 → 변경 전/후 비교로 도출
  return sa_compareType(before, after);
}

function sa_compareType(beforeRaw, afterRaw) {
  // 겸직 표기 증감 우선 판별 ("... 담당자 (겸)" 나열형은 소속 파싱이 불가능)
  var bc = (String(beforeRaw || '').match(/\(\s*겸\s*\)/g) || []).length;
  var ac = (String(afterRaw  || '').match(/\(\s*겸\s*\)/g) || []).length;
  if (bc !== ac) return ac > bc ? '겸직' : '겸직해제';

  var b = sa_splitOrgPos(beforeRaw), a = sa_splitOrgPos(afterRaw);
  var parts = [];
  if (b.org && a.org && sa_flat(b.org) !== sa_flat(a.org)) parts.push('부서이동');
  if (sa_flat(b.pos) !== sa_flat(a.pos)) {
    if (a.pos) parts.push(sa_isPromotion(beforeRaw, afterRaw) ? '직책승격' : '직책변경');
  }
  return parts.length ? parts.join('/') : '인사발령';
}

function sa_isPromotion(beforeRaw, afterRaw) {
  return sa_rank(sa_splitOrgPos(afterRaw).pos) > sa_rank(sa_splitOrgPos(beforeRaw).pos);
}

function sa_rank(pos) {
  var s = String(pos || ''), best = 0;
  Object.keys(SA_POS_RANK).forEach(function(k) {
    if (s.indexOf(k) >= 0 && SA_POS_RANK[k] > best) best = SA_POS_RANK[k];
  });
  return best;
}


// ================================================================
//  ② 입사 발령 자동 생성
// ================================================================
function sa_buildHireRecords(roster, existing, warn) {
  var out = [], noDate = [];
  Object.keys(roster).forEach(function(sabun) {
    var p = roster[sabun];
    if (existing.hasHire[sabun]) return;
    if (!SA_INCLUDE_LEAVER_HIRE && p.left) return;
    if (!p.join) { if (!p.left) noDate.push(sabun + (p.nick ? '(' + p.nick + ')' : '')); return; }

    out.push({
      sabun: sabun,
      nick: p.nick,
      type: '입사',
      date: p.join,
      beforeOrg: '', beforePos: '',
      afterOrg: sa_orgPath(p),
      afterPos: p.직책,
      note: '자동생성 (' + (p.left ? '퇴직자' : '조직도') + ' 입사일 기준)',
      src: 'hire'
    });
  });
  if (noDate.length) warn.push('입사일 없음 → 입사 발령 생성 불가: ' + noDate.join(', '));
  return out;
}


// ================================================================
//  ③ 퇴사 발령 자동 생성
// ================================================================
function sa_buildLeaveRecords(roster, existing, warn) {
  var out = [], noDate = [];
  Object.keys(roster).forEach(function(sabun) {
    var p = roster[sabun];
    if (!p.left) return;                       // 퇴직 확정자만 (퇴직예정은 제외)
    if (existing.hasLeave[sabun]) return;
    if (!p.leave) {
      noDate.push(sabun + (p.nick ? '(' + p.nick + ')' : ''));
      return;
    }
    out.push({
      sabun: sabun,
      nick: p.nick,
      type: '퇴사',
      date: p.leave,
      beforeOrg: sa_orgPath(p),
      beforePos: p.직책,
      afterOrg: '', afterPos: '',
      note: sa_join(['자동생성', p.leaveKind], ' · '),
      src: 'leave'
    });
  });
  if (noDate.length) {
    warn.push('조직도 재직상태 퇴직 표기이나 [퇴직자] 퇴직일 없음 → 퇴사 발령 생성 불가: ' + noDate.join(', '));
  }
  return out;
}


// ================================================================
//  명부 로드: [조직도] + [퇴직자]
// ================================================================
function sa_loadRoster(ss, warn) {
  var roster = {};

  // ── 조직도 ──
  var org = ss.getSheetByName('조직도');
  if (!org) { warn.push('[조직도] 시트 없음'); }
  else {
    var od = org.getDataRange().getValues();
    var oh = sa_findHeader(od, ['사번']);
    if (!oh) warn.push('[조직도] 헤더 탐지 실패');
    else {
      var c = {
        사번:   sa_findCol(oh.hdr, ['사번']),
        닉네임: sa_findCol(oh.hdr, ['닉네임(영문)', '닉네임']),
        성명:   sa_findCol(oh.hdr, ['한국이름', '성명']),
        상태:   sa_findCol(oh.hdr, ['재직상태', '재직여부']),
        본부:   sa_findCol(oh.hdr, ['본부']),
        부서:   sa_findCol(oh.hdr, ['부서']),
        팀:     sa_findCol(oh.hdr, ['팀']),
        직책:   sa_findCol(oh.hdr, ['직책']),
        입사일: sa_findCol(oh.hdr, ['입사일'])
      };
      for (var i = oh.idx + 1; i < od.length; i++) {
        var r = od[i];
        var sb = c.사번 >= 0 ? String(r[c.사번] || '').trim() : '';
        if (!sb || !/^\d{6,}$/.test(sb)) continue;
        var st = c.상태 >= 0 ? String(r[c.상태] || '').trim() : '';
        // 재직 판정 (CLAUDE.md 3-6): '퇴' 포함 && '예정' 미포함 = 퇴직
        var left = st.indexOf('퇴') >= 0 && st.indexOf('예정') < 0;
        roster[sb] = {
          nick: c.닉네임 >= 0 ? String(r[c.닉네임] || '').trim() : '',
          name: c.성명   >= 0 ? String(r[c.성명]   || '').trim() : '',
          본부: c.본부   >= 0 ? String(r[c.본부]   || '').trim() : '',
          부서: c.부서   >= 0 ? String(r[c.부서]   || '').trim() : '',
          팀:   c.팀     >= 0 ? String(r[c.팀]     || '').trim() : '',
          직책: c.직책   >= 0 ? String(r[c.직책]   || '').trim() : '',
          join: c.입사일 >= 0 ? sa_toDateStr(r[c.입사일]) : '',
          status: st,
          left: left,
          leave: '', leaveKind: ''
        };
      }
    }
  }

  // ── 퇴직자 (조직도에 없는 과거 인원 포함) ──
  var lv = ss.getSheetByName('퇴직자');
  if (lv) {
    var ld = lv.getDataRange().getValues();
    var lh = sa_findHeader(ld, ['퇴직일', '퇴사일']);
    if (lh) {
      var lc = {
        사번:   sa_findCol(lh.hdr, ['사번']),
        닉네임: sa_findCol(lh.hdr, ['닉네임']),
        성명:   sa_findCol(lh.hdr, ['한국이름', '성명', '이름']),
        본부:   sa_findCol(lh.hdr, ['본부']),
        팀:     sa_findCol(lh.hdr, ['팀', '부서']),
        직책:   sa_findCol(lh.hdr, ['직책']),
        입사일: sa_findCol(lh.hdr, ['입사일']),
        퇴직일: sa_findCol(lh.hdr, ['퇴직일', '퇴사일']),
        구분:   sa_findCol(lh.hdr, ['퇴직구분', '구분'])
      };
      for (var j = lh.idx + 1; j < ld.length; j++) {
        var lr = ld[j];
        var lsb = lc.사번 >= 0 ? String(lr[lc.사번] || '').trim() : '';
        if (!lsb || !/^\d{6,}$/.test(lsb)) continue;
        var leaveStr = lc.퇴직일 >= 0 ? sa_toDateStr(lr[lc.퇴직일]) : '';
        if (!leaveStr) continue;

        var p = roster[lsb];
        if (!p) {
          p = roster[lsb] = { nick: '', name: '', 본부: '', 부서: '', 팀: '', 직책: '', join: '', status: '퇴직' };
        }
        p.left = true;
        // 재입사 대비: 가장 늦은 퇴직일 채택
        if (!p.leave || leaveStr > p.leave) {
          p.leave     = leaveStr;
          p.leaveKind = lc.구분 >= 0 ? String(lr[lc.구분] || '').trim() : '';
        }
        // 조직도에 없는 인원은 퇴직자 시트 값으로 보완
        if (!p.nick && lc.닉네임 >= 0) p.nick = String(lr[lc.닉네임] || '').trim();
        if (!p.name && lc.성명   >= 0) p.name = String(lr[lc.성명]   || '').trim();
        if (!p.본부 && lc.본부   >= 0) p.본부 = String(lr[lc.본부]   || '').trim();
        if (!p.팀   && lc.팀     >= 0) p.팀   = String(lr[lc.팀]     || '').trim();
        if (!p.직책 && lc.직책   >= 0) p.직책 = String(lr[lc.직책]   || '').trim();
        if (!p.join && lc.입사일 >= 0) p.join = sa_toDateStr(lr[lc.입사일]);
      }
    } else {
      warn.push('[퇴직자] 헤더 탐지 실패');
    }
  }

  return roster;
}


// ================================================================
//  기존 [인사카드_발령] 인덱싱
// ================================================================
function sa_readExisting(sheet, warn) {
  var data = sheet.getDataRange().getValues();
  var h = sa_findHeader(data, ['발령유형']) || sa_findHeader(data, ['사번']);
  if (!h) {
    warn.push('[' + SA_APPT_SHEET + '] 헤더 탐지 실패 — 1행에 사번/발령유형 헤더가 필요합니다.');
    return { hdr: [], hdrIdx: 0, col: {}, keys: {}, loose: {}, hasHire: {}, hasLeave: {}, count: 0 };
  }

  var col = {
    사번:       sa_findCol(h.hdr, ['사번']),
    닉네임:     sa_findCol(h.hdr, ['닉네임']),
    발령유형:   sa_findCol(h.hdr, ['발령유형']),
    발령일:     sa_findCol(h.hdr, ['발령일']),
    변경전_소속: sa_findCol(h.hdr, ['변경전_소속', '변경전 소속']),
    변경전_직책: sa_findCol(h.hdr, ['변경전_직책', '변경전 직책']),
    변경후_소속: sa_findCol(h.hdr, ['변경후_소속', '변경후 소속']),
    변경후_직책: sa_findCol(h.hdr, ['변경후_직책', '변경후 직책']),
    비고:       sa_findCol(h.hdr, ['비고'])
  };

  var keys = {}, loose = {}, hasHire = {}, hasLeave = {}, count = 0;
  for (var i = h.idx + 1; i < data.length; i++) {
    var r = data[i];
    var sb = col.사번 >= 0 ? String(r[col.사번] || '').trim() : '';
    if (!sb || !/^\d{6,}$/.test(sb)) continue;   // 설명행·공란 스킵
    var type  = col.발령유형   >= 0 ? sa_norm(r[col.발령유형]) : '';
    var date  = col.발령일     >= 0 ? sa_toDateStr(r[col.발령일]) : '';
    var after = col.변경후_소속 >= 0 ? sa_norm(r[col.변경후_소속]) : '';
    keys[sa_apptKey(sb, date, type, after)] = true;
    loose[sa_looseKey(sb, date, type)] = after;
    if (type.indexOf('입사') >= 0) hasHire[sb] = true;
    if (type.indexOf('퇴사') >= 0) hasLeave[sb] = date;
    count++;
  }

  return { hdr: h.hdr, hdrIdx: h.idx, col: col, keys: keys, loose: loose,
           hasHire: hasHire, hasLeave: hasLeave, count: count };
}

/** 레코드 → 시트 행 배열 (헤더 순서에 맞춰 배치, 미매칭 컬럼은 공란 유지) */
function sa_toSheetRow(r, hdr, col) {
  var row = [];
  for (var i = 0; i < hdr.length; i++) row.push('');
  var put = function(idx, v) { if (idx >= 0 && idx < row.length) row[idx] = v; };
  put(col.사번, r.sabun);
  put(col.닉네임, r.nick || '');
  put(col.발령유형, r.type);
  put(col.발령일, sa_toDateStr(r.date));
  put(col.변경전_소속, r.beforeOrg || '');
  put(col.변경전_직책, r.beforePos || '');
  put(col.변경후_소속, r.afterOrg || '');
  put(col.변경후_직책, r.afterPos || '');
  put(col.비고, r.note || '');
  return row;
}


// ================================================================
//  공통 유틸 (sa_ 접두어 — 다른 .gs 파일의 전역 함수와 충돌 방지)
// ================================================================
function sa_pad(n) { return String(n).length < 2 ? '0' + n : String(n); }

/** 날짜 값 → "YYYY.MM.DD" (Date / "2026. 5. 17" / "2026-05-17" / "2026년 5월 17일" 수용) */
function sa_toDateStr(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getFullYear() + '.' + sa_pad(v.getMonth() + 1) + '.' + sa_pad(v.getDate());
  }
  return sa_extractDate(v);
}

/** 문자열에서 첫 날짜만 추출 — "2025. 8. 22 (Jake 퇴사일)" → "2025.08.22" */
function sa_extractDate(v) {
  if (v instanceof Date) return sa_toDateStr(v);
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/(\d{4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/);
  return m ? (m[1] + '.' + sa_pad(m[2]) + '.' + sa_pad(m[3])) : '';
}

/** 줄바꿈 제거 + " - " → " " + 공백 정규화 */
function sa_norm(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s*[-–—]\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** 비교·키 생성용: 공백·기호 제거 + 소문자화 */
function sa_flat(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/[\s\-–—·/()（）,.]/g, '')
    .toLowerCase();
}

/** 본부 > 부서 > 팀 을 소속 경로 문자열로 (중복 값은 1회만 — '본부/부서' 단일 컬럼 대응) */
function sa_orgPath(p) {
  var parts = [], seen = {};
  [p.본부, p.부서, p.팀].forEach(function(v) {
    var t = String(v || '').trim();
    if (!t) return;
    var f = sa_flat(t);
    if (seen[f]) return;
    seen[f] = true;
    parts.push(t);
  });
  return parts.join(' ');
}

function sa_join(arr, sep) {
  return arr.filter(function(x) { return x && String(x).trim(); })
            .map(function(x) { return String(x).trim(); })
            .join(sep);
}

function sa_summary(s, max) {
  var t = sa_norm(s);
  return t.length > max ? t.slice(0, max) + '…' : t;
}

/** "이윤철 (Charles)" / "곽인욱(Roni)" → {name, nick} */
function sa_parseName(s) {
  var t = sa_norm(s);
  var m = t.match(/^(.*?)\s*[\(（]([^)）]*)[\)）]\s*$/);
  return m ? { name: sa_norm(m[1]), nick: sa_norm(m[2]) } : { name: t, nick: '' };
}

function sa_apptKey(sabun, date, type, afterOrg) {
  return [String(sabun).trim(), sa_toDateStr(date), sa_flat(type), sa_flat(afterOrg)].join('|');
}

function sa_looseKey(sabun, date, type) {
  return [String(sabun).trim(), sa_toDateStr(date), sa_flat(type)].join('|');
}

/** 1~3행 내 헤더 탐지 */
function sa_findHeader(data, kws) {
  for (var i = 0; i < Math.min(3, data.length); i++) {
    var row = data[i].map(function(c) { return String(c === null || c === undefined ? '' : c).trim(); });
    var ok = row.some(function(c) {
      return kws.some(function(k) { return c === k || c.indexOf(k) >= 0; });
    });
    if (ok) return { idx: i, hdr: row };
  }
  return null;
}

/** 헤더에서 컬럼 인덱스 (완전일치 우선 → 포함 매칭) */
function sa_findCol(hdr, kws) {
  for (var k = 0; k < kws.length; k++) {
    var i = hdr.indexOf(kws[k]);
    if (i >= 0) return i;
  }
  for (var k2 = 0; k2 < kws.length; k2++) {
    for (var j = 0; j < hdr.length; j++) if (hdr[j].indexOf(kws[k2]) >= 0) return j;
  }
  return -1;
}

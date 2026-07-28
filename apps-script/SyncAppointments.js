// ================================================================
// 윌로그 인사발령 자동 동기화 — SyncAppointments.gs
// v1.1 (2026-07-28) : 중복 판정 2단계화 + [발령_프리뷰] 시트 출력
// v1.0 (2026-07-28) : CLAUDE.md 6-1 + 6-2 구현
//
//  syncAppointments()      … 트리거 진입점. 아래 3종을 한 배치로 처리 후 append
//    ① 공고 원천 시트 → [인사카드_발령]  (6-1)
//    ② [조직도]/[퇴직자] 입사일 → '입사' 발령 자동 생성 (6-2)
//    ③ [퇴직자] 퇴직일 → '퇴사' 발령 자동 생성 (6-2)
//  previewAppointments()   … ★ 쓰기 없이 전체 결과를 [발령_프리뷰] 시트에 기록 (첫 실행 전 검증용)
//  previewAppointmentsLogOnly() … 프리뷰 시트 없이 로그로만
//  syncBalryeong()         … 공고 동기화(①)만 단독 실행
//  listGonggoTabs()        … 공고 시트 탭 목록 (제외 탭 판별용)
//
//  ★ 중복 판정 2단계 (v1.1) — [인사카드_발령] 기존 행은 사람이 검수한 정본이라
//     파서 표기와 어휘·띄어쓰기가 다르다 (시트 "팀 이동/팀 개편/팀명 변경" ↔ 파서 "부서이동",
//     시트 "직책변경" ↔ 파서 "직책승격", "Pulse1" ↔ "Pulse 1", "Streamline Work" ↔ "Streamline Works").
//     1차(엄격) 사번|발령일|발령유형|변경후_소속 → 일치하면 이미 반영된 것으로 보고 조용히 스킵
//     2차(느슨) 사번|발령일               → 기존 행과 일치하면 **자동 반영하지 않고**
//                                            '표기 상이 — 수동확인' 목록으로만 올림
//     같은 날 같은 사람에게 복수 발령(겸직 등)이 실제로 있으므로 2차 매칭 건은 전부 사람이 판단한다.
//     ※ 2차 매칭은 **기존 시트 행에 대해서만** 적용 — 배치 내 후보끼리는 1차 키로만 구분
//    + '입사'·'퇴사'는 사번당 1건으로 추가 제한 (수기 입력분 보호)
//  안전장치:
//    - [조직도] 또는 [퇴직자]에 존재하는 사번만 반영 (공고 양식의 더미 사번 200000000 차단)
//    - LockService 로 트리거·수동 실행 동시성 차단
//    - 기존 행은 절대 수정·삭제하지 않음 (append 전용). 프리뷰는 [발령_프리뷰] 시트에만 씀
//
//  ※ 웹앱 라우트와 무관한 내부 배치 → clasp push 만으로 반영 (새 버전 배포 불필요)
// ================================================================

// ── 상수 ────────────────────────────────────────────────────────
var SA_GONGGO_SS_ID  = '1-pZ0qKTddrINBjWKfrs9kRytcMunr1xOq3vnJu_aPy4'; // 인사발령 공고 원천 시트
var SA_APPT_SHEET    = '인사카드_발령';
var SA_PREVIEW_SHEET = '발령_프리뷰';   // previewAppointments() 결과 출력용 임시 시트 (자동 생성/전체 갱신)
var SA_PREVIEW_BANNER = '⚡ previewAppointments() 자동 생성 시트입니다. 직접 편집하지 마세요 (실행할 때마다 전체 덮어씀). 검수 후 삭제해도 무방합니다.';

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

/**
 * 쓰기 없이 결과만 확인 — 최초 실행 전 반드시 이걸로 먼저 검수
 * [인사카드_발령] 에는 손대지 않고, 결과 전체를 [발령_프리뷰] 시트에 구분 컬럼과 함께 기록한다.
 * (실행 로그는 길면 잘리므로 검수는 시트에서)
 */
function previewAppointments() {
  return sa_run({ dryRun: true, previewSheet: true, gonggo: true, hire: true, leave: true });
}

/** 프리뷰 시트 없이 로그로만 확인하고 싶을 때 */
function previewAppointmentsLogOnly() {
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
//  [유지보수] 기존 [인사카드_발령] 정본 유형 표기 일괄 정규화
//  2026-07-28 Ted 승인 계획 — [발령_프리뷰] ③ 검수 결과 반영
//
//  ① "팀 이동 / 팀 개편 / 팀명 변경"  → "부서이동"   (복합유형은 구성요소 단위로 치환 후 중복 제거)
//     "직책변경" 이면서 변경전_직책=팀원, 변경후_직책=팀장 → "직책승격"
//     유형이 바뀐 행은 비고 **앞**에 "[원표기: 원래유형]" 을 붙인다
//  ② 수정한 행 비고 **끝**에 " #확정" 추가 (이후 syncAppointments 의 ③ 알림에서 제외됨)
//  ③ 250010047 / 2026.01.26 (전략 고문) — 유형은 그대로 두고 #확정만
//  ④ 240010004 · 210010008 / 2026.03.16 (26-03호 TFT) — 유형을 "겸직"으로
//  + 사번·발령일·정규화유형이 같은 행이 이미 있으면 비고에 "중복의심(N행과 동일)" 표시 (삭제는 하지 않음)
//
//  실행:  normalizeAppointmentTypes()       … dry-run. 변경 예정 목록만 로그 (쓰기 없음)
//         normalizeAppointmentTypesApply()  … 실제 반영
//  안전:  - 발령유형·비고 **두 셀만** 쓴다. 나머지 컬럼·행은 손대지 않고 삭제도 없음
//         - 비고에 이미 #확정 이 있는 행은 건너뜀 → 재실행해도 중복 적용 안 됨(멱등)
//         - 실행 전 [인사카드_발령] 사본 백업 권장
// ================================================================
var SA_TEAM_MOVE_TOKENS = ['팀이동', '팀명변경', '팀개편', '팀명개편', '팀재편'];
var SA_RULE_TFT  = { '240010004|2026.03.16': 1, '210010008|2026.03.16': 1 }; // ④ → 겸직
var SA_RULE_KEEP = { '250010047|2026.01.26': 1 };                            // ③ 유형 유지

function normalizeAppointmentTypes()      { return sa_normalizeTypes(true); }
function normalizeAppointmentTypesApply() { return sa_normalizeTypes(false); }

function sa_normalizeTypes(dryRun) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('⛔ 다른 실행이 진행 중입니다. 중단.'); return; }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SA_APPT_SHEET);
    if (!sh) { Logger.log('[' + SA_APPT_SHEET + '] 시트 없음'); return; }

    var data = sh.getDataRange().getValues();
    var h = sa_findHeader(data, ['발령유형']) || sa_findHeader(data, ['사번']);
    if (!h) { Logger.log('[' + SA_APPT_SHEET + '] 헤더 탐지 실패'); return; }

    var c = {
      사번:   sa_findCol(h.hdr, ['사번']),
      닉네임: sa_findCol(h.hdr, ['닉네임']),
      유형:   sa_findCol(h.hdr, ['발령유형']),
      발령일: sa_findCol(h.hdr, ['발령일']),
      전직책: sa_findCol(h.hdr, ['변경전_직책', '변경전 직책']),
      후직책: sa_findCol(h.hdr, ['변경후_직책', '변경후 직책']),
      비고:   sa_findCol(h.hdr, ['비고'])
    };
    if (c.사번 < 0 || c.유형 < 0 || c.발령일 < 0 || c.비고 < 0) {
      Logger.log('필수 컬럼(사번/발령유형/발령일/비고) 탐지 실패 → 중단'); return;
    }

    var g = function(row, idx) { return idx >= 0 ? String(row[idx] || '').trim() : ''; };

    // ── 1차 스캔: 모든 데이터 행의 정규화 유형을 구해 중복(사번|발령일|정규화유형) 최초 행 기록
    var firstSeen = {}, items = [];
    for (var i = h.idx + 1; i < data.length; i++) {
      var row = data[i];
      var sabun = g(row, c.사번);
      if (!sabun || !/^\d{6,}$/.test(sabun)) continue;    // 설명행·공란 스킵

      var it = {
        sheetRow: i + 1,
        sabun: sabun,
        nick: g(row, c.닉네임),
        type: g(row, c.유형),
        date: sa_toDateStr(row[c.발령일]),
        bPos: g(row, c.전직책),
        aPos: g(row, c.후직책),
        note: g(row, c.비고)
      };
      var res = sa_normType(it);
      it.newType = res.type;
      it.why = res.why;

      var dk = it.sabun + '|' + it.date + '|' + it.newType;
      if (firstSeen[dk] === undefined) firstSeen[dk] = it.sheetRow;
      else it.dupOf = firstSeen[dk];

      items.push(it);
    }

    // ── 2차: 실제 변경 계획 수립
    var plan = [], settled = 0;
    items.forEach(function(it) {
      if (it.note.indexOf('#확정') >= 0) { settled++; return; }   // 이미 처리된 행

      var typeChanged = (it.newType !== it.type);
      var isKeep = !!SA_RULE_KEEP[it.sabun + '|' + it.date];
      if (!typeChanged && !isKeep && !it.dupOf) return;           // 손댈 이유 없음

      var note = it.note;
      if (typeChanged) note = '[원표기: ' + it.type + '] ' + note;
      if (it.dupOf) note = note + (note.slice(-1) === ' ' ? '' : ' ') + '중복의심(' + it.dupOf + '행과 동일)';
      note = (note + ' #확정').replace(/\s+/g, ' ').trim();

      plan.push({ it: it, newNote: note, typeChanged: typeChanged });
    });

    // ── 쓰기 (발령유형·비고 두 셀만) ─────────────────────────
    if (!dryRun) {
      plan.forEach(function(p) {
        if (p.typeChanged) sh.getRange(p.it.sheetRow, c.유형 + 1).setValue(p.it.newType);
        sh.getRange(p.it.sheetRow, c.비고 + 1).setValue(p.newNote);
      });
    }

    // ── 로그 ────────────────────────────────────────────────
    var log = [];
    log.push(dryRun ? '🔍 [DRY-RUN] 쓰기 없음 — 실제 반영은 normalizeAppointmentTypesApply()'
                    : '✍️ [' + SA_APPT_SHEET + '] 유형 표기 정규화 반영 완료');
    log.push('데이터 ' + items.length + '행 / 변경 ' + plan.length + '행 / 이미 #확정 ' + settled + '행 건너뜀');
    log.push('');
    plan.forEach(function(p) {
      var it = p.it;
      log.push(it.sheetRow + '행 ' + it.sabun + ' ' + (it.nick || '') + ' ' + it.date +
               ' | ' + it.type + (p.typeChanged ? ' → ' + it.newType : ' (유형 유지)') +
               (it.why.length ? '  [' + it.why.join('; ') + ']' : '') +
               (it.dupOf ? '  [중복의심 ' + it.dupOf + '행]' : ''));
      log.push('      비고: "' + it.note + '"  →  "' + p.newNote + '"');
    });
    Logger.log('발령유형 표기 정규화\n' + log.join('\n'));

    return { total: items.length, changed: plan.length, settledSkipped: settled, dryRun: !!dryRun };

  } finally {
    lock.releaseLock();
  }
}

/** 한 행의 정규화 유형 계산 → {type, why} */
function sa_normType(it) {
  var key = it.sabun + '|' + it.date;
  if (SA_RULE_TFT[key])  return { type: '겸직', why: ['④ 26-03호 TFT 겸직'] };
  if (SA_RULE_KEEP[key]) return { type: it.type, why: ['③ 유형 유지 (#확정만)'] };

  var why = [], parts = String(it.type).split('/'), out = [];
  parts.forEach(function(raw) {
    var p = raw.trim();
    if (!p) return;
    var flat = p.replace(/\s/g, '');
    if (SA_TEAM_MOVE_TOKENS.indexOf(flat) >= 0) {
      out.push('부서이동');
      why.push('①a "' + p + '" → 부서이동');
    } else if (flat === '직책변경' && it.bPos === '팀원' && it.aPos === '팀장') {
      out.push('직책승격');
      why.push('①b 직책변경(팀원→팀장) → 직책승격');
    } else if (flat === '직책변경') {
      out.push('직책변경');   // 공백 표기 "직책 변경" 정규화
    } else {
      out.push(p);
    }
  });

  var dedup = [];
  out.forEach(function(p) { if (dedup.indexOf(p) < 0) dedup.push(p); });
  return { type: dedup.join('/'), why: why };
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

    var cand = [];      // 후보 레코드 {sabun,nick,type,date,beforeOrg,beforePos,afterOrg,afterPos,note,tab}
    var rejected = [];  // 파싱 단계에서 제외된 것 {rec, reason}
    if (opts.gonggo) cand = cand.concat(sa_parseGonggo(roster, warn, rejected));
    if (opts.hire)   cand = cand.concat(sa_buildHireRecords(roster, existing, warn));
    if (opts.leave)  cand = cand.concat(sa_buildLeaveRecords(roster, existing, warn));

    // ── 2단계 중복 판정 ────────────────────────────────────────
    //  1차(엄격): 사번|발령일|발령유형|변경후_소속 → 같으면 이미 반영된 것으로 보고 조용히 스킵
    //  2차(느슨): 사번|발령일 만으로 기존 정본 행과 대조 → 걸리면 자동 반영하지 않고
    //             '표기 상이 — 수동확인' 목록으로만 올림
    //             (시트 "팀 이동/팀 개편" vs 파서 "부서이동", "Pulse1" vs "Pulse 1" 같은
    //              표기 차이로 정본이 중복되는 것을 막는다)
    //  ※ 느슨 매칭은 **기존 시트 행에 대해서만** 적용한다. 같은 날 같은 사람에게 복수 발령
    //     (겸직 등)이 실제로 있으므로 배치 내 후보끼리는 1차 키로만 구분한다.
    var batchKeys = {};
    var fresh = [], dupExact = [], dupLoose = [], dupOnce = [];
    cand.forEach(function(r) {
      var key = sa_apptKey(r.sabun, r.date, r.type, r.afterOrg);

      if (existing.keys[key]) { dupExact.push({ rec: r, reason: '기존 행과 정확 일치' }); return; }
      if (batchKeys[key])     { dupExact.push({ rec: r, reason: '이번 배치 내 중복 (공고 2벌 등)' }); return; }

      // 사번당 1건 제한 유형 (수기 입력분 보호)
      if (r.type === '입사' && existing.hasHire[r.sabun]) {
        dupOnce.push({ rec: r, reason: '이미 입사 발령 보유 (사번당 1건 제한)' }); return;
      }
      if (r.type === '퇴사' && existing.hasLeave[r.sabun]) {
        dupOnce.push({ rec: r, reason: '이미 퇴사 발령 보유 (기존 ' + existing.hasLeave[r.sabun] + ')' }); return;
      }

      // ★ 2차: 사번 + 발령일이 같은 기존 행이 있으면 자동 반영 금지
      var hits = existing.byDate[sa_dateKey(r.sabun, r.date)];
      if (hits && hits.length && sa_isSettled(hits)) {
        // 기존 행 비고에 #확정 → 검수 끝난 건. 매 실행마다 재알림하지 않고 조용히 스킵
        dupExact.push({ rec: r, reason: '기존 행 비고 #확정 — 검수 완료 처리됨' });
        return;
      }
      if (hits && hits.length) {
        dupLoose.push({
          rec: r,
          reason: '기존 ' + hits.length + '행과 사번·발령일 동일 → ' +
                  hits.map(function(h) {
                    return '[' + h.row + '행] ' + (h.type || '(유형없음)') + ' / ' +
                           (h.afterOrg || h.beforeOrg || '(소속없음)');
                  }).join(' , ') +
                  '   ※ 검수 후 이 알림을 끄려면 기존 행 비고에 #확정 추가'
        });
        return;
      }

      batchKeys[key] = true;
      if (r.type === '입사') existing.hasHire[r.sabun] = true;
      if (r.type === '퇴사') existing.hasLeave[r.sabun] = sa_toDateStr(r.date);
      fresh.push(r);
    });
    var skipped = dupExact.length + dupOnce.length + dupLoose.length;

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

    // ── 프리뷰 시트 출력 ─────────────────────────────────────
    var unresolved = fresh.filter(function(r) { return r.type === '인사발령'; }).length;
    var previewMsg = '';
    if (opts.previewSheet) {
      previewMsg = sa_writePreviewSheet(ss, {
        fresh: fresh, dupExact: dupExact, dupOnce: dupOnce, dupLoose: dupLoose,
        rejected: rejected, warn: warn, existingCount: existing.count
      });
    }

    // ── 로그 (길면 잘리므로 요약만 — 상세는 [발령_프리뷰] 시트에서) ──
    var byType = {};
    fresh.forEach(function(r) { byType[r.type] = (byType[r.type] || 0) + 1; });
    var log = [];
    log.push(opts.dryRun ? '🔍 [DRY-RUN] [' + SA_APPT_SHEET + '] 쓰기 없음' : '✍️ [' + SA_APPT_SHEET + '] 반영 완료');
    log.push('기존 ' + existing.count + '행 (수정·삭제 없음) / 후보 ' + cand.length + '건');
    log.push('  ① 신규' + (opts.dryRun ? ' 예정' : '') + '            ' + fresh.length + '건' +
             (unresolved ? '   (그중 유형 판별 실패 "인사발령" ' + unresolved + '건)' : ''));
    log.push('  ② 중복 스킵(정확일치)  ' + dupExact.length + '건');
    log.push('  ③ 표기 상이 — 수동확인 ' + dupLoose.length + '건  ← 사번+발령일이 기존 행과 같아 자동 반영 안 함');
    log.push('  ④ 사번당 1건 제한 스킵 ' + dupOnce.length + '건');
    log.push('  ⑤ 파싱 단계 제외        ' + rejected.length + '건');
    if (fresh.length) {
      log.push('신규 유형별: ' + Object.keys(byType).map(function(k) { return k + ' ' + byType[k]; }).join(', '));
    }
    if (previewMsg) { log.push(''); log.push(previewMsg); }
    if (warn.length) {
      log.push('');
      log.push('⚠️ 점검 필요 ' + warn.length + '건');
      warn.forEach(function(w) { log.push('  - ' + w); });
    }
    Logger.log('인사발령 자동 동기화\n' + log.join('\n'));

    return {
      added: opts.dryRun ? 0 : fresh.length,
      pending: fresh.length,
      dupExact: dupExact.length,
      needsReview: dupLoose.length,
      dupOnce: dupOnce.length,
      rejected: rejected.length,
      warnings: warn.length
    };

  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  프리뷰 시트 출력 — [발령_프리뷰]
//  ※ [인사카드_발령] 은 절대 건드리지 않는다. 이 시트만 매 실행 전체 갱신.
// ================================================================
function sa_writePreviewSheet(ss, r) {
  var sh = ss.getSheetByName(SA_PREVIEW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SA_PREVIEW_SHEET);
    sh.setTabColor('#F59E0B');
  } else {
    // 사람이 만든 동명 시트를 덮어쓰지 않도록 배너로 소유권 확인
    var a1 = String(sh.getRange(1, 1).getValue() || '');
    if (sh.getLastRow() > 0 && a1.indexOf('previewAppointments()') < 0) {
      return '⛔ [' + SA_PREVIEW_SHEET + '] 시트에 다른 내용이 있어 덮어쓰지 않았습니다. ' +
             '해당 시트를 삭제하거나 이름을 바꾼 뒤 다시 실행하세요.';
    }
  }

  var CAT = {
    fresh:    '① 신규 예정',
    unres:    '① 신규 예정 (유형 판별 실패)',
    loose:    '③ 표기 상이 — 수동확인',
    once:     '④ 사번당 1건 제한 스킵',
    exact:    '② 중복 스킵 (정확일치)',
    rejected: '⑤ 파싱 단계 제외',
    warn:     '⑨ 점검 필요'
  };
  var COLOR = {};
  COLOR[CAT.fresh]    = '#ECFDF5';
  COLOR[CAT.unres]    = '#FEF3C7';
  COLOR[CAT.loose]    = '#FEE2E2';
  COLOR[CAT.once]     = '#F8FAFC';
  COLOR[CAT.exact]    = '#F8FAFC';
  COLOR[CAT.rejected] = '#F1F5F9';
  COLOR[CAT.warn]     = '#EFF6FF';

  var rows = [];
  var push = function(cat, rec, reason) {
    rec = rec || {};
    rows.push([
      cat,
      rec.sabun || '', rec.nick || '',
      rec.type || '', sa_toDateStr(rec.date),
      rec.beforeOrg || '', rec.beforePos || '',
      rec.afterOrg || '',  rec.afterPos || '',
      rec.note || '', rec.tab || '', reason || ''
    ]);
  };

  // 검수 우선순위 순서로 배치: 수동확인 → 신규(판별실패) → 신규 → 제외 → 스킵 → 경고
  r.dupLoose.forEach(function(x) { push(CAT.loose, x.rec, x.reason); });
  r.fresh.filter(function(x) { return x.type === '인사발령'; })
         .forEach(function(x) { push(CAT.unres, x, '발령사항·전후 비교로 유형을 특정하지 못함 → 유형 직접 지정 필요'); });
  r.fresh.filter(function(x) { return x.type !== '인사발령'; })
         .forEach(function(x) { push(CAT.fresh, x, ''); });
  r.rejected.forEach(function(x) { push(CAT.rejected, x.rec, x.reason); });
  r.dupOnce.forEach(function(x) { push(CAT.once, x.rec, x.reason); });
  r.dupExact.forEach(function(x) { push(CAT.exact, x.rec, x.reason); });
  r.warn.forEach(function(w) { push(CAT.warn, null, w); });

  var HDR = ['구분', '사번', '닉네임', '발령유형', '발령일',
             '변경전_소속', '변경전_직책', '변경후_소속', '변경후_직책', '비고', '출처탭', '판정사유'];
  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');

  sh.clear();
  sh.getRange(1, 1, 1, HDR.length).merge()
    .setValue(SA_PREVIEW_BANNER + '   (갱신: ' + now + ' · 기존 [' + SA_APPT_SHEET + '] ' + r.existingCount + '행은 무변경)')
    .setFontSize(9).setFontStyle('italic').setFontColor('#B45309').setBackground('#FFFBEB');
  sh.setRowHeight(1, 22);

  sh.getRange(2, 1, 1, HDR.length).setValues([HDR])
    .setFontFamily('Arial').setFontWeight('bold').setFontSize(10)
    .setFontColor('#FFFFFF').setBackground('#1E293B')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(2, 26);

  if (rows.length) {
    var rng = sh.getRange(3, 1, rows.length, HDR.length);
    rng.setValues(rows).setFontFamily('Arial').setFontSize(9).setVerticalAlignment('top');
    // 구분별 배경색 (연속 구간 단위로 묶어서 setBackground 호출 최소화)
    var start = 0;
    for (var i = 1; i <= rows.length; i++) {
      if (i === rows.length || rows[i][0] !== rows[start][0]) {
        var bg = COLOR[rows[start][0]];
        if (bg) sh.getRange(start + 3, 1, i - start, HDR.length).setBackground(bg);
        start = i;
      }
    }
  } else {
    sh.getRange(3, 1, 1, HDR.length).merge()
      .setValue('추가할 발령이 없습니다 (모두 기존 행과 중복).')
      .setFontSize(9).setFontStyle('italic').setFontColor('#94A3B8');
  }

  [110, 90, 80, 130, 85, 200, 100, 200, 100, 220, 110, 420]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(2);

  return '📋 [' + SA_PREVIEW_SHEET + '] 시트에 ' + rows.length + '행 기록 완료 — 여기서 검수하세요.';
}


// ================================================================
//  ① 공고 원천 시트 파싱 (CLAUDE.md 5장 규칙)
// ================================================================
function sa_parseGonggo(roster, warn, rejected) {
  rejected = rejected || [];
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
    sa_parseGonggoSheet(sheet, roster, rows, allRows, unknown, warn, nos, tab, rejected);

    // 양식·데모 탭 자동 격리 — 한 사람이 같은 날 여러 발령유형을 받는 구성은 실제 공고에 없음
    // ※ 교차검증에서 이미 걸러진 행까지 포함해 판정해야 양식 탭의 특징이 드러난다
    var demo = sa_looksLikeTemplate(allRows);
    if (demo) {
      warn.push('양식·데모 탭으로 판단해 ' + allRows.length + '건 전체 제외: [' + tab + '] — ' + demo +
                ' (실제 공고라면 SA_SKIP_TAB_RE 조정 필요)');
      allRows.forEach(function(x) { rejected.push({ rec: x, reason: '양식·데모 탭으로 판정되어 제외 — ' + demo }); });
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
              us.map(function(k) { return k + '(' + unknown[k].name + ')'; }).join(', '));
    us.forEach(function(k) {
      rejected.push({
        rec: { sabun: k, nick: unknown[k].name, tab: unknown[k].tab },
        reason: '[조직도]/[퇴직자] 명부에 없는 사번 (' + unknown[k].count + '행) — 공고 양식의 더미 사번이거나 오기입'
      });
    });
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

function sa_parseGonggoSheet(sheet, roster, out, allRows, unknown, warn, nos, tab, rejected) {
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
      if (!unknown[sabun]) unknown[sabun] = { name: person.name || person.nick || '', tab: tab, count: 0 };
      unknown[sabun].count++;
      continue;
    }

    var dateStr = sa_extractDate(g(map.발령일)) || ctx.date;
    if (!dateStr) {
      warn.push('발령일자 파싱 실패 → 제외: [' + tab + '] ' + sabun + ' ' + person.name);
      rejected.push({
        rec: { sabun: sabun, nick: emp.nick || person.nick, tab: tab, note: sa_norm(g(map.비고)) },
        reason: '발령일자를 파싱하지 못함 (원본 셀: "' + sa_summary(g(map.발령일), 40) + '")'
      });
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
    rec.tab = tab;
    allRows.push(rec);

    // ── 교차검증: 명부와 모순되는 입·퇴사 발령은 반영하지 않음 (양식 행·오기입 차단)
    if (rec.type === '퇴사' && !emp.left && String(emp.status || '').indexOf('퇴') < 0) {
      var m1 = '공고에는 퇴사 발령이나 명부상 재직 중 (' + (emp.status || '재직') + ')';
      warn.push(m1 + ' → 제외: [' + tab + '] ' + sabun + ' ' + (emp.nick || person.name) + ' ' + dateStr);
      rejected.push({ rec: rec, reason: m1 });
      continue;
    }
    if (rec.type === '입사' && emp.join && sa_toDateStr(dateStr) !== emp.join) {
      var m2 = '공고 입사일(' + dateStr + ')이 명부 입사일(' + emp.join + ')과 불일치';
      warn.push(m2 + ' → 제외: [' + tab + '] ' + sabun + ' ' + (emp.nick || person.name));
      rejected.push({ rec: rec, reason: m2 });
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
    return { hdr: [], hdrIdx: 0, col: {}, keys: {}, byDate: {}, hasHire: {}, hasLeave: {}, count: 0 };
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

  var keys = {}, byDate = {}, hasHire = {}, hasLeave = {}, count = 0, noDate = 0;
  for (var i = h.idx + 1; i < data.length; i++) {
    var r = data[i];
    var sb = col.사번 >= 0 ? String(r[col.사번] || '').trim() : '';
    if (!sb || !/^\d{6,}$/.test(sb)) continue;   // 설명행·공란 스킵
    var type   = col.발령유형    >= 0 ? sa_norm(r[col.발령유형])   : '';
    var date   = col.발령일      >= 0 ? sa_toDateStr(r[col.발령일]) : '';
    var before = col.변경전_소속 >= 0 ? sa_norm(r[col.변경전_소속]) : '';
    var after  = col.변경후_소속 >= 0 ? sa_norm(r[col.변경후_소속]) : '';

    keys[sa_apptKey(sb, date, type, after)] = true;

    // ★ 2차(느슨) 인덱스: 사번 + 발령일 → 기존 행 목록 (유형·소속 표기 차이 무시)
    if (date) {
      var dk = sa_dateKey(sb, date);
      if (!byDate[dk]) byDate[dk] = [];
      byDate[dk].push({
        row: i + 1, type: type, beforeOrg: before, afterOrg: after,
        note: col.비고 >= 0 ? String(r[col.비고] || '') : ''
      });
    } else {
      noDate++;
    }

    if (type.indexOf('입사') >= 0) hasHire[sb] = true;
    if (type.indexOf('퇴사') >= 0) hasLeave[sb] = date;
    count++;
  }
  if (noDate) {
    warn.push('[' + SA_APPT_SHEET + '] 발령일을 읽지 못한 기존 행 ' + noDate +
              '건 — 해당 행은 느슨 매칭(사번+발령일) 대상에서 빠지므로 중복 반영될 수 있습니다. 발령일 형식 확인 필요.');
  }

  return { hdr: h.hdr, hdrIdx: h.idx, col: col, keys: keys, byDate: byDate,
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

/** 2차(느슨) 키 — 사번 + 발령일만. 유형·소속 표기 차이를 무시하고 같은 발령으로 본다 */
function sa_dateKey(sabun, date) {
  return String(sabun).trim() + '|' + sa_toDateStr(date);
}

/**
 * 느슨 매칭에 걸린 기존 행들이 이미 검수 완료 처리됐는지 (비고에 #확정)
 * → 매 실행마다 같은 건이 수동확인 목록에 반복 등장하는 것을 끄는 스위치
 */
function sa_isSettled(hits) {
  return hits.some(function(h) { return String(h.note || '').indexOf('#확정') >= 0; });
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

// ================================================================
// 윌로그 조직도 월간 스냅샷 — SnapshotOrgChart.gs
// v1.0 (2026-08-03) : 매월 1일 조직도 스냅샷 아카이브 + PDF 보관
//
//  snapshotOrgChart()            … 트리거 진입점 (매월 1일 07시). [조직도_YYYY.MM.01] 탭 생성 + PDF
//  snapshotOrgChartTest()        … ★ 최초 검수용. [조직도_TEST] 탭에만 기록 (재실행 시 덮어씀)
//  snapshotOrgChartPreviewLog()  … 아무것도 쓰지 않고 집계 결과만 로그로 확인
//  createSnapshotTrigger()       … 매월 1일 07시 트리거 등록 (기존 일일 4건 불침범)
//  listAllTriggers()             … 프로젝트 트리거 전체 목록 (충돌 점검용)
//
//  ★ 데이터 원천 — 대시보드와 동일함을 '구조적으로' 보장한다
//    별도 파서를 만들지 않고 Code.gs 의 readOrgStructured(ss) 를 그대로 호출한다.
//    이 함수가 웹앱 action=read 의 emps 필드를 만드는 바로 그 함수이므로,
//    스냅샷 명단 = 대시보드 조직도 탭 명단 이 코드 수준에서 보장된다.
//    (6-A 전사 요약 스트립에서 empSal()·aiCostFor() 를 재사용한 것과 같은 원칙)
//    재직 판정도 프론트와 동일하게 재직상태 !== '퇴직' → **퇴직예정 포함**
//    (index.html empSt(e)!=='퇴직' · CLAUDE.md 3장 6번과 동일 결과)
//
//  ★ 쓰기 범위 — 원본 시트는 절대 건드리지 않는다
//    · [조직도] · [인사카드_발령] 등 raw 데이터 스프레드시트에는 **읽기만** 한다
//    · 쓰기는 전용 아카이브 스프레드시트 "윌로그 조직도 스냅샷 아카이브" 한 곳뿐
//      (없으면 최초 1회 생성 → ID를 Script Properties SNAP_ARCHIVE_SS_ID 에 저장)
//    · readOrgStructured 내부의 getOrCreate 가 시트를 생성하지 않도록
//      호출 전에 [조직도] 존재를 직접 확인한다 (없으면 중단)
//
//  ★ 멱등성
//    · 같은 이름 탭이 이미 있으면 덮어쓰지 않고 로그 경고 후 종료 (재실행 안전)
//    · PDF 도 동일 파일명이 이미 있으면 만들지 않고 경고만 (파일 삭제·덮어쓰기 없음)
//    · TEST 탭만 예외적으로 덮어쓴다 (검수 중 반복 실행 대비, 배너로 소유권 확인)
//
//  ※ 웹앱 라우트와 무관한 내부 배치 → clasp push 만으로 반영 (새 버전 배포 불필요)
//  ※ 단 PDF 저장 때문에 Drive·외부요청 권한이 추가된다 → 최초 수동 실행 시 재승인 1회 필요
// ================================================================

// ── 상수 ────────────────────────────────────────────────────────
var SNAP_PROP_KEY      = 'SNAP_ARCHIVE_SS_ID';               // Script Properties 키
var SNAP_ARCHIVE_NAME  = '윌로그 조직도 스냅샷 아카이브';      // 아카이브 스프레드시트 파일명
var SNAP_TAB_PREFIX    = '조직도_';
var SNAP_TEST_TAB      = '조직도_TEST';
var SNAP_PDF_PREFIX    = '윌로그_조직도_';
var SNAP_TZ            = 'Asia/Seoul';
var SNAP_TRIGGER_FN    = 'snapshotOrgChart';
var SNAP_TRIGGER_DAY   = 1;   // 매월 1일
var SNAP_TRIGGER_HOUR  = 7;   // 오전 7시 (일일 트리거 8·9시보다 앞 — 겹치지 않음)

var SNAP_TEST_BANNER   = '⚡ snapshotOrgChartTest() 자동 생성 검수용 탭입니다. 실행할 때마다 전체 덮어씁니다. 검수 후 삭제해도 무방합니다.';
var SNAP_README_TAB    = '_README';

// 스냅샷 컬럼 — 대시보드 조직도 렌더링 필드와 1:1
//   ※ '재직상태' 는 퇴직예정자가 명단에 포함되므로 구분용으로 함께 기록 (마지막 컬럼)
var SNAP_HDR = ['사번', '이름', '닉네임', '본부', '부서', '팀', '직책', '레벨', '입사일', '재직상태'];
var SNAP_COL_W = [95, 90, 110, 150, 150, 150, 110, 90, 100, 90];

// 정렬용 직책 서열 — SyncAppointments.gs 의 정의를 재사용 (없으면 전원 동순위)
function snap_posRank(pos) {
  var tbl = (typeof SA_POS_RANK !== 'undefined') ? SA_POS_RANK : {};
  var s = String(pos || '').trim();
  var best = 0;
  for (var k in tbl) { if (s.indexOf(k) >= 0 && tbl[k] > best) best = tbl[k]; }
  return best;
}


// ================================================================
//  진입점
// ================================================================

/** 트리거 진입점 — 매월 1일 07시. 인자(트리거 이벤트 객체)는 의도적으로 무시한다. */
function snapshotOrgChart() {
  return snap_run({ test: false, pdf: true });
}

/**
 * ★ 최초 검수용 — 실제 월 탭([조직도_2026.08.01]) 대신 [조직도_TEST] 에만 기록한다.
 * 아카이브 스프레드시트는 (없으면) 생성되지만 원본 raw 시트에는 쓰지 않는다.
 * PDF 도 함께 생성해 경로·권한까지 미리 검증한다 (파일명에 실행시각 suffix → 덮어쓰기 없음).
 */
function snapshotOrgChartTest() {
  return snap_run({ test: true, pdf: true });
}

/** 아무것도 쓰지 않고 집계 결과(재직자 수·본부별 인원)만 로그로 확인 */
function snapshotOrgChartPreviewLog() {
  return snap_run({ test: true, pdf: false, logOnly: true });
}


// ================================================================
//  본체
// ================================================================
function snap_run(opts) {
  opts = opts || {};

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('⛔ 다른 실행이 진행 중입니다. 중단.');
    return;
  }

  try {
    var now   = new Date();
    var stamp = Utilities.formatDate(now, SNAP_TZ, 'yyyy-MM-dd HH:mm');

    // ── ① 기준일·탭 이름 결정 ──────────────────────────────────
    // 기준일 = 스냅샷이 대표하는 월의 1일. 매월 1일 트리거이므로 통상 '실행 당월 1일'.
    // 과거 월 백필이 필요하면 snap_run({ test:false, pdf:true, ym:'2026.09' }) 형태로 호출.
    var ym = String(opts.ym || '').trim();
    if (ym && !/^\d{4}[.\-]\d{1,2}$/.test(ym)) {
      Logger.log('⛔ ym 형식 오류: "' + ym + '" (예: 2026.08) → 중단');
      return;
    }
    var refYear, refMonth;
    if (ym) {
      var p = ym.split(/[.\-]/);
      refYear = Number(p[0]); refMonth = Number(p[1]);
    } else {
      refYear = Number(Utilities.formatDate(now, SNAP_TZ, 'yyyy'));
      refMonth = Number(Utilities.formatDate(now, SNAP_TZ, 'MM'));
    }
    var refLabel = refYear + '.' + snap_pad(refMonth) + '.01';        // 2026.08.01
    var tabName  = opts.test ? SNAP_TEST_TAB : (SNAP_TAB_PREFIX + refLabel);

    // ── ② 원본 조직도 읽기 (읽기 전용) ─────────────────────────
    var srcSs = SpreadsheetApp.getActiveSpreadsheet();
    // readOrgStructured 내부의 getOrCreate 가 빈 시트를 만들지 않도록 사전 확인
    if (!srcSs.getSheetByName(SHEET_ORG)) {
      Logger.log('⛔ [' + SHEET_ORG + '] 시트를 찾을 수 없습니다 → 중단 (아무것도 쓰지 않음)');
      return;
    }

    var all = readOrgStructured(srcSs);   // ★ 대시보드 action=read 의 emps 와 동일한 함수
    if (!all || !all.length) {
      Logger.log('⛔ [' + SHEET_ORG + '] 에서 읽은 인원이 0명 → 중단 (아무것도 쓰지 않음)');
      return;
    }

    // 재직 판정 = 프론트와 동일 (재직상태 !== '퇴직' → 퇴직예정 포함)
    var warn   = [];
    var active = all.filter(function(e) { return snap_isActive(e, warn); });
    if (!active.length) {
      Logger.log('⛔ 재직자 0명으로 집계됨 → 중단 (재직상태 컬럼 확인 필요)');
      return;
    }

    var retiring = active.filter(function(e) { return snap_status(e) === '퇴직예정'; }).length;
    var buStat   = snap_countByBu(active);

    // 본부 > 부서 > 팀 > 직책서열(높은 순) > 입사일(오래된 순) 로 정렬
    //   ※ 입사일은 snap_normDate 로 0 패딩한 값으로 비교한다.
    //     readOrgStructured 는 원본이 문자열이면 패딩하지 않아 ("2025. 3. 2" → "2025-3-2")
    //     그대로 문자열 비교하면 3월이 8월보다 뒤로 밀린다.
    var rows = active.slice().sort(function(a, b) {
      return String(a.본부).localeCompare(String(b.본부))
          || String(a.본부부서).localeCompare(String(b.본부부서))
          || String(a.팀).localeCompare(String(b.팀))
          || (snap_posRank(b.직책) - snap_posRank(a.직책))
          || snap_normDate(a.입사일).localeCompare(snap_normDate(b.입사일));
    }).map(function(e) {
      return [
        e.사번, e.성명, e.닉네임,
        e.본부, e.본부부서, e.팀,
        e.직책, e.직급, snap_normDate(e.입사일),
        snap_status(e)
      ];
    });

    var summary = '재직자 ' + active.length + '명' +
                  (retiring ? ' (퇴직예정 ' + retiring + '명 포함)' : '') +
                  ' · ' + buStat.length + '개 본부/부서';
    var buLine = buStat.map(function(b) { return b.name + ' ' + b.count + '명'; }).join(' · ');

    // ── ③ 로그 전용 모드는 여기서 종료 ─────────────────────────
    if (opts.logOnly) {
      Logger.log('[스냅샷 프리뷰 — 쓰기 없음]\n' +
                 '기준일: ' + refLabel + '  (탭 예정 이름: ' + (SNAP_TAB_PREFIX + refLabel) + ')\n' +
                 summary + '\n본부별: ' + buLine + '\n' +
                 '전체 조직도 행 ' + all.length + '명 중 퇴직 제외 ' + (all.length - active.length) + '명' +
                 (warn.length ? '\n\n⚠ 경고 ' + warn.length + '건\n  - ' + warn.join('\n  - ') : ''));
      return { refLabel: refLabel, count: active.length, byBu: buStat, warn: warn, wrote: false };
    }

    // ── ④ 아카이브 스프레드시트 확보 ───────────────────────────
    var arc = snap_openArchive();
    if (!arc.ss) { Logger.log(arc.error); return; }
    var arcSs = arc.ss;

    // ── ⑤ 멱등성 — 같은 이름 탭이 있으면 덮어쓰지 않는다 ───────
    var sh = arcSs.getSheetByName(tabName);
    if (sh) {
      if (!opts.test) {
        Logger.log('⚠ [' + tabName + '] 탭이 이미 존재합니다 → 덮어쓰지 않고 종료 (멱등).\n' +
                   '   같은 달에 다시 스냅샷을 만들려면 기존 탭을 직접 이름 변경/삭제한 뒤 재실행하세요.\n' +
                   '   아카이브: ' + arcSs.getUrl());
        return { refLabel: refLabel, skipped: true, reason: '동일 이름 탭 존재', wrote: false };
      }
      // TEST 탭은 반복 검수를 위해 덮어쓴다 — 단 사람이 만든 동명 탭은 보호
      var a1 = String(sh.getRange(1, 1).getValue() || '');
      if (sh.getLastRow() > 0 && a1.indexOf('snapshotOrgChartTest()') < 0) {
        Logger.log('⛔ [' + tabName + '] 탭에 다른 내용이 있어 덮어쓰지 않았습니다. ' +
                   '해당 탭을 삭제하거나 이름을 바꾼 뒤 다시 실행하세요.');
        return { refLabel: refLabel, skipped: true, reason: 'TEST 탭 소유권 불일치', wrote: false };
      }
      sh.clear();
    } else {
      sh = arcSs.insertSheet(tabName, 0);
      sh.setTabColor(opts.test ? '#F59E0B' : '#2563EB');
    }

    // ── ⑥ 기록 ────────────────────────────────────────────────
    snap_writeTab(sh, {
      test:     !!opts.test,
      refLabel: refLabel,
      stamp:    stamp,
      summary:  summary,
      buLine:   buLine,
      rows:     rows
    });
    SpreadsheetApp.flush();

    var msg = ['✅ 스냅샷 기록 완료',
               '  아카이브 : ' + arcSs.getName() + (arc.created ? '  (★ 이번 실행에서 새로 생성)' : ''),
               '  URL      : ' + arcSs.getUrl(),
               '  탭       : [' + tabName + ']' + (opts.test ? '  ※ 검수용 TEST 탭' : ''),
               '  기준일   : ' + refLabel,
               '  ' + summary,
               '  본부별   : ' + buLine];

    // ── ⑦ PDF export (실패해도 시트 기록은 유지 — 로그만) ──────
    var pdf = null;
    if (opts.pdf) {
      try {
        pdf = snap_exportPdf(arcSs, sh, refLabel, opts.test);
        msg.push(pdf.skipped ? '  PDF      : ⚠ ' + pdf.message
                             : '  PDF      : ' + pdf.name + '  (' + pdf.folder + ')');
      } catch (e) {
        msg.push('  PDF      : ⚠ 실패 — ' + e.message +
                 '\n             (시트 기록은 정상 유지됨. 권한 재승인 또는 수동 내보내기로 대응)');
      }
    }

    if (warn.length) msg.push('', '⚠ 경고 ' + warn.length + '건', '  - ' + warn.join('\n  - '));
    Logger.log(msg.join('\n'));

    return {
      refLabel: refLabel, tab: tabName, count: active.length,
      byBu: buStat, warn: warn, wrote: true,
      archiveUrl: arcSs.getUrl(), pdf: pdf
    };

  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  아카이브 스프레드시트 확보 (없으면 최초 1회 생성 + ID 저장)
// ================================================================
function snap_openArchive() {
  var props = PropertiesService.getScriptProperties();
  var id    = String(props.getProperty(SNAP_PROP_KEY) || '').trim();

  if (id) {
    try {
      return { ss: SpreadsheetApp.openById(id), created: false };
    } catch (e) {
      // 휴지통 이동·권한 상실 등. 조용히 새로 만들면 과거 스냅샷이 고아가 되므로 중단한다.
      return { ss: null, error:
        '⛔ 저장된 아카이브 스프레드시트를 열 수 없습니다.\n' +
        '   ID   : ' + id + '\n' +
        '   원인 : ' + e.message + '\n' +
        '   대응 : 파일이 휴지통에 있으면 복원하세요. 정말 새로 만들려면 Script Properties 의 ' +
        SNAP_PROP_KEY + ' 값을 지운 뒤 재실행하세요 (과거 스냅샷은 기존 파일에 그대로 남습니다).' };
    }
  }

  // 최초 1회 생성 — 내 드라이브 루트에 생성된다 (원하는 폴더로 옮겨도 ID 기반이라 계속 동작)
  var ss = SpreadsheetApp.create(SNAP_ARCHIVE_NAME);
  props.setProperty(SNAP_PROP_KEY, ss.getId());

  // 기본 빈 시트를 안내 탭으로 재활용 (빈 '시트1' 방치 방지)
  try {
    var first = ss.getSheets()[0];
    first.setName(SNAP_README_TAB);
    first.getRange(1, 1).setValue(
      '윌로그 조직도 스냅샷 아카이브\n\n' +
      '· SnapshotOrgChart.gs 의 snapshotOrgChart() 가 매월 1일 07시에 [조직도_YYYY.MM.01] 탭을 자동 생성합니다.\n' +
      '· 각 탭은 그 달 1일 기준 재직자 명단(퇴직예정 포함)이며, 생성 후에는 자동으로 수정되지 않습니다.\n' +
      '· 같은 이름 탭이 있으면 덮어쓰지 않습니다. 수정이 필요하면 탭을 직접 편집하세요.\n' +
      '· 같은 폴더에 ' + SNAP_PDF_PREFIX + 'YYYY.MM.01.pdf 가 함께 저장됩니다.\n' +
      '· 이 파일의 ID 는 Apps Script 프로젝트의 Script Properties(' + SNAP_PROP_KEY + ')에 저장돼 있습니다.\n' +
      '  이 파일을 삭제하면 다음 실행에서 오류가 나므로, 옮기더라도 삭제하지 마세요.'
    ).setWrap(true).setVerticalAlignment('top');
    first.setColumnWidth(1, 720);
    first.setRowHeight(1, 160);
  } catch (e) { /* 안내 탭 실패는 무해 */ }

  return { ss: ss, created: true };
}


// ================================================================
//  탭 기록 — 상단 요약 3행 + 헤더 + 명단
// ================================================================
function snap_writeTab(sh, d) {
  var W = SNAP_HDR.length;
  var r = 1;

  if (d.test) {
    sh.getRange(r, 1, 1, W).merge()
      .setValue(SNAP_TEST_BANNER)
      .setFontSize(9).setFontStyle('italic').setFontColor('#B45309').setBackground('#FFFBEB');
    sh.setRowHeight(r, 22);
    r++;
  }

  // 제목
  sh.getRange(r, 1, 1, W).merge()
    .setValue('📌 윌로그 조직도 스냅샷 — 기준일 ' + d.refLabel)
    .setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setFontColor('#0F172A').setBackground('#E2E8F0')
    .setVerticalAlignment('middle');
  sh.setRowHeight(r, 30);
  r++;

  // 요약
  sh.getRange(r, 1, 1, W).merge()
    .setValue(d.summary + '  ·  생성 ' + d.stamp + ' (KST)')
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setFontColor('#334155')
    .setVerticalAlignment('middle');
  sh.setRowHeight(r, 22);
  r++;

  // 본부별 인원
  sh.getRange(r, 1, 1, W).merge()
    .setValue('본부별 인원 — ' + d.buLine)
    .setFontFamily('Arial').setFontSize(9).setFontColor('#475569')
    .setWrap(true).setVerticalAlignment('middle');
  r++;

  r++; // 빈 행

  // 헤더
  var hdrRow = r;
  sh.getRange(r, 1, 1, W).setValues([SNAP_HDR])
    .setFontFamily('Arial').setFontWeight('bold').setFontSize(10)
    .setFontColor('#FFFFFF').setBackground('#1E293B')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r, 26);
  r++;

  // 명단 — 사번·입사일이 숫자/날짜로 자동 변환되지 않도록 먼저 텍스트 서식 지정
  if (d.rows.length) {
    var rng = sh.getRange(r, 1, d.rows.length, W);
    sh.getRange(r, 1, d.rows.length, 1).setNumberFormat('@');   // 사번
    sh.getRange(r, 9, d.rows.length, 1).setNumberFormat('@');   // 입사일
    rng.setValues(d.rows).setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle');
    sh.getRange(r, 1, d.rows.length, 1).setHorizontalAlignment('center');
    sh.getRange(r, 8, d.rows.length, 3).setHorizontalAlignment('center'); // 레벨·입사일·재직상태
    // 퇴직예정자만 옅은 붉은 배경으로 구분
    for (var i = 0; i < d.rows.length; i++) {
      if (d.rows[i][W - 1] === '퇴직예정') {
        sh.getRange(r + i, 1, 1, W).setBackground('#FEF2F2');
      }
    }
  }

  SNAP_COL_W.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(hdrRow);
}


// ================================================================
//  PDF export → 아카이브 스프레드시트와 같은 Drive 폴더에 저장
//  실패 시 예외를 던지고 호출부가 로그로만 처리한다 (시트 기록은 유지)
// ================================================================
function snap_exportPdf(arcSs, sh, refLabel, isTest) {
  var name = isTest
    ? SNAP_PDF_PREFIX + 'TEST_' + Utilities.formatDate(new Date(), SNAP_TZ, 'yyyyMMdd-HHmm') + '.pdf'
    : SNAP_PDF_PREFIX + refLabel + '.pdf';

  // 저장 폴더 = 아카이브 스프레드시트의 부모 폴더 (없으면 내 드라이브 루트)
  var parents = DriveApp.getFileById(arcSs.getId()).getParents();
  var folder  = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  // 동일 파일명이 이미 있으면 만들지 않는다 (삭제·덮어쓰기 없음)
  var dup = folder.getFilesByName(name);
  if (dup.hasNext()) {
    return { skipped: true, name: name, folder: folder.getName(),
             message: '동일 파일명이 이미 있어 PDF 를 만들지 않았습니다: ' + name +
                      ' (' + folder.getName() + ') — 시트 기록은 정상' };
  }

  var url = 'https://docs.google.com/spreadsheets/d/' + arcSs.getId() + '/export?' + [
    'format=pdf',
    'gid=' + sh.getSheetId(),
    'portrait=false',     // 가로 — 컬럼 10개
    'size=A4',
    'fitw=true',
    'scale=4',            // 페이지 너비에 맞춤
    'top_margin=0.4', 'bottom_margin=0.4', 'left_margin=0.4', 'right_margin=0.4',
    'gridlines=false',
    'printtitle=false',
    'sheetnames=false',
    'pagenumbers=true',
    'fzr=true'            // 각 페이지에 헤더 반복
  ].join('&');

  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('PDF export HTTP ' + code + ' — ' +
                    String(resp.getContentText() || '').slice(0, 200));
  }

  var file = folder.createFile(resp.getBlob().setName(name));
  return { skipped: false, name: name, folder: folder.getName(), url: file.getUrl() };
}


// ================================================================
//  트리거 — 매월 1일 07시
//  ※ HrCard_Setup.gs 의 createDailyTrigger() 는 자기 4개 핸들러
//    (autoConvertProbation / syncAppointments / refreshHrCardEvents / refreshHeadcount)
//    만 삭제하므로 이 월간 트리거는 그 함수를 재실행해도 살아남는다.
//    반대로 이 함수도 snapshotOrgChart 핸들러만 삭제한다 → 양방향 무충돌.
// ================================================================
function createSnapshotTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === SNAP_TRIGGER_FN) { ScriptApp.deleteTrigger(t); removed++; }
  });

  ScriptApp.newTrigger(SNAP_TRIGGER_FN)
    .timeBased()
    .onMonthDay(SNAP_TRIGGER_DAY)
    .atHour(SNAP_TRIGGER_HOUR)
    .create();

  var others = ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() !== SNAP_TRIGGER_FN; })
    .map(function(t) { return t.getHandlerFunction(); });

  var msg = '✅ 조직도 스냅샷 트리거 등록 완료 — 매월 ' + SNAP_TRIGGER_DAY + '일 오전 ' +
            SNAP_TRIGGER_HOUR + '시 (' + SNAP_TRIGGER_FN + ')\n' +
            (removed ? '  기존 동일 트리거 ' + removed + '건 삭제 후 재등록\n' : '') +
            '  함께 유지되는 기존 트리거 ' + others.length + '건: ' + (others.join(', ') || '(없음)') + '\n' +
            '  ※ 시간 기반 트리거는 지정 시각부터 1시간 이내에 실행됩니다 (07~08시).';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}

/** 프로젝트 트리거 전체 목록 — 충돌·중복 점검용 */
function listAllTriggers() {
  var out = ScriptApp.getProjectTriggers().map(function(t) {
    var src = String(t.getEventType());
    return '  · ' + t.getHandlerFunction() + '  [' + src + ']';
  });
  var msg = '프로젝트 트리거 ' + out.length + '건\n' + out.join('\n');
  Logger.log(msg);
  return msg;
}


// ================================================================
//  헬퍼
// ================================================================

/** 재직상태 문자열 — 공란은 '재직' (readOrgStructured 기본값과 동일) */
function snap_status(e) {
  return String(e.재직여부 || '재직').trim();
}

/**
 * 재직 판정 — index.html empSt(e)!=='퇴직' 과 동일 (퇴직예정 포함).
 * 문서에 없는 재직상태 값이 나오면 경고를 남긴다 ('퇴사' 같은 표기 변경을 조용히 재직으로
 * 세지 않도록 — CLAUDE.md 3장 6번 규칙과 어긋나는 값 감지).
 */
function snap_isActive(e, warn) {
  var st = snap_status(e);
  if (st !== '재직' && st !== '퇴직예정' && st !== '퇴직') {
    var who = e.사번 || e.닉네임 || e.성명;
    var hit = String(st).indexOf('퇴') >= 0;
    warn.push('알 수 없는 재직상태 "' + st + '" (' + who + ') → ' +
              (hit ? '⚠ 퇴직 계열로 보이지만 프론트 기준(=="퇴직")으로는 재직으로 집계됩니다. 조직도 표기 확인 필요'
                   : '재직으로 집계'));
  }
  return st !== '퇴직';
}

/** 본부별 인원수 — 많은 순, 동수면 이름순 */
function snap_countByBu(list) {
  var m = {};
  list.forEach(function(e) {
    var k = String(e.본부 || '기타').trim() || '기타';
    m[k] = (m[k] || 0) + 1;
  });
  return Object.keys(m).map(function(k) { return { name: k, count: m[k] }; })
    .sort(function(a, b) { return (b.count - a.count) || a.name.localeCompare(b.name); });
}

function snap_pad(n) {
  return (Number(n) < 10 ? '0' : '') + Number(n);
}

/**
 * 입사일 표기 통일 — 'yyyy-MM-dd'
 * readOrgStructured 는 Date 셀만 0 패딩하고 문자열 셀은 구분자만 바꾼다
 * ("2026. 5. 17" → "2026-5-17"). 아카이브·정렬 일관성을 위해 여기서 패딩한다.
 * 형식을 못 알아보면 원본을 그대로 남긴다 (임의 가공 금지).
 */
function snap_normDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, SNAP_TZ, 'yyyy-MM-dd');
  }
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/);
  return m ? (m[1] + '-' + snap_pad(m[2]) + '-' + snap_pad(m[3])) : s;
}

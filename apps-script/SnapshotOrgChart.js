// ================================================================
// 윌로그 조직도 월간 스냅샷 — SnapshotOrgChart.gs
// v1.1 (2026-08-03) : 저장 위치를 공유 드라이브 폴더로 (Script Properties SNAPSHOT_FOLDER_ID)
// v1.0 (2026-08-03) : 매월 1일 조직도 스냅샷 아카이브 + PDF 보관
//
//  checkSnapshotFolder()         … ★ 먼저 실행. 저장 폴더 접근 진단 (읽기만)
//  setSnapshotFolderId()         … 저장 폴더 ID 를 Script Properties 에 기록 (최초 1회)
//  snapshotOrgChart()            … 트리거 진입점 (매월 1일 07시). [조직도_YYYY.MM.01] 탭 생성 + PDF
//  snapshotOrgChartTest()        … ★ 최초 검수용. [조직도_TEST] 탭에만 기록 (재실행 시 덮어씀)
//  snapshotOrgChartPreviewLog()  … 아무것도 쓰지 않고 집계 결과만 로그로 확인
//  migrateSnapshotFilesToFolder()  … 기존 파일을 지정 폴더로 이동 (dry-run) / ...Apply() 로 반영
//  createSnapshotTrigger()       … 매월 1일 07시 트리거 등록 (기존 일일 4건 불침범)
//  listAllTriggers()             … 프로젝트 트리거 전체 목록 (충돌 점검용)
//
//  ★ 저장 위치 (v1.1) — 공유 드라이브 폴더
//    · 폴더 ID 는 코드에 하드코딩하지 않고 **Script Properties SNAPSHOT_FOLDER_ID** 에서만 읽는다
//      → 폴더를 바꿀 때 코드 수정 불필요 ([프로젝트 설정 → 스크립트 속성] 에서 값만 변경)
//    · SpreadsheetApp.create() 는 항상 내 드라이브 루트에 만들므로 **생성 직후 이동**한다
//    · 이동·생성은 DriveApp 대신 **Drive 고급 서비스(Drive API v3)** 사용 —
//      DriveApp.moveTo() 는 공유 드라이브 대상에서 실패할 수 있다.
//      files.update + addParents/removeParents + supportsAllDrives:true
//      (appsscript.json 의 enabledAdvancedServices 에 Drive v3 활성화 필요)
//    · PDF 도 같은 폴더에 Drive API files.create + supportsAllDrives 로 생성
//    · **위치는 하드 전제조건** — 폴더 접근 실패, 폴더 아님, 휴지통, 권한 없음,
//      아카이브가 폴더 밖에 있음 → 전부 **로그 남기고 중단**. 내 드라이브에 조용히 만들지 않는다
//      ("잘못된 위치에 쌓이는 게 더 나쁘다")
//
//  ★ 트리거 실행 실패 알림 (v1.1) — snap_abort()
//    월 1회 작업이라 조용히 실패하면 몇 달 뒤에나 알게 되는 것이 가장 위험하다.
//    시간 기반 트리거는 e.triggerUid 를 넘기고 수동 실행은 e 가 undefined 이므로
//    이걸로 컨텍스트를 판별해 **트리거 중단 시에만 예외를 throw** → 실패 알림 메일 발송.
//    throw 안 하는 것: 같은 이름 탭 존재(정상 멱등) · PDF 중복 스킵 · PDF export 실패
//    (마지막 건은 "시트 기록 유지 + 로그만" 기존 스펙 유지 → PDF 지속 실패는 메일이 오지 않음)
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
//    · 스프레드시트 쓰기는 전용 아카이브 "윌로그 조직도 스냅샷 아카이브" 한 곳뿐
//      (없으면 최초 1회 생성 → ID를 Script Properties SNAP_ARCHIVE_SS_ID 에 저장)
//    · Drive 쓰기는 지정 폴더 안에서만 — 파일 **생성·이동**뿐이고 삭제·덮어쓰기는 없다
//    · readOrgStructured 내부의 getOrCreate 가 시트를 생성하지 않도록
//      호출 전에 [조직도] 존재를 직접 확인한다 (없으면 중단)
//
//  ★ 멱등성
//    · 같은 이름 탭이 이미 있으면 덮어쓰지 않고 로그 경고 후 종료 (재실행 안전)
//    · PDF 도 동일 파일명이 이미 있으면 만들지 않고 경고만 (파일 삭제·덮어쓰기 없음)
//    · TEST 탭만 예외적으로 덮어쓴다 (검수 중 반복 실행 대비, 배너로 소유권 확인)
//    · 이미 대상 폴더에 있는 파일은 이동하지 않는다 (migrate 재실행 안전)
//
//  ※ 웹앱 라우트와 무관한 내부 배치 → clasp push 만으로 반영 (새 버전 배포 불필요)
//  ※ 단 PDF 저장·폴더 이동 때문에 Drive·외부요청 권한이 추가되고, v1.1 에서
//    Drive 고급 서비스가 켜지므로 **최초 수동 실행 시 권한 재승인**이 필요하다
// ================================================================

// ── 상수 ────────────────────────────────────────────────────────
var SNAP_PROP_KEY      = 'SNAP_ARCHIVE_SS_ID';               // Script Properties 키 (아카이브 파일 ID)
var SNAP_FOLDER_PROP   = 'SNAPSHOT_FOLDER_ID';               // ★ Script Properties 키 (저장 폴더 ID)
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

/**
 * 트리거 진입점 — 매월 1일 07시.
 * ★ 시간 기반 트리거는 e.triggerUid 를 넘기고 편집기 수동 실행은 e 가 undefined 다.
 *   이 차이로 실행 컨텍스트를 판별해, **트리거 실행이 전제조건 실패로 중단될 때는
 *   예외를 throw** 한다 → Apps Script 실패 알림 메일이 발송된다.
 *   월 1회 작업이라 조용히 실패하면 몇 달 뒤에나 알게 되는 것이 가장 위험하므로.
 *   (수동 실행은 지금처럼 로그만 남긴다 — 편집기에서 바로 보이므로)
 */
function snapshotOrgChart(e) {
  return snap_run({ test: false, pdf: true, fromTrigger: !!(e && e.triggerUid) });
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
//  저장 폴더 설정 — Script Properties SNAPSHOT_FOLDER_ID
//
//  ★ 런타임은 폴더 ID 를 **Script Properties 에서만** 읽는다.
//    코드 상수로 두지 않으므로, 폴더를 바꿀 때는 코드를 고칠 필요 없이
//    [프로젝트 설정 → 스크립트 속성] 에서 SNAPSHOT_FOLDER_ID 값만 바꾸면 된다.
//    (아래 setSnapshotFolderId() 는 최초 1회 입력 편의용 부트스트랩일 뿐,
//     실행 경로에서는 이 함수를 호출하지 않는다)
// ================================================================

/**
 * ★ 먼저 이걸 실행하세요 — 폴더 접근 진단 (읽기만, 쓰기·이동·생성 전부 없음)
 * 폴더가 존재하는지 / 공유 드라이브인지 / 파일을 넣을 권한이 있는지 확인한다.
 */
function checkSnapshotFolder() {
  var out = [];
  var id = String(PropertiesService.getScriptProperties().getProperty(SNAP_FOLDER_PROP) || '').trim();

  out.push('── 스냅샷 저장 폴더 진단 (읽기 전용) ──');
  out.push('Script Properties [' + SNAP_FOLDER_PROP + '] = ' + (id || '(미설정)'));

  if (!id) {
    out.push('');
    out.push('⛔ 폴더 ID 가 설정되지 않았습니다. 아래 중 하나로 설정하세요.');
    out.push('   ① setSnapshotFolderId() 실행 (한 번만)');
    out.push('   ② Apps Script → 프로젝트 설정 → 스크립트 속성 →');
    out.push('      속성 ' + SNAP_FOLDER_PROP + ' / 값 <폴더 ID> 추가');
    Logger.log(out.join('\n'));
    return { ok: false, reason: 'SNAPSHOT_FOLDER_ID 미설정' };
  }

  // ── Drive 고급 서비스 확인 ──
  if (typeof Drive === 'undefined' || !Drive.Files) {
    out.push('');
    out.push('⛔ Drive 고급 서비스(Drive API v3)가 활성화되지 않았습니다.');
    out.push('   appsscript.json 의 enabledAdvancedServices 를 반영(clasp push)했는지 확인하세요.');
    Logger.log(out.join('\n'));
    return { ok: false, reason: 'Drive 고급 서비스 비활성' };
  }

  // ── 폴더 메타데이터 조회 ──
  var f;
  try {
    f = Drive.Files.get(id, {
      fields: 'id,name,mimeType,driveId,trashed,capabilities(canAddChildren,canEdit,canListChildren)',
      supportsAllDrives: true
    });
  } catch (e) {
    out.push('');
    out.push('⛔ 폴더를 조회할 수 없습니다: ' + e.message);
    out.push('   확인 사항: ① 폴더 ID 오타 ② 이 계정(' + snap_effectiveUser() + ')이 해당 공유 드라이브 멤버인지');
    out.push('             ③ Drive API 권한 재승인 여부');
    Logger.log(out.join('\n'));
    return { ok: false, reason: '폴더 조회 실패: ' + e.message };
  }

  var isFolder = f.mimeType === 'application/vnd.google-apps.folder';
  var cap = f.capabilities || {};
  out.push('');
  out.push('폴더명       : ' + f.name);
  out.push('타입         : ' + f.mimeType + (isFolder ? '  ✅ 폴더' : '  ⛔ 폴더가 아님'));
  out.push('휴지통       : ' + (f.trashed ? '⛔ 휴지통에 있음' : '아님'));
  out.push('공유 드라이브: ' + (f.driveId ? '✅ 예 (driveId=' + f.driveId + ')' : '아니오 (내 드라이브)'));
  out.push('파일 추가 권한: ' + (cap.canAddChildren ? '✅ 가능' : '⛔ 없음 — 콘텐츠 관리자 이상 권한 필요'));
  out.push('목록 조회 권한: ' + (cap.canListChildren ? '✅ 가능' : '⛔ 없음'));
  out.push('실행 계정     : ' + snap_effectiveUser());

  // ── 아카이브 파일 현재 위치 ──
  var arcId = String(PropertiesService.getScriptProperties().getProperty(SNAP_PROP_KEY) || '').trim();
  out.push('');
  out.push('아카이브 파일 [' + SNAP_PROP_KEY + '] = ' + (arcId || '(미설정 — 아직 생성 안 됨)'));
  if (arcId) {
    try {
      var a = Drive.Files.get(arcId, { fields: 'id,name,parents,trashed', supportsAllDrives: true });
      var par = (a.parents || []).join(',');
      out.push('  파일명   : ' + a.name + (a.trashed ? '  ⛔ 휴지통' : ''));
      out.push('  현재 위치: ' + (par || '(부모 없음)'));
      out.push('  대상 폴더: ' + id);
      out.push('  판정     : ' + (par === id ? '✅ 이미 대상 폴더에 있음'
                                            : '⚠ 위치 불일치 → migrateSnapshotFilesToFolder() 로 이동 필요'));
    } catch (e) {
      out.push('  ⛔ 아카이브 파일 조회 실패: ' + e.message);
    }
  }

  var ok = isFolder && !f.trashed && !!cap.canAddChildren;
  out.push('');
  out.push(ok ? '✅ 결론: 이 폴더에 스냅샷을 저장할 수 있습니다.'
              : '⛔ 결론: 아직 저장할 수 없습니다. 위 ⛔ 항목을 해결하세요.');
  Logger.log(out.join('\n'));
  return { ok: ok, folderId: id, name: f.name, sharedDrive: !!f.driveId, canAddChildren: !!cap.canAddChildren };
}

/**
 * 저장 폴더 ID 를 Script Properties 에 기록한다 — **최초 1회 부트스트랩용**.
 *
 * ★★ 이 함수는 사용 후 삭제 예정이다 (2026-08-03 Ted 결정).
 *   순서: setSnapshotFolderId() 실행 → checkSnapshotFolder() 확인
 *        → migrateSnapshotFilesToFolder()/...Apply() 로 이동 완료
 *        → **이 함수 전체 삭제 + clasp push**
 *   이유: 코드에 남은 리터럴이 나중에 실제 스크립트 속성 값과 어긋나면,
 *        문서/코드의 값을 믿고 판단하다 틀리는 사고가 난다
 *        (AI 비용 `ai_` 단위 오기가 10배 표시 버그의 원인이었던 것과 같은 종류).
 *   삭제 후 폴더 변경은 [프로젝트 설정 → 스크립트 속성] 에서 SNAPSHOT_FOLDER_ID 를
 *   직접 수정한다 — 런타임은 속성만 읽으므로 코드 수정이 필요 없다.
 */
function setSnapshotFolderId() {
  var FOLDER_ID = '1gJvL171EeBnZTz6psOdsNpOo29_D2Y9z';   // 2026-08-03 Ted 지정 (공유 드라이브 폴더)
  PropertiesService.getScriptProperties().setProperty(SNAP_FOLDER_PROP, FOLDER_ID);
  var msg = '✅ ' + SNAP_FOLDER_PROP + ' = ' + FOLDER_ID + ' 저장 완료\n' +
            '   이어서 checkSnapshotFolder() 로 접근 가능한지 확인하세요.\n' +
            '   ※ 마이그레이션까지 끝나면 이 함수(setSnapshotFolderId)는 삭제할 예정입니다.';
  Logger.log(msg);
  return msg;
}


// ================================================================
//  본체
// ================================================================
function snap_run(opts) {
  opts = opts || {};

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    // 트리거 실행이 락을 못 잡으면 그 달 스냅샷이 조용히 누락된다 → 알림 대상
    return snap_abort('⛔ 다른 실행이 진행 중입니다. 중단.', opts);
  }

  try {
    var now   = new Date();
    var stamp = Utilities.formatDate(now, SNAP_TZ, 'yyyy-MM-dd HH:mm');

    // ── ① 기준일·탭 이름 결정 ──────────────────────────────────
    // 기준일 = 스냅샷이 대표하는 월의 1일. 매월 1일 트리거이므로 통상 '실행 당월 1일'.
    // 과거 월 백필이 필요하면 snap_run({ test:false, pdf:true, ym:'2026.09' }) 형태로 호출.
    var ym = String(opts.ym || '').trim();
    if (ym && !/^\d{4}[.\-]\d{1,2}$/.test(ym)) {
      return snap_abort('⛔ ym 형식 오류: "' + ym + '" (예: 2026.08) → 중단', opts);
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
      return snap_abort('⛔ [' + SHEET_ORG + '] 시트를 찾을 수 없습니다 → 중단 (아무것도 쓰지 않음)', opts);
    }

    var all = readOrgStructured(srcSs);   // ★ 대시보드 action=read 의 emps 와 동일한 함수
    if (!all || !all.length) {
      return snap_abort('⛔ [' + SHEET_ORG + '] 에서 읽은 인원이 0명 → 중단 (아무것도 쓰지 않음)', opts);
    }

    // 재직 판정 = 프론트와 동일 (재직상태 !== '퇴직' → 퇴직예정 포함)
    var warn   = [];
    var active = all.filter(function(e) { return snap_isActive(e, warn); });
    if (!active.length) {
      return snap_abort('⛔ 재직자 0명으로 집계됨 → 중단 (재직상태 컬럼 확인 필요)', opts);
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

    // ── ④ 저장 폴더 확인 (하드 전제조건) → 아카이브 스프레드시트 확보 ──
    //     폴더에 접근할 수 없으면 내 드라이브에 조용히 만들지 않고 중단한다
    var folder = snap_resolveFolder();
    if (folder.error) { return snap_abort(folder.error, opts); }

    var arc = snap_openArchive(folder);
    if (!arc.ss) { return snap_abort(arc.error, opts); }
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
               '  저장 폴더: ' + folder.name + '  (' + folder.id + ')' +
                              (folder.shared ? '  ※ 공유 드라이브' : '  ※ 내 드라이브'),
               '  URL      : ' + arcSs.getUrl(),
               '  탭       : [' + tabName + ']' + (opts.test ? '  ※ 검수용 TEST 탭' : ''),
               '  기준일   : ' + refLabel,
               '  ' + summary,
               '  본부별   : ' + buLine];

    // ── ⑦ PDF export (실패해도 시트 기록은 유지 — 로그만) ──────
    var pdf = null;
    if (opts.pdf) {
      try {
        pdf = snap_exportPdf(arcSs, sh, refLabel, opts.test, folder);
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
//  저장 폴더 해석 — 실패하면 **절대 내 드라이브에 만들지 않고 중단**
//  ("잘못된 위치에 쌓이는 게 더 나쁘다" — 위치는 하드 전제조건)
// ================================================================
function snap_resolveFolder() {
  var id = String(PropertiesService.getScriptProperties().getProperty(SNAP_FOLDER_PROP) || '').trim();

  if (!id) {
    return { error:
      '⛔ 저장 폴더가 설정되지 않았습니다 (Script Properties ' + SNAP_FOLDER_PROP + ').\n' +
      '   setSnapshotFolderId() 를 1회 실행하거나 [프로젝트 설정 → 스크립트 속성] 에서 직접 추가하세요.\n' +
      '   → 내 드라이브에 임의로 만들지 않고 중단합니다.' };
  }
  if (typeof Drive === 'undefined' || !Drive.Files) {
    return { error:
      '⛔ Drive 고급 서비스(Drive API v3)가 비활성 상태입니다. appsscript.json 반영(clasp push) 여부를 확인하세요.\n' +
      '   → 공유 드라이브 이동이 불가능하므로 중단합니다.' };
  }

  var f;
  try {
    f = Drive.Files.get(id, {
      fields: 'id,name,mimeType,driveId,trashed,capabilities(canAddChildren)',
      supportsAllDrives: true
    });
  } catch (e) {
    return { error:
      '⛔ 저장 폴더를 조회할 수 없습니다 (' + id + '): ' + e.message + '\n' +
      '   폴더 ID·공유 드라이브 멤버십·권한 승인을 확인하세요. checkSnapshotFolder() 로 진단 가능.\n' +
      '   → 내 드라이브에 임의로 만들지 않고 중단합니다.' };
  }

  if (f.mimeType !== 'application/vnd.google-apps.folder') {
    return { error: '⛔ 지정된 ID 가 폴더가 아닙니다 (' + f.mimeType + '): ' + id + ' → 중단' };
  }
  if (f.trashed) {
    return { error: '⛔ 저장 폴더가 휴지통에 있습니다: ' + f.name + ' (' + id + ') → 중단' };
  }
  if (!f.capabilities || !f.capabilities.canAddChildren) {
    return { error:
      '⛔ 저장 폴더에 파일을 추가할 권한이 없습니다: ' + f.name + ' (' + id + ')\n' +
      '   실행 계정 ' + snap_effectiveUser() + ' 에게 콘텐츠 관리자 이상 권한이 필요합니다. → 중단' };
  }

  return { id: id, name: f.name, driveId: f.driveId || '', shared: !!f.driveId };
}

/** 파일의 현재 부모 폴더 ID 목록 */
function snap_parentsOf(fileId) {
  var f = Drive.Files.get(fileId, { fields: 'parents', supportsAllDrives: true });
  return f.parents || [];
}

/**
 * 부모 폴더 ID 목록 → 사람이 읽을 수 있는 "이름 (ID)" 문자열.
 * 마이그레이션 dry-run 에서 무관한 파일이 섞였는지 판단하려면 ID 보다 폴더명이 필요하다.
 * 조회 실패하면 ID 를 그대로 보여준다 (진단을 막지 않도록).
 */
var SNAP_FOLDER_NAME_CACHE = {};
function snap_folderLabel(parents) {
  if (!parents || !parents.length) return '(부모 없음)';
  return parents.map(function(pid) {
    if (SNAP_FOLDER_NAME_CACHE[pid]) return SNAP_FOLDER_NAME_CACHE[pid];
    var label;
    try {
      var f = Drive.Files.get(pid, { fields: 'id,name,driveId', supportsAllDrives: true });
      label = '내 드라이브 루트';
      // 루트 폴더는 이름이 'My Drive'/'내 드라이브' 로 오거나 driveId 가 없다
      if (f.name && f.name !== 'My Drive' && f.name !== '내 드라이브') {
        label = f.name + (f.driveId ? ' [공유 드라이브]' : '');
      }
      label += ' (' + pid + ')';
    } catch (e) {
      label = '(조회 실패: ' + pid + ')';
    }
    SNAP_FOLDER_NAME_CACHE[pid] = label;
    return label;
  }).join(' , ');
}

/**
 * 파일을 대상 폴더로 이동 — 공유 드라이브 대응.
 * DriveApp.moveTo() 는 공유 드라이브 대상에서 실패할 수 있어 Drive API v3 를 쓴다.
 * 내 드라이브 → 공유 드라이브 이동은 기존 부모를 removeParents 로 함께 제거해야 한다.
 */
function snap_moveToFolder(fileId, folderId) {
  var cur = snap_parentsOf(fileId);
  if (cur.length === 1 && cur[0] === folderId) return { moved: false, already: true, from: cur };

  Drive.Files.update({}, fileId, null, {
    addParents: folderId,
    removeParents: cur.join(','),
    supportsAllDrives: true,
    fields: 'id,parents'
  });

  var after = snap_parentsOf(fileId);
  if (after.indexOf(folderId) < 0) {
    throw new Error('이동 후에도 대상 폴더가 부모에 없습니다 (현재: ' + after.join(',') + ')');
  }
  return { moved: true, from: cur, to: after };
}


// ================================================================
//  아카이브 스프레드시트 확보 (없으면 최초 1회 생성 → 지정 폴더로 이동)
//  ※ 호출 전에 snap_resolveFolder() 로 폴더 접근이 확인돼 있어야 한다
// ================================================================
function snap_openArchive(folder) {
  var props = PropertiesService.getScriptProperties();
  var id    = String(props.getProperty(SNAP_PROP_KEY) || '').trim();

  if (id) {
    var ss;
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      // 휴지통 이동·권한 상실 등. 조용히 새로 만들면 과거 스냅샷이 고아가 되므로 중단한다.
      return { ss: null, error:
        '⛔ 저장된 아카이브 스프레드시트를 열 수 없습니다.\n' +
        '   ID   : ' + id + '\n' +
        '   원인 : ' + e.message + '\n' +
        '   대응 : 파일이 휴지통에 있으면 복원하세요. 정말 새로 만들려면 Script Properties 의 ' +
        SNAP_PROP_KEY + ' 값을 지운 뒤 재실행하세요 (과거 스냅샷은 기존 파일에 그대로 남습니다).' };
    }

    // ★ 위치는 하드 전제조건 — 대상 폴더 밖이면 쓰지 않고 중단한다 (자동 이동도 하지 않음)
    var par;
    try { par = snap_parentsOf(id); }
    catch (e) {
      return { ss: null, error: '⛔ 아카이브 파일의 위치를 확인할 수 없습니다 (' + id + '): ' + e.message + ' → 중단' };
    }
    if (par.indexOf(folder.id) < 0) {
      return { ss: null, error:
        '⛔ 아카이브 파일이 지정 폴더 밖에 있습니다 → 기록하지 않고 중단합니다.\n' +
        '   파일     : ' + SNAP_ARCHIVE_NAME + ' (' + id + ')\n' +
        '   현재 위치: ' + (par.join(',') || '(부모 없음)') + '\n' +
        '   대상 폴더: ' + folder.name + ' (' + folder.id + ')\n' +
        '   대응     : migrateSnapshotFilesToFolder() 로 확인 후 ...Apply() 로 이동하세요.' };
    }

    return { ss: ss, created: false };
  }

  // 최초 1회 생성 — SpreadsheetApp.create() 는 항상 내 드라이브 루트에 만들므로 즉시 이동한다
  var ss = SpreadsheetApp.create(SNAP_ARCHIVE_NAME);
  props.setProperty(SNAP_PROP_KEY, ss.getId());   // 실패해도 중복 생성되지 않도록 먼저 저장

  try {
    snap_moveToFolder(ss.getId(), folder.id);
  } catch (e) {
    return { ss: null, error:
      '⛔ 아카이브 파일을 지정 폴더로 이동하지 못했습니다 → 기록하지 않고 중단합니다.\n' +
      '   파일 : ' + SNAP_ARCHIVE_NAME + ' (' + ss.getId() + ')  ※ 지금은 내 드라이브 루트에 있습니다\n' +
      '   대상 : ' + folder.name + ' (' + folder.id + ')\n' +
      '   원인 : ' + e.message + '\n' +
      '   대응 : 권한 확인 후 migrateSnapshotFilesToFolder() → ...Apply() 로 이동하세요.\n' +
      '         (파일 ID 는 Script Properties 에 저장됐으므로 재실행해도 중복 생성되지 않습니다)' };
  }

  // 기본 빈 시트를 안내 탭으로 재활용 (빈 '시트1' 방치 방지)
  try {
    var first = ss.getSheets()[0];
    first.setName(SNAP_README_TAB);
    first.getRange(1, 1).setValue(
      '윌로그 조직도 스냅샷 아카이브\n\n' +
      '· SnapshotOrgChart.gs 의 snapshotOrgChart() 가 매월 1일 07시에 [조직도_YYYY.MM.01] 탭을 자동 생성합니다.\n' +
      '· 각 탭은 그 달 1일 기준 재직자 명단(퇴직예정 포함)이며, 생성 후에는 자동으로 수정되지 않습니다.\n' +
      '· 같은 이름 탭이 있으면 덮어쓰지 않습니다. 수정이 필요하면 탭을 직접 편집하세요.\n' +
      '· 같은 폴더(' + folder.name + ')에 ' + SNAP_PDF_PREFIX + 'YYYY.MM.01.pdf 가 함께 저장됩니다.\n' +
      '· 이 파일의 ID 는 Apps Script 프로젝트의 Script Properties(' + SNAP_PROP_KEY + ')에 저장돼 있습니다.\n' +
      '· ★ 이 파일을 다른 폴더로 옮기거나 삭제하면 다음 실행이 중단됩니다.\n' +
      '  저장 위치를 바꾸려면 스크립트 속성 ' + SNAP_FOLDER_PROP + ' 를 수정하고\n' +
      '  migrateSnapshotFilesToFolder() 로 기존 파일을 함께 옮기세요.'
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
function snap_exportPdf(arcSs, sh, refLabel, isTest, folder) {
  var name = isTest
    ? SNAP_PDF_PREFIX + 'TEST_' + Utilities.formatDate(new Date(), SNAP_TZ, 'yyyyMMdd-HHmm') + '.pdf'
    : SNAP_PDF_PREFIX + refLabel + '.pdf';

  // 저장 폴더 = SNAPSHOT_FOLDER_ID 로 지정된 폴더 (아카이브 파일과 동일 폴더)
  //   DriveApp.getFolderById().createFile() 대신 Drive API v3 를 쓴다 — 공유 드라이브 보장
  // 동일 파일명이 이미 있으면 만들지 않는다 (삭제·덮어쓰기 없음)
  var dup = Drive.Files.list({
    q: "'" + folder.id + "' in parents and name = '" + name.replace(/'/g, "\\'") + "' and trashed = false",
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    fields: 'files(id,name)',
    pageSize: 2
  });
  if (dup && dup.files && dup.files.length) {
    return { skipped: true, name: name, folder: folder.name,
             message: '동일 파일명이 이미 있어 PDF 를 만들지 않았습니다: ' + name +
                      ' (' + folder.name + ') — 시트 기록은 정상' };
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

  var file = Drive.Files.create(
    { name: name, parents: [folder.id], mimeType: 'application/pdf' },
    resp.getBlob(),
    { supportsAllDrives: true, fields: 'id,name,webViewLink' }
  );
  return { skipped: false, name: name, folder: folder.name, url: file.webViewLink || '' };
}


// ================================================================
//  [유지보수] 기존 파일을 지정 폴더로 이동
//    내 드라이브 루트에 이미 만들어진 아카이브 스프레드시트와 PDF 들을
//    SNAPSHOT_FOLDER_ID 폴더로 옮긴다. 아카이브는 **새로 만들지 않고 이동**하므로
//    Script Properties 의 파일 ID 가 그대로 유지된다.
//    파일 삭제·복사·내용 변경은 하지 않는다 (부모 폴더만 교체).
// ================================================================
function migrateSnapshotFilesToFolder()      { return snap_migrate(true); }
function migrateSnapshotFilesToFolderApply() { return snap_migrate(false); }

function snap_migrate(dryRun) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('⛔ 다른 실행이 진행 중입니다. 중단.'); return; }

  try {
    var out = ['── 스냅샷 파일 이동 ' + (dryRun ? '(dry-run — 실제 이동 없음)' : '★ 실제 반영') + ' ──'];

    var folder = snap_resolveFolder();
    if (folder.error) { Logger.log(folder.error); return; }
    out.push('대상 폴더: ' + folder.name + ' (' + folder.id + ')' + (folder.shared ? ' ※ 공유 드라이브' : ''));
    out.push('');

    var targets = [];

    // ① 아카이브 스프레드시트 (Script Properties 의 ID — 새로 만들지 않고 이동)
    var arcId = String(PropertiesService.getScriptProperties().getProperty(SNAP_PROP_KEY) || '').trim();
    if (!arcId) {
      out.push('아카이브 스프레드시트: (' + SNAP_PROP_KEY + ' 미설정 — 아직 생성 안 됨, 이동 대상 아님)');
    } else {
      try {
        var a = Drive.Files.get(arcId, { fields: 'id,name,parents,trashed,mimeType', supportsAllDrives: true });
        targets.push({ kind: '아카이브', id: a.id, name: a.name, parents: a.parents || [], trashed: a.trashed });
      } catch (e) {
        out.push('⛔ 아카이브 파일 조회 실패 (' + arcId + '): ' + e.message);
      }
    }

    // ② PDF — **내 드라이브 루트에 있는 것만** 대상 ('root' in parents)
    //    드라이브 전체 검색은 수동으로 내보낸 동명 PDF까지 끌어올 수 있어 범위를 좁혔다.
    //    루트가 아닌 곳에 있는 PDF는 의도적으로 건드리지 않는다.
    try {
      var q = "'root' in parents" +
              " and name contains '" + SNAP_PDF_PREFIX.replace(/'/g, "\\'") + "'" +
              " and mimeType = 'application/pdf' and trashed = false";
      var res = Drive.Files.list({
        q: q, supportsAllDrives: true, includeItemsFromAllDrives: true,
        fields: 'files(id,name,parents)', pageSize: 100
      });
      (res.files || []).forEach(function(f) {
        if (f.name.indexOf(SNAP_PDF_PREFIX) !== 0) return;   // contains → 접두사 일치만 채택
        targets.push({ kind: 'PDF', id: f.id, name: f.name, parents: f.parents || [], trashed: false });
      });
      out.push('PDF 검색 범위: 내 드라이브 루트, 이름이 "' + SNAP_PDF_PREFIX + '" 로 시작하는 PDF → ' +
               (res.files || []).length + '건 발견');
      out.push('');
    } catch (e) {
      out.push('⛔ PDF 검색 실패: ' + e.message);
    }

    if (!targets.length) {
      out.push('이동할 파일이 없습니다.');
      Logger.log(out.join('\n'));
      return { moved: 0, already: 0, failed: 0 };
    }

    var moved = 0, already = 0, failed = 0;
    targets.forEach(function(t) {
      var inTarget = t.parents.length === 1 && t.parents[0] === folder.id;
      var label = '[' + t.kind + '] ' + t.name;

      if (t.trashed)  { out.push('  ⚠ 건너뜀 (휴지통): ' + label); failed++; return; }
      if (inTarget)   { out.push('  · 이미 대상 폴더: ' + label); already++; return; }

      if (dryRun) {
        // 현재 위치를 폴더 '이름' 으로 함께 보여준다 — 무관한 파일이 섞였는지 사람이 판단할 수 있게
        out.push('  → 이동 예정: ' + label);
        out.push('      현재 위치: ' + snap_folderLabel(t.parents) + '  →  ' + folder.name);
        moved++;
        return;
      }
      try {
        var r = snap_moveToFolder(t.id, folder.id);
        out.push('  ✅ 이동 완료: ' + label + '  (' + (r.from.join(',') || '-') + ' → ' + folder.id + ')');
        moved++;
      } catch (e) {
        out.push('  ⛔ 이동 실패: ' + label + ' — ' + e.message);
        failed++;
      }
    });

    out.push('');
    out.push((dryRun ? '이동 예정 ' : '이동 완료 ') + moved + '건 · 이미 대상 폴더 ' + already + '건 · 실패/건너뜀 ' + failed + '건');
    if (dryRun && moved) out.push('→ 실제로 옮기려면 migrateSnapshotFilesToFolderApply() 를 실행하세요.');
    if (!dryRun && moved) out.push('→ checkSnapshotFolder() 로 위치를 재확인하세요.');
    Logger.log(out.join('\n'));
    return { moved: moved, already: already, failed: failed };

  } finally {
    lock.releaseLock();
  }
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
 * 중단 처리 — 로그를 남기고, **트리거 실행이면 예외를 throw** 한다.
 *   throw 하면 Apps Script 가 실패 알림 메일을 보내므로 조용한 실패를 막을 수 있다.
 *   (월 1회 작업이라 몇 달 뒤에 발견되는 것이 가장 위험)
 *   수동 실행(fromTrigger 아님)은 편집기에서 로그가 바로 보이므로 throw 하지 않는다.
 * ※ '같은 이름 탭 존재'(정상 멱등)·'PDF 중복 스킵'·'PDF export 실패' 는 중단이 아니므로
 *   이 함수를 거치지 않는다 — 기존 스펙대로 로그 경고만 남는다.
 */
function snap_abort(msg, opts) {
  Logger.log(msg);
  if (opts && opts.fromTrigger) {
    // 메일 제목·본문에서 읽기 쉽도록 1줄로 압축 (전문은 위 로그에 남아 있다)
    throw new Error('[조직도 스냅샷 중단] ' + String(msg).replace(/\s+/g, ' ').trim().slice(0, 400));
  }
  return { ok: false, aborted: true, wrote: false, reason: String(msg).split('\n')[0] };
}

/** 실행 계정 — 권한 문제 진단용 (조회 실패 시 빈 문자열 대신 안내 문구) */
function snap_effectiveUser() {
  try { return Session.getEffectiveUser().getEmail() || '(확인 불가)'; }
  catch (e) { return '(확인 불가)'; }
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

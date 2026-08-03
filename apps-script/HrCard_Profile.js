// ═══════════════════════════════════════════════════════
//  HrCard_Profile.gs  — 인사카드 프로필 데이터 조회
//  수정 이력:
//    2026-06 : joinDate 정규화 (Date 객체 → "YYYY.MM.DD" 문자열)
//              고용형태(empType) 필드 추가 (인사카드_기본추가정보 시트)
//              최근 평가 2개 종합등급 evals에 포함 (기존 동일)
//    2026-08-03 : 발령 이력 정렬 추가 (발령일 내림차순 + 유형 서열)
//                 — hp_sortApptHistory(). 시트는 재정렬하지 않는다.
// ═══════════════════════════════════════════════════════
 
/**
 * 날짜 값을 "YYYY.MM.DD" 형식으로 정규화
 * - Google Sheets에서 날짜 셀이 Apps Script로 전달될 때
 *   Date 객체 또는 직렬화된 문자열로 올 수 있음
 */
function normalizeDateStr(val) {
  if (!val) return '';
  // 이미 "YYYY.MM.DD" 형식이면 그대로
  if (typeof val === 'string' && /^\d{4}\.\d{2}\.\d{2}$/.test(val.trim())) return val.trim();
  // Date 객체 또는 파싱 가능한 값
  var d;
  if (val instanceof Date) {
    d = val;
  } else {
    // 숫자형(Excel 시리얼) 처리는 Sheets에서 getDisplayValue() 사용 권장
    d = new Date(val);
  }
  if (isNaN(d.getTime())) return String(val);
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + '.' + m + '.' + day;
}
 
/**
 * 메인 함수: 사번 기준으로 인사카드 9개 시트 조회
 */
function getProfileData(empId) {
  var ss = SpreadsheetApp.openById('196QxIjHyD2KhQ7Q-hvCrtRAKRKZc6UZypGcE9EtyNwg');
  empId = String(empId).trim();
 
  // ── 1. 기본추가정보 ──────────────────────────────────
  var sheetBasic = ss.getSheetByName('인사카드_기본추가정보');
  var basicData = {};
  if (sheetBasic) {
    var bVals = sheetBasic.getDataRange().getValues();
    var bDisp = sheetBasic.getDataRange().getDisplayValues(); // 날짜 표시값
    var bHead = bVals[0].map(function(h){ return String(h).trim(); });
    for (var i = 1; i < bVals.length; i++) {
      if (String(bVals[i][0]).trim() !== empId) continue;
      var row = {};
      bHead.forEach(function(h, hi) {
        // 날짜 컬럼은 DisplayValue(표시값) 우선 사용
        var raw  = bVals[i][hi];
        var disp = bDisp[i][hi];
        if (raw instanceof Date) {
          row[h] = normalizeDateStr(raw);
        } else {
          row[h] = disp || raw || '';
        }
      });
      basicData = row;
      break;
    }
  }
 
  // ── 2. 조직도 시트에서 기본 인사 정보 보완 ──────────
  var sheetOrg = ss.getSheetByName('조직도');
  var orgRow = {};
  if (sheetOrg) {
    var oVals = sheetOrg.getDataRange().getValues();
    var oDisp = sheetOrg.getDataRange().getDisplayValues();
    var oHead = oVals[0].map(function(h){ return String(h).trim(); });
    for (var oi = 1; oi < oVals.length; oi++) {
      if (String(oVals[oi][0]).trim() !== empId) continue;
      oHead.forEach(function(h, hi) {
        var raw  = oVals[oi][hi];
        var disp = oDisp[oi][hi];
        orgRow[h] = (raw instanceof Date) ? normalizeDateStr(raw) : (disp || raw || '');
      });
      break;
    }
  }
 
  // basic 객체 조합
  // 조직도 컬럼명: 사번, 한국이름, 닉네임(영문), 본부/부서, 부서, 팀, 직책, 레벨, 재직상태, 월고정총액(만원), 입사일, 연봉총액(만원)
  var nick      = basicData['닉네임(영문)'] || basicData['닉네임'] || orgRow['닉네임(영문)'] || orgRow['닉네임'] || '';
  var korName   = basicData['한국이름']     || basicData['성명']   || orgRow['한국이름']     || orgRow['성명']   || '';
  var status    = basicData['재직여부']     || orgRow['재직상태']  || orgRow['재직여부']     || '재직';
  var dept      = basicData['본부']         || orgRow['본부/부서'] || orgRow['본부']         || '';
  var subDept   = basicData['본부부서']     || orgRow['부서']      || orgRow['본부부서']     || '';
  var team      = basicData['팀']           || orgRow['팀']        || '';
  var position  = basicData['직책']         || orgRow['직책']      || '';
  // Job-Level: 조직도의 '레벨' 컬럼에서 읽음
  var level     = basicData['직급']         || orgRow['레벨']      || orgRow['직급']         || '';
  // 연봉/월고정: 조직도의 실제 컬럼명 반영
  var annualSal = basicData['연봉총액_만원']|| orgRow['연봉총액(만원)'] || orgRow['연봉총액만원'] || '';
  var monthly   = basicData['월고정_만원']  || orgRow['월고정총액(만원)'] || orgRow['월고정만'] || '';
 
  // ★ 입사일
  var joinDateRaw = basicData['입사일'] || orgRow['입사일'] || '';
  var joinDate    = normalizeDateStr(joinDateRaw);
 
  // ★ 고용형태
  var empType = String(basicData['고용형태'] || '').trim();
 
  var basic = {
    nick        : nick,
    korName     : korName,
    status      : status,
    dept        : dept,
    subDept     : subDept,
    team        : team,
    position    : position,
    level       : level,
    annualSal   : annualSal,
    monthlySal  : monthly,
    joinDate    : joinDate,
    empType     : empType,
    jobDesc     : String(basicData['직무'] || orgRow['직무'] || ''),
    birthDate   : normalizeDateStr(basicData['생년월일'] || ''),
    gender      : String(basicData['성별'] || ''),
    nationality : String(basicData['국적'] || ''),
    contact     : String(basicData['연락처'] || ''),
    hireRoute   : String(basicData['입사경로'] || ''),
    // ★ 입사 전 경력: 시트 컬럼명은 '입사전경력_연' (년 아님)
    prevCareerY : String(basicData['입사전경력_연'] || basicData['입사전경력_년'] || '0'),
    prevCareerM : String(basicData['입사전경력_월'] || '0'),
    school      : String(basicData['최종학력_학교'] || ''),
    major       : String(basicData['최종학력_전공'] || ''),
    marriageStatus : String(basicData['결혼여부'] || ''),
    weddingDate : normalizeDateStr(basicData['결혼식일자'] || basicData['결혼기념일'] || ''),
    visaType    : String(basicData['비자유형'] || orgRow['비자유형'] || ''),
    visaExpire  : normalizeDateStr(basicData['비자만료일'] || orgRow['비자만료일'] || ''),
    stockOption : String(basicData['스톡옵션'] || ''),
    photoUrl    : String(basicData['프로필사진_URL'] || ''),
    photoEmoji  : String(basicData['프로필이모지'] || ''),
    probEndDate : normalizeDateStr(basicData['수습종료일'] || ''),
    memo        : String(basicData['메모'] || ''),
    asanaUrl    : String(basicData['Asana_URL'] || orgRow['asana (링크 삽입)'] || orgRow['asanaUrl'] || ''),
    profileUrl  : String(basicData['Profile_URL'] || orgRow['인사노트'] || orgRow['profileUrl'] || ''),
  };
 
  // ── 3. 가족사항 ──────────────────────────────────────
  var family = readSheet(ss, '인사카드_가족', empId, ['입사일']);
 
  // ── 4. 면접 기록 ────────────────────────────────────
  var interview = readSheet(ss, '인사카드_면접', empId, ['면접일']);
 
  // ── 5. 수습 면담 ────────────────────────────────────
  var probInt = readSheet(ss, '인사카드_수습면담', empId, ['면담일']);
 
  // ── 6. 수습 평가 ────────────────────────────────────
  var probEval = readSheet(ss, '인사카드_수습평가', empId, ['평가일']);
 
  // ── 7. 발령 이력 ────────────────────────────────────
  // 시트 행 순서는 사람마다 다르다 (append 전용 원장) → 여기서 최신순으로 정렬한다.
  // 시트 자체는 재정렬하지 않는다 (syncAppointments 멱등성 보호)
  var history = hp_sortApptHistory(readSheet(ss, '인사카드_발령', empId, ['발령일']), empId);
 
  // ── 8. 연봉 이력 ────────────────────────────────────
  var salary = readSheet(ss, '인사카드_연봉', empId, ['계약일','계약만료일','특별보수지급일']);
  // 최신순 정렬 (계약일 기준 내림차순)
  salary.sort(function(a, b){
    return (b['계약일'] || '').localeCompare(a['계약일'] || '');
  });
 
  // ── 9. 평가 이력 ────────────────────────────────────
  var evals = readSheet(ss, '인사카드_평가이력', empId, ['평가기간']);
  // 최신순 정렬 (평가기간 기준 내림차순)
  evals.sort(function(a, b){
    return (b['평가기간'] || '').localeCompare(a['평가기간'] || '');
  });
 
  // ── 10. 이벤트 알림 ─────────────────────────────────
  var events = readSheet(ss, '인사카드_이벤트알림', empId, []);
 
  return {
    empId     : empId,
    basic     : basic,
    family    : family,
    interview : interview,
    probInt   : probInt,
    probEval  : probEval,
    history   : history,
    salary    : salary,
    evals     : evals,
    events    : events,
  };
}
 
/**
 * 공통 시트 읽기 헬퍼
 * @param {Spreadsheet} ss
 * @param {string} sheetName
 * @param {string} empId
 * @param {string[]} dateCols - 날짜 정규화할 컬럼명 목록
 */
function readSheet(ss, sheetName, empId, dateCols) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var vals = sheet.getDataRange().getValues();
  var disp = sheet.getDataRange().getDisplayValues();
  if (vals.length < 2) return [];
  var headers = vals[0].map(function(h){ return String(h).trim(); });
  var dateSet = {};
  (dateCols || []).forEach(function(c){ dateSet[c] = true; });
  var result = [];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== empId) continue;
    var row = {};
    headers.forEach(function(h, hi) {
      var raw = vals[i][hi];
      if (raw instanceof Date || dateSet[h]) {
        // 날짜 컬럼은 DisplayValue 사용
        row[h] = normalizeDateStr(raw);
      } else {
        // 텍스트 컬럼은 raw 값 사용 (줄바꿈 보존)
        row[h] = raw !== undefined && raw !== null ? String(raw) : '';
      }
    });
    result.push(row);
  }
  return result;
}


// ═══════════════════════════════════════════════════════
//  발령 이력 정렬 (2026-08-03 신규)
//
//  [인사카드_발령] 은 append 전용 원장이라 행 순서가 사람마다 다르다.
//    · 정본 77행 = 수기 입력분이라 대체로 시간순
//    · syncAppointments 백필분 = 공고 탭 순서. 배치 내에서만 발령일 오름차순으로
//      정렬해 append하므로(SyncAppointments.gs), 회차가 갈리거나 입사/퇴사 발령이
//      나중 회차에 붙으면 개인 단위 순서가 깨진다
//  → 시트는 절대 재정렬하지 않고(syncAppointments 의 행 기반 멱등성 보호),
//    응답을 조립할 때만 정렬한다. 정렬은 여기 한 곳에서만 하고
//    프론트(index.html)는 서버가 준 순서를 그대로 렌더한다.
// ═══════════════════════════════════════════════════════

/**
 * 같은 발령일에 복수 발령이 있을 때(겸직 등)의 2차 정렬 서열.
 * 발령일과 함께 내림차순 정렬하므로 숫자가 큰 쪽이 화면 위로 온다
 * (요구 순서: 입사 < 부서이동 < 직책승격 < 겸직 < 겸직해제 < 레벨업 < 퇴사
 *  → 화면에서는 퇴사가 위, 입사가 아래).
 */
var HP_APPT_RANK = {
  '입사'    : 10,
  '부서이동': 20,
  '직책승격': 30,
  '겸직'    : 40,
  '겸직해제': 50,
  '레벨업'  : 60,
  '퇴사'    : 70
};

/**
 * 위 7종 외에 실제로 쓰이는 발령유형 어휘 (SyncAppointments.gs 참조).
 * 서열이 명시된 7종 사이에 끼워 넣은 값이라 조정해도 무방하다.
 * 여기에도 없는 유형(판별 실패분 '인사발령' 등)은 0 → 같은 날짜 그룹의 맨 아래.
 */
var HP_APPT_RANK_EXT = {
  '복직'      : 12,
  '정규직전환': 15,
  '직무변경'  : 25,
  '직책변경'  : 30,  // 직책승격과 동급 (승격/강등 여부는 서열로 구분하지 않음)
  '휴직'      : 55,
  '징계'      : 65
};

/**
 * 발령일 → 정렬키 "YYYYMMDD". 못 읽으면 '' 을 돌려준다.
 *
 * readSheet() 가 normalizeDateStr() 로 "YYYY.MM.DD" 를 만들어 주지만,
 * V8 Date 가 못 읽는 값("2026년 5월 17일", "미정" 등)은 원문 그대로 통과시킨다.
 * 문자열 비교만 하면 그런 값이 조용히 섞이므로 여기서 다시 파싱한다.
 * new Date() 로 재파싱하지 않는 이유: 로케일·타임존에 따라 결과가 흔들린다.
 */
function hp_apptDateKey(val) {
  if (val === null || val === undefined) return '';

  // Date 객체 (getValues() 는 날짜 셀을 Date 로 넘긴다).
  // instanceof 만 보면 다른 실행 컨텍스트에서 넘어온 Date 를 놓치므로 getTime 유무도 함께 본다
  if (val instanceof Date || (typeof val === 'object' && val && typeof val.getTime === 'function')) {
    if (isNaN(val.getTime())) return '';
    return String(val.getFullYear())
         + String(val.getMonth() + 1).padStart(2, '0')
         + String(val.getDate()).padStart(2, '0');
  }

  var s = String(val).trim();
  if (!s) return '';

  // "2026.05.17" / "2026. 5. 17" / "2026-5-17" / "2026/05/17" / "2026년 5월 17일"
  // + 뒤에 부연이 붙는 경우("2025. 8. 22 (Jake 퇴사일)") 앞쪽 날짜만 사용
  var m = s.match(/(\d{4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/);
  if (!m) return '';

  var mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
  if (!(mo >= 1 && mo <= 12) || !(d >= 1 && d <= 31)) return '';

  return m[1] + String(mo).padStart(2, '0') + String(d).padStart(2, '0');
}

/**
 * 발령유형 → 서열. 조합형("부서이동/직책승격")은 구성요소 중 가장 높은 서열을 쓴다.
 * 토큰 단위 정확일치로 판정한다 — 부분일치를 쓰면 '겸직해제' 가 '겸직' 으로 잡힌다.
 */
function hp_apptTypeRank(type) {
  var s = String(type == null ? '' : type);
  if (!s) return 0;
  var best = 0;
  s.split(/[\/,·+&]/).forEach(function(tok) {
    var k = tok.replace(/\s+/g, '');
    if (!k) return;
    var r = HP_APPT_RANK[k];
    if (r === undefined) r = HP_APPT_RANK_EXT[k];
    if (r !== undefined && r > best) best = r;
  });
  return best;
}

/**
 * 발령 이력을 화면 표시 순서로 정렬해 새 배열로 돌려준다 (입력 배열 불변).
 *   1순위 발령일 내림차순 (최신이 맨 위)
 *   2순위 발령유형 서열 내림차순 (같은 날 복수 발령)
 *   3순위 원래 시트 순서 (안정 정렬 — 동순위끼리 순서가 흔들리지 않게)
 * 발령일이 비었거나 파싱 불가한 행은 맨 아래에 시트 순서대로 붙이고 로그 경고를 남긴다.
 */
function hp_sortApptHistory(rows, empId) {
  if (!rows || !rows.length) return rows || [];

  var dated = [], undated = [];
  rows.forEach(function(r, i) {
    var key = hp_apptDateKey(r['발령일']);
    var item = { row: r, idx: i, key: key, rank: hp_apptTypeRank(r['발령유형']) };
    if (key) dated.push(item); else undated.push(item);
  });

  dated.sort(function(a, b) {
    if (a.key !== b.key)   return a.key < b.key ? 1 : -1;  // 발령일 내림차순
    if (a.rank !== b.rank) return b.rank - a.rank;          // 유형 서열 내림차순
    return a.idx - b.idx;                                   // 시트 순서 유지
  });

  if (undated.length) {
    try {
      Logger.log('⚠ [인사카드_발령] 사번 ' + empId + ' — 발령일 없음/파싱불가 '
        + undated.length + '건 → 목록 맨 아래 배치: '
        + undated.map(function(x) {
            return '(' + (x.row['발령유형'] || '유형없음') + ' / 발령일="'
                 + String(x.row['발령일'] == null ? '' : x.row['발령일']) + '")';
          }).join(', '));
    } catch (e) {}  // Logger 없는 컨텍스트 방어
  }

  return dated.concat(undated).map(function(x) { return x.row; });
}
 
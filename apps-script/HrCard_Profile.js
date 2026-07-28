// ═══════════════════════════════════════════════════════
//  HrCard_Profile.gs  — 인사카드 프로필 데이터 조회
//  수정 이력:
//    2026-06 : joinDate 정규화 (Date 객체 → "YYYY.MM.DD" 문자열)
//              고용형태(empType) 필드 추가 (인사카드_기본추가정보 시트)
//              최근 평가 2개 종합등급 evals에 포함 (기존 동일)
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
  var history = readSheet(ss, '인사카드_발령', empId, ['발령일']);
 
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
 
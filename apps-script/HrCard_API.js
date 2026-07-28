// ================================================================
// 윌로그 인사카드 API — HrCard_API.gs
// Code.gs 의 doGet() 에 'getHrCard' 액션을 추가하는 파일
//
// 실행 방법:
//   1. Apps Script 에서 새 파일(+) → 스크립트 → 이름: HrCard_API
//   2. 이 코드 전체 붙여넣기 → 저장(Ctrl+S)
//   3. Code.gs 의 doGet() 함수 안에 아래 한 줄 추가:
//
//      if (action === 'getHrCard') {
//        return jsonOut(getHrCardData(ss, e.parameter.sabun));
//      }
//
//      위치: "if (action === 'read')" 블록 바로 위에 추가
// ================================================================
 
 
// ── 메인 조회 함수 ──────────────────────────────────────────────
function getHrCardData(ss, sabun) {
  if (!sabun) return { error: '사번이 없습니다.' };
 
  sabun = String(sabun).trim();
 
  try {
    return {
      basic:      getBasicInfo(ss, sabun),
      family:     getSheetRows(ss, '인사카드_가족',           sabun),
      interview:  getSheetRows(ss, '인사카드_면접',           sabun),
      probation:  getSheetRows(ss, '인사카드_수습면담',       sabun),
      probEval:   getSheetRows(ss, '인사카드_수습평가',       sabun),
      history:    getSheetRows(ss, '인사카드_발령',           sabun),
      salary:     getSheetRows(ss, '인사카드_연봉',           sabun),
      evals:      getSheetRows(ss, '인사카드_평가이력',       sabun),
      events:     getSheetRows(ss, '인사카드_이벤트알림',     sabun),
    };
  } catch(e) {
    return { error: e.message };
  }
}
 
 
// ── 기본 정보: 조직도 + 기본추가정보 합치기 ──────────────────────
function getBasicInfo(ss, sabun) {
  // 1) 조직도 시트에서 해당 사번 행 찾기
  var orgSheet = ss.getSheetByName('조직도');
  if (!orgSheet) return {};
 
  var orgData   = orgSheet.getDataRange().getValues();
  var orgHeader = orgData[0]; // 조직도는 1행이 헤더
  var orgRow    = null;
 
  for (var i = 1; i < orgData.length; i++) {
    if (String(orgData[i][0]).trim() === sabun) {
      orgRow = orgData[i];
      break;
    }
  }
  if (!orgRow) return { error: '조직도에서 해당 사번을 찾을 수 없습니다.' };
 
  // 조직도 컬럼 매핑
  function orgVal(colName) {
    var idx = orgHeader.indexOf(colName);
    return idx >= 0 ? String(orgRow[idx] || '').trim() : '';
  }
 
  // 2) 기본추가정보 시트
  var extraSheet = ss.getSheetByName('인사카드_기본추가정보');
  var extraRow   = null;
  if (extraSheet) {
    var extraData   = extraSheet.getDataRange().getValues();
    var extraHeader = extraData[0];
    for (var i = 2; i < extraData.length; i++) { // 2행(index 1)은 설명행 → 3행부터
      if (String(extraData[i][0]).trim() === sabun) {
        extraRow = extraData[i];
        break;
      }
    }
    function extraVal(colName) {
      if (!extraRow) return '';
      var idx = extraHeader.indexOf(colName);
      return idx >= 0 ? String(extraRow[idx] || '').trim() : '';
    }
  } else {
    function extraVal() { return ''; }
  }
 
  // 입사일 정규화 (조직도는 "2024. 3. 6" 형식)
  function normalizeDate(raw) {
    if (!raw) return '';
    return String(raw).replace(/\.\s*/g, '.').replace(/\s/g, '').trim();
  }
 
  return {
    sabun:       sabun,
    korName:     orgVal('한국이름'),
    nick:        orgVal('닉네임(영문)'),
    dept:        orgVal('본부/부서'),
    subDept:     orgVal('부서'),
    team:        orgVal('팀'),
    position:    orgVal('직책'),
    level:       orgVal('레벨'),
    status:      orgVal('재직상태'),
    joinDate:    normalizeDate(orgVal('입사일')),
    monthlySal:  orgVal('월고정총액(만원)'),
    annualSal:   orgVal('연봉총액(만원)'),
    asanaUrl:    orgVal('asana (링크 삽입)'),
    hrNote:      orgVal('인사노트'),
    // 기본추가정보
    photoUrl:    extraVal('프로필사진_URL'),
    emoji:       extraVal('프로필이모지'),
    birthDate:   extraVal('생년월일'),
    gender:      extraVal('성별'),
    marriageStatus: extraVal('결혼여부'),
    weddingDate: extraVal('결혼식일자'),
    school:      extraVal('최종학력_학교'),
    major:       extraVal('최종학력_전공'),
    nationality: extraVal('국적'),
    visaType:    extraVal('비자유형'),
    visaExpire:  extraVal('비자만료일'),
    contact:     extraVal('연락처'),
    jobDesc:     extraVal('직무'),
    stockOption: extraVal('스톡옵션'),
    hireRoute:   extraVal('입사경로'),
    prevCareerY: extraVal('입사전경력_연'),
    prevCareerM: extraVal('입사전경력_월'),
    memo:        extraVal('메모'),
  };
}
 
 
// ── 범용 시트 행 조회: 사번으로 해당 행들 전부 반환 ───────────────
function getSheetRows(ss, sheetName, sabun) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
 
  var data   = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
 
  var header = data[0]; // 1행이 헤더
  var result = [];
 
  // 설명행(2행, index 1) 건너뜀 → 3행(index 2)부터 실제 데이터
  var startRow = (data[1] && String(data[1][0]).trim() === '') ? 2 : 1;
  // 2행 첫 셀이 비어있으면 설명행으로 간주하고 건너뜀
 
  for (var i = startRow; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || String(row[0]).trim() === '') continue;
    if (String(row[0]).trim() !== sabun) continue;
 
    // 헤더 키: 값 객체로 변환
    var obj = {};
    for (var j = 0; j < header.length; j++) {
      var key = String(header[j]).trim();
      if (!key) continue;
      var val = row[j];
      // Date 객체는 문자열로 변환
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Seoul', 'yyyy.MM.dd');
      }
      obj[key] = val !== undefined && val !== null ? String(val).trim() : '';
    }
    result.push(obj);
  }
 
  return result;
}
 
 
// ================================================================
// Code.gs 에 추가할 한 줄 (복사해서 Code.gs doGet 함수 안에 붙여넣기)
// ================================================================
//
//  if (action === 'getHrCard') {
//    return jsonOut(getHrCardData(ss, e.parameter.sabun));
//  }
//
// 위치: Code.gs의 doGet(e) 함수 안,
//       "if (action === 'read')" 블록 바로 위에 추가
// ================================================================
 
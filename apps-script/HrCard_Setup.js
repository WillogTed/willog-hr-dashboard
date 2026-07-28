// ================================================================
// 윌로그 인사카드 — 신규 시트 자동 생성 + 이벤트 알림 자동 갱신
// v2.5 (2026-07): 입사 N주년 이벤트 (3/5/7/10주년 🏆 리프레시 강조) +
//   수습→정규직 자동 전환 autoConvertProbation() (종료일 경과+수습평가 '통과' 시에만)
//   createDailyTrigger: 트리거 3건 (오전 8시 전환 → 오전 9시 알림·집계)
// v2.4.1 (2026-07): [인건비] 헤더 탐지 정확일치로 수정 (1행 배너 오인 방지),
//   자발/비자발 분류를 대시보드 퇴직자 KPI와 동일 규칙으로 + 미분류 경고 로그
// v2.4 (2026-07): 월별 입퇴사·재직자 수 자동 집계 — refreshHeadcount()
//   - [인건비] 월초/월말/입사/자발/비자발 자동 계산 (당월+전월만, 그 이전 월 불가침)
//   - createDailyTrigger가 이벤트 알림 + 집계 트리거 2건을 함께 등록
// v2.3 (2026-07): 퇴사예정 이벤트에 '최종출근일' 반영
//   - [조직도] '최종출근일' 컬럼 신설 (연차 소진 케이스 대응)
//   - D-day 기준 = 최종출근일 우선, 배지에 오피셜 퇴사일 병기
//   - 목록 유지: 오피셜 퇴사일 경과 7일까지
// v2.2 (2026-07): 이벤트 유형 재정비
//   - 근로계약종료 신설: [인사카드_기본추가정보] '근로계약종료일' (계약직, 60일 전~경과 7일)
//   - 퇴사예정 신설: [조직도] 재직상태='퇴직예정' + '퇴사예정일' 컬럼 (구 asana 열 재활용)
//   - 수습종료임박: 고용형태='수습' 인원만 생성 (정규직 전환자 제외)
//   - 퇴직예정자도 이벤트 생성 대상에 포함 (재직 중이므로)
// v2.1 (2026-07): D-day 표기 개선 + 출산예정 알림 신설
//   - 당일 이벤트는 "D-0" 대신 "D-day"로 표기 (전 유형 공통)
//   - 출산예정 알림 신설: [인사카드_가족] '출산예정일' 컬럼 (맨 끝 추가)
//     → 60일 이내 표시, 예정일 경과 14일까지 D+N 표시 (👶)
// v2 (2026-07): refreshHrCardEvents() 전면 재작성
//   - 조직도 헤더 1~3행 자동 탐지 + 키워드 매칭 (Code.gs와 동일 방식)
//   - 날짜 셀이 Date 객체여도 정상 파싱 (기존: "YYYY.MM.DD" 텍스트만 지원)
//   - 가족 시트 결혼기념일 컬럼을 헤더 이름으로 탐지 (기존: 고정 인덱스)
//   - 수습종료: 기본추가정보 '수습종료일' 우선, 없으면 입사일+3개월
//   - 알림 0건일 때 "알림 없음 + 갱신일시" 상태 행 기록 (미실행과 구분)
//   - 재직 판정: 재직상태에 '퇴' 포함 시 제외 (공란은 재직 간주)
// 실행 방법:
//   1. 이 파일 전체를 HrCard_Setup.gs 에 교체 붙여넣기 → 저장
//   2. 함수 선택: refreshHrCardEvents → 실행 (1회 수동 확인)
//   3. 함수 선택: createDailyTrigger → 실행 (매일 오전 9시 자동 갱신 등록)
//   ※ 웹앱 배포와 무관한 내부 함수라 "새 버전 배포"는 불필요
// ================================================================

function setupHrCardSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  var SHEETS = [
    {
      name: "인사카드_기본추가정보",
      color: "#4A90D9",
      cols: [
        ["사번", "조직도 시트 사번과 동일", 12],
        ["닉네임(영문)", "", 16],
        ["프로필사진_URL", "구글드라이브 썸네일 URL (가이드 참조)", 44],
        ["프로필이모지", "사진 없을 때 대체 이모지 (예: 👨)", 14],
        ["생년월일", "YYYY.MM.DD", 14],
        ["성별", "남/여", 8],
        ["결혼여부", "미혼/기혼/기타", 14],
        ["결혼식일자", "YYYY.MM.DD — 기혼자만 입력", 16],
        ["최종학력_학교", "", 22],
        ["최종학력_전공", "", 20],
        ["국적", "", 12],
        ["비자유형", "외국인만 입력 (예: E-7)", 18],
        ["비자만료일", "YYYY.MM.DD", 16],
        ["연락처", "", 18],
        ["직무", "", 18],
        ["스톡옵션", "미부여 / 부여 (N주)", 18],
        ["입사경로", "공개채용/헤드헌팅/내부추천 등", 18],
        ["입사전경력_연", "숫자", 12],
        ["입사전경력_월", "숫자", 12],
        ["메모", "자유 입력", 30],
      ],
      sample: ["240010008","Ted","","👨","1993.03.05","남","기혼","2026.05.17",
               "고려대학교","경영학과","대한민국","","","010-9062-7896",
               "인사, 인사기획","미부여","내부추천 (Peter)",3,9,""]
    },
    {
      name: "인사카드_가족",
      color: "#E06C75",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["관계", "배우자/부/모/장녀/장남 등", 14],
        ["이름", "이니셜 가능 (예: 박OO)", 14],
        ["생년월일", "YYYY.MM.DD — 생일 알림 자동 계산", 16],
        ["결혼기념일", "YYYY.MM.DD — 기혼자 배우자에만 입력", 16],
        ["거주지", "", 18],
        ["학교_직장", "미취학/유치원/초등 등", 18],
        ["건강특이사항", "알러지·지병 등 CEO 챙김 시 참고", 22],
        ["이모지", "카드 표시용 (예: 💍👧)", 10],
        ["메모", "CEO가 챙겨줄 내용 등", 30],
      ],
      sample: ["240010008","Ted","배우자","박OO","1995.02.18","2026.05.17","서울","","","💍","신혼"]
    },
    {
      name: "인사카드_면접",
      color: "#56B6C2",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["면접단계", "서류전형/1차면접/기술면접/최종면접", 18],
        ["면접일", "YYYY.MM.DD", 14],
        ["면접관", "이름(직책) — 복수 시 쉼표 구분", 24],
        ["결과", "통과/보류/탈락", 10],
        ["등급_점수", "A+/A/B+ 등 자유 입력", 14],
        ["면접코멘트", "자유 입력", 44],
        ["그리팅URL", "그리팅 해당 지원자 면접 기록 페이지 URL", 44],
      ],
      sample: ["240010008","Ted","서류전형","2024.02.10","HR팀","통과","",
               "이랜드그룹 HR 경력 3년9개월. 서류 통과.",""]
    },
    {
      name: "인사카드_수습면담",
      color: "#C678DD",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["면담차수", "1 / 2 / 3", 10],
        ["면담일", "YYYY.MM.DD", 14],
        ["면담자", "이름(직책)", 20],
        ["면담내용", "자유 입력", 44],
        ["특이사항_우려점", "없으면 '없음' 입력", 30],
        ["결과", "면담완료/수습연장", 14],
      ],
      sample: ["240010008","Ted",1,"2024.04.15","Peter (CHO)",
               "온보딩 적응도 良好. OKR 이해도 빠름.","사내 시스템 셋업 조율 필요.","면담완료"]
    },
    {
      name: "인사카드_수습평가",
      color: "#D19A66",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["평가차수", "1 / 2", 10],
        ["평가일", "YYYY.MM.DD", 14],
        ["평가자1", "이름(직책)", 18],
        ["평가자1_등급", "S/A/B/C 또는 자유", 14],
        ["평가자1_의견", "자유 입력", 30],
        ["평가자2", "이름(직책) — 없으면 공란", 18],
        ["평가자2_등급", "", 14],
        ["평가자2_의견", "", 30],
        ["평가자3", "필요 시 입력", 18],
        ["평가자3_등급", "", 14],
        ["평가자3_의견", "", 30],
        ["종합의견", "HR 또는 CHRO 종합 의견", 36],
        ["최종결과", "수습통과/수습연장/수습해지", 14],
      ],
      sample: ["240010008","Ted",1,"2024.06.01",
               "Peter (CHO)","A","HR 기획 역량 우수.",
               "Ben (CEO)","A","조직 문화 핏 良好.",
               "","","","전반적 기대치 충족. 수습 통과 권고.","수습통과"]
    },
    {
      name: "인사카드_발령",
      color: "#98C379",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["발령유형", "입사/부서이동/직책변경/직책승격/겸직/파견/복직", 20],
        ["발령일", "YYYY.MM.DD", 14],
        ["변경전_소속", "본부/부서/팀", 24],
        ["변경전_직책", "", 18],
        ["변경후_소속", "본부/부서/팀", 24],
        ["변경후_직책", "", 18],
        ["비고", "발령 사유 등", 24],
      ],
      sample: ["240010008","Ted","입사","2024.03.06","","",
               "BOOST 본부 인사기획 팀","팀장(Team Manager)",""]
    },
    {
      name: "인사카드_연봉",
      color: "#E5C07B",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["계약일", "YYYY.MM.DD", 14],
        ["계약유형", "입사계약/연봉근로계약/직책변경계약", 20],
        ["연봉총액_만원", "숫자만 입력", 16],
        ["월고정_만원", "숫자만 입력", 14],
        ["인상률", "예: +4.0% / 신규", 12],
        ["계약만료일", "YYYY.MM.DD — 만료 알림 자동 계산", 16],
        ["특별보수유형", "사이닝보너스/성과보수/리텐션보너스/없음", 20],
        ["특별보수금액_만원", "숫자", 16],
        ["특별보수지급일", "YYYY.MM.DD", 16],
        ["특별보수조건", "지급 조건 (예: 6개월 재직 시)", 30],
        ["비고", "", 20],
      ],
      sample: ["240010008","Ted","2026.03.01","연봉근로계약",7529,627,"+4.0%",
               "2027.03.01","없음","","","",""]
    },
    {
      name: "인사카드_평가이력",
      color: "#ABB2BF",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["평가기간", "예: 25년 하반기", 16],
        ["1차평가자", "이름(직책)", 18],
        ["2차평가자", "이름(직책)", 18],
        ["종합등급", "Outstanding/Superior/Adequate/Below", 16],
        ["혁신성과", "동일 등급 체계", 14],
        ["표준성과", "동일 등급 체계", 14],
        ["AI성과", "26년 상반기 신설 — 이전은 공란", 14],
        ["협업", "동일 등급 체계", 12],
        ["학습_성장_변화수용", "동일 등급 체계", 18],
        ["총평_코멘트", "자유 입력", 44],
      ],
      sample: ["240010008","Ted","25년 하반기","Peter (CHO)","Ben (CEO)",
               "Superior","Superior","Outstanding","","Adequate","Adequate",
               "HR 인프라 전체 구축 완료. OKR 시스템 설계 및 운영 주도."]
    },
    {
      name: "인사카드_이벤트알림",
      color: "#BE5046",
      cols: [
        ["사번", "", 12],
        ["닉네임", "참조용", 12],
        ["알림유형", "생일임박/결혼기념일임박/연봉만료임박/비자만료/수습종료임박", 26],
        ["기준일자", "YYYY.MM.DD", 14],
        ["D_day", "Apps Script 자동 계산", 12],
        ["표시여부", "Y/N — 자동 설정", 10],
        ["배지텍스트", "예: 🎂 생일 D-7", 24],
        ["최종갱신일시", "자동 기록", 20],
      ],
      sample: [] // 자동 생성 시트이므로 샘플 없음
    }
  ];

  SHEETS.forEach(function(def) {
    // 이미 있으면 건너뜀 (기존 시트 보호)
    var existing = ss.getSheetByName(def.name);
    if (existing) {
      results.push("⚠️ 이미 존재 (건너뜀): " + def.name);
      return;
    }

    var ws = ss.insertSheet(def.name);
    ws.setTabColor(def.color);

    var ncols = def.cols.length;
    var headers = def.cols.map(function(c){ return c[0]; });
    var notes   = def.cols.map(function(c){ return c[1]; });
    var widths  = def.cols.map(function(c){ return c[2]; });

    // 1행: 헤더
    var hRow = ws.getRange(1, 1, 1, ncols);
    hRow.setValues([headers]);
    hRow.setFontFamily("Arial");
    hRow.setFontWeight("bold");
    hRow.setFontSize(10);
    hRow.setFontColor("#FFFFFF");
    hRow.setBackground("#1E293B");
    hRow.setHorizontalAlignment("center");
    hRow.setVerticalAlignment("middle");
    ws.setRowHeight(1, 28);

    // 2행: 설명(주석)
    var nRow = ws.getRange(2, 1, 1, ncols);
    nRow.setValues([notes]);
    nRow.setFontFamily("Arial");
    nRow.setFontSize(8);
    nRow.setFontStyle("italic");
    nRow.setFontColor("#64748B");
    nRow.setBackground("#FFFBEB");
    ws.setRowHeight(2, 18);

    // 3행: 샘플 데이터
    if (def.sample && def.sample.length > 0) {
      var sRow = ws.getRange(3, 1, 1, ncols);
      var padded = def.sample.slice(0, ncols);
      while (padded.length < ncols) padded.push("");
      sRow.setValues([padded]);
      sRow.setFontFamily("Arial");
      sRow.setFontSize(9);
      sRow.setBackground("#F1F5F9");
      sRow.setFontColor("#475569");
      ws.setRowHeight(3, 20);
    }

    // 열 너비 설정
    widths.forEach(function(w, i) {
      ws.setColumnWidth(i+1, w * 7);
    });

    // 틀 고정 (3행부터 입력)
    ws.setFrozenRows(2);
    ws.setFrozenColumns(1);

    // 이벤트알림 시트에는 특별 안내 추가
    if (def.name === "인사카드_이벤트알림") {
      var infoRow = ws.getRange(3, 1, 1, ncols);
      infoRow.merge();
      infoRow.setValue("⚡ 이 시트는 refreshHrCardEvents() 함수가 자동으로 채워줍니다. 직접 편집하지 마세요.");
      infoRow.setFontSize(9);
      infoRow.setFontStyle("italic");
      infoRow.setFontColor("#2563EB");
      infoRow.setBackground("#EFF6FF");
    }

    results.push("✅ 생성 완료: " + def.name + " (" + ncols + "개 컬럼)");
  });

  SpreadsheetApp.getUi().alert(
    "인사카드 시트 생성 결과\n\n" + results.join("\n")
  );
}


// ================================================================
// 이벤트 알림 자동 갱신 — v2 전면 재작성
// 트리거 설정: createDailyTrigger() 1회 실행 → 매일 오전 9시 자동
// ================================================================
function refreshHrCardEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orgSheet    = ss.getSheetByName("조직도");
  var basicSheet  = ss.getSheetByName("인사카드_기본추가정보");
  var salarySheet = ss.getSheetByName("인사카드_연봉");
  var famSheet    = ss.getSheetByName("인사카드_가족");
  var eventSheet  = ss.getSheetByName("인사카드_이벤트알림");

  if (!orgSheet || !basicSheet || !salarySheet || !eventSheet) {
    Logger.log("필수 시트 없음. setupHrCardSheets() 먼저 실행하세요.");
    return;
  }

  var today = new Date();
  today.setHours(0,0,0,0);

  // ── 공통 헬퍼 ─────────────────────────────────────────
  // 셀 값 → "YYYY.MM.DD" 문자열 정규화 (Date 객체 / 텍스트 모두 지원)
  function toDateStr(v) {
    if (!v && v !== 0) return "";
    if (v instanceof Date) {
      return v.getFullYear() + "." +
             String(v.getMonth()+1).padStart(2,"0") + "." +
             String(v.getDate()).padStart(2,"0");
    }
    var s = String(v).trim();
    // "2026. 5. 17" / "2026-05-17" / "2026.05.17" 모두 수용
    s = s.replace(/[\-\/]/g, ".").replace(/\.\s*/g, ".").replace(/\.$/,"").replace(/\s/g,"");
    return s;
  }
  // "YYYY.MM.DD" → Date (실패 시 null)
  function parseDate(s) {
    s = toDateStr(s);
    if (!s) return null;
    var p = s.split(".");
    if (p.length < 3) return null;
    var y = parseInt(p[0]), m = parseInt(p[1]), d = parseInt(p[2]);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
    var dt = new Date(y, m-1, d);
    return isNaN(dt.getTime()) ? null : dt;
  }
  // 절대 D-day (만료일 등)
  function dDayCalc(dt) {
    if (!dt) return null;
    dt.setHours(0,0,0,0);
    return Math.ceil((dt - today) / 86400000);
  }
  // 매년 반복 D-day (생일·기념일): 올해 지나갔으면 내년 기준
  function annualDDay(s) {
    var dt = parseDate(s);
    if (!dt) return null;
    var t = new Date(today.getFullYear(), dt.getMonth(), dt.getDate());
    if (t < today) t = new Date(today.getFullYear()+1, dt.getMonth(), dt.getDate());
    return Math.ceil((t - today) / 86400000);
  }
  // D-day 라벨: 당일은 "D-day", 미래 "D-N", 경과 "D+N"
  function ddLabel(dd) {
    if (dd === 0) return "D-day";
    return dd > 0 ? "D-" + dd : "D+" + (-dd);
  }
  // 헤더 자동 탐지: 1~3행 내에서 mustHave 키워드가 포함된 행을 헤더로
  function findHeader(data, mustHave) {
    for (var i = 0; i < Math.min(3, data.length); i++) {
      var row = data[i].map(function(c){ return String(c).trim(); });
      var ok = row.some(function(c){
        return mustHave.some(function(kw){ return c === kw || c.indexOf(kw) >= 0; });
      });
      if (ok) return { idx: i, hdr: row };
    }
    return null;
  }
  // 헤더에서 컬럼 인덱스 찾기 (완전일치 우선 → 포함 매칭)
  function findCol(hdr, kws) {
    for (var k = 0; k < kws.length; k++) {
      var idx = hdr.indexOf(kws[k]);
      if (idx >= 0) return idx;
    }
    for (var k = 0; k < kws.length; k++) {
      for (var i = 0; i < hdr.length; i++) {
        if (hdr[i].indexOf(kws[k]) >= 0) return i;
      }
    }
    return -1;
  }

  // ── 조직도: 재직자 목록 (헤더 1~3행 자동 탐지) ──────────
  var orgData = orgSheet.getDataRange().getValues();
  var orgHdr = findHeader(orgData, ["사번"]);
  if (!orgHdr) {
    Logger.log("조직도 시트에서 '사번' 헤더를 찾지 못했습니다.");
    return;
  }
  var iSabun  = findCol(orgHdr.hdr, ["사번"]);
  var iNick   = findCol(orgHdr.hdr, ["닉네임(영문)", "닉네임"]);
  var iStatus = findCol(orgHdr.hdr, ["재직상태", "재직여부"]);
  var iJoin   = findCol(orgHdr.hdr, ["입사일"]);
  var iLeave  = findCol(orgHdr.hdr, ["퇴사예정일", "퇴사 예정일", "퇴사일"]); // ★ v2.2 (구 asana 컬럼 재활용)
  var iLast   = findCol(orgHdr.hdr, ["최종출근일", "최종 출근일", "마지막출근일"]); // ★ v2.3

  // ── 기본추가정보 맵 ──────────────────────────────────
  var basicData = basicSheet.getDataRange().getValues();
  var basicHdrInfo = findHeader(basicData, ["사번"]);
  var basicMap = {};
  if (basicHdrInfo) {
    var bHdr    = basicHdrInfo.hdr;
    var bSabun  = findCol(bHdr, ["사번"]);
    var bBday   = findCol(bHdr, ["생년월일"]);
    var bVisa   = findCol(bHdr, ["비자만료일"]);
    var bMarry  = findCol(bHdr, ["결혼여부"]);
    var bWed    = findCol(bHdr, ["결혼식일자", "결혼기념일"]);
    var bProbEnd= findCol(bHdr, ["수습종료일"]);
    var bEmpTy  = findCol(bHdr, ["고용형태"]);
    var bCtrEnd = findCol(bHdr, ["근로계약종료일", "근로계약 종료일"]);
    for (var bi = basicHdrInfo.idx + 1; bi < basicData.length; bi++) {
      var brow = basicData[bi];
      var bs = bSabun >= 0 ? String(brow[bSabun] || "").trim() : "";
      if (!bs || !/\d/.test(bs)) continue; // 설명행/공란 제외 (숫자 미포함 사번 스킵)
      basicMap[bs] = {
        bday:    bBday    >= 0 ? toDateStr(brow[bBday])    : "",
        visa:    bVisa    >= 0 ? toDateStr(brow[bVisa])    : "",
        marry:   bMarry   >= 0 ? String(brow[bMarry]||"").trim() : "",
        wed:     bWed     >= 0 ? toDateStr(brow[bWed])     : "",
        probEnd: bProbEnd >= 0 ? toDateStr(brow[bProbEnd]) : "",
        empType: bEmpTy   >= 0 ? String(brow[bEmpTy]||"").trim() : "",
        ctrEnd:  bCtrEnd  >= 0 ? toDateStr(brow[bCtrEnd])  : ""
      };
    }
  }

  // ── 가족 시트: 결혼기념일 + 출산예정일 (헤더 이름으로 탐지) ──
  var famMap = {}; // {사번: {anni:"", dues:[]}}
  if (famSheet) {
    var famData = famSheet.getDataRange().getValues();
    var famHdrInfo = findHeader(famData, ["사번"]);
    if (famHdrInfo) {
      var fSabun = findCol(famHdrInfo.hdr, ["사번"]);
      var fAnni  = findCol(famHdrInfo.hdr, ["결혼기념일"]);
      var fDue   = findCol(famHdrInfo.hdr, ["출산예정일", "출산예정"]);
      for (var fi = famHdrInfo.idx + 1; fi < famData.length; fi++) {
        var frow = famData[fi];
        var fs = fSabun >= 0 ? String(frow[fSabun] || "").trim() : "";
        if (!fs || !/\d/.test(fs)) continue;
        if (!famMap[fs]) famMap[fs] = { anni: "", dues: [] };
        if (fAnni >= 0) {
          var anni = toDateStr(frow[fAnni]);
          if (anni) famMap[fs].anni = anni;
        }
        if (fDue >= 0) {
          var due = toDateStr(frow[fDue]);
          if (due) famMap[fs].dues.push(due);
        }
      }
    }
  }

  // ── 연봉 시트: 계약만료일 (사번별 가장 늦은 만료일 채택) ──
  var salaryMap = {};
  var salData = salarySheet.getDataRange().getValues();
  var salHdrInfo = findHeader(salData, ["사번"]);
  if (salHdrInfo) {
    var sSabun  = findCol(salHdrInfo.hdr, ["사번"]);
    var sExpire = findCol(salHdrInfo.hdr, ["계약만료일"]);
    if (sExpire >= 0) {
      for (var si = salHdrInfo.idx + 1; si < salData.length; si++) {
        var srow = salData[si];
        var ssb = sSabun >= 0 ? String(srow[sSabun] || "").trim() : "";
        if (!ssb || !/\d/.test(ssb)) continue;
        var exp = toDateStr(srow[sExpire]);
        if (!exp) continue;
        // 계약이 여러 건이면 만료일이 가장 늦은(현행) 계약 기준
        if (!salaryMap[ssb] || exp > salaryMap[ssb]) salaryMap[ssb] = exp;
      }
    }
  }

  // ── 재직자 순회 ──────────────────────────────────────
  var rows = [];
  var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");

  for (var oi = orgHdr.idx + 1; oi < orgData.length; oi++) {
    var orow = orgData[oi];
    var sabun = iSabun >= 0 ? String(orow[iSabun] || "").trim() : "";
    if (!sabun || !/\d/.test(sabun)) continue;
    // 재직 판정: '퇴직예정'은 아직 재직 중이므로 포함, 그 외 '퇴' 포함(퇴직/퇴사)은 제외
    var st = iStatus >= 0 ? String(orow[iStatus] || "").trim() : "";
    var isLeaving = st.indexOf("예정") >= 0 && st.indexOf("퇴") >= 0; // 퇴직예정
    if (st && st.indexOf("퇴") >= 0 && !isLeaving) continue;

    var nick  = iNick >= 0 ? String(orow[iNick] || "").trim() : "";
    var binfo = basicMap[sabun] || {};

    // 1. 생일 (30일 이내)
    if (binfo.bday) {
      var dd1 = annualDDay(binfo.bday);
      if (dd1 !== null && dd1 <= 30) {
        rows.push([sabun, nick, "생일임박", binfo.bday, dd1, "Y", "🎂 생일 " + ddLabel(dd1), now]);
      }
    }

    // 2. 결혼기념일 (기혼자, 30일 이내) — 가족 시트 우선, 없으면 기본추가정보 결혼식일자
    if (binfo.marry && binfo.marry.indexOf("기혼") >= 0) {
      var fam2 = famMap[sabun] || {};
      var anni = fam2.anni || binfo.wed;
      if (anni) {
        var dd2 = annualDDay(anni);
        if (dd2 !== null && dd2 <= 30) {
          rows.push([sabun, nick, "결혼기념일임박", anni, dd2, "Y", "💐 기념일 " + ddLabel(dd2), now]);
        }
      }
    }

    // 3. 연봉만료 (30일 이내, 지난 만료일도 D+표시 위해 -30일까지 허용)
    var salExpire = salaryMap[sabun];
    if (salExpire) {
      var dd3 = dDayCalc(parseDate(salExpire));
      if (dd3 !== null && dd3 <= 30 && dd3 >= -30) {
        var t3 = dd3 >= 0 ? "📋 연봉갱신 " + ddLabel(dd3) : "📋 연봉만료 " + ddLabel(dd3);
        rows.push([sabun, nick, "연봉만료임박", salExpire, dd3, "Y", t3, now]);
      }
    }

    // 4. 비자만료 (외국인만, 항상 표시)
    if (binfo.visa) {
      var dd4 = dDayCalc(parseDate(binfo.visa));
      if (dd4 !== null) {
        rows.push([sabun, nick, "비자만료", binfo.visa, dd4, "Y", "⚠️ 비자만료 " + ddLabel(dd4), now]);
      }
    }

    // 5. 수습종료 (30일 이내) — 고용형태가 '수습'인 인원만 (v2.2)
    //    기본추가정보 '수습종료일' 우선, 없으면 입사일+3개월
    if (binfo.empType && binfo.empType.indexOf("수습") >= 0) {
      var probEndStr = binfo.probEnd || "";
      if (!probEndStr && iJoin >= 0 && orow[iJoin]) {
        var joinDate = parseDate(orow[iJoin]);
        if (joinDate) {
          var pe = new Date(joinDate);
          pe.setMonth(pe.getMonth() + 3);
          probEndStr = toDateStr(pe);
        }
      }
      if (probEndStr) {
        var dd5 = dDayCalc(parseDate(probEndStr));
        if (dd5 !== null && dd5 <= 30 && dd5 >= 0) {
          rows.push([sabun, nick, "수습종료임박", probEndStr, dd5, "Y", "🔔 수습종료 " + ddLabel(dd5), now]);
        }
      }
    }

    // 6. 출산예정 ★ v2.1 신규 — [인사카드_가족] '출산예정일' 컬럼
    //    60일 전부터 표시, 예정일 경과 후 14일까지 D+ 유지 (예정일 변동 대비)
    var fam6 = famMap[sabun];
    if (fam6 && fam6.dues && fam6.dues.length) {
      fam6.dues.forEach(function(dueStr){
        var dd6 = dDayCalc(parseDate(dueStr));
        if (dd6 !== null && dd6 <= 60 && dd6 >= -14) {
          rows.push([sabun, nick, "출산예정", dueStr, dd6, "Y", "👶 출산예정 " + ddLabel(dd6), now]);
        }
      });
    }

    // 7. 근로계약 종료 ★ v2.2 신규 — [인사카드_기본추가정보] '근로계약종료일' (계약직)
    //    재계약/종료 의사결정 리드타임 확보 위해 60일 전부터 표시, 경과 7일까지 유지
    if (binfo.ctrEnd) {
      var dd7 = dDayCalc(parseDate(binfo.ctrEnd));
      if (dd7 !== null && dd7 <= 60 && dd7 >= -7) {
        rows.push([sabun, nick, "근로계약종료", binfo.ctrEnd, dd7, "Y", "📄 근로계약종료 " + ddLabel(dd7), now]);
      }
    }

    // 8. 퇴사예정 ★ v2.2 신규 / v2.3 최종출근일 반영 — [조직도] 재직상태='퇴직예정'
    //    D-day 기준 = '최종출근일' 우선 (없으면 '퇴사예정일')
    //    최종출근일 ≠ 퇴사일이면 배지에 오피셜 퇴사일 병기 (연차 소진 케이스)
    //    표시 유지: 오피셜 퇴사일 경과 7일까지
    if (isLeaving) {
      var officialStr = (iLeave >= 0 && orow[iLeave]) ? toDateStr(orow[iLeave]) : "";
      var lastStr     = (iLast  >= 0 && orow[iLast])  ? toDateStr(orow[iLast])  : "";
      var baseStr = lastStr || officialStr; // D-day 기준일
      if (baseStr) {
        var ddBase = dDayCalc(parseDate(baseStr));
        var endStr = officialStr || lastStr; // 목록 유지 판정은 더 늦은(공식) 일자
        var ddEnd  = dDayCalc(parseDate(endStr));
        if (ddBase !== null && ddEnd !== null && ddEnd >= -7) {
          var badge8 = (lastStr && officialStr && lastStr !== officialStr)
            ? "🚪 최종출근 " + ddLabel(ddBase) + " (퇴사일 " + officialStr + ")"
            : "🚪 퇴사예정 " + ddLabel(ddBase);
          rows.push([sabun, nick, "퇴사예정", baseStr, ddBase, "Y", badge8, now]);
        }
      }
    }

    // 9. 입사 N주년 ★ v2.5 신규 — [조직도] 입사일 기준, 30일 이내 표시
    //    3/5/7/10주년은 리프레시 휴가·휴가비 대상 → 🏆 강조 표기
    if (iJoin >= 0 && orow[iJoin]) {
      var jd9 = parseDate(orow[iJoin]);
      if (jd9) {
        var anniv = new Date(today.getFullYear(), jd9.getMonth(), jd9.getDate());
        if (anniv < today) anniv = new Date(today.getFullYear() + 1, jd9.getMonth(), jd9.getDate());
        var yrs = anniv.getFullYear() - jd9.getFullYear();
        var dd9 = Math.ceil((anniv - today) / 86400000);
        if (yrs >= 1 && dd9 !== null && dd9 <= 30) {
          var isMilestone = (yrs === 3 || yrs === 5 || yrs === 7 || yrs === 10);
          var badge9 = isMilestone
            ? "🏆 입사 " + yrs + "주년 " + ddLabel(dd9) + " (리프레시 휴가·휴가비 대상)"
            : "🎉 입사 " + yrs + "주년 " + ddLabel(dd9);
          rows.push([sabun, nick, "입사기념일", toDateStr(jd9), dd9, "Y", badge9, now]);
        }
      }
    }
  }

  // D-day 오름차순 정렬 (급한 것부터)
  rows.sort(function(a, b){ return a[4] - b[4]; });

  // ── 이벤트알림 시트 갱신 ─────────────────────────────
  var lastRow = eventSheet.getLastRow();
  if (lastRow >= 3) {
    // 병합 잔여(과거 안내행) 대비 병합 해제 후 삭제
    try { eventSheet.getRange(3, 1, lastRow-2, 8).breakApart(); } catch(e) {}
    eventSheet.getRange(3, 1, lastRow-2, 8).clearContent()
              .setBackground(null).setFontColor(null);
  }

  if (rows.length > 0) {
    var writeRange = eventSheet.getRange(3, 1, rows.length, 8);
    writeRange.setValues(rows);
    writeRange.setFontFamily("Arial");
    writeRange.setFontSize(9);

    // D-day 5 이하 빨강 / 14 이하 주황 강조
    for (var ri = 0; ri < rows.length; ri++) {
      var dd = rows[ri][4];
      if (typeof dd === "number" && dd <= 5) {
        eventSheet.getRange(ri+3, 1, 1, 8).setBackground("#FEF2F2").setFontColor("#DC2626");
      } else if (typeof dd === "number" && dd <= 14) {
        eventSheet.getRange(ri+3, 1, 1, 8).setBackground("#FFFBEB").setFontColor("#D97706");
      }
    }
  } else {
    // 0건이어도 상태 행을 남겨 "미실행"과 구분
    eventSheet.getRange(3, 1, 1, 8).setValues([[
      "", "", "알림없음", "", "", "N", "오늘 기준 표시할 알림 없음", now
    ]]);
    eventSheet.getRange(3, 1, 1, 8).setFontSize(9).setFontStyle("italic").setFontColor("#94A3B8");
  }

  Logger.log("이벤트 알림 갱신 완료. 생성된 알림: " + rows.length + "건 (기준: " + now + ")");
}


// ================================================================
// 월별 입퇴사·재직자 수 자동 집계 ★ v2.4 신규
// [인건비] 시트의 월초/월말/입사/자발/비자발 5개 컬럼을 자동 계산
//   - 원천: [조직도] (입사일·재직상태) + [퇴직자] (입사일·퇴직일·퇴직구분)
//   - 정책: **당월 + 전월만 갱신** — 그 이전 월은 절대 건드리지 않음
//   - 직접비/간접비/매출/특이사항 등 다른 컬럼은 절대 건드리지 않음 (헤더 매칭 셀 단위 쓰기)
//   - 당월 '월말'은 오늘 기준 재직자 수 (월이 끝나면 실제 월말 값으로 수렴)
//   - 퇴직일 = 마지막 재직일로 간주 (퇴직일 당일까지 재직 카운트)
//   - 해당 연도/월 행이 없으면 연도·월·계산 5컬럼만 채워 새 행 추가
// 실행: refreshHeadcount 수동 실행 또는 createDailyTrigger로 매일 오전 9시 자동
// ================================================================
function refreshHeadcount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var costSheet = ss.getSheetByName("인건비");
  var orgSheet  = ss.getSheetByName("조직도");
  var lvSheet   = ss.getSheetByName("퇴직자");
  if (!costSheet || !orgSheet) { Logger.log("필수 시트 없음 ([인건비]/[조직도])"); return; }

  var today = new Date(); today.setHours(0,0,0,0);

  // ── 헬퍼 ──
  function toDateStr(v) {
    if (!v && v !== 0) return "";
    if (v instanceof Date) {
      return v.getFullYear() + "." + String(v.getMonth()+1).padStart(2,"0") + "." + String(v.getDate()).padStart(2,"0");
    }
    var s = String(v).trim().replace(/[\-\/]/g,".").replace(/\.\s*/g,".").replace(/\.$/,"").replace(/\s/g,"");
    return s;
  }
  function parseDate(s) {
    s = toDateStr(s);
    if (!s) return null;
    var p = s.split(".");
    if (p.length < 3) return null;
    var d = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
    return isNaN(d.getTime()) ? null : d;
  }
  function findHeader(data, kw) {
    for (var i = 0; i < Math.min(3, data.length); i++) {
      var row = data[i].map(function(c){ return String(c).trim(); });
      if (row.some(function(c){ return c === kw || c.indexOf(kw) >= 0; })) return { idx: i, hdr: row };
    }
    return null;
  }
  function findCol(hdr, kws) {
    for (var k = 0; k < kws.length; k++) { var i = hdr.indexOf(kws[k]); if (i >= 0) return i; }
    for (var k = 0; k < kws.length; k++) {
      for (var i = 0; i < hdr.length; i++) { if (hdr[i].indexOf(kws[k]) >= 0) return i; }
    }
    return -1;
  }

  // ── 통합 명부 구성: 조직도(재직·퇴직예정 포함) + 퇴직자 ──
  // { 사번: {join:Date, leave:Date|null, kind:''} }
  var roster = {};

  var orgData = orgSheet.getDataRange().getValues();
  var oh = findHeader(orgData, "사번");
  if (!oh) { Logger.log("[조직도] 헤더 탐지 실패"); return; }
  var oSabun = findCol(oh.hdr, ["사번"]);
  var oJoin  = findCol(oh.hdr, ["입사일"]);
  var oSt    = findCol(oh.hdr, ["재직상태", "재직여부"]);
  for (var i = oh.idx + 1; i < orgData.length; i++) {
    var r = orgData[i];
    var sb = String(r[oSabun] || "").trim();
    if (!sb || !/\d/.test(sb)) continue;
    var st = oSt >= 0 ? String(r[oSt] || "").trim() : "";
    // 조직도에 '퇴직' 상태 행이 남아있으면 퇴직일은 [퇴직자] 시트에서 채움 (아래 병합)
    roster[sb] = {
      join: parseDate(oJoin >= 0 ? r[oJoin] : ""),
      leave: null,
      orgLeft: (st.indexOf("퇴") >= 0 && st.indexOf("예정") < 0), // 조직도상 이미 퇴직 표기
    };
  }

  if (lvSheet) {
    var lvData = lvSheet.getDataRange().getValues();
    var lh = findHeader(lvData, "퇴직일") || findHeader(lvData, "퇴사일");
    if (lh) {
      var lSabun = findCol(lh.hdr, ["사번"]);
      var lJoin  = findCol(lh.hdr, ["입사일"]);
      var lLeave = findCol(lh.hdr, ["퇴직일", "퇴사일"]);
      var lKind  = findCol(lh.hdr, ["퇴직구분", "구분"]);
      for (var i = lh.idx + 1; i < lvData.length; i++) {
        var r = lvData[i];
        var sb = String(r[lSabun] || "").trim();
        var lv = parseDate(lLeave >= 0 ? r[lLeave] : "");
        if (!lv) continue;
        var key = (sb && /\d/.test(sb)) ? sb : ("LV_" + i); // 사번 없으면 임시 키
        if (!roster[key]) roster[key] = { join: null, leave: null };
        if (!roster[key].join) roster[key].join = parseDate(lJoin >= 0 ? r[lJoin] : "");
        // 재입사자 대비: 더 늦은 퇴직일 우선 (단순화)
        if (!roster[key].leave || lv > roster[key].leave) roster[key].leave = lv;
        roster[key].kind = lKind >= 0 ? String(r[lKind] || "").trim() : "";
      }
    }
  }

  // 조직도상 '퇴직' 표기인데 [퇴직자]에 퇴직일이 없는 인원 → 시점 계산 불가, 집계 제외 + 로그
  var skipped = [];
  Object.keys(roster).forEach(function(k){
    var p = roster[k];
    if (p.orgLeft && !p.leave) { skipped.push(k); p.exclude = true; }
  });

  // ── 특정 날짜 재직 여부 (퇴직일 = 마지막 재직일) ──
  function employedOn(p, d) {
    if (p.exclude || !p.join) return false;
    if (p.join > d) return false;
    if (p.leave && p.leave < d) return false;
    return true;
  }

  // ── 대상: 당월 + 전월 ──
  var targets = [];
  var cy = today.getFullYear(), cm = today.getMonth() + 1;
  targets.push({ y: cy, m: cm });
  targets.push(cm === 1 ? { y: cy - 1, m: 12 } : { y: cy, m: cm - 1 });

  // ── 인건비 시트 헤더/컬럼 위치 ──
  // ※ 1행 배너("...연도+월 입력...")가 부분일치로 오인되지 않도록
  //    '연도'와 '월' 셀이 정확히 존재하는 행만 헤더로 인정 (v2.4.1)
  var costData = costSheet.getDataRange().getValues();
  var ch = null;
  for (var i = 0; i < Math.min(3, costData.length); i++) {
    var row = costData[i].map(function(c){ return String(c).trim(); });
    if (row.indexOf("연도") >= 0 && row.indexOf("월") >= 0) { ch = { idx: i, hdr: row }; break; }
  }
  if (!ch) { Logger.log("[인건비] 헤더 탐지 실패 ('연도'/'월' 정확일치 셀 필요)"); return; }
  var cYear = findCol(ch.hdr, ["연도"]);
  var cMon  = findCol(ch.hdr, ["월"]);
  var col = {
    s:   findCol(ch.hdr, ["월초"]),
    e:   findCol(ch.hdr, ["월말"]),
    j:   findCol(ch.hdr, ["입사"]),
    v:   findCol(ch.hdr, ["자발"]),
    inv: findCol(ch.hdr, ["비자발"]),
  };
  // '자발' 키워드가 '비자발'에 먼저 걸리지 않도록 보정
  if (col.v === col.inv) {
    col.v = -1;
    for (var i = 0; i < ch.hdr.length; i++) {
      if (ch.hdr[i].indexOf("자발") >= 0 && ch.hdr[i].indexOf("비자발") < 0) { col.v = i; break; }
    }
  }

  var log = [];
  var unclassified = [];
  targets.forEach(function(t){
    var first = new Date(t.y, t.m - 1, 1);
    var last  = new Date(t.y, t.m, 0);
    var effEnd = last > today ? today : last; // 당월 월말은 오늘 기준

    var startCnt = 0, endCnt = 0, joinCnt = 0, volCnt = 0, invCnt = 0;
    Object.keys(roster).forEach(function(k){
      var p = roster[k];
      if (employedOn(p, first))  startCnt++;
      if (employedOn(p, effEnd)) endCnt++;
      if (p.join && !p.exclude && p.join >= first && p.join <= last) joinCnt++;
      if (p.leave && p.leave >= first && p.leave <= last) {
        var kind = String(p.kind || "");
        // 대시보드 퇴직자 KPI와 동일 규칙: '비자발' 포함 → 비자발 / '자발' 포함 → 자발
        if (kind.indexOf("비자발") >= 0) invCnt++;
        else if (kind.indexOf("자발") >= 0) volCnt++;
        else { volCnt++; unclassified.push(k + "(" + (kind || "구분없음") + ")"); } // 미분류는 자발로 집계 + 경고
      }
    });

    // 해당 연도/월 행 찾기
    var rowIdx = -1;
    for (var i = ch.idx + 1; i < costData.length; i++) {
      if (parseInt(String(costData[i][cYear]).trim()) === t.y &&
          parseInt(String(costData[i][cMon]).trim())  === t.m) { rowIdx = i; break; }
    }
    if (rowIdx < 0) {
      // 새 행 추가: 연도·월 + 계산 5컬럼만
      var newRow = Array(ch.hdr.length).fill("");
      newRow[cYear] = t.y; newRow[cMon] = t.m;
      if (col.s   >= 0) newRow[col.s]   = startCnt;
      if (col.e   >= 0) newRow[col.e]   = endCnt;
      if (col.j   >= 0) newRow[col.j]   = joinCnt;
      if (col.v   >= 0) newRow[col.v]   = volCnt;
      if (col.inv >= 0) newRow[col.inv] = invCnt;
      costSheet.appendRow(newRow);
      log.push(t.y + "." + t.m + " 신규행: 월초 " + startCnt + " / 월말 " + endCnt + " / 입사 " + joinCnt + " / 자발 " + volCnt + " / 비자발 " + invCnt);
    } else {
      // 기존 행: 계산 5컬럼 셀만 개별 갱신 (다른 컬럼 불가침)
      var sheetRow = rowIdx + 1;
      if (col.s   >= 0) costSheet.getRange(sheetRow, col.s   + 1).setValue(startCnt);
      if (col.e   >= 0) costSheet.getRange(sheetRow, col.e   + 1).setValue(endCnt);
      if (col.j   >= 0) costSheet.getRange(sheetRow, col.j   + 1).setValue(joinCnt);
      if (col.v   >= 0) costSheet.getRange(sheetRow, col.v   + 1).setValue(volCnt);
      if (col.inv >= 0) costSheet.getRange(sheetRow, col.inv + 1).setValue(invCnt);
      log.push(t.y + "." + t.m + " 갱신: 월초 " + startCnt + " / 월말 " + endCnt + " / 입사 " + joinCnt + " / 자발 " + volCnt + " / 비자발 " + invCnt);
    }
  });

  if (skipped.length) log.push("⚠️ 집계 제외 (조직도 퇴직 표기 + [퇴직자] 퇴직일 없음): " + skipped.join(", "));
  if (unclassified.length) log.push("⚠️ 퇴직구분 미분류 → 자발로 집계됨: " + unclassified.join(", ") + " — [퇴직자] 퇴직구분에 자발/비자발 명시 필요");
  Logger.log("입퇴사·재직자 자동 집계 완료\n" + log.join("\n"));
}


// ================================================================
// 수습 → 정규직 자동 전환 ★ v2.5 신규
// 조건: [인사카드_기본추가정보] 고용형태='수습' AND 수습종료일 경과
//       AND [인사카드_수습평가] 해당 사번의 최종결과에 '통과' 포함
// v2.5.1: ① 수습종료일 미입력 시 입사일+3개월 폴백 (알림·현황과 동일 규칙 — '통과' 게이트가 있어 안전)
//         ② 조직도 재직상태 '퇴직/퇴직예정' 인원 점검 대상에서 제외
//         ③ 수습 기간 중인 인원은 로그 생략 (종료일 경과+미통과만 보류 표시)
// ※ 종료일만 지나고 통과 기록이 없으면 전환하지 않고 로그에 보류 명단 기록
//    (연장/해지/평가 미입력 케이스의 오전환 방지 — 처리 누락 감지 역할)
// 실행: autoConvertProbation 수동 실행 또는 createDailyTrigger로 매일 오전 9시
// ================================================================
function autoConvertProbation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bs = ss.getSheetByName("인사카드_기본추가정보");
  var es = ss.getSheetByName("인사카드_수습평가");
  var os = ss.getSheetByName("조직도");
  if (!bs || !os) { Logger.log("[인사카드_기본추가정보]/[조직도] 시트 없음"); return; }

  var today = new Date(); today.setHours(0,0,0,0);
  function toDt(v) {
    if (!v && v !== 0) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    var s = String(v).trim().replace(/[\-\/]/g,".").replace(/\.\s*/g,".").replace(/\.$/,"").replace(/\s/g,"");
    var p = s.split(".");
    if (p.length < 3) return null;
    var d = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
    return isNaN(d.getTime()) ? null : d;
  }
  function findHdr(data, kw) {
    for (var i = 0; i < Math.min(3, data.length); i++) {
      var row = data[i].map(function(c){ return String(c).trim(); });
      if (row.some(function(c){ return c === kw || c.indexOf(kw) >= 0; })) return { idx: i, hdr: row };
    }
    return null;
  }

  // ── 조직도: 사번 → {입사일, 재직상태} — 퇴사자 제외 + 입사일+3개월 폴백용 (v2.5.1)
  var orgMap = {};
  var od = os.getDataRange().getValues();
  var oHdrInfo = findHdr(od, "사번");
  if (oHdrInfo) {
    var oSab = oHdrInfo.hdr.findIndex(function(h){ return h === "사번" || h.indexOf("사번") >= 0; });
    var oJn  = oHdrInfo.hdr.findIndex(function(h){ return h.indexOf("입사일") >= 0; });
    var oSt  = oHdrInfo.hdr.findIndex(function(h){ return h.indexOf("재직상태") >= 0 || h.indexOf("재직여부") >= 0; });
    for (var i = oHdrInfo.idx + 1; i < od.length; i++) {
      var sb = String(od[i][oSab] || "").trim();
      if (!sb || !/\d/.test(sb)) continue;
      orgMap[sb] = {
        join: oJn >= 0 ? toDt(od[i][oJn]) : null,
        status: oSt >= 0 ? String(od[i][oSt] || "").trim() : ""
      };
    }
  }

  // ── 수습평가: 사번별 최종결과 (아래 행일수록 최신으로 간주해 덮어씀)
  var resultMap = {};
  if (es) {
    var ed = es.getDataRange().getValues();
    var eHdrInfo = findHdr(ed, "사번");
    if (eHdrInfo) {
      var eSabun = eHdrInfo.hdr.findIndex(function(h){ return h === "사번" || h.indexOf("사번") >= 0; });
      var eRes   = eHdrInfo.hdr.findIndex(function(h){ return h.indexOf("최종결과") >= 0; });
      if (eRes >= 0) {
        for (var i = eHdrInfo.idx + 1; i < ed.length; i++) {
          var sb = String(ed[i][eSabun] || "").trim();
          if (!sb || !/\d/.test(sb)) continue;
          var res = String(ed[i][eRes] || "").trim();
          if (res) resultMap[sb] = res;
        }
      }
    }
  }

  // ── 기본추가정보 순회
  var bd = bs.getDataRange().getValues();
  var bHdrInfo = findHdr(bd, "사번");
  if (!bHdrInfo) { Logger.log("[인사카드_기본추가정보] 헤더 탐지 실패"); return; }
  var bHdr = bHdrInfo.hdr;
  var bSabun = bHdr.findIndex(function(h){ return h === "사번" || h.indexOf("사번") >= 0; });
  var bType  = bHdr.findIndex(function(h){ return h.indexOf("고용형태") >= 0; });
  var bProb  = bHdr.findIndex(function(h){ return h.indexOf("수습종료일") >= 0; });
  if (bType < 0) { Logger.log("고용형태 컬럼 없음"); return; }

  var converted = [], held = [];
  for (var i = bHdrInfo.idx + 1; i < bd.length; i++) {
    var sb = String(bd[i][bSabun] || "").trim();
    if (!sb || !/\d/.test(sb)) continue;
    var typ = String(bd[i][bType] || "").trim();
    if (typ.indexOf("수습") < 0) continue;

    // 퇴사자·퇴사예정자 제외 (v2.5.1) — 재직 중인 수습만 점검
    var org = orgMap[sb] || {};
    var st = String(org.status || "");
    if (st.indexOf("퇴") >= 0) continue;

    // 수습종료일: 시트 값 우선, 없으면 입사일+3개월 폴백 (v2.5.1 — 알림·현황과 동일 규칙)
    var pe = bProb >= 0 ? toDt(bd[i][bProb]) : null;
    if (!pe && org.join) {
      pe = new Date(org.join);
      pe.setMonth(pe.getMonth() + 3);
    }
    if (!pe) { held.push(sb + "(종료일·입사일 모두 확인 불가)"); continue; }
    if (pe >= today) continue; // 아직 수습 기간 중 — 로그 생략

    var res = resultMap[sb] || "";
    if (res.indexOf("통과") >= 0) {
      bs.getRange(i + 1, bType + 1).setValue("정규직");
      converted.push(sb + " (종료 " + Utilities.formatDate(pe, "Asia/Seoul", "yyyy.MM.dd") + ", " + res + ")");
    } else {
      held.push(sb + " (종료 " + Utilities.formatDate(pe, "Asia/Seoul", "yyyy.MM.dd") + " 경과, " + (res || "수습평가 미기록") + ")");
    }
  }

  var log = [];
  if (converted.length) log.push("✅ 정규직 자동 전환 " + converted.length + "명: " + converted.join(", "));
  if (held.length)      log.push("⏸ 전환 보류 (종료일 경과, 통과 기록 없음 → 연장이면 수습종료일 갱신 / 해지면 퇴사 처리): " + held.join(", "));
  if (!log.length)      log.push("전환 대상 없음 (전원 수습 기간 중이거나 처리 완료)");
  Logger.log("수습→정규직 자동 전환 점검 완료\n" + log.join("\n"));
}


// ================================================================
// (선택) 매일 자동 실행 트리거 등록 — 한 번만 실행
// v2.6: 인사발령 자동 동기화(syncAppointments) 추가 → 총 4건
// v2.5: 이벤트 알림 + 입퇴사 집계 + 수습 자동 전환, 3건 모두 등록
// ※ 재실행해도 안전 (동일 함수의 기존 트리거를 지우고 새로 등록)
// ================================================================
function createDailyTrigger() {
  var HANDLERS = ["autoConvertProbation", "refreshHrCardEvents", "refreshHeadcount", "syncAppointments"];

  // 기존 트리거 중복 방지
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (HANDLERS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("autoConvertProbation").timeBased().everyDays(1).atHour(8).create(); // 전환 먼저
  ScriptApp.newTrigger("syncAppointments").timeBased().everyDays(1).atHour(9).create();     // ★ v2.6
  ScriptApp.newTrigger("refreshHrCardEvents").timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger("refreshHeadcount").timeBased().everyDays(1).atHour(9).create();
  var msg = "✅ 자동 갱신 트리거 4건이 등록되었습니다.\n" +
            "  오전 8시 — 수습→정규직 전환\n" +
            "  오전 9시 — 인사발령 동기화 / 이벤트 알림 / 입퇴사·재직자 집계";
  // 스프레드시트 UI가 연결되지 않은 컨텍스트(편집기 단독 실행 등)에서는 알림창 대신 로그로 안내
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
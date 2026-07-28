// ============================================================
//  윌로그 HR 대시보드 — Google Apps Script v3.7
//  [인건비] 시트: 연도+월 구조 (2025/2026 통합)
//  [HR인사이트] 시트: 연월+코멘트
//  [HR월간리포트] 시트: 연월+제목+본문
//  [조직도] 시트: 기존 동일
//  [설정] 시트: 목표값 + 부서별 예산 (budget_부서명) + AI 비용 (ai_본부명) ★
//  [외부공유용 조직도] 시트: 외부 공개용
//  v3.7   변경: readProbation에 수습평가 '최종결과' 포함 (수습통과/연장/해지 표시용)
//  v3.6   변경: 부서별 확장 월비용 항목 — readDeptMonthlyCosts()
//               [설정] '월비용_항목명_본부명' 키 자동 인식 (야근식대·법인카드 등 코드 수정 없이 확장)
//  v3.5.1 변경: readProbation — 고용형태='수습' 인원만 반환 (정규직 전환자 제외),
//               날짜 기반 자동 제외 삭제 (고용형태가 유일한 목록 기준)
//  v3.5   변경: 조직도 탭 노출용 데이터 2종 추가
//               ① hrEvents  — [인사카드_이벤트알림] 시트 읽기 (readHrEvents)
//               ② probation — 수습 직원 현황 자동 계산 (readProbation)
//                  조직도(입사일) + 기본추가정보(수습종료일) + 수습면담/평가/설문 집계
//  v3.4   변경: [인건비] '특이사항' 컬럼 읽기 추가 — 월별 주요 지표 각주/📌 표시용
//               (시트 맨 끝 컬럼 권장 — saveCostRow 위치 기반 쓰기와 충돌 방지)
//  v3.3.1 수정: getProfile 라우트에 설문 첨부(인사카드 실제 라우트), 매칭·메타 제외 보강
//  v3.3 변경: [설문_개인진단] 수습 목표 설문 → 인사카드 수습평가 탭 연동 — readProbSurvey()
//             [퇴직자] 직책 컬럼 반환 추가
//  v3.2 변경: 퇴직자 누적 관리 시트 연동 — readLeavers()
//  v3.1 변경: 부서별 AI 사용 비용(Claude) 반환 추가 — readDeptAiCost()
// ============================================================

const SHEET_ORG      = '조직도';
const SHEET_COST     = '인건비';
const SHEET_REPORT   = 'HR리포트';
const SHEET_CFG      = '설정';
const SHEET_EXT_ORG  = '외부공유용 조직도';
const SHEET_LEAVERS  = '퇴직자';
const SHEET_PROB_SURVEY = '설문_개인진단';

function doGet(e) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const action = e.parameter.action || 'read';

    if (action === 'getHrCard') {
      const card = getHrCardData(ss, e.parameter.sabun) || {};
      try { card.probSurvey = readProbSurvey(ss, e.parameter.sabun); }
      catch (err) { card.probSurvey = []; }
      return jsonOut(card);
    }

    if (action === 'getProfile') {
      const prof = getProfileData(e.parameter.empId) || {};
      try { prof.probSurvey = readProbSurvey(ss, e.parameter.empId); }
      catch (err) { prof.probSurvey = []; }
      return jsonOut(prof);
    }

    if (action === 'read') {
      const curYear = new Date().getFullYear();
      const costData = readCost(ss, curYear);
      const reportData = readReport(ss, curYear);
      return jsonOut({
        emps:            readOrgStructured(ss),
        orgRaw:          readOrgRaw(ss),
        months:          costData.current,
        prevMonths:      costData.prev,
        insights:        reportData.insights,
        report:          reportData.latest,
        revenue:         readConfig(ss, 'revenue', 150),
        targetCostRatio: readConfig(ss, 'targetCostRatio', 60),
        targetPerPerson: readConfig(ss, 'targetPerPerson', 2.0),
        annualCostTarget:readConfig(ss, 'ANNUAL_COST_TARGET', 75),
        deptBudgets:     readDeptBudgets(ss),  // ★ 부서별 예산
        deptAiCost:      readDeptAiCost(ss),   // ★ 부서별 AI 사용 비용 (Claude) v3.1 신규
        deptMoCosts:     readDeptMonthlyCosts(ss), // ★ 부서별 확장 월비용 항목 v3.6 신규
        leavers:         readLeavers(ss),       // ★ 퇴직자 누적 관리 v3.2 신규
        hrEvents:        readHrEvents(ss),      // ★ 임박 이벤트 (조직도 탭 노출) v3.5 신규
        probation:       readProbation(ss),     // ★ 수습 직원 현황 자동 계산 v3.5 신규
      });
    }

    if (action === 'write') {
      const type = e.parameter.type;
      const data = e.parameter.data
        ? JSON.parse(decodeURIComponent(e.parameter.data))
        : null;
      if (type === 'org'             && data)         saveOrg(ss, data.raw, data.emps);
      if (type === 'cost'            && data)         saveCostRow(ss, data);
      if (type === 'revenue'         && data != null) saveConfig(ss, 'revenue', data);
      if (type === 'targetCostRatio' && data != null) saveConfig(ss, 'targetCostRatio', data);
      if (type === 'targetPerPerson' && data != null) saveConfig(ss, 'targetPerPerson', data);
      return jsonOut({ ok: true });
    }

    if (action === 'readExtOrg') {
      const extNodes = readExtOrg(ss);
      const refDate = (()=>{
        const sh = getOrCreate(ss, SHEET_CFG);
        const data = sh.getDataRange().getValues();
        for(const row of data){
          if(String(row[0]).trim()==='extOrgRefDate') return String(row[1]||'').trim();
        }
        return '';
      })();
      return jsonOut({ nodes: extNodes, ref_date: refDate });
    }

    return jsonOut({ error: '알 수 없는 action: ' + action });

  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ════════════════════════════════════════════════════════════
//  인건비: 연도+월 구조 읽기
//  v3.4: '특이사항' 컬럼 추가 읽기 (없으면 빈 문자열 — 기존 시트와 호환)
// ════════════════════════════════════════════════════════════
function readCost(ss, curYear) {
  const sh   = getOrCreate(ss, SHEET_COST);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { current: emptyMonths(), prev: emptyMonths() };

  let hdrRowIdx = 0;
  for (let i = 0; i < Math.min(3, data.length); i++) {
    const row = data[i].map(c => String(c).trim());
    if (row.some(c => c === '연도' || c === '월')) { hdrRowIdx = i; break; }
  }
  const hdr = data[hdrRowIdx].map(c => String(c).trim());
  const fi  = (...kws) => {
    for (const kw of kws) {
      const idx = hdr.findIndex(h => h === kw || h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const col = {
    year: fi('연도'), month: fi('월'),
    d:   fi('직접비', '급여'),
    i1:  fi('간접비1', '4대보험'),
    i2:  fi('간접비2', '퇴직'),
    s:   fi('월초'),
    e:   fi('월말'),
    j:   fi('입사'),
    v:   fi('자발'),
    inv: fi('비자발'),
    rev: fi('매출'),
    note: fi('특이사항', '비고'),   // ★ v3.4 신규
  };

  const g   = (row, c) => c >= 0 ? String(row[c] !== undefined ? row[c] : '').trim() : '';
  const num = (row, c) => { const v = parseFloat(g(row, c)); return isNaN(v) ? 0 : v; };

  const current = emptyMonths();
  const prev    = emptyMonths();

  data.slice(hdrRowIdx + 1).forEach(row => {
    const yr = parseInt(g(row, col.year));
    const mn = parseInt(g(row, col.month));
    if (isNaN(yr) || isNaN(mn) || mn < 1 || mn > 12) return;

    const entry = {
      직접비:  g(row, col.d),
      간접비1: g(row, col.i1),
      간접비2: g(row, col.i2),
      월초:    g(row, col.s),
      월말:    g(row, col.e),
      입사:    g(row, col.j),
      자발:    g(row, col.v),
      비자발:  g(row, col.inv),
      매출:    g(row, col.rev),
      매출억:  num(row, col.rev),
      특이사항: g(row, col.note),   // ★ v3.4 신규
    };

    if (yr === curYear)     current[mn - 1] = entry;
    if (yr === curYear - 1) prev[mn - 1]    = entry;
  });

  return { current, prev };
}

function emptyMonths() {
  return Array(12).fill(null).map(() => ({
    직접비:'', 간접비1:'', 간접비2:'', 월초:'', 월말:'',
    입사:'', 자발:'', 비자발:'', 매출:'', 매출억:0, 특이사항:'',
  }));
}

// ════════════════════════════════════════════════════════════
//  HR 리포트 통합 읽기 (인사이트 + 월간리포트)
// ════════════════════════════════════════════════════════════
function readReport(ss, curYear) {
  const sh = getOrCreate(ss, SHEET_REPORT);
  const data = sh.getDataRange().getValues();
  const empty = { insights: Array(12).fill(null).map(() => ({})), latest: null };
  if (data.length < 2) return empty;

  let hdrRowIdx = 0;
  for (let i = 0; i < Math.min(3, data.length); i++) {
    const row = data[i].map(c => String(c).trim());
    if (row.some(c => c === '연도' || c === '월')) { hdrRowIdx = i; break; }
  }
  const hdr = data[hdrRowIdx].map(c => String(c).trim());
  const fi  = (...kws) => {
    for (const kw of kws) {
      let idx = hdr.findIndex(h => h === kw);
      if (idx >= 0) return idx;
    }
    for (const kw of kws) {
      let idx = hdr.findIndex(h => h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const col = {
    year:    fi('연도'),
    month:   fi('월'),
    title:   fi('리포트 제목', '제목'),
    body:    fi('리포트 본문', '본문'),
    hc_c:   fi('인원수 코멘트', '인원수'),
    cr_c:   fi('인건비율 코멘트', '인건비율'),
    tr_c:   fi('퇴사율 코멘트', '퇴사율'),
    cr_tgt: fi('목표 인건비율', '인건비율(%)'),
    tr_tgt: fi('목표 퇴사율',   '퇴사율(%)'),
  };

  const g   = (row, c) => c >= 0 ? String(row[c] || '').trim() : '';
  const insights = Array(12).fill(null).map(() => ({}));
  let latest = null;

  data.slice(hdrRowIdx + 1).forEach(row => {
    const yr = parseInt(g(row, col.year));
    const mn = parseInt(g(row, col.month));
    if (isNaN(yr) || isNaN(mn) || mn < 1 || mn > 12) return;

    if (yr === curYear) {
      insights[mn - 1] = {
        hc_comment:   g(row, col.hc_c),
        cr_comment:   g(row, col.cr_c),
        tr_comment:   g(row, col.tr_c),
        cr_target:    g(row, col.cr_tgt),
        tr_target:    g(row, col.tr_tgt),
        report_title: g(row, col.title),
      };
    }

    const body = g(row, col.body);
    if (body && (!latest || yr > latest.year || (yr === latest.year && mn > latest.month))) {
      latest = { year: yr, month: mn, title: g(row, col.title), body };
    }
  });

  return { insights, latest };
}

// ════════════════════════════════════════════════════════════
//  인건비 행 저장 (단일 행 추가)
//  ※ 위치 기반(1~11열) 쓰기 — '특이사항'은 12열 이후에 두면 이 함수가 건드리지 않음
// ════════════════════════════════════════════════════════════
function saveCostRow(ss, data) {
  const sh = getOrCreate(ss, SHEET_COST);
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(data.year) &&
        String(rows[i][1]).trim() === String(data.month)) {
      const vals = [data.year, data.month, data.직접비||'', data.간접비1||'', data.간접비2||'',
                    data.월초||'', data.월말||'', data.입사||'', data.자발||'', data.비자발||'', data.매출||''];
      sh.getRange(i + 1, 1, 1, vals.length).setValues([vals]);
      return;
    }
  }
  const lastRow = sh.getLastRow();
  sh.getRange(lastRow + 1, 1, 1, 11).setValues([[
    data.year, data.month, data.직접비||'', data.간접비1||'', data.간접비2||'',
    data.월초||'', data.월말||'', data.입사||'', data.자발||'', data.비자발||'', data.매출||''
  ]]);
}

// ════════════════════════════════════════════════════════════
//  조직도 읽기 — 연봉총액 필드 추가 ★
// ════════════════════════════════════════════════════════════
function readOrgStructured(ss) {
  const sh   = getOrCreate(ss, SHEET_ORG);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  let hdrIdx = -1, hdr = [];
  for (let i = 0; i < Math.min(3, data.length); i++) {
    const row = data[i].map(c => String(c).trim());
    if (row.some(c => c.includes('닉네임') || c.includes('한국이름') || c === '성명')) {
      hdrIdx = i; hdr = row; break;
    }
  }
  if (hdrIdx < 0) return [];

  const fi = (...kws) => {
    for (const kw of kws) {
      const idx = hdr.findIndex(h => h === kw || h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const col = {
    사번:       fi('사번'),
    닉네임:     fi('닉네임'),
    성명:       fi('한국이름', '성명'),
    재직여부:   fi('재직상태', '재직여부'),
    본부:       fi('본부'),
    본부부서:   fi('부서'),
    팀:         fi('팀'),
    직책:       fi('직책'),
    직급:       fi('레벨', '직급'),
    입사일:     fi('입사일'),
    월고정만:   fi('월고정총액', '월고정'),   // ★ 컬럼명 우선순위 조정
    연봉총액만: fi('연봉총액'),               // ★ 신규
    status:     fi('status', '신호등'),
    asanaUrl:   fi('asana', 'Asana'),
    profileUrl: fi('인사노트', 'profile'),
  };

  const parseNum = v => parseFloat(String(v).replace(/[,\s₩원]/g,'')) || 0;
  const parseLevel = v => {
    if (!v) return '';
    v = String(v).trim();
    if (v.includes('Level') || v === 'CEO') return v;
    if (/^\d+$/.test(v)) return 'Level ' + v;
    return v;
  };
  const parseDt = v => {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return String(v).replace(/\.\s*/g,'-').replace(/-+$/,'').trim();
  };
  const g = (row, c) => c >= 0 ? String(row[c] || '').trim() : '';

  return data.slice(hdrIdx + 1).map(row => ({
    사번:       g(row, col.사번),
    닉네임:     g(row, col.닉네임),
    성명:       g(row, col.성명),
    재직여부:   g(row, col.재직여부) || '재직',
    본부:       g(row, col.본부) || '기타',
    본부부서:   g(row, col.본부부서) || g(row, col.본부),
    팀:         g(row, col.팀),
    직책:       g(row, col.직책),
    직급:       parseLevel(g(row, col.직급)),
    입사일:     parseDt(col.입사일 >= 0 ? row[col.입사일] : ''),
    월고정만:   col.월고정만   >= 0 ? parseNum(row[col.월고정만])   : 0,
    연봉총액만: col.연봉총액만 >= 0 ? parseNum(row[col.연봉총액만]) : 0,  // ★ 신규
    status:     g(row, col.status),
    asanaUrl:   g(row, col.asanaUrl),
    profileUrl: g(row, col.profileUrl),
  })).filter(r => (r.닉네임 || r.성명).trim());
}

function readOrgRaw(ss) {
  const sh  = getOrCreate(ss, SHEET_ORG);
  const v   = sh.getRange('A1').getValue();
  const str = String(v || '').trim();
  return (str.includes('\t') && str.split('\n').length > 2) ? str : '';
}

function saveOrg(ss, raw, emps) {
  const sh = getOrCreate(ss, SHEET_ORG);
  sh.clearContents();
  if (raw) sh.getRange('A1').setValue(raw);
  if (emps && emps.length > 0) {
    const keys = ['사번','닉네임','성명','재직여부','본부','본부부서','팀','직책','직급',
                  '입사일','월고정만','status','asanaUrl','profileUrl'];
    sh.getRange(1, 3, 1, keys.length).setValues([keys]);
    sh.getRange(2, 3, emps.length, keys.length).setValues(emps.map(e => keys.map(k => e[k] || '')));
  }
}

// ════════════════════════════════════════════════════════════
//  설정값
// ════════════════════════════════════════════════════════════
function readConfig(ss, key, defaultVal) {
  const sh   = getOrCreate(ss, SHEET_CFG);
  const data = sh.getDataRange().getValues();
  for (const row of data) {
    if (String(row[0]).trim() === key) {
      const v = row[1]; return (v !== '' && v !== null) ? Number(v) : defaultVal;
    }
  }
  return defaultVal;
}

function saveConfig(ss, key, value) {
  const sh   = getOrCreate(ss, SHEET_CFG);
  const data = sh.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) { sh.getRange(i+1,2).setValue(value); return; }
  }
  const last = sh.getLastRow();
  sh.getRange(last+1,1).setValue(key); sh.getRange(last+1,2).setValue(value);
}

// ════════════════════════════════════════════════════════════
//  부서별 예산 읽기 ★
//  설정 시트에서 'budget_부서명' 키로 된 행을 모두 읽어 반환
//  예: budget_CEO → 3115 (만원 단위)
// ════════════════════════════════════════════════════════════
function readDeptBudgets(ss) {
  const sh   = getOrCreate(ss, SHEET_CFG);
  const data = sh.getDataRange().getValues();
  const result = {};
  for (const row of data) {
    const key = String(row[0]).trim();
    if (key.startsWith('budget_')) {
      const val = parseFloat(String(row[1]).replace(/[,\s₩원]/g,''));
      if (!isNaN(val)) result[key] = val;
    }
  }
  return result;
}

// ════════════════════════════════════════════════════════════
//  부서별 AI 사용 비용 읽기 ★ v3.1 신규 (Claude 한정)
//  설정 시트에서 'ai_본부명' 키로 된 행을 모두 읽어 반환
//  예: ai_BOOST 본부 → 45 (만원 단위, 월 기준)
//  ※ 키의 본부명은 [조직도] 시트 '본부' 컬럼 값과 정확히 일치해야 함
//  ※ 인건비 예산(budget_)과 완전히 분리된 별도 객체로 반환
// ════════════════════════════════════════════════════════════
function readDeptAiCost(ss) {
  const sh   = getOrCreate(ss, SHEET_CFG);
  const data = sh.getDataRange().getValues();
  const result = {};
  for (const row of data) {
    const key = String(row[0]).trim();
    if (key.startsWith('ai_')) {
      const val = parseFloat(String(row[1]).replace(/[,\s₩원]/g,''));
      if (!isNaN(val)) result[key.substring(3)] = val;  // 'ai_' 접두어 제거 → 본부명만 키로
    }
  }
  return result;
}

// ════════════════════════════════════════════════════════════
//  부서별 확장 월비용 항목 읽기 ★ v3.6 신규
//  [설정] 시트 '월비용_항목명_본부명' 키 (만원/월 단위)
//  예: 월비용_야근식대_BOOST 본부 → 35
//      월비용_법인카드_Experience 부서 → 120
//  → 항목명은 자유 — 새 비용 항목 추가 시 코드 수정 없이 설정 키만 추가
//  ※ 본부명은 [조직도] '본부' 값과 일치 (공백 차이는 프론트에서 무시 매칭)
// ════════════════════════════════════════════════════════════
function readDeptMonthlyCosts(ss) {
  const sh   = getOrCreate(ss, SHEET_CFG);
  const data = sh.getDataRange().getValues();
  const result = {};
  for (const row of data) {
    const key = String(row[0]).trim();
    if (!key.startsWith('월비용_')) continue;
    const rest = key.substring(4); // '월비용_' 제거 → '항목명_본부명'
    const sep  = rest.indexOf('_');
    if (sep <= 0) continue;
    const item = rest.substring(0, sep).trim();
    const bu   = rest.substring(sep + 1).trim();
    if (!item || !bu) continue;
    const val = parseFloat(String(row[1]).replace(/[,\s₩원]/g,''));
    if (isNaN(val)) continue;
    if (!result[bu]) result[bu] = {};
    result[bu][item] = val;
  }
  return result;
}

// ════════════════════════════════════════════════════════════
//  퇴직자 누적 관리 읽기 ★ v3.2 신규
//  [퇴직자] 시트: 사번|닉네임|성명|본부|팀|입사일|퇴직일|퇴직구분|
//                 퇴사사유|상세사유|퇴직평가|재입사가능여부|비고
//  헤더는 1~3행 내 자동 탐지, 컬럼 순서 무관 (이름 매칭)
// ════════════════════════════════════════════════════════════
function readLeavers(ss) {
  const sh = ss.getSheetByName(SHEET_LEAVERS);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  let hdrIdx = -1, hdr = [];
  for (let i = 0; i < Math.min(3, data.length); i++) {
    const row = data[i].map(c => String(c).trim());
    if (row.some(c => c.includes('퇴직일') || c.includes('퇴사일'))) {
      hdrIdx = i; hdr = row; break;
    }
  }
  if (hdrIdx < 0) return [];

  const fi = (...kws) => {
    for (const kw of kws) {
      const idx = hdr.findIndex(h => h === kw);
      if (idx >= 0) return idx;
    }
    for (const kw of kws) {
      const idx = hdr.findIndex(h => h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const col = {
    사번:     fi('사번'),
    닉네임:   fi('닉네임'),
    성명:     fi('한국이름', '성명', '이름'),
    본부:     fi('본부'),
    팀:       fi('팀', '부서'),
    직책:     fi('직책'),
    입사일:   fi('입사일'),
    퇴직일:   fi('퇴직일', '퇴사일'),
    퇴직구분: fi('퇴직구분', '구분'),
    퇴사사유: fi('퇴사사유', '사유'),
    상세사유: fi('상세사유', '상세'),
    퇴직평가: fi('퇴직평가', '평가'),
    재입사:   fi('재입사가능여부', '재입사'),
    비고:     fi('비고'),
  };

  const parseDt = v => {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return String(v).trim();
  };
  const g = (row, c) => c >= 0 ? String(row[c] || '').trim() : '';

  return data.slice(hdrIdx + 1).map(row => ({
    사번:          g(row, col.사번),
    닉네임:        g(row, col.닉네임),
    성명:          g(row, col.성명),
    본부:          g(row, col.본부),
    팀:            g(row, col.팀),
    직책:          g(row, col.직책),
    입사일:        parseDt(col.입사일 >= 0 ? row[col.입사일] : ''),
    퇴직일:        parseDt(col.퇴직일 >= 0 ? row[col.퇴직일] : ''),
    퇴직구분:      g(row, col.퇴직구분),
    퇴사사유:      g(row, col.퇴사사유),
    상세사유:      g(row, col.상세사유),
    퇴직평가:      g(row, col.퇴직평가),
    재입사가능여부: g(row, col.재입사),
    비고:          g(row, col.비고),
  })).filter(r => r.퇴직일 && (r.닉네임 || r.성명));
}

// ════════════════════════════════════════════════════════════
//  수습 목표 개인진단 설문 읽기 ★ v3.3 신규
//  [설문_개인진단] 시트: 사번 컬럼 기준으로 해당 인원의 응답을 반환.
//  질문 컬럼 구성이 바뀌어도 동작하는 범용 구조 —
//  헤더행을 그대로 질문으로 사용해 [질문, 답변] 쌍 목록으로 반환.
//  (사번/타임스탬프/이메일/이름 등 메타 컬럼은 자동 제외)
// ════════════════════════════════════════════════════════════
function readProbSurvey(ss, sabun) {
  if (!sabun) return [];
  const sh = ss.getSheetByName(SHEET_PROB_SURVEY);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  // 사번 미입력 응답 대비 폴백 매칭용: 조직도에서 해당 사번의 성명/닉네임 조회
  let empName = '', empNick = '';
  try {
    const emp = readOrgStructured(ss).find(x => String(x.사번 || '').trim() === String(sabun).trim());
    if (emp) { empName = String(emp.성명 || '').trim(); empNick = String(emp.닉네임 || '').trim(); }
  } catch (e) {}

  // 헤더행 탐지 (1~3행 내 '사번' 포함 행)
  let hdrIdx = -1, hdr = [];
  for (let i = 0; i < Math.min(3, data.length); i++) {
    const row = data[i].map(c => String(c).trim());
    if (row.some(c => c === '사번' || c.includes('사번'))) { hdrIdx = i; hdr = row; break; }
  }
  if (hdrIdx < 0) return [];

  const sabunCol = hdr.findIndex(h => h === '사번' || h.includes('사번'));
  const tsCol = hdr.findIndex(h => h.includes('타임스탬프') || h.includes('제출일') || h.toLowerCase().includes('timestamp'));

  // 메타 컬럼(질문 아님) 판별 — 설문 응답 시트의 신원/시스템 컬럼은 Q&A에서 제외
  const isMeta = h => !h
    || h === '사번' || h.includes('사번')
    || h.includes('타임스탬프') || h.includes('제출일') || h.toLowerCase().includes('timestamp')
    || h.toLowerCase().includes('이메일') || h.toLowerCase().includes('email')
    || h === '이름' || h.includes('성명') || h.includes('닉네임')
    || h.includes('입사일')
    || h.includes('함수') || h.includes('자동세팅') || h.includes('공유용');

  const parseDt = v => {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return String(v).trim();
  };

  const nameCol = hdr.findIndex(h => h.includes('성명') || h === '이름');
  const nickCol = hdr.findIndex(h => h.includes('닉네임'));

  const out = [];
  data.slice(hdrIdx + 1).forEach(row => {
    // ① 사번 일치 우선 → ② 사번 미입력 행은 성명/닉네임 포함 매칭 폴백
    const rowSabun = String(row[sabunCol] !== undefined && row[sabunCol] !== null ? row[sabunCol] : '').trim();
    let hit = rowSabun !== '' && rowSabun === String(sabun).trim();
    if (!hit && rowSabun === '' && (empName || empNick)) {
      const rn = nameCol >= 0 ? String(row[nameCol] || '') : '';
      const rk = nickCol >= 0 ? String(row[nickCol] || '') : '';
      hit = (empName && rn.indexOf(empName) >= 0) || (empNick && rk.indexOf(empNick) >= 0);
    }
    if (!hit) return;
    const items = [];
    hdr.forEach((h, c) => {
      if (isMeta(h)) return;
      const v = String(row[c] !== undefined && row[c] !== null ? row[c] : '').trim();
      if (v) items.push([h, v]);
    });
    if (items.length) out.push({
      제출일: tsCol >= 0 ? parseDt(row[tsCol]) : '',
      항목: items,
    });
  });
  return out;
}

// ════════════════════════════════════════════════════════════
//  임박 이벤트 읽기 ★ v3.5 신규
//  [인사카드_이벤트알림] 시트(HrCard_Setup.gs의 refreshHrCardEvents가
//  매일 자동 갱신)를 그대로 읽어 조직도 탭 "임박 이벤트" 섹션에 공급.
//  상태 행(알림없음)과 설명행은 제외.
// ════════════════════════════════════════════════════════════
function readHrEvents(ss) {
  const sh = ss.getSheetByName('인사카드_이벤트알림');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  let hdrIdx = -1, hdr = [];
  for (let i = 0; i < Math.min(3, data.length); i++) {
    const row = data[i].map(c => String(c).trim());
    if (row.some(c => c === '사번' || c.includes('사번'))) { hdrIdx = i; hdr = row; break; }
  }
  if (hdrIdx < 0) return [];

  const fi = (...kws) => {
    for (const kw of kws) {
      const idx = hdr.findIndex(h => h === kw || h.includes(kw));
      if (idx >= 0) return idx;
    }
    return -1;
  };
  const col = {
    사번: fi('사번'), 닉네임: fi('닉네임'), 유형: fi('알림유형'),
    기준일: fi('기준일자'), dday: fi('D_day', 'D-day'), 표시: fi('표시여부'),
    배지: fi('배지텍스트'), 갱신: fi('최종갱신일시'),
  };
  const parseDt = v => {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy.MM.dd');
    return String(v).trim();
  };
  const g = (row, c) => c >= 0 ? String(row[c] !== undefined && row[c] !== null ? row[c] : '').trim() : '';

  return data.slice(hdrIdx + 1).map(row => ({
    사번:   g(row, col.사번),
    닉네임: g(row, col.닉네임),
    유형:   g(row, col.유형),
    기준일: parseDt(col.기준일 >= 0 ? row[col.기준일] : ''),
    dday:   (() => { const v = parseFloat(g(row, col.dday)); return isNaN(v) ? null : v; })(),
    배지:   g(row, col.배지),
    갱신:   g(row, col.갱신),
  })).filter(r => r.사번 && /\d/.test(r.사번) && r.배지); // 설명행·상태행 제외
}

// ════════════════════════════════════════════════════════════
//  수습 직원 현황 자동 계산 ★ v3.5 신규 / v3.5.1 필터 변경
//  대상: 재직자(퇴직예정 포함) 중 [인사카드_기본추가정보] 고용형태='수습' 인원
//        → 수습 통과 시 고용형태를 '정규직'으로 바꾸면 목록에서 자동 제외
//  수습종료일: [인사카드_기본추가정보] '수습종료일' 우선, 없으면 입사일+3개월
//  집계: 경과 주차, 수습면담/수습평가 기록 수, 개인진단 설문 제출 수
//  → 별도 수기 관리 시트 없이 대시보드에서 자동 갱신되는 현황판
// ════════════════════════════════════════════════════════════
function readProbation(ss) {
  const today = new Date(); today.setHours(0,0,0,0);

  const toDt = v => {
    if (!v && v !== 0) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    let s = String(v).trim().replace(/[\-\/]/g,'.').replace(/\.\s*/g,'.').replace(/\.$/,'').replace(/\s/g,'');
    const p = s.split('.');
    if (p.length < 3) return null;
    const d = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
    return isNaN(d.getTime()) ? null : d;
  };
  const fmt = d => d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy.MM.dd') : '';

  // 사번별 행 수 집계 헬퍼 (헤더 1~3행 자동 탐지)
  const countBySabun = (sheetName) => {
    const map = {};
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return map;
    const data = sh.getDataRange().getValues();
    let hdrIdx = -1, sabunCol = -1;
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const row = data[i].map(c => String(c).trim());
      const idx = row.findIndex(h => h === '사번' || h.includes('사번'));
      if (idx >= 0) { hdrIdx = i; sabunCol = idx; break; }
    }
    if (hdrIdx < 0) return map;
    for (let i = hdrIdx + 1; i < data.length; i++) {
      const s = String(data[i][sabunCol] || '').trim();
      if (!s || !/\d/.test(s)) continue;
      map[s] = (map[s] || 0) + 1;
    }
    return map;
  };

  const intCnt  = countBySabun('인사카드_수습면담');
  const evalCnt = countBySabun('인사카드_수습평가');
  const survCnt = countBySabun('설문_개인진단');

  // 수습평가 최종결과 (사번별, 아래 행일수록 최신으로 간주) ★ v3.7
  const evalResMap = {};
  const esh = ss.getSheetByName('인사카드_수습평가');
  if (esh) {
    const ed = esh.getDataRange().getValues();
    let eIdx = -1, eHdr = [];
    for (let i = 0; i < Math.min(3, ed.length); i++) {
      const row = ed[i].map(c => String(c).trim());
      if (row.some(c => c === '사번' || c.includes('사번'))) { eIdx = i; eHdr = row; break; }
    }
    if (eIdx >= 0) {
      const sc = eHdr.findIndex(h => h === '사번' || h.includes('사번'));
      const rc = eHdr.findIndex(h => h.includes('최종결과'));
      if (rc >= 0) {
        for (let i = eIdx + 1; i < ed.length; i++) {
          const s = String(ed[i][sc] || '').trim();
          if (!s || !/\d/.test(s)) continue;
          const r = String(ed[i][rc] || '').trim();
          if (r) evalResMap[s] = r;
        }
      }
    }
  }

  // 기본추가정보 → 수습종료일 + 고용형태 맵
  const probEndMap = {};
  const empTypeMap = {};
  const bs = ss.getSheetByName('인사카드_기본추가정보');
  if (bs) {
    const bd = bs.getDataRange().getValues();
    let hIdx = -1, hdr = [];
    for (let i = 0; i < Math.min(3, bd.length); i++) {
      const row = bd[i].map(c => String(c).trim());
      if (row.some(c => c === '사번' || c.includes('사번'))) { hIdx = i; hdr = row; break; }
    }
    if (hIdx >= 0) {
      const sc = hdr.findIndex(h => h === '사번' || h.includes('사번'));
      const pc = hdr.findIndex(h => h.includes('수습종료일'));
      const ec = hdr.findIndex(h => h.includes('고용형태'));
      for (let i = hIdx + 1; i < bd.length; i++) {
        const s = String(bd[i][sc] || '').trim();
        if (!s || !/\d/.test(s)) continue;
        if (pc >= 0) { const d = toDt(bd[i][pc]); if (d) probEndMap[s] = d; }
        if (ec >= 0) empTypeMap[s] = String(bd[i][ec] || '').trim();
      }
    }
  }

  // 조직도 재직자 순회 — 고용형태 '수습'인 인원만 (v3.5.1)
  const out = [];
  readOrgStructured(ss).forEach(e => {
    const sabun = String(e.사번 || '').trim();
    if (!sabun) return;
    const st = String(e.재직여부 || '재직');
    if (st.indexOf('퇴') >= 0 && st.indexOf('예정') < 0) return;

    // ★ [인사카드_기본추가정보] 고용형태에 '수습' 표기된 인원만 목록에 포함
    const empType = empTypeMap[sabun] || '';
    if (empType.indexOf('수습') < 0) return;

    const joinD = toDt(e.입사일);
    let probEnd = probEndMap[sabun] || null;
    if (!probEnd && joinD) {
      probEnd = new Date(joinD);
      probEnd.setMonth(probEnd.getMonth() + 3);
    }
    if (!probEnd) return;

    const dday = Math.ceil((probEnd - today) / 86400000);
    // 고용형태='수습'이 목록 기준이므로 날짜로 자동 제외하지 않음
    // (종료일 경과했는데 '수습'이면 = 연장 중이거나 전환 처리 누락 → 오히려 보여야 함)

    const weeks = joinD ? Math.floor((today - joinD) / (86400000 * 7)) + 1 : null;

    out.push({
      사번:     sabun,
      닉네임:   e.닉네임 || '',
      성명:     e.성명 || '',
      본부:     e.본부 || '',
      부서:     e.본부부서 || '',
      팀:       e.팀 || '',
      입사일:   e.입사일 || '',
      수습종료일: fmt(probEnd),
      dday:     dday,
      주차:     weeks,
      면담수:   intCnt[sabun]  || 0,
      평가수:   evalCnt[sabun] || 0,
      평가결과: evalResMap[sabun] || '',
      설문수:   survCnt[sabun] || 0,
    });
  });

  out.sort((a, b) => a.dday - b.dday); // 종료 임박 순
  return out;
}

// ════════════════════════════════════════════════════════════
//  외부 공개용 조직도
// ════════════════════════════════════════════════════════════
function readExtOrg(ss) {
  const sh = getOrCreate(ss, SHEET_EXT_ORG);
  const data = sh.getDataRange().getValues();
  if (data.length < 3) return [];
  const hdr = data[1].map(c => String(c).trim());
  const fi  = k => hdr.findIndex(h => h === k || h.includes(k));
  const col = { id:fi('id'), level:fi('level'), parent:fi('parent_id'),
                name:fi('조직명'), count:fi('인원'), title:fi('리더 직함'),
                leader:fi('리더 이름'), note:fi('비고'), side:fi('side','위치') };
  const g = (row, c) => c >= 0 ? String(row[c] || '').trim() : '';
  return data.slice(2).map(row => ({
    id: +g(row,col.id)||0, level: +g(row,col.level)||1,
    parent_id: g(row,col.parent) ? +g(row,col.parent) : null,
    name: g(row,col.name), count: +g(row,col.count)||0,
    leader_title: g(row,col.title), leader: g(row,col.leader), note: g(row,col.note),
    side: g(row,col.side),
  })).filter(r => r.id && r.name);
}

// ════════════════════════════════════════════════════════════
//  공통 유틸
// ════════════════════════════════════════════════════════════
function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
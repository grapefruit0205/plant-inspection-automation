/**
 * 현장 점검표 자동 집계·주간보고 시스템 — 통합 설치 파일
 * ------------------------------------------------------------
 * 사용법:
 *   1) 구글시트(sheets.new)를 새로 만들고 이름을 '점검시스템_포트폴리오'로 지정
 *   2) 확장 프로그램 → Apps Script → 기본 Code.gs 내용을 전부 지우고 이 파일 전체를 붙여넣기
 *   3) 저장(Ctrl+S) → 함수 선택에서 '설치_전체' 선택 → 실행 → 권한 허용
 *   4) (선택) '설치_샘플데이터' 실행 → 4주치 합성 데이터 생성
 *
 * 이후에는 시트 상단 [점검시스템] 메뉴에서 실행합니다.
 *
 * 데이터 고지: 전부 합성 데이터입니다. 실제 회사 자료를 쓰지 않았습니다.
 */

/* ============================================================
 *  1부 — 본체 (판정·알림·집계·PDF)
 * ============================================================ */
/**
 * 현장 점검표 자동 집계·주간보고 시스템
 * ------------------------------------------------------------
 * 실행 환경: Google Sheets + Apps Script (무료 계정)
 * 데이터: 전부 합성(가상) 데이터. 실제 회사 자료 사용 금지.
 *
 * 트리거 대상 함수는 영문명으로 둔다(트리거 목록에서 한글이 깨질 수 있음).
 *   - onFormSubmit  : 양식 제출 시
 *   - weeklyReport  : 시간 기반(월요일 오전 7시)
 *
 * 시트 탭: 원본응답 / 기준값 / 설비목록 / 일일집계 / 이상이력 / 주간요약 / 설정
 */

const TZ = 'Asia/Seoul';

const SH = {
  원본: '원본응답',
  기준: '기준값',
  설비: '설비목록',
  일일: '일일집계',
  이력: '이상이력',
  주간: '주간요약',
  설정: '설정',
};

/* ============================ 공통 도구 ============================ */

function _시트(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('탭을 찾을 수 없습니다: ' + name);
  return sh;
}

function _헤더(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

function _열(sh, name) {
  const i = _헤더(sh).indexOf(name);
  if (i < 0) throw new Error('헤더 없음: "' + name + '" in ' + sh.getName());
  return i + 1;
}

/**
 * 타임스탬프 열 위치(0부터). 시트 로케일에 따라 머리글이 '타임스탬프'/'Timestamp' 로
 * 갈리므로 둘 다 보고, 그래도 없으면 첫 열을 쓴다(폼 응답 시트는 첫 열이 제출 시각).
 */
function _타임스탬프열(h) {
  const i = h.indexOf('타임스탬프');
  if (i !== -1) return i;
  const j = h.indexOf('Timestamp');
  return j === -1 ? 0 : j;
}

/**
 * 응답 한 행을 헤더 이름으로 읽는다.
 * 폼 문항을 다시 만들면 같은 문항 열이 시트에 새로 추가될 수 있어(이름 중복),
 * 그때 비어 있지 않은 첫 값을 쓴다. 이 처리가 없으면 뒤쪽 빈 열을 읽어
 * 항상 "정상" 으로 판정된다.
 */
function _응답만들기(h, r) {
  const 응답 = {};
  h.forEach((name, j) => {
    const v = r[j];
    const 값있음 = v !== '' && v !== null && v !== undefined;
    if (!(name in 응답)) 응답[name] = 값있음 ? v : '';
    else if (값있음 && String(응답[name]).trim() === '') 응답[name] = v;
  });
  return 응답;
}

/**
 * 점검일 셀 값을 'yyyy-MM-dd' 문자열로 맞춘다.
 * 시트가 날짜로 바꿔 저장하면 Date 객체가 되고, 문자열이면 '2026-09-14 08:23:41'
 * 처럼 시각이 붙을 수 있어, 그대로 두면 날짜 비교('2026-09-14' <= d)가 어긋난다.
 */
function _날짜문자열(raw) {
  if (raw === null || raw === undefined) return '';
  if (Object.prototype.toString.call(raw) === '[object Date]') return _날짜(raw);
  const s = String(raw).trim();
  return s.length > 10 ? s.slice(0, 10) : s;
}

function 설정(key) {
  const v = _시트(SH.설정).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) return v[i][1];
  }
  throw new Error('설정 탭에 없는 키: ' + key);
}

function _날짜(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function _오늘0시_기준(d) {
  const s = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return new Date(s + 'T00:00:00+09:00');
}

/* ============================ 1. 기준값 읽기 ============================ */

function 기준값읽기() {
  const sh = _시트(SH.기준);
  const v = sh.getDataRange().getValues();
  const h = v.shift().map(String);
  const i = (k) => h.indexOf(k);

  return v
    .filter((r) => String(r[0]).trim() !== '')
    .map((r) => ({
      항목명: String(r[i('항목명')] || '').trim(),
      컬럼명: String(r[i('원본컬럼명')] || '').trim(),
      유형: String(r[i('판정유형')] || '').trim(),
      기준: r[i('기준값')],
      이상값: String(r[i('이상값목록')] || '').trim(),
      심각도: String(r[i('심각도')] || '').trim(),
      메일: String(r[i('담당자이메일')] || '').trim(),
      사용여부: String(r[i('사용여부')] || '').toUpperCase() !== 'FALSE',
    }))
    .filter((k) => k.항목명 && k.컬럼명);
}

/* ============================ 2. 이상치 판정 ============================ */

function 이상치판정(응답, 기준목록) {
  const 결과 = [];
  기준목록
    .filter((k) => k.사용여부)
    .forEach((k) => {
      const raw = 응답[k.컬럼명];
      const v = raw === undefined || raw === null ? '' : String(raw).trim();
      if (v === '') return; // 빈 값은 정상 처리

      let 이상 = false;
      if (k.유형 === '초과') {
        이상 = Number(v) > Number(k.기준);
      } else if (k.유형 === '미만') {
        이상 = Number(v) < Number(k.기준);
      } else if (k.유형 === '값일치') {
        const 목록 = k.이상값.split(',').map((s) => s.trim()).filter(String);
        이상 = 목록.indexOf(v) > -1;
      }

      if (이상) {
        결과.push({
          항목명: k.항목명,
          측정값: v,
          기준: k.유형 === '값일치' ? k.이상값 : String(k.기준),
          심각도: k.심각도,
          메일: k.메일,
        });
      }
    });
  return 결과;
}

/** 개발용: 편집기에서 실행해 판정 로직만 확인 */
function 테스트_판정() {
  const 샘플 = {
    '압축기 토출 압력(bar)': '7.5',
    '오일 레벨': '보충필요',
    '펌프 진동·소음': '정상',
    '베어링 온도(℃)': '62',
    '탱크 액위(%)': '55',
    '배관 누설': '없음',
    '밸브 잠금 상태': '정상',
    '안전 커버·방호': '정상',
    '윤활 급유': '완료',
  };
  Logger.log(JSON.stringify(이상치판정(샘플, 기준값읽기()), null, 2));
}

/* ============================ 3. 이상이력 기록 ============================ */

function 이상이력기록(응답, 이상목록, 기준일) {
  if (!이상목록.length) return;
  const sh = _시트(SH.이력);
  const h = _헤더(sh);
  const now = 기준일 || new Date();

  const rows = 이상목록.map((x) =>
    h.map((name) => {
      switch (name) {
        case '발생시각':
          return Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss');
        case '점검일':
          return _날짜(now);
        case '설비':
          return 응답['설비'] || '';
        case '점검자':
          return 응답['점검자'] || '';
        case '항목명':
          return x.항목명;
        case '측정값':
          return x.측정값;
        case '기준':
          return x.기준;
        case '심각도':
          return x.심각도;
        case '알림발송':
          return 'Y';
        default:
          return '';
      }
    })
  );

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, h.length).setValues(rows);
}

/* ============================ 4. 메일 알림 ============================ */

function 메일알림(응답, 이상목록, 기준일) {
  if (!이상목록.length) return 0;
  if (String(설정('알림활성화')).toUpperCase() !== 'TRUE') return 0;

  // 메일테스트모드: 부서 주소 대신 관리자에게만 보내고, 제목에 원래 수신자를 남긴다.
  // (기준값 탭의 담당자 주소가 자리표시자일 때 데모·검증용)
  const 테스트모드 = String(설정('메일테스트모드')).toUpperCase() === 'TRUE';
  const 관리자 = String(설정('관리자이메일'));

  // 담당자별 묶음
  const 묶음 = {};
  이상목록.forEach((x) => {
    const to = x.메일 || 관리자;
    (묶음[to] = 묶음[to] || []).push(x);
  });

  const 설비 = 응답['설비'] || '';
  const 점검자 = 응답['점검자'] || '';
  let 발송 = 0;

  Object.keys(묶음).forEach((to) => {
    const 목록 = 묶음[to];
    const 받는사람 = 테스트모드 ? 관리자 : to;
    if (!받는사람) return;   // 수신 주소를 못 읽으면 건너뜀
    const 제목 = (테스트모드 ? '[테스트→' + to + '] ' : '') +
      '[설비이상] ' + 설비 + ' ' + 목록[0].항목명 + (목록.length > 1 ? ' 외 ' + (목록.length - 1) + '건' : '');
    const 행 = 목록
      .map(
        (x) =>
          '<tr><td>' + x.항목명 + '</td><td>' + x.측정값 + '</td><td>' + x.기준 +
          '</td><td>' + x.심각도 + '</td></tr>'
      )
      .join('');
    const 본문 =
      '<p><b>' + 설비 + '</b> 점검에서 기준 초과 항목이 발견되었습니다.</p>' +
      '<p>점검자: ' + 점검자 + ' / 점검일시: ' + Utilities.formatDate(기준일 || new Date(), TZ, 'yyyy-MM-dd HH:mm') + '</p>' +
      '<table border="1" cellpadding="6" cellspacing="0"><tr><th>항목</th><th>측정값</th><th>기준</th><th>심각도</th></tr>' +
      행 +
      '</table><p>조치 후 이상이력 탭의 조치상태를 갱신해 주세요.</p>';

    GmailApp.sendEmail(받는사람, 제목, 제목, { htmlBody: 본문 });
    발송++;
  });

  return 발송;
}

/* ============================ 5. 폼 제출 진입점 ============================ */

function onFormSubmit(e) {
  if (!e || !e.range) {
    throw new Error('이 함수는 폼 제출로만 실행됩니다. 편집기에서 테스트하려면 테스트_판정()을 쓰세요.');
  }

  // 이벤트의 namedValues 대신 시트에 기록된 그 행을 직접 읽는다.
  // 폼 문항을 다시 만들면 같은 이름의 열이 늘어날 수 있고, 그때 이벤트 값은
  // 뒤쪽 빈 열에 가려질 수 있어서다. 행을 직접 읽으면 항상 실제 값이 잡힌다.
  const sh = _시트(SH.원본);
  const 행 = e.range.getRow();
  const h = _헤더(sh);
  const r = sh.getRange(행, 1, 1, h.length).getValues()[0];
  const 응답 = _응답만들기(h, r);

  const 기준목록 = 기준값읽기();
  const 이상목록 = 이상치판정(응답, 기준목록);

  if (이상목록.length) {
    이상이력기록(응답, 이상목록, new Date());
    메일알림(응답, 이상목록, new Date());
  }
  원본응답판정기록(행, 이상목록);
  일일집계갱신(new Date());
  Logger.log(행 + '행 판정: ' + (이상목록.length ? '이상 ' + 이상목록.length + '건' : '정상'));
}

/** 원본응답 탭 N열(판정결과), O열(이상항목수) 기입 */
function 원본응답판정기록(row, 이상목록) {
  const sh = _시트(SH.원본);
  const n = _열(sh, '판정결과');
  sh.getRange(row, n, 1, 2).setValues([[이상목록.length ? '이상' : '정상', 이상목록.length]]);
}

/* ============================ 6. 일일 집계 ============================ */

function 일일집계갱신(기준일) {
  const day = _날짜(기준일 || new Date());
  const sh = _시트(SH.원본);
  const v = sh.getDataRange().getValues();
  const h = v.shift().map(String);
  const c = {
    ts: _타임스탬프열(h),
    설비: h.indexOf('설비'),
    판정: h.indexOf('판정결과'),
  };

  const rows = v.filter((r) => r[c.ts] && _날짜(new Date(r[c.ts])) === day);
  const 점검설비 = {};
  const 이상설비 = {};
  const 항목수 = {};

  rows.forEach((r) => {
    const 설비 = String(r[c.설비]);
    점검설비[설비] = true;
    if (String(r[c.판정]) === '이상') 이상설비[설비] = true;
  });

  // 이상이력에서 그날 항목별 건수
  const ih = _시트(SH.이력).getDataRange().getValues();
  const eh = ih.shift().map(String);
  const e일 = eh.indexOf('점검일');
  const e항목 = eh.indexOf('항목명');
  ih.forEach((r) => {
    if (String(r[e일]) === day) {
      const a = String(r[e항목]);
      항목수[a] = (항목수[a] || 0) + 1;
    }
  });

  const 설비목록 = _시트(SH.설비).getDataRange().getValues().slice(1).map((r) => String(r[0]));
  const 미점검 = 설비목록.filter((x) => !점검설비[x]).length;
  const 최다 = Object.keys(항목수).sort((a, b) => 항목수[b] - 항목수[a])[0] || '';

  const 요약 = [
    day,
    rows.length,
    Object.keys(점검설비).length,
    미점검,
    Object.keys(이상설비).length,
    Object.keys(이상설비).join(', '),
    최다 ? 최다 + '(' + 항목수[최다] + ')' : '',
  ];

  const out = _시트(SH.일일);
  const ov = out.getDataRange().getValues();
  for (let i = 1; i < ov.length; i++) {
    if (_날짜(new Date(ov[i][0])) === day) {
      out.getRange(i + 1, 1, 1, 요약.length).setValues([요약]);
      return;
    }
  }
  out.appendRow(요약);
}

/* ============================ 7. 주간 요약 ============================ */

function 주간요약갱신(시작일) {
  const start = _오늘0시_기준(시작일);
  const end = new Date(start.getTime() + 6 * 86400000);
  const 시작 = _날짜(start);
  const 종료 = _날짜(end);

  const v = _시트(SH.원본).getDataRange().getValues();
  const h = v.shift().map(String);
  const ts = _타임스탬프열(h);
  const 설비열 = h.indexOf('설비');

  let 총제출 = 0;
  v.forEach((r) => {
    if (!r[ts]) return;
    const d = _날짜(new Date(r[ts]));
    if (d >= 시작 && d <= 종료) 총제출++;
  });

  const ih = _시트(SH.이력).getDataRange().getValues();
  const eh = ih.shift().map(String);
  const e일 = eh.indexOf('점검일');
  const e항목 = eh.indexOf('항목명');
  const e설비 = eh.indexOf('설비');
  const e조치 = eh.indexOf('조치상태');

  let 이상건수 = 0;
  let 미조치 = 0;
  const 항목카운트 = {};
  // TOP3 설비는 제출 횟수가 아니라 '이상이 난 횟수' 기준이어야 의미가 있다
  const 설비카운트 = {};
  ih.forEach((r) => {
    const d = _날짜문자열(r[e일]);
    if (d >= 시작 && d <= 종료) {
      이상건수++;
      const a = String(r[e항목]);
      항목카운트[a] = (항목카운트[a] || 0) + 1;
      const sf = String(r[e설비]);
      설비카운트[sf] = (설비카운트[sf] || 0) + 1;
      if (String(r[e조치] || '').trim() !== '완료') 미조치++;
    }
  });

  const top = (obj) =>
    Object.keys(obj)
      .sort((a, b) => obj[b] - obj[a])
      .slice(0, 3)
      .map((k) => k + '(' + obj[k] + ')')
      .join(', ');

  const 주차 = 시작 + ' ~ ' + 종료;
  const 요약 = {
    주차: 주차,
    기간: 주차,
    시작일: 시작,
    종료일: 종료,
    총제출: 총제출,
    이상건수: 이상건수,
    이상률: 총제출 ? ((이상건수 / 총제출) * 100).toFixed(1) : '0.0',
    TOP3설비: top(설비카운트),
    TOP3항목: top(항목카운트),
    미조치: 미조치,
    생성일시: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'),
  };

  const sh = _시트(SH.주간);
  const 행 = [
    요약.주차,
    요약.시작일,
    요약.종료일,
    요약.총제출,
    요약.이상건수,
    요약.이상률,
    요약.TOP3설비,
    요약.TOP3항목,
    요약.미조치,
    '',
  ];

  // 같은 주가 이미 있으면 새 행을 만들지 않고 갱신한다(PDF링크 열은 보존)
  const 기존 = sh.getDataRange().getValues();
  for (let i = 1; i < 기존.length; i++) {
    if (String(기존[i][0]) === 요약.주차) {
      sh.getRange(i + 1, 1, 1, 행.length).setValues([행]);
      return 요약;
    }
  }
  sh.appendRow(행);

  return 요약;
}

/* ============================ 8. 주간 PDF 보고서 ============================ */

/** 트리거 대상(영문명) */
function weeklyReport() {
  return 주간PDF생성();
}

function 지난주월요일() {
  const now = _오늘0시_기준(new Date());
  const day = Number(Utilities.formatDate(now, TZ, 'u')); // 월=1
  return new Date(now.getTime() - (day - 1 + 7) * 86400000);
}

/* 보고서 서식값. A4(595pt) - 좌우 여백 2cm(56.7pt×2) = 본문 폭 약 481pt 기준으로 열 너비를 잡는다 */
const 보고서양식 = {
  글꼴: 'Noto Sans KR',
  여백: 56.7,
  진남: '#1F3864',
  회색: '#666666',
  본문: '#202124',
  라벨배경: '#F1F3F4',
  빨강: '#C00000',
  음영: '#F8F9FA',
  선: '#DADCE0',
  정상배경: '#E2EFDA',
  정상글: '#274E13',
  바닥글: '#999999',
  심각도배경: { 상: '#F4CCCC', 중: '#FCE5CD', 하: '#EFEFEF' },
  최대행: 20,
};

/**
 * 기간 내 이상이력을 보고서용으로 읽는다.
 * 심각도(상→중→하)·점검일 순으로 정렬해, 20건에서 잘려도 중요한 것이 먼저 남게 한다.
 */
function _이상상세(시작일, 종료일) {
  const v = _시트(SH.이력).getDataRange().getValues();
  const h = v.shift().map(String);
  const 열 = {};
  ['점검일', '설비', '항목명', '측정값', '기준', '심각도', '조치상태'].forEach((k) => (열[k] = h.indexOf(k)));
  if (열.점검일 < 0) throw new Error('헤더 없음: "점검일" in ' + SH.이력);

  const 값 = (r, k) => (열[k] < 0 || r[열[k]] === null ? '' : String(r[열[k]]).trim());
  const 순서 = { 상: 0, 중: 1, 하: 2 };
  const 등급 = (s) => (s in 순서 ? 순서[s] : 9);

  const 목록 = [];
  v.forEach((r) => {
    // 시트가 날짜 문자열을 날짜형으로 바꿔 저장했을 수도 있어 둘 다 받는다
    const d = _날짜문자열(r[열.점검일]);
    if (!d || d < 시작일 || d > 종료일) return;
    목록.push({
      점검일: d,
      설비: 값(r, '설비'),
      항목: 값(r, '항목명'),
      측정값: 값(r, '측정값'),
      기준: 값(r, '기준'),
      심각도: 값(r, '심각도'),
      조치상태: 값(r, '조치상태') || '미조치',
    });
  });

  목록.sort((a, b) => 등급(a.심각도) - 등급(b.심각도) || (a.점검일 < b.점검일 ? -1 : a.점검일 > b.점검일 ? 1 : 0));

  // 심각도 분포 — 표는 심각도 순이라 상만 보이므로, 나머지 등급이 있다는 것을 한 줄로 보여준다
  const 분포 = { 상: 0, 중: 0, 하: 0 };
  목록.forEach((x) => {
    if (x.심각도 in 분포) 분포[x.심각도]++;
  });

  return { 행: 목록.slice(0, 보고서양식.최대행), 전체: 목록.length, 분포 };
}

function 주간PDF생성() {
  const 요약 = 주간요약갱신(지난주월요일());
  const 폴더ID = String(설정('PDF폴더ID'));
  const 관리자 = String(설정('관리자이메일'));
  const 상세 = _이상상세(요약.시작일, 요약.종료일);

  // 템플릿 사본 대신 문서를 코드로 새로 만든다(서식 통제 + 자리표시자 불일치 방지)
  const doc = DocumentApp.create('임시_' + 요약.주차);
  const 임시파일 = DriveApp.getFileById(doc.getId());
  let pdf, 파일;
  try {
    _보고서작성(doc, 요약, 상세);
    doc.saveAndClose();
    pdf = 임시파일.getAs('application/pdf').setName('주간설비점검보고_' + 요약.시작일 + '.pdf');
    파일 = DriveApp.getFolderById(폴더ID).createFile(pdf);
  } finally {
    // 중간에 실패해도 임시 문서가 드라이브에 쌓이지 않게
    try { 임시파일.setTrashed(true); } catch (e) { Logger.log('임시 문서 휴지통 이동 실패: ' + e.message); }
  }

  GmailApp.sendEmail(관리자, '[주간보고] ' + 요약.주차, '주간 설비 점검 보고서를 첨부합니다.', {
    attachments: [pdf],
  });

  // 주간요약 탭 PDF링크 기록(마지막 행)
  const sh = _시트(SH.주간);
  sh.getRange(sh.getLastRow(), _열(sh, 'PDF링크')).setValue(파일.getUrl());

  return 파일.getUrl();
}

/** 보고서 본문 구성. 서식은 실패해도 내용은 남도록 구간마다 _서식시도로 감싼다 */
function _보고서작성(doc, 요약, 상세) {
  const S = 보고서양식;
  const body = doc.getBody();
  _서식시도('여백', () => body.setMarginTop(S.여백).setMarginBottom(S.여백).setMarginLeft(S.여백).setMarginRight(S.여백));

  // 1) 머리글 — 새 문서는 빈 문단 하나로 시작하므로 그것을 제목으로 쓴다
  const 제목 = body.getParagraphs()[0] || body.appendParagraph('');
  제목.setText('주간 설비 점검 보고서');
  _글자(제목, 20, true, S.진남);
  _문단간격(제목, 0, 2);

  const 기간 = body.appendParagraph('기간 ' + 요약.시작일 + ' ~ ' + 요약.종료일);
  _글자(기간, 11, false, S.회색);
  _문단간격(기간, 0, 6);

  _구분선(body, S.진남);
  _간격(body, 10);

  // 2) 요약 카드
  const 이상 = Number(요약.이상건수) || 0;
  const 미조치 = Number(요약.미조치) || 0;
  const 카드 = _표(body, [
    ['총 제출', '이상 건수', '이상률', '미조치'],
    [String(요약.총제출) + '건', 이상 + '건', String(요약.이상률) + '%', 미조치 + '건'],
  ], [120, 120, 120, 120], 7);
  for (let c = 0; c < 4; c++) {
    _셀(카드.getCell(0, c), { 크기: 9, 색: S.회색, 배경: S.라벨배경, 가운데: true });
    const 경고 = (c === 1 && 이상 > 0) || (c === 3 && 미조치 > 0);
    _셀(카드.getCell(1, c), { 크기: 14, 굵게: true, 색: 경고 ? S.빨강 : S.본문, 가운데: true });
  }
  _간격(body, 8);

  // 3) TOP3 + 심각도 분포
  const 심각도줄 = '상 ' + 상세.분포.상 + '건   ·   중 ' + 상세.분포.중 + '건   ·   하 ' + 상세.분포.하 + '건';
  const 순위 = _표(body, [
    ['TOP3 설비', _top3표시(요약.TOP3설비)],
    ['TOP3 항목', _top3표시(요약.TOP3항목)],
    ['심각도 분포', 상세.전체 ? 심각도줄 : '—'],
  ], [96, 384], 5);
  for (let r = 0; r < 3; r++) {
    _셀(순위.getCell(r, 0), { 크기: 9, 색: S.회색, 배경: S.라벨배경 });
    _셀(순위.getCell(r, 1), { 크기: 10, 색: S.본문 });
  }

  // 4) 이상 상세
  const 소제목 = body.appendParagraph('▶ 이상 상세');
  _글자(소제목, 13, true, S.진남);
  _문단간격(소제목, 16, 6);

  if (!상세.전체) {
    // 5) 0건이면 빈 표 대신 안내
    const 안내 = _표(body, [['이번 주 기준 초과 항목이 없습니다.']], [480], 10, S.정상배경);
    _셀(안내.getCell(0, 0), { 크기: 10, 색: S.정상글, 배경: S.정상배경, 가운데: true });
  } else {
    const 머리 = ['점검일', '설비', '항목', '측정값', '기준', '심각도', '조치상태'];
    const 가운데열 = [0, 1, 5, 6];
    const 행들 = [머리].concat(상세.행.map((x) => [x.점검일, x.설비, x.항목, x.측정값, x.기준, x.심각도, x.조치상태]));
    const 표 = _표(body, 행들, [64, 46, 92, 90, 80, 44, 64], 4);

    머리.forEach((_, c) => _셀(표.getCell(0, c), { 크기: 9.5, 굵게: true, 색: '#FFFFFF', 배경: S.진남, 가운데: true }));
    상세.행.forEach((x, i) => {
      const r = i + 1;
      const 줄배경 = r % 2 === 0 ? S.음영 : null;
      머리.forEach((_, c) => {
        const o = { 크기: 9, 색: S.본문, 배경: 줄배경, 가운데: 가운데열.indexOf(c) > -1 };
        if (c === 5) { o.배경 = S.심각도배경[x.심각도] || 줄배경; o.굵게 = x.심각도 === '상'; }
        if (c === 6 && x.조치상태 === '미조치') { o.색 = S.빨강; o.굵게 = true; }
        _셀(표.getCell(r, c), o);
      });
    });

    // 5) 20건 초과분 안내. 셀 병합(merge)은 동작이 불확실해 표 바로 아래 문단으로 둔다
    if (상세.전체 > 상세.행.length) {
      const 외 = body.appendParagraph('외 ' + (상세.전체 - 상세.행.length) + '건 — 이상이력 탭 참조');
      _글자(외, 9, false, S.회색);
      _문단간격(외, 4, 0);
      _서식시도('외 N건 정렬', () => 외.setAlignment(DocumentApp.HorizontalAlignment.RIGHT));
    }
  }

  // 6) 바닥글 — 문서 바닥글 영역이 안 되면 본문 끝에 붙인다
  let 시트명 = '점검시스템_포트폴리오';
  try { 시트명 = SpreadsheetApp.getActiveSpreadsheet().getName() || 시트명; } catch (e) {}
  const 문구 = '생성 ' + 요약.생성일시 + ' · ' + 시트명 + ' · 데이터: 합성(가상)';
  let 바닥 = null;
  try {
    const f = doc.getFooter() || doc.addFooter();
    바닥 = f.getParagraphs()[0] || f.appendParagraph('');
    바닥.setText(문구);
  } catch (e) {
    Logger.log('[서식 건너뜀] 바닥글 영역: ' + e.message);
    바닥 = null;
  }
  if (!바닥) {
    바닥 = body.appendParagraph(문구);
    _문단간격(바닥, 18, 0);
  }
  _글자(바닥, 8, false, S.바닥글);
  _서식시도('바닥글 정렬', () => 바닥.setAlignment(DocumentApp.HorizontalAlignment.RIGHT));
}

/** 서식 한 구간 실행. 실패해도 보고서 생성은 계속한다 */
function _서식시도(이름, 작업) {
  try {
    작업();
  } catch (e) {
    Logger.log('[서식 건너뜀] ' + 이름 + ': ' + e.message);
  }
}

/** 문단·셀 글자 서식. 글꼴·크기는 따로 감싸 하나가 실패해도 굵기·색은 들어가게 한다 */
function _글자(el, 크기, 굵게, 색) {
  _서식시도('글자', () => {
    const t = el.editAsText();
    try { t.setFontFamily(보고서양식.글꼴); } catch (e) {}
    try { t.setFontSize(크기); } catch (e) { t.setFontSize(Math.round(크기)); } // 9.5 같은 소수 크기 거부 대비
    t.setBold(!!굵게);
    t.setForegroundColor(색 || 보고서양식.본문);
  });
}

function _문단간격(p, 앞, 뒤) {
  _서식시도('문단 간격', () => p.setSpacingBefore(앞).setSpacingAfter(뒤).setLineSpacing(1.15));
}

/** 표와 표 사이를 원하는 높이로 띄우는 빈 문단(기본 11pt 줄 높이를 없애려고 글자 크기를 1로) */
function _간격(body, pt) {
  const p = body.appendParagraph('');
  _서식시도('간격', () => {
    p.setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 1 });
    p.setSpacingBefore(0).setSpacingAfter(pt).setLineSpacing(1);
  });
  return p;
}

function _셀문단(cell) {
  return cell.getChild(0).asParagraph();
}

/** 표를 붙이고 공통 서식(테두리·열 너비·안쪽 여백)을 준다 */
function _표(body, 행들, 열너비, 안쪽, 테두리색) {
  const t = body.appendTable(행들.map((r) => r.map((x) => String(x === null || x === undefined ? '' : x))));
  const 여백 = 안쪽 || 4;
  const 셀마다 = (fn) => {
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) fn(row.getCell(c));
    }
  };
  _서식시도('표 테두리', () => { t.setBorderColor(테두리색 || 보고서양식.선); t.setBorderWidth(0.75); });
  _서식시도('표 열 너비', () => 열너비.forEach((w, i) => t.setColumnWidth(i, w)));
  _서식시도('셀 여백', () => 셀마다((cell) => cell.setPaddingTop(여백).setPaddingBottom(여백).setPaddingLeft(6).setPaddingRight(6)));
  _서식시도('셀 세로 정렬', () => 셀마다((cell) => cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER)));
  _서식시도('셀 문단 간격', () => 셀마다((cell) => _셀문단(cell).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.1)));
  return t;
}

/** 셀 하나 서식. o = { 크기, 굵게, 색, 배경, 가운데 } */
function _셀(cell, o) {
  _글자(cell, o.크기, o.굵게, o.색);
  if (o.배경) _서식시도('셀 배경', () => cell.setBackgroundColor(o.배경));
  if (o.가운데) _서식시도('셀 가로 정렬', () => _셀문단(cell).setAlignment(DocumentApp.HorizontalAlignment.CENTER));
}

/** 머리글 아래 굵은 줄. 문단 테두리는 DocumentApp에서 못 쓰니 색 채운 1칸 표로 대신한다 */
function _구분선(body, 색) {
  const t = body.appendTable([[' ']]);
  const cell = t.getCell(0, 0);
  _서식시도('구분선 색', () => { t.setBorderColor(색); cell.setBackgroundColor(색); });
  _서식시도('구분선 두께', () => {
    cell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
    _셀문단(cell).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    cell.editAsText().setFontSize(2);
  });
  _서식시도('구분선 테두리', () => t.setBorderWidth(0));
  return t;
}

/** 'PU-03(5), AC-02(4)' → 'PU-03 5건 · AC-02 4건' */
function _top3표시(s) {
  const v = String(s || '').trim();
  if (!v) return '—';
  return v.replace(/\((\d+)\)/g, ' $1건').split(', ').join('   ·   ');
}

/* ============================ 9. 트리거 설정 ============================ */

function 트리거설정() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    const f = t.getHandlerFunction();
    if (f === 'onFormSubmit' || f === 'weeklyReport' || f === '미판정행처리') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('weeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();

  // 안전망: 폼 제출 트리거가 실패해도 5분 안에 미판정 행을 따라잡는다
  ScriptApp.newTrigger('미판정행처리')
    .timeBased()
    .everyMinutes(5)
    .create();

  const 목록 = ScriptApp.getProjectTriggers().map((t) => t.getHandlerFunction() + ' [' + t.getEventType() + ']');
  Logger.log('트리거 ' + 목록.length + '개: ' + 목록.join(' / '));
}

/* ============================ 10. 가상 데이터 ============================ */

/** 개발 전용. 실수로 실행하지 않도록 접두어 _dev_ */
function _dev_가상데이터생성(일수) {
  const 일 = Number(일수 || 28);
  const sh = _시트(SH.원본);
  const h = _헤더(sh);
  const ts열 = _타임스탬프열(h);

  const 설비 = _시트(SH.설비).getDataRange().getValues().slice(1)
    .map((r) => String(r[0])).filter(String);
  const 점검자 = ['김민수', '이서연', '박지훈', '최유진', '정도현'];

  const rnd = (a, b) => a + Math.random() * (b - a);
  const 이상몰림 = ['PU-03', 'AC-02'];

  const out = [];
  const 오늘 = _오늘0시_기준(new Date());

  for (let d = 일 - 1; d >= 0; d--) {
    const base = new Date(오늘.getTime() - d * 86400000);
    설비.forEach((s, si) => {
      const 이상확률 = 이상몰림.indexOf(s) > -1 ? 0.3 : 0.1;
      const 이상함 = Math.random() < 이상확률;

      const 압력 = 이상함 && Math.random() < 0.4 ? rnd(7.1, 8.5).toFixed(1) : rnd(5.5, 6.8).toFixed(1);
      const 온도 = 이상함 && Math.random() < 0.5 ? Math.round(rnd(71, 90)) : Math.round(rnd(45, 65));
      const 액위 = 이상함 && Math.random() < 0.4 ? Math.round(rnd(5, 19)) : Math.round(rnd(40, 85));

      const 고르기 = (정상, 이상) => (이상함 && Math.random() < 0.35 ? 이상 : 정상);
      const 값 = {
        [h[ts열]]: Utilities.formatDate(new Date(base.getTime() + rnd(8 * 3600, 9.5 * 3600) * 1000), TZ, 'yyyy-MM-dd HH:mm:ss'),
        '점검자': 점검자[(si + d) % 점검자.length],
        '설비': s,
        '압축기 토출 압력(bar)': 압력,
        '오일 레벨': 고르기('정상', '보충필요'),
        '펌프 진동·소음': 고르기('정상', Math.random() < 0.5 ? '주의' : '이상'),
        '베어링 온도(℃)': 온도,
        '탱크 액위(%)': 액위,
        '배관 누설': 고르기('없음', '있음'),
        '밸브 잠금 상태': 고르기('정상', '해제됨'),
        '안전 커버·방호': 고르기('정상', '파손'),
        '윤활 급유': 고르기('완료', '미실시'),
        '특이사항': 이상함 ? '소음 증가 확인' : '',
      };

      out.push(h.map((name) => (값[name] !== undefined ? 값[name] : name === '판정결과' ? '' : name === '이상항목수' ? '' : '')));
    });
  }

  sh.getRange(sh.getLastRow() + 1, 1, out.length, h.length).setValues(out);
  Logger.log('가상 데이터 ' + out.length + '행 생성. 이제 전체재판정()을 실행하세요.');
}

/** 기준값을 바꾼 뒤 원본응답 전체를 다시 판정하고 이상이력을 재구성 */
function 전체재판정() {
  const sh = _시트(SH.원본);
  const v = sh.getDataRange().getValues();
  const h = v.shift().map(String);
  const 기준목록 = 기준값읽기();
  const nCol = h.indexOf('판정결과') + 1;

  // 이상이력 비우기(헤더 유지)
  const ih = _시트(SH.이력);
  if (ih.getLastRow() > 1) ih.getRange(2, 1, ih.getLastRow() - 1, ih.getLastColumn()).clearContent();

  const 판정열 = [];
  let 건수 = 0;
  v.forEach((r, i) => {
    const 응답 = _응답만들기(h, r);
    const 이상목록 = 이상치판정(응답, 기준목록);
    판정열.push([이상목록.length ? '이상' : '정상', 이상목록.length]);
    if (이상목록.length) {
      이상이력기록(응답, 이상목록, new Date(r[_타임스탬프열(h)]));
      건수++;
    }
  });

  if (판정열.length) sh.getRange(2, nCol, 판정열.length, 2).setValues(판정열);
  Logger.log('재판정 완료: ' + v.length + '행 중 이상 ' + 건수 + '행');

  // 일일집계도 다시 계산한다(설치를 다시 하면 비워지므로). 날짜 수만큼 시간이 걸린다.
  const ts열 = _타임스탬프열(h);
  const 날짜모음 = {};
  v.forEach((r) => {
    const d = _날짜문자열(r[ts열]);
    if (d) 날짜모음[d] = true;
  });
  const 날짜들 = Object.keys(날짜모음).sort();
  날짜들.forEach((d) => 일일집계갱신(new Date(d + 'T00:00:00+09:00')));
  Logger.log('일일집계 ' + 날짜들.length + '일 재계산 완료');
}

/* ============================ 8. 시트 메뉴 ============================ */

/**
 * 시트를 열면 상단에 [점검시스템] 메뉴를 만든다.
 * 편집기 실행은 입력창(getUi)을 못 띄우는 문맥이 있어, 시트에서 바로 실행할 수 있게 한다.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('점검시스템')
    .addItem('알림 메일 주소 바꾸기', '설정_관리자메일바꾸기')
    .addItem('주간 PDF 보고서 만들기', '메뉴_주간PDF')
    .addItem('밀린 응답 지금 처리', '메뉴_미판정처리')
    .addSeparator()
    .addItem('판정 로직 테스트', '테스트_판정')
    .addItem('이상이력 전체 재판정', '전체재판정')
    .addItem('중복 열 정리', '메뉴_중복열정리')
    .addToUi();
}

/** 알림 메일 받을 주소를 입력창으로 바꾼다 */
function 설정_관리자메일바꾸기() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    const 안내 = '이 실행 문맥에서는 입력창을 띄울 수 없습니다. 시트를 새로고침(F5)한 뒤 상단 [점검시스템] 메뉴에서 실행하세요.';
    Logger.log(안내);
    throw new Error(안내);
  }

  const 지금 = String(설정('관리자이메일') || '');
  const 응답 = ui.prompt(
    '알림 메일 받을 주소',
    '현재: ' + (지금 || '(없음)') + '\n\n새 주소를 입력하고 확인을 누르세요.',
    ui.ButtonSet.OK_CANCEL
  );
  if (응답.getSelectedButton() !== ui.Button.OK) return '취소됨';

  const 메일 = 응답.getResponseText().trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(메일)) {
    ui.alert('이메일 형식이 아닙니다: ' + 메일);
    return '형식 오류';
  }

  const sh = _시트(SH.설정);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === '관리자이메일') {
      sh.getRange(i + 1, 2).setValue(메일);
      Logger.log('관리자이메일 = ' + 메일);
      ui.alert('저장했습니다.\n\n알림 메일 주소: ' + 메일);
      return 메일;
    }
  }
  throw new Error('설정 탭에 관리자이메일 키가 없습니다. 설치_전체를 먼저 실행하세요.');
}

/**
 * 메일 발송 경로 진단. 설정값을 로그로 남기고 실제로 1통 보낸다.
 * 메일이 안 올 때 어느 주소로 가는지, 발송이 막히는지 구분하기 위한 것.
 */
function 진단_메일보내기() {
  const 관리자 = String(설정('관리자이메일') || '').trim();
  const 테스트모드 = String(설정('메일테스트모드') || '').toUpperCase() === 'TRUE';
  const 알림 = String(설정('알림활성화') || '').toUpperCase() === 'TRUE';

  Logger.log('관리자이메일 = "' + 관리자 + '"');
  Logger.log('메일테스트모드 = ' + 테스트모드 + ' / 알림활성화 = ' + 알림);
  Logger.log('폼 제출 알림은 알림활성화가 TRUE 여야 발송됩니다.');

  if (!관리자) {
    throw new Error('관리자이메일이 비어 있습니다. 시트 상단 [점검시스템] → 알림 메일 주소 바꾸기 로 설정하세요.');
  }

  GmailApp.sendEmail(
    관리자,
    '[진단] 알림 메일 경로 테스트',
    '이 메일이 보이면 발송 경로는 정상입니다.\n수신 주소: ' + 관리자 + '\n보낸 계정: ' + Session.getEffectiveUser().getEmail()
  );
  Logger.log('발송 완료 → ' + 관리자 + '  (받은편지함과 스팸함을 함께 확인하세요)');
  return 관리자;
}

/** 주간 집계가 왜 0건으로 나오는지 확인하기 위한 진단 */
function 진단_주간집계() {
  const ih = _시트(SH.이력).getDataRange().getValues();
  const h = ih.shift().map(String);
  const e일 = h.indexOf('점검일');
  Logger.log('이상이력 ' + ih.length + '행 / "점검일" 열 위치: ' + e일 + ' / 헤더: ' + JSON.stringify(h));

  const 종류 = {};
  const 정규화 = [];
  ih.forEach((r) => {
    const raw = r[e일];
    const t = Object.prototype.toString.call(raw);
    종류[t] = (종류[t] || 0) + 1;
    정규화.push(_날짜문자열(raw));
  });
  Logger.log('점검일 값 종류: ' + JSON.stringify(종류));
  Logger.log('앞 3개 원본: ' + JSON.stringify(ih.slice(0, 3).map((r) => String(r[e일]))));

  정규화.sort();
  Logger.log('정규화 범위: "' + 정규화[0] + '" ~ "' + 정규화[정규화.length - 1] + '"');

  const 시작 = _날짜(지난주월요일());
  const 종료 = _날짜(new Date(지난주월요일().getTime() + 6 * 86400000));
  Logger.log('지난주 ' + 시작 + ' ~ ' + 종료 + ' : ' + 정규화.filter((d) => d >= 시작 && d <= 종료).length + '건');

  const 주 = _시트(SH.주간).getDataRange().getValues();
  Logger.log('주간요약 ' + (주.length - 1) + '행 / 마지막 행: ' + JSON.stringify(주[주.length - 1].slice(0, 9)));
}

/** 원본응답 판정결과 열 색 규칙 — 이상=빨강, 정상=초록, 이상항목수>0=굵게 */
function _판정색규칙(응답탭) {
  const h = 응답탭.getRange(1, 1, 1, 응답탭.getLastColumn()).getValues()[0].map(String);
  const n = h.indexOf('판정결과') + 1;
  const m = h.indexOf('이상항목수') + 1;
  if (!n) return;
  const 행수 = Math.max(응답탭.getMaxRows() - 1, 1);
  const 판정범위 = 응답탭.getRange(2, n, 행수, 1);
  const 규칙 = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('이상')
      .setBackground('#f4cccc').setFontColor('#990000').setRanges([판정범위]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('정상')
      .setBackground('#d9ead3').setFontColor('#274e13').setRanges([판정범위]).build(),
  ];
  if (m) {
    규칙.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBold(true).setBackground('#fce5cd')
        .setRanges([응답탭.getRange(2, m, 행수, 1)]).build()
    );
  }
  응답탭.setConditionalFormatRules(규칙);
}

/**
 * 원본응답 탭에 같은 이름의 문항 열이 여러 번 늘어났을 때, 첫 번째 열만 남기고
 * 뒤 중복 열을 정리한다. (설치를 반복해 폼 문항이 다시 만들어질 때 생기는 문제)
 *
 * 폼이 관리하는 열은 구글이 삭제를 막으므로("Cannot delete column with form data"),
 * 지울 수 있는 열은 지우고 나머지는 숨긴다. 숨겨도 값은 남아 있고,
 * 코드는 중복 열에 안전하게 읽도록 되어 있다.
 */
function 정리_중복열() {
  const sh = _시트(SH.원본);
  const h = _헤더(sh);
  const 처음 = {};
  const 중복 = [];
  h.forEach((name, i) => {
    if (!name) return;
    if (처음[name] === undefined) 처음[name] = i;
    else 중복.push(i);
  });

  if (!중복.length) {
    Logger.log('중복 열이 없습니다. 현재 ' + h.length + '열.');
    return 0;
  }

  let 삭제 = 0;
  let 숨김 = 0;
  중복
    .sort((a, b) => b - a) // 뒤에서부터 처리해야 앞 열 번호가 밀리지 않는다
    .forEach((i) => {
      try {
        sh.deleteColumn(i + 1);
        삭제++;
      } catch (e) {
        sh.hideColumns(i + 1);
        숨김++;
      }
    });

  _판정색규칙(sh);
  Logger.log('중복 열 처리: 삭제 ' + 삭제 + '개, 숨김 ' + 숨김 + '개');
  Logger.log('전체 ' + h.length + '열 (숨긴 열 포함). 화면에는 첫 문항 열과 판정 열만 보입니다.');
  Logger.log('헤더: ' + JSON.stringify(_헤더(sh)));
  return 삭제 + 숨김;
}

/**
 * 마지막 응답 1건을 기준값과 대조해, 왜 이상으로 잡히거나 안 잡히는지 보여준다.
 * 폼 제출 자동 판정이 안 될 때 원인을 찾기 위한 진단.
 */
function 진단_마지막응답() {
  const sh = _시트(SH.원본);
  const v = sh.getDataRange().getValues();
  const h = v[0].map(String);
  const r = v[v.length - 1];
  Logger.log('응답 데이터 ' + (v.length - 1) + '행 / 검사 대상은 마지막 행');

  const 응답 = _응답만들기(h, r);

  const 기준 = 기준값읽기();
  Logger.log('기준값 탭 컬럼명: ' + JSON.stringify(기준.map((k) => k.컬럼명)));
  Logger.log('시트 헤더: ' + JSON.stringify(h));

  let 없는것 = 0;
  기준.forEach((k) => {
    if (응답[k.컬럼명] === undefined) {
      Logger.log('⚠️ 응답에 없는 컬럼명: "' + k.컬럼명 + '"');
      없는것++;
    }
  });
  if (없는것) Logger.log('→ 컬럼명이 어긋나 있습니다. 기준값 탭 원본컬럼명을 시트 헤더와 똑같이 맞추세요.');

  Logger.log('응답값: ' + JSON.stringify(응답));
  const 이상목록 = 이상치판정(응답, 기준);
  Logger.log('판정 결과 ' + 이상목록.length + '건: ' + JSON.stringify(이상목록));
  Logger.log('시트에 기록된 값 — 판정결과="' + String(r[h.indexOf('판정결과')]) + '" 이상항목수="' + String(r[h.indexOf('이상항목수')]) + '"');
  return 이상목록;
}

/**
 * 판정결과가 비어 있는 응답 행을 찾아 판정·기록·알림을 수행한다.
 * 폼 제출 트리거가 실패해도 5분 안에 따라잡게 하는 안전망(시간 기반 트리거 대상).
 */
function 미판정행처리() {
  const sh = _시트(SH.원본);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return 0;

  const h = v.shift().map(String);
  const 판정열 = h.indexOf('판정결과');
  const ts열 = _타임스탬프열(h);
  const 기준목록 = 기준값읽기();

  const 처리할것 = [];
  v.forEach((r, i) => {
    if (String(r[판정열] || '').trim() !== '') return;
    if (String(r[ts열] || '').trim() === '') return; // 빈 행
    처리할것.push({ 행: i + 2, 응답: r });
  });

  처리할것.forEach(({ 행, 응답: r }) => {
    const 응답 = _응답만들기(h, r);
    const 시각 = new Date(r[ts열]);
    const 이상목록 = 이상치판정(응답, 기준목록);
    if (이상목록.length) {
      이상이력기록(응답, 이상목록, 시각);
      메일알림(응답, 이상목록, 시각);
    }
    sh.getRange(행, 판정열 + 1, 1, 2).setValues([[이상목록.length ? '이상' : '정상', 이상목록.length]]);
  });

  if (처리할것.length) {
    일일집계갱신(new Date());
    Logger.log('미판정 행 ' + 처리할것.length + '건 처리·알림 완료');
  }
  return 처리할것.length;
}

function 메뉴_중복열정리() {
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  const n = 정리_중복열();
  const 말 = n
    ? '중복 열 ' + n + '개를 정리했습니다.\n\n폼이 관리하는 열은 삭제할 수 없어 숨김 처리했습니다.\n값은 그대로 있고, 판정 코드는 중복 열에 안전하게 동작합니다.'
    : '중복 열이 없습니다.';
  if (ui) ui.alert(말);
  else Logger.log(말);
  return n;
}

function 메뉴_미판정처리() {
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  const n = 미판정행처리();
  const 말 = n ? '밀린 응답 ' + n + '건을 처리했습니다.' : '처리할 응답이 없습니다.';
  if (ui) ui.alert(말);
  else Logger.log(말);
  return n;
}

function 메뉴_주간PDF() {
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  const url = 주간PDF생성();
  if (ui) ui.alert('주간 보고서를 만들었습니다.\n\n' + url);
  else Logger.log('주간 보고서: ' + url);
  return url;
}
/* ============================================================
 *  2부 — 자동 설치 (폼·시트·기준값·PDF폴더·트리거 생성)
 * ============================================================ */
/**
 * 자동 설치 스크립트
 * ------------------------------------------------------------------
 * 손으로 만들던 작업(폼 12문항, 시트 탭 7개, 기준값, 설비목록, 설정,
 * PDF 저장 폴더, 트리거)을 코드로 한 번에 만듭니다.
 * (주간보고 문서는 주간PDF생성()이 매번 코드로 새로 만들므로 템플릿이 필요 없습니다)
 *
 * 사용법:
 *   1) 이 스크립트가 연결된 구글시트를 엽니다.
 *   2) 확장 프로그램 → Apps Script → 함수 선택에서 `설치_전체` 선택 → 실행 → 권한 허용
 *   3) 실행 로그에 폼 URL과 시트 URL이 출력됩니다.
 *   4) (선택) `설치_샘플데이터` 실행 → 4주치 합성 데이터 생성
 *
 * ※ 실제 회사 자료는 쓰지 않습니다. 전부 합성 데이터입니다.
 */

const 설비태그 = [
  'AC-01', 'AC-02', 'AC-03', 'AC-04', 'AC-05',   // 공기압축기
  'PU-01', 'PU-02', 'PU-03', 'PU-04', 'PU-05', 'PU-06', 'PU-07', 'PU-08', // 펌프
  'TK-01', 'TK-02', 'TK-03', 'TK-04',            // 탱크
  'BL-01', 'BL-02', 'BL-03',                     // 송풍기
];

const 점검자목록 = ['김민수', '이서연', '박지훈', '최유진', '정도현'];

const 탭정의 = {
  '기준값': ['항목명', '원본컬럼명', '판정유형', '기준값', '이상값목록', '심각도', '담당자이메일', '사용여부'],
  '설비목록': ['설비태그', '설비명', '구역', '담당자', '담당자이메일'],
  '일일집계': ['점검일', '제출건수', '점검설비수', '미점검설비수', '이상건수', '이상설비목록', '최다이상항목'],
  '이상이력': ['발생시각', '점검일', '설비', '점검자', '항목명', '측정값', '기준', '심각도', '알림발송', '조치상태', '조치내용'],
  '주간요약': ['주차', '시작일', '종료일', '총제출', '이상건수', '이상률(%)', 'TOP3설비', 'TOP3항목', '미조치건수', 'PDF링크'],
  '설정': ['키', '값'],
};

const 기준값데이터 = [
  ['토출압력', '압축기 토출 압력(bar)', '초과', 7.0, '', '상', 'mech@example.com', true],
  ['오일레벨', '오일 레벨', '값일치', '', '보충필요', '중', 'mech@example.com', true],
  ['진동소음', '펌프 진동·소음', '값일치', '', '주의,이상', '상', 'mech@example.com', true],
  ['베어링온도', '베어링 온도(℃)', '초과', 70, '', '상', 'mech@example.com', true],
  ['탱크액위', '탱크 액위(%)', '미만', 20, '', '중', 'ops@example.com', true],
  ['배관누설', '배관 누설', '값일치', '', '있음', '상', 'ops@example.com', true],
  ['밸브잠금', '밸브 잠금 상태', '값일치', '', '해제됨', '상', 'safety@example.com', true],
  ['안전커버', '안전 커버·방호', '값일치', '', '파손', '상', 'safety@example.com', true],
  ['윤활급유', '윤활 급유', '값일치', '', '미실시', '하', 'mech@example.com', true],
];

/** 숫자 범위 검증(안내문구는 지원 여부가 불확실해 실패해도 무시) */
function _숫자검증(최소, 최대, 안내) {
  const b = FormApp.createTextValidation().requireNumberBetween(최소, 최대);
  try {
    if (typeof b.setHelpText === 'function') b.setHelpText(안내);
  } catch (e) {}
  return b.build();
}

/* ============================ 전체 설치 ============================ */

function 설치_전체() {
  const 로그 = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('점검시스템_포트폴리오');
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  ss.setSpreadsheetLocale('ko_KR');
  로그.push('시트: ' + ss.getUrl());

  // 재설치해도 기존 설정값은 유지한다(탭을 비우기 전에 미리 읽어 둔다)
  const 기존설정 = {};
  const 설정탭기존 = ss.getSheetByName('설정');
  if (설정탭기존 && 설정탭기존.getLastRow() > 1) {
    설정탭기존.getRange(2, 1, 설정탭기존.getLastRow() - 1, 2).getValues().forEach((r) => {
      const k = String(r[0]).trim();
      if (k) 기존설정[k] = r[1];
    });
  }
  const 기존관리자메일 = String(기존설정['관리자이메일'] || '').trim();

  // 1) 탭 준비 — 재실행해도 안전하게: 없는 탭만 만들고, 기본 시트(Sheet1)는 재활용
  const 탭이름들 = Object.keys(탭정의);
  const 우리탭전체 = 탭이름들.concat(['원본응답']);
  탭이름들.forEach((이름) => {
    if (ss.getSheetByName(이름)) return;
    const 재활용 = ss.getSheets().filter((s) => 우리탭전체.indexOf(s.getName()) === -1)[0];
    if (재활용) 재활용.setName(이름);
    else ss.insertSheet(이름);
  });
  탭이름들.forEach((이름) => {
    const sh = ss.getSheetByName(이름);
    const 정의 = 탭정의[이름];
    const 현재 = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    const 헤더같음 = 정의.every((v, i) => 현재[i] === v);

    // 헤더가 다를 때만 비운다. 같을 때도 비우면 재설치 때 이상이력·일일집계·주간요약이 사라진다.
    if (!헤더같음) {
      sh.clear();
      sh.getRange(1, 1, 1, 정의.length).setValues([정의]).setFontWeight('bold');
    }
    sh.setFrozenRows(1);
  });

  // 2) 폼 생성 + 문항 12개 -------------------------------------
  let form = null;
  const 기존 = DriveApp.getFilesByName('유틸리티동 일일 설비 점검표');
  if (기존.hasNext()) {
    form = FormApp.openById(기존.next().getId());
    form.setTitle('유틸리티동 일일 설비 점검표').setDescription('설비 1대당 1회 제출. 숫자는 게이지 표시값 그대로 입력.');
  } else {
    form = FormApp.create('유틸리티동 일일 설비 점검표');
    form.setDescription('설비 1대당 1회 제출. 숫자는 게이지 표시값 그대로 입력.');
  }

  // 문항이 이미 있으면 다시 만들지 않는다.
  // 폼 문항을 다시 만들면 연결된 시트에 같은 이름의 열이 새로 늘어나서
  // 판정이 뒤쪽 빈 열을 읽게 되는 문제가 생긴다.
  const 문항이있음 = form.getItems().length === 12;
  if (문항이있음) {
    로그.push('폼 문항이 이미 있어 그대로 사용합니다(중복 열 방지)');
  } else {
    form.getItems().forEach((it) => form.deleteItem(it));
  }

  if (!문항이있음) {
    form.addListItem().setTitle('점검자').setChoiceValues(점검자목록).setRequired(true);
    form.addListItem().setTitle('설비').setChoiceValues(설비태그).setRequired(true);
    form.addTextItem().setTitle('압축기 토출 압력(bar)').setRequired(true)
      .setValidation(_숫자검증(0, 15, '0~15 사이 숫자로 입력하세요.'));
    form.addMultipleChoiceItem().setTitle('오일 레벨').setChoiceValues(['정상', '보충필요']).setRequired(true);
    form.addMultipleChoiceItem().setTitle('펌프 진동·소음').setChoiceValues(['정상', '주의', '이상']).setRequired(true);
    form.addTextItem().setTitle('베어링 온도(℃)').setRequired(true)
      .setValidation(_숫자검증(0, 150, '0~150 사이 숫자로 입력하세요.'));
    form.addTextItem().setTitle('탱크 액위(%)').setRequired(true)
      .setValidation(_숫자검증(0, 100, '0~100 사이 숫자로 입력하세요.'));
    form.addMultipleChoiceItem().setTitle('배관 누설').setChoiceValues(['없음', '있음']).setRequired(true);
    form.addMultipleChoiceItem().setTitle('밸브 잠금 상태').setChoiceValues(['정상', '해제됨']).setRequired(true);
    form.addMultipleChoiceItem().setTitle('안전 커버·방호').setChoiceValues(['정상', '파손']).setRequired(true);
    form.addMultipleChoiceItem().setTitle('윤활 급유').setChoiceValues(['완료', '미실시']).setRequired(true);
    form.addParagraphTextItem().setTitle('특이사항').setRequired(false);
  }


  if (form.getDestinationId() !== ss.getId()) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  }
  로그.push('폼: ' + form.getPublishedUrl());
  로그.push('폼 편집: ' + form.getEditUrl());

  // 3) 응답 탭 이름 정리 --------------------------------------
  const 우리탭 = Object.keys(탭정의).concat(['원본응답']);
  let 응답탭 = ss.getSheets().filter((s) => 우리탭.indexOf(s.getName()) === -1)[0];
  if (!응답탭) 응답탭 = ss.getSheetByName('원본응답');   // 두 번째 실행부터는 이름으로 찾음
  if (응답탭) {
    응답탭.setName('원본응답');
    const h = 응답탭.getRange(1, 1, 1, Math.max(응답탭.getLastColumn(), 1)).getValues()[0];
    if (h.indexOf('판정결과') === -1) {
      응답탭.getRange(1, h.length + 1, 1, 2).setValues([['판정결과', '이상항목수']]);
    }
    응답탭.setFrozenRows(1);
    _판정색규칙(응답탭);
    로그.push('응답 탭: 원본응답 (판정결과/이상항목수 + 이상=빨강 색규칙)');
  } else {
    로그.push('⚠️ 응답 탭을 찾지 못했습니다. 폼 문항을 한 번 제출한 뒤 다시 실행하세요.');
  }

  // 4) 데이터 채우기 ------------------------------------------
  // 사용자가 임계값을 고쳤을 수 있으므로 비어 있을 때만 기본값을 채운다
  const 기준시트 = ss.getSheetByName('기준값');
  if (기준시트.getLastRow() <= 1) {
    기준시트.getRange(2, 1, 기준값데이터.length, 8).setValues(기준값데이터);
  } else {
    로그.push('기준값: 기존 값 유지 (' + (기준시트.getLastRow() - 1) + '행)');
  }

  const 설비명 = { AC: '공기압축기', PU: '펌프', TK: '탱크', BL: '송풍기' };
  const 설비행 = 설비태그.map((t) => {
    const 접두 = t.split('-')[0];
    const 이상 = 접두 === 'TK' ? 'ops@example.com' : 접두 === 'BL' ? 'safety@example.com' : 'mech@example.com';
    return [t, 설비명[접두] + ' ' + t.split('-')[1], 접두 === 'TK' ? '탱크구역' : '유틸리티동', '담당자' + t.split('-')[1], 이상];
  });
  const 설비시트 = ss.getSheetByName('설비목록');
  if (설비시트.getLastRow() <= 1) 설비시트.getRange(2, 1, 설비행.length, 5).setValues(설비행);

  let 내메일 = 기존관리자메일;
  if (내메일) {
    로그.push('관리자 이메일: 기존 값 유지 (' + 내메일 + ')');
  } else {
    try { 내메일 = Session.getActiveUser().getEmail(); } catch (e) {}
    if (!내메일) { try { 내메일 = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  }
  if (!내메일) 로그.push('⚠️ 관리자 이메일을 자동으로 읽지 못했습니다. 시트 상단 [점검시스템] → 알림 메일 주소 바꾸기 로 설정하세요.');
  // 5) PDF 저장 폴더 ------------------------------------------
  const 폴더검색 = DriveApp.getFoldersByName('점검시스템_포트폴리오');
  const 폴더 = 폴더검색.hasNext() ? 폴더검색.next() : DriveApp.createFolder('점검시스템_포트폴리오');

  if (기존설정['알림활성화'] === undefined) 기존설정['알림활성화'] = 'TRUE';
  if (기존설정['메일테스트모드'] === undefined) 기존설정['메일테스트모드'] = 'TRUE';
  기존설정['관리자이메일'] = 내메일;
  기존설정['PDF폴더ID'] = 폴더.getId();

  const 설정시트2 = ss.getSheetByName('설정');
  설정시트2.clear();
  설정시트2.getRange(1, 1, 1, 2).setValues([['키', '값']]).setFontWeight('bold');
  const 설정행 = Object.keys(기존설정).map((k) => [k, 기존설정[k]]);
  설정시트2.getRange(2, 1, 설정행.length, 2).setValues(설정행);
  로그.push('PDF 저장 폴더: ' + 폴더.getUrl());

  // 6) 트리거 -------------------------------------------------
  트리거설정();
  로그.push('트리거: onFormSubmit(폼 제출 시), weeklyReport(월요일 07시)');

  // 7) 결과 ---------------------------------------------------
  로그.push('');
  로그.push('※ 다음: (선택) 설치_샘플데이터 실행 → 4주치 합성 데이터 생성');
  Logger.log(로그.join('\n'));
  return 로그.join('\n');
}

/* ============================ 샘플 데이터 ============================ */

function 설치_샘플데이터() {
  _dev_가상데이터생성(28);
  전체재판정();

  // 데모 시트가 바로 완성되게 일일집계(28일)·주간요약(4주)까지 채운다.
  // (1~2분 걸릴 수 있습니다)
  const 오늘 = _오늘0시_기준(new Date());
  for (let d = 27; d >= 0; d--) 일일집계갱신(new Date(오늘.getTime() - d * 86400000));

  const 이번주월요일 = new Date(오늘.getTime() - ((오늘.getDay() + 6) % 7) * 86400000);
  for (let w = 3; w >= 0; w--) 주간요약갱신(new Date(이번주월요일.getTime() - w * 7 * 86400000));

  const n = _시트(SH.이력).getLastRow() - 1;
  Logger.log('샘플 데이터 생성 완료. 이상이력 ' + n + '건 / 일일집계 28일 / 주간요약 ' + (_시트(SH.주간).getLastRow() - 1) + '주.');
  return '이상이력 ' + n + '건';
}

/* ============================ 설치 상태 점검 ============================ */

function 설치_확인() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 있는탭 = ss.getSheets().map((s) => s.getName());
  const 필요 = Object.keys(탭정의).concat(['원본응답']);
  const 없는탭 = 필요.filter((t) => 있는탭.indexOf(t) === -1);
  const 트리거 = ScriptApp.getProjectTriggers().map((t) => t.getHandlerFunction());
  const 결과 = {
    시트탭: 없는탭.length ? '누락: ' + 없는탭.join(', ') : '정상 (' + 필요.length + '개)',
    트리거: 트리거.length ? 트리거.join(', ') : '없음 — 트리거설정() 실행 필요',
    기준값행: _시트(SH.기준).getLastRow() - 1 + '행',
    설비수: _시트(SH.설비).getLastRow() - 1 + '대',
    관리자메일: String(설정('관리자이메일')),
    메일테스트모드: String(설정('메일테스트모드')),
    PDF폴더: String(설정('PDF폴더ID')) ? '설정됨' : '비어 있음',
  };
  Logger.log(JSON.stringify(결과, null, 2));
  return 결과;
}

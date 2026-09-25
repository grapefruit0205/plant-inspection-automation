/**
 * 현장 점검표 자동 집계·주간보고 시스템 — 통합 설치 파일
 * ------------------------------------------------------------
 * 사용법:
 *   1) 구글시트(sheets.new)를 새로 만들고 이름을 '점검시스템_포트폴리오'로 지정
 *   2) 확장 프로그램 → Apps Script → 기본 Code.gs 내용을 전부 지우고 이 파일 전체를 붙여넣기
 *   3) 저장(Ctrl+S) → 함수 선택에서 '설치_전체' 선택 → 실행 → 권한 허용
 *   4) (선택) '설치_샘플데이터' 실행 → 4주치 합성 데이터 생성
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
  if (!e || !e.namedValues) {
    throw new Error('이 함수는 폼 제출로만 실행됩니다. 편집기에서 테스트하려면 테스트_판정()을 쓰세요.');
  }
  const 응답 = {};
  Object.keys(e.namedValues).forEach((k) => (응답[k] = e.namedValues[k][0]));

  const 기준목록 = 기준값읽기();
  const 이상목록 = 이상치판정(응답, 기준목록);

  if (이상목록.length) {
    이상이력기록(응답, 이상목록, new Date());
    메일알림(응답, 이상목록, new Date());
  }
  원본응답판정기록(e.range.getRow(), 이상목록);
  일일집계갱신(new Date());
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
  const 설비카운트 = {};
  v.forEach((r) => {
    if (!r[ts]) return;
    const d = _날짜(new Date(r[ts]));
    if (d >= 시작 && d <= 종료) {
      총제출++;
      const s = String(r[설비열]);
      설비카운트[s] = (설비카운트[s] || 0) + 1;
    }
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
  ih.forEach((r) => {
    const d = String(r[e일]);
    if (d >= 시작 && d <= 종료) {
      이상건수++;
      const a = String(r[e항목]);
      항목카운트[a] = (항목카운트[a] || 0) + 1;
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
  sh.appendRow([
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
  ]);

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

function 이상표데이터(시작일, 종료일) {
  const ih = _시트(SH.이력).getDataRange().getValues();
  const h = ih.shift().map(String);
  const e일 = h.indexOf('점검일');
  const e설비 = h.indexOf('설비');
  const e항목 = h.indexOf('항목명');
  const e값 = h.indexOf('측정값');
  const e기준 = h.indexOf('기준');
  const e심 = h.indexOf('심각도');

  const rows = [['점검일', '설비', '항목', '측정값', '기준', '심각도']];
  ih.forEach((r) => {
    const d = String(r[e일]);
    if (d >= 시작일 && d <= 종료일) {
      rows.push([d, String(r[e설비]), String(r[e항목]), String(r[e값]), String(r[e기준]), String(r[e심])]);
    }
  });
  return rows.slice(0, 21); // 표가 너무 길어지지 않게 최대 20행
}

function 주간PDF생성() {
  const 요약 = 주간요약갱신(지난주월요일());
  const 템플릿ID = String(설정('템플릿문서ID'));
  const 폴더ID = String(설정('PDF폴더ID'));
  const 관리자 = String(설정('관리자이메일'));

  const 사본 = DriveApp.getFileById(템플릿ID).makeCopy('임시_' + 요약.주차);
  const doc = DocumentApp.openById(사본.getId());
  const body = doc.getBody();

  Object.keys(요약).forEach((k) => body.replaceText('{{' + k + '}}', String(요약[k])));

  // {{이상표}} 자리에 표 삽입
  const 찾기 = body.findText('{{이상표}}');
  if (찾기) {
    const el = 찾기.getElement();
    const 부모 = el.getParent();
    const idx = body.getChildIndex(부모);
    body.insertTable(idx, 이상표데이터(요약.시작일, 요약.종료일));
    부모.removeFromParent();
  }

  doc.saveAndClose();

  const pdf = DriveApp.getFileById(사본.getId())
    .getAs('application/pdf')
    .setName('주간설비점검보고_' + 요약.시작일 + '.pdf');
  const 파일 = DriveApp.getFolderById(폴더ID).createFile(pdf);
  사본.setTrashed(true);

  GmailApp.sendEmail(관리자, '[주간보고] ' + 요약.주차, '주간 설비 점검 보고서를 첨부합니다.', {
    attachments: [pdf],
  });

  // 주간요약 탭 PDF링크 기록(마지막 행)
  const sh = _시트(SH.주간);
  sh.getRange(sh.getLastRow(), _열(sh, 'PDF링크')).setValue(파일.getUrl());

  return 파일.getUrl();
}

/* ============================ 9. 트리거 설정 ============================ */

function 트리거설정() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    const f = t.getHandlerFunction();
    if (f === 'onFormSubmit' || f === 'weeklyReport') ScriptApp.deleteTrigger(t);
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

  Logger.log('트리거 2개 생성 완료: onFormSubmit, weeklyReport');
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
    const 응답 = {};
    h.forEach((name, j) => (응답[name] = r[j]));
    const 이상목록 = 이상치판정(응답, 기준목록);
    판정열.push([이상목록.length ? '이상' : '정상', 이상목록.length]);
    if (이상목록.length) {
      이상이력기록(응답, 이상목록, new Date(r[_타임스탬프열(h)]));
      건수++;
    }
  });

  if (판정열.length) sh.getRange(2, nCol, 판정열.length, 2).setValues(판정열);
  Logger.log('재판정 완료: ' + v.length + '행 중 이상 ' + 건수 + '행');
}
/* ============================================================
 *  2부 — 자동 설치 (폼·시트·기준값·템플릿·트리거 생성)
 * ============================================================ */
/**
 * 자동 설치 스크립트
 * ------------------------------------------------------------------
 * 손으로 만들던 작업(폼 12문항, 시트 탭 7개, 기준값, 설비목록, 설정,
 * 주간보고 템플릿, 트리거)을 코드로 한 번에 만듭니다.
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

const 템플릿본문 = [
  '주간 설비 점검 보고서',
  '',
  '기간: {{기간}}',
  '총 제출: {{총제출}}건',
  '이상 건수: {{이상건수}}건 (이상률 {{이상률}}%)',
  'TOP3 설비: {{TOP3설비}}',
  'TOP3 항목: {{TOP3항목}}',
  '미조치: {{미조치}}건',
  '',
  '[이상 상세]',
  '{{이상표}}',
  '',
  '생성일시: {{생성일시}}',
];

/** 숫자 범위 검증(안내문구는 지원 여부가 불확실해 실패해도 무시) */
function _숫자검증(최소, 최대, 안내) {
  const b = FormApp.createTextValidation().requireNumberBetween(최소, 최대);
  try {
    if (typeof b.setHelpText === 'function') b.setHelpText(안내);
  } catch (e) {}
  return b.build();
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

/* ============================ 전체 설치 ============================ */

function 설치_전체() {
  const 로그 = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('점검시스템_포트폴리오');
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  ss.setSpreadsheetLocale('ko_KR');
  로그.push('시트: ' + ss.getUrl());

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
    sh.clear();
    sh.getRange(1, 1, 1, 탭정의[이름].length).setValues([탭정의[이름]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });

  // 2) 폼 생성 + 문항 12개 -------------------------------------
  let form = null;
  const 기존 = DriveApp.getFilesByName('유틸리티동 일일 설비 점검표');
  if (기존.hasNext()) {
    const id = 기존.next().getId();
    form = FormApp.openById(id);
    form.getItems().forEach((it) => form.deleteItem(it));
    form.setTitle('유틸리티동 일일 설비 점검표').setDescription('설비 1대당 1회 제출. 숫자는 게이지 표시값 그대로 입력.');
  } else {
    form = FormApp.create('유틸리티동 일일 설비 점검표');
    form.setDescription('설비 1대당 1회 제출. 숫자는 게이지 표시값 그대로 입력.');
  }

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
  const 기준시트 = ss.getSheetByName('기준값');
  기준시트.getRange(2, 1, 기준값데이터.length, 8).setValues(기준값데이터);

  const 설비명 = { AC: '공기압축기', PU: '펌프', TK: '탱크', BL: '송풍기' };
  const 설비행 = 설비태그.map((t) => {
    const 접두 = t.split('-')[0];
    const 이상 = 접두 === 'TK' ? 'ops@example.com' : 접두 === 'BL' ? 'safety@example.com' : 'mech@example.com';
    return [t, 설비명[접두] + ' ' + t.split('-')[1], 접두 === 'TK' ? '탱크구역' : '유틸리티동', '담당자' + t.split('-')[1], 이상];
  });
  ss.getSheetByName('설비목록').getRange(2, 1, 설비행.length, 5).setValues(설비행);

  let 내메일 = '';
  try { 내메일 = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!내메일) { try { 내메일 = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  if (!내메일) 로그.push('⚠️ 관리자 이메일을 자동으로 읽지 못했습니다. 설정 탭 관리자이메일에 직접 입력하세요.');
  ss.getSheetByName('설정').getRange(2, 1, 4, 2).setValues([
    ['관리자이메일', 내메일],
    ['알림활성화', 'TRUE'],
    ['메일테스트모드', 'TRUE'],
    ['템플릿문서ID', ''],
  ]);

  // 5) 주간보고 템플릿 문서 -----------------------------------
  let 폴더;
  const 폴더검색 = DriveApp.getFoldersByName('점검시스템_포트폴리오');
  폴더 = 폴더검색.hasNext() ? 폴더검색.next() : DriveApp.createFolder('점검시스템_포트폴리오');

  const 기존문서 = 폴더.getFilesByName('주간보고_템플릿');
  const 템플릿 = 기존문서.hasNext()
    ? DocumentApp.openById(기존문서.next().getId())
    : DocumentApp.create('주간보고_템플릿');
  const body = 템플릿.getBody();
  body.clear();
  템플릿본문.forEach((줄, i) => (i === 0 ? body.appendParagraph(줄).setHeading(DocumentApp.ParagraphHeading.HEADING1) : body.appendParagraph(줄)));
  body.setFontFamily('Noto Sans KR');
  템플릿.saveAndClose();
  try { DriveApp.getFileById(템플릿.getId()).moveTo(폴더); } catch (e) {}

  const 설정시트 = ss.getSheetByName('설정');
  const 설정값 = 설정시트.getDataRange().getValues();
  for (let i = 1; i < 설정값.length; i++) {
    if (설정값[i][0] === '템플릿문서ID') 설정시트.getRange(i + 1, 2).setValue(템플릿.getId());
  }
  설정시트.appendRow(['PDF폴더ID', 폴더.getId()]);
  로그.push('템플릿 문서: ' + 템플릿.getUrl());
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
    템플릿ID: String(설정('템플릿문서ID')) ? '설정됨' : '비어 있음',
  };
  Logger.log(JSON.stringify(결과, null, 2));
  return 결과;
}

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

  // 담당자별 묶음
  const 묶음 = {};
  이상목록.forEach((x) => {
    const to = x.메일 || String(설정('관리자이메일'));
    (묶음[to] = 묶음[to] || []).push(x);
  });

  const 설비 = 응답['설비'] || '';
  const 점검자 = 응답['점검자'] || '';
  let 발송 = 0;

  Object.keys(묶음).forEach((to) => {
    const 목록 = 묶음[to];
    const 제목 = '[설비이상] ' + 설비 + ' ' + 목록[0].항목명 + (목록.length > 1 ? ' 외 ' + (목록.length - 1) + '건' : '');
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

    GmailApp.sendEmail(to, 제목, 제목, { htmlBody: 본문 });
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
    ts: h.indexOf('타임스탬프'),
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
  const ts = h.indexOf('타임스탬프');
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
        '타임스탬프': Utilities.formatDate(new Date(base.getTime() + rnd(8 * 3600, 9.5 * 3600) * 1000), TZ, 'yyyy-MM-dd HH:mm:ss'),
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
      이상이력기록(응답, 이상목록, new Date(r[h.indexOf('타임스탬프')]));
      건수++;
    }
  });

  if (판정열.length) sh.getRange(2, nCol, 판정열.length, 2).setValues(판정열);
  Logger.log('재판정 완료: ' + v.length + '행 중 이상 ' + 건수 + '행');
}
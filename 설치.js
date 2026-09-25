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
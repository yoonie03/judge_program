// Google Sheets와 Google Forms 연동 자동 채점 시스템
// 메인 스프레드시트 ID를 여기에 설정하세요
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
const FORM_ID = 'YOUR_FORM_ID';

// 시트 이름 상수
const SHEET_NAMES = {
  PARTICIPANTS: '참가자명단',
  RESPONSES: '설문응답',
  SCORES: '점수계산',
  SUMMARY: '요약',
  MVP_SCORES: 'MVP점수'
};

// 채점 항목 구조
const SCORING_ITEMS = [
  // 창의성 (30점)
  {
    category: '창의성',
    items: [
      { name: '아이디어와 결과물이 얼마나 창의적인가? (현재 시장에서 찾아보기 어려운 독특하고 새로운 아이디어)', maxScore: 10, description: '현재 시장에서 찾아보기 어려운 독특하고 새로운 아이디어' },
      { name: '제시된 아이디어가 실행 가능하면서도 독창적인가?', maxScore: 10, description: '제시된 아이디어가 실행 가능하면서도 독창적인가?' },
      { name: '기존에 있는 유사한 제품과 차별성이 있는가?', maxScore: 10, description: '기존에 있는 유사한 제품과 차별성이 있는가?' }
    ]
  },
  // 구현 완성도 (30점)
  {
    category: '구현 완성도',
    items: [
      { name: '프로젝트의 필수 기능이 구현되어있는가?', maxScore: 10, description: '프로젝트의 필수 기능이 구현되어있는가?' },
      { name: '예상치 못한 상황과 오류를 얼마나 잘 처리 하였는가?', maxScore: 10, description: '예상치 못한 상황과 오류를 얼마나 잘 처리 하였는가?' },
      { name: '활용한 데이터와 기술에 대한 이해도가 충분했는가?', maxScore: 10, description: '활용한 데이터와 기술에 대한 이해도가 충분했는가?' }
    ]
  },
  // 필요성 (30점)
  {
    category: '필요성',
    items: [
      { name: '확장성 - 해당 기술이 타 제품에도 응용될 수 있는가?', maxScore: 10, description: '해당 기술이 타제품에도 응용될 수 있는가?' },
      { name: '주제 적합성 - 사전에 공지된 주제에 적합한 아이디어를 구현하였는가?', maxScore: 10, description: '사전에 공지된 주제에 적합한 아이디어를 구현하였는가?' },
      { name: '상품성 - 해당 기술이 우리 사회에 충분한 수요가 있는가?', maxScore: 10, description: '해당 기술이 우리 사회에 충분한 수요가 있는가?' }
    ]
  },
  // 발표 전달력 (10점)
  {
    category: '발표 전달력',
    items: [
      { name: '발표 속도가 적절하고 적절한 표현으로 발표하였는가?', maxScore: 5, description: '발표 속도가 적절하고 적절한 표현으로 발표하였는가?', scaleMax: 5 },
      { name: '발표자료가 직관적이고 이해하기 쉽게 구성되었는가?', maxScore: 5, description: '발표자료가 직관적이고 이해하기 쉽게 구성되었는가?', scaleMax: 5 }
    ]
  }
];

// 전체 항목을 평탄화한 리스트 (응답 처리용)
function getAllScoringItems() {
  const allItems = [];
  SCORING_ITEMS.forEach(category => {
    category.items.forEach(item => {
      allItems.push({
        name: item.name,
        maxScore: item.maxScore,
        category: category.category,
        scaleMax: item.scaleMax || 10  // 기본값 10, 발표 전달력은 5
      });
    });
  });
  return allItems;
}

// 설문지 설명 텍스트
const FORM_DESCRIPTION = '반드시 이름과 팀을 정확히 기록해 주셔야 합니다.\n\n본인 팀을 제외한 모든 팀을 평가해주세요.\n\n채점 항목:\n- 창의성 (30점): 아이디어 창의성, 실행 가능성, 차별성\n- 구현 완성도 (30점): 필수 기능, 오류 처리, 기술 이해도\n- 필요성 (30점): 확장성, 주제 적합성, 상품성\n- 발표 전달력 (10점): 발표 속도/표현, 발표자료 구성\n- 총 100점 만점\n\n※ 본인 팀 평가는 자동으로 제외되므로, 본인 팀 섹션은 건너뛰셔도 됩니다.\n※ 5팀씩 나누어 평가하실 수 있습니다.';

/**
 * 초기 설정 함수 - 시트 생성 및 구조 설정
 */
function initializeSheets() {
  try {
    // SPREADSHEET_ID 확인
    if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID') {
      throw new Error('SPREADSHEET_ID를 먼저 설정해주세요. Code.gs 파일 상단의 SPREADSHEET_ID를 Google Sheets ID로 변경해주세요.');
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    if (!ss) {
      throw new Error('스프레드시트를 찾을 수 없습니다. SPREADSHEET_ID가 올바른지 확인해주세요.');
    }
  
  // 참가자명단 시트 생성
  let sheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PARTICIPANTS);
    sheet.getRange(1, 1, 1, 3).setValues([['이름', '팀명', '완료여부']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  
  // 설문응답 시트 생성
  sheet = ss.getSheetByName(SHEET_NAMES.RESPONSES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.RESPONSES);
    const headers = ['타임스탬프', '평가자이름', '평가자팀', '평가받은팀'];
    SCORING_ITEMS.forEach(item => headers.push(item));
    headers.push('총점', 'MVP점수');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  
  // MVP점수 시트 생성
  sheet = ss.getSheetByName(SHEET_NAMES.MVP_SCORES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.MVP_SCORES);
    sheet.getRange(1, 1, 1, 3).setValues([['팀명', 'MVP점수', '입력일시']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  
  // 점수계산 시트 생성
  sheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SCORES);
    sheet.getRange(1, 1, 1, 5).setValues([['팀명', '총점수', '팀원수', '평균점수', 'MVP점수']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  
  // 요약 시트 생성
  sheet = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SUMMARY);
  }
  
  Logger.log('시트 초기화 완료');
  Logger.log('생성된 시트: 참가자명단, 설문응답, 점수계산, 요약, MVP점수');
  return '시트 초기화가 완료되었습니다!';
  } catch (error) {
    Logger.log(`시트 초기화 오류: ${error.toString()}`);
    throw error;
  }
}

/**
 * 설문 응답을 처리하는 함수 (트리거로 사용)
 */
function processFormResponse(e) {
  try {
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    const form = FormApp.openById(FORM_ID);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responsesSheet = ss.getSheetByName(SHEET_NAMES.RESPONSES);
    
    // 설문 응답에서 데이터 추출
    // 0: 설명 (무시)
    // 1: 평가자 이름 (드롭다운)
    // 2: 평가자 팀명 (드롭다운)
    // 3 이후: 각 팀별 채점 항목 11개 + MVP 점수
    
    let evaluatorName = '';
    let evaluatorTeam = '';
    const teamScores = {}; // {팀명: {항목별점수: [], 총점: 0, MVP: 0}}
    
    // 참가자 정보 추출
    if (itemResponses.length >= 2) {
      evaluatorName = itemResponses[0].getResponse();
      evaluatorTeam = itemResponses[1].getResponse();
    }
    
    // 참가자 명단 확인
    const participants = getParticipants();
    const participant = participants.find(p => 
      p.name === evaluatorName && p.team === evaluatorTeam
    );
    
    if (!participant) {
      Logger.log(`참가자 명단에 없음: ${evaluatorName} (${evaluatorTeam})`);
      return;
    }
    
    // 모든 팀 목록 가져오기
    const allTeams = [...new Set(participants.map(p => p.team))];
    
    // 각 팀별로 채점 항목 + MVP 점수 추출
    // 설문지에는 모든 팀이 포함되어 있으므로, 모든 팀의 응답을 읽어야 함
    const allItems = getAllScoringItems();
    const totalItemsPerTeam = allItems.length + 1; // 채점 항목 + MVP
    
    let currentIndex = 2; // 이름, 팀명 다음부터
    
    allTeams.forEach(team => {
      // 본인 팀은 제외
      if (team === evaluatorTeam) {
        // 본인 팀 섹션 건너뛰기
        currentIndex += totalItemsPerTeam;
        return;
      }
      
      const itemScores = [];
      let totalScore = 0;
      
      // 모든 채점 항목 추출
      allItems.forEach((item, index) => {
        if (currentIndex < itemResponses.length) {
          // Google Forms에서 척도로 입력받은 값을 실제 점수로 변환
          const rawScore = parseFloat(itemResponses[currentIndex].getResponse()) || 0;
          const scaleMax = item.scaleMax || 10;  // 발표 전달력은 5, 나머지는 10
          // 실제 점수 = (입력값 / 척도최대값) * 최대점수
          const actualScore = (rawScore / scaleMax) * item.maxScore;
          itemScores.push(actualScore);
          totalScore += actualScore;
          currentIndex++;
        }
      });
      
      // MVP 점수는 현재 설문지에 포함되지 않으므로 제외
      // 필요시 아래 주석을 해제하고 설문지 생성 부분도 수정
      // let mvpScore = 0;
      // if (currentIndex < itemResponses.length) {
      //   mvpScore = parseFloat(itemResponses[currentIndex].getResponse()) || 0;
      //   currentIndex++;
      // }
      const mvpScore = 0;  // MVP 점수 미사용
      
      if (itemScores.length === allItems.length) {
        teamScores[team] = {
          itemScores: itemScores,
          totalScore: totalScore,
          mvpScore: mvpScore
        };
      }
    });
    
    // 응답 저장
    Object.keys(teamScores).forEach(team => {
      const scoreData = teamScores[team];
      const row = [
        new Date(),
        evaluatorName,
        evaluatorTeam,
        team
      ];
      
      // 모든 항목별 점수
      scoreData.itemScores.forEach(score => row.push(score));
      
      // 총점 및 MVP 점수 (MVP는 0으로 저장)
      row.push(scoreData.totalScore);
      row.push(0);  // MVP 점수 미사용
      
      responsesSheet.appendRow(row);
    });
    
    // 참가자 완료 여부 업데이트
    updateParticipantStatus(evaluatorName, evaluatorTeam, true);
    
    // 점수 재계산
    calculateScores();
    
    Logger.log(`응답 처리 완료: ${evaluatorName}`);
  } catch (error) {
    Logger.log(`오류 발생: ${error.toString()}`);
  }
}

/**
 * 참가자 명단 가져오기
 */
function getParticipants() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    if (!ss) {
      throw new Error('스프레드시트를 찾을 수 없습니다. SPREADSHEET_ID를 확인해주세요.');
    }
    
    const sheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    
    if (!sheet) {
      throw new Error(`'${SHEET_NAMES.PARTICIPANTS}' 시트를 찾을 수 없습니다. 먼저 'initializeSheets' 함수를 실행해주세요.`);
    }
    
    const dataRange = sheet.getDataRange();
    
    if (!dataRange) {
      return [];  // 빈 시트인 경우 빈 배열 반환
    }
    
    const data = dataRange.getValues();
    
    const participants = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        participants.push({
          name: data[i][0],
          team: data[i][1],
          completed: data[i][2] === '완료' || data[i][2] === true
        });
      }
    }
    
    return participants;
  } catch (error) {
    Logger.log(`참가자 명단 가져오기 오류: ${error.toString()}`);
    throw error;
  }
}

/**
 * 참가자 완료 여부 업데이트
 */
function updateParticipantStatus(name, team, completed) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name && data[i][1] === team) {
      sheet.getRange(i + 1, 3).setValue(completed ? '완료' : '미완료');
      break;
    }
  }
}

/**
 * 점수 계산 함수
 */
function calculateScores() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const responsesSheet = ss.getSheetByName(SHEET_NAMES.RESPONSES);
  const scoresSheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  const mvpSheet = ss.getSheetByName(SHEET_NAMES.MVP_SCORES);
  
  // 기존 점수 데이터 삭제 (헤더 제외)
  const lastRow = scoresSheet.getLastRow();
  if (lastRow > 1) {
    scoresSheet.deleteRows(2, lastRow - 1);
  }
  
  // 모든 응답 가져오기
  const responses = responsesSheet.getDataRange().getValues();
  const teamScores = {};
  const teamMembers = {};
  const teamMvpScores = {};
  
  // 팀별 구성원 수 계산
  const participants = getParticipants();
  participants.forEach(p => {
    if (!teamMembers[p.team]) {
      teamMembers[p.team] = 0;
    }
    teamMembers[p.team]++;
  });
  
  // 응답 데이터 처리 (헤더 제외)
  // 컬럼 구조: 타임스탬프, 평가자이름, 평가자팀, 평가받은팀, 11개항목점수..., 총점, MVP점수
  const totalScoreIndex = 3 + SCORING_ITEMS.length + 1; // 평가받은팀 다음 + 11개 항목 + 1
  const mvpScoreIndex = totalScoreIndex + 1;
  
  for (let i = 1; i < responses.length; i++) {
    const team = responses[i][3]; // 평가받은팀
    const totalScore = parseFloat(responses[i][totalScoreIndex]) || 0; // 총점
    const mvpScore = parseFloat(responses[i][mvpScoreIndex]) || 0; // MVP 점수
    
    if (!teamScores[team]) {
      teamScores[team] = 0;
      teamMvpScores[team] = 0;
    }
    teamScores[team] += totalScore;
    teamMvpScores[team] += mvpScore;
  }
  
  // 점수 계산 및 저장
  const scoreData = [];
  Object.keys(teamScores).forEach(team => {
    const totalScore = teamScores[team];
    const memberCount = teamMembers[team] || 1;
    const averageScore = totalScore / memberCount;
    const mvpScore = teamMvpScores[team] / memberCount; // MVP 평균 점수
    
    scoreData.push([team, totalScore, memberCount, averageScore, mvpScore]);
  });
  
  // 점수순으로 정렬
  scoreData.sort((a, b) => b[3] - a[3]);
  
  // 시트에 저장
  if (scoreData.length > 0) {
    scoresSheet.getRange(2, 1, scoreData.length, 5).setValues(scoreData);
  }
  
  // 요약 시트 업데이트
  updateSummary();
}

/**
 * 요약 시트 업데이트
 */
function updateSummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const summarySheet = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
  const scoresSheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  
  summarySheet.clear();
  
  // 헤더
  summarySheet.getRange(1, 1, 1, 2).setValues([['항목', '값']]);
  summarySheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  
  // 참가자 통계
  const participants = getParticipants();
  const totalParticipants = participants.length;
  const completedParticipants = participants.filter(p => p.completed).length;
  
  const summaryData = [
    ['총 참가자 수', totalParticipants],
    ['완료한 참가자 수', completedParticipants],
    ['미완료 참가자 수', totalParticipants - completedParticipants],
    ['완료율', `${((completedParticipants / totalParticipants) * 100).toFixed(1)}%`],
    ['', ''],
    ['팀별 최종 점수', '']
  ];
  
  // 팀별 점수 추가
  const scores = scoresSheet.getDataRange().getValues();
  for (let i = 1; i < scores.length; i++) {
    summaryData.push([scores[i][0], `${scores[i][3].toFixed(2)}점 (MVP: ${scores[i][4].toFixed(2)}점)`]);
  }
  
  summarySheet.getRange(2, 1, summaryData.length, 2).setValues(summaryData);
}

/**
 * 실시간 데이터 가져오기 (웹 앱용)
 */
function getDashboardData() {
  const participants = getParticipants();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const scoresSheet = ss.getSheetByName(SHEET_NAMES.SCORES);
  
  const scores = [];
  const scoreData = scoresSheet.getDataRange().getValues();
  for (let i = 1; i < scoreData.length; i++) {
    scores.push({
      team: scoreData[i][0],
      totalScore: scoreData[i][1],
      memberCount: scoreData[i][2],
      averageScore: scoreData[i][3],
      mvpScore: scoreData[i][4] || 0
    });
  }
  
  return {
    participants: participants,
    scores: scores,
    timestamp: new Date().toISOString()
  };
}

/**
 * 웹 앱 doGet 함수
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('자동 채점 시스템 대시보드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTML 파일에서 스크립트/스타일 포함을 위한 함수
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 수동으로 모든 설문 응답을 다시 처리하는 함수
 */
function reprocessAllResponses() {
  try {
    const form = FormApp.openById(FORM_ID);
    const formResponses = form.getResponses();
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const responsesSheet = ss.getSheetByName(SHEET_NAMES.RESPONSES);
    
    // 기존 응답 데이터 삭제 (헤더 제외)
    const lastRow = responsesSheet.getLastRow();
    if (lastRow > 1) {
      responsesSheet.deleteRows(2, lastRow - 1);
    }
    
    // 모든 참가자 완료 여부 초기화
    const participantsSheet = ss.getSheetByName(SHEET_NAMES.PARTICIPANTS);
    const participantData = participantsSheet.getDataRange().getValues();
    for (let i = 1; i < participantData.length; i++) {
      participantsSheet.getRange(i + 1, 3).setValue('미완료');
    }
    
    // 모든 응답 재처리
    formResponses.forEach(formResponse => {
      const event = {
        response: formResponse
      };
      processFormResponse(event);
    });
    
    Logger.log('모든 응답 재처리 완료');
    return '모든 응답이 성공적으로 재처리되었습니다.';
  } catch (error) {
    Logger.log(`재처리 오류: ${error.toString()}`);
    return `오류 발생: ${error.toString()}`;
  }
}

/**
 * 특정 참가자의 완료 여부를 초기화하는 함수
 */
function resetParticipantStatus(name, team) {
  updateParticipantStatus(name, team, false);
  return '완료 여부가 초기화되었습니다.';
}

/**
 * Google Forms 자동 생성 함수
 * 참가자 명단을 기반으로 설문지를 자동으로 생성합니다.
 */
function createFormAutomatically() {
  try {
    // SPREADSHEET_ID 확인
    if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID') {
      throw new Error('SPREADSHEET_ID를 먼저 설정해주세요. Code.gs 파일 상단의 SPREADSHEET_ID를 Google Sheets ID로 변경해주세요.');
    }
    
    // 참가자 명단 가져오기
    let participants;
    try {
      participants = getParticipants();
    } catch (error) {
      if (error.message.includes('시트를 찾을 수 없습니다')) {
        throw new Error('참가자명단 시트가 없습니다. 먼저 "initializeSheets" 함수를 실행해주세요.');
      }
      throw error;
    }
    
    if (participants.length === 0) {
      throw new Error('참가자 명단이 비어있습니다. 먼저 "참가자명단" 시트에 참가자 정보를 입력해주세요.');
    }
    
    // 모든 팀 목록 추출
    const allTeams = [...new Set(participants.map(p => p.team))];
    const allNames = [...new Set(participants.map(p => p.name))];
    
    // 새 설문지 생성
    const form = FormApp.create('팀 평가 설문지');
    form.setDescription(FORM_DESCRIPTION);
    
    // 1. 설명 섹션 (이미 위에서 설정)
    
    // 2. 이름 선택 (드롭다운)
    const nameItem = form.addListItem();
    nameItem.setTitle('이름을 선택해주세요');
    nameItem.setRequired(true);
    nameItem.setChoiceValues(allNames);
    
    // 3. 팀 선택 (드롭다운)
    const teamItem = form.addListItem();
    teamItem.setTitle('소속 팀을 선택해주세요');
    teamItem.setRequired(true);
    teamItem.setChoiceValues(allTeams);
    
    // 4. 각 팀별 채점 항목 생성 (5팀씩 그룹으로 나누기)
    const allItems = getAllScoringItems();
    const teamsPerPage = 5;  // 한 번에 평가할 팀 수
    
    // 팀을 5팀씩 그룹으로 나누기
    for (let i = 0; i < allTeams.length; i += teamsPerPage) {
      const teamGroup = allTeams.slice(i, i + teamsPerPage);
      const pageNumber = Math.floor(i / teamsPerPage) + 1;
      const totalPages = Math.ceil(allTeams.length / teamsPerPage);
      
      // 페이지 구분 섹션 (첫 번째 그룹이 아니면)
      if (i > 0) {
        const pageBreak = form.addPageBreakItem();
        pageBreak.setTitle(`--- ${pageNumber}페이지 / 총 ${totalPages}페이지 ---`);
        pageBreak.setHelpText(`다음 ${teamGroup.length}개 팀을 평가해주세요.`);
      }
      
      // 각 팀별 평가 항목 추가
      teamGroup.forEach(team => {
        // 팀별 섹션 헤더
        const teamSection = form.addSectionHeaderItem();
        teamSection.setTitle(`${team} 평가 문항입니다.`);
        teamSection.setHelpText('창의성, 구현 완성도, 필요성, 발표 전달력 부문이 있습니다.');
        
        // 각 카테고리별로 항목 추가
        SCORING_ITEMS.forEach(category => {
          // 카테고리별 항목 추가
          category.items.forEach(item => {
            const scaleItem = form.addScaleItem();
            const scaleMax = item.scaleMax || 10;  // 발표 전달력은 5, 나머지는 10
            
            // 제목 형식: [카테고리] 항목번호. 항목명
            const itemIndex = category.items.indexOf(item) + 1;
            const categoryNum = SCORING_ITEMS.indexOf(category) + 1;
            scaleItem.setTitle(`[${category.category}] ${categoryNum}-${itemIndex}. ${item.name}`);
            
            scaleItem.setBounds(1, scaleMax);
            scaleItem.setRequired(true);
            scaleItem.setLabels('매우 그렇지 않다', '매우 그렇다');
            
            if (item.description) {
              scaleItem.setHelpText(item.description);
            }
          });
        });
      });
    }
    
    // 설문지 설정
    form.setCollectEmail(false);
    form.setAllowResponseEdits(true);
    form.setShowLinkToRespondAgain(false);
    
    const formUrl = form.getPublishedUrl();
    const formId = form.getId();
    
    Logger.log(`설문지 생성 완료!`);
    Logger.log(`URL: ${formUrl}`);
    Logger.log(`ID: ${formId}`);
    
    // Code.gs의 FORM_ID 업데이트 안내
    Logger.log(`\n중요: Code.gs 파일의 FORM_ID를 다음으로 업데이트하세요: ${formId}`);
    
    return {
      success: true,
      formUrl: formUrl,
      formId: formId,
      message: `설문지가 성공적으로 생성되었습니다!\n\nURL: ${formUrl}\n\n중요: Code.gs 파일의 FORM_ID를 "${formId}"로 업데이트하세요.`
    };
  } catch (error) {
    const errorMessage = error.toString();
    Logger.log(`설문지 생성 오류: ${errorMessage}`);
    
    // 사용자 친화적인 오류 메시지
    let userMessage = `오류 발생: ${errorMessage}`;
    
    if (errorMessage.includes('SPREADSHEET_ID')) {
      userMessage = 'SPREADSHEET_ID를 먼저 설정해주세요.\nCode.gs 파일 상단의 SPREADSHEET_ID를 Google Sheets ID로 변경해주세요.';
    } else if (errorMessage.includes('시트를 찾을 수 없습니다')) {
      userMessage = '참가자명단 시트가 없습니다.\n먼저 "initializeSheets" 함수를 실행해주세요.';
    } else if (errorMessage.includes('참가자 명단이 비어있습니다')) {
      userMessage = '참가자 명단이 비어있습니다.\n"참가자명단" 시트에 참가자 정보를 입력해주세요.';
    }
    
    return {
      success: false,
      message: userMessage
    };
  }
}

/**
 * 기존 설문지에 팀별 채점 항목 추가 (이미 설문지가 있는 경우)
 */
function updateExistingForm() {
  try {
    if (FORM_ID === 'YOUR_FORM_ID') {
      throw new Error('FORM_ID를 먼저 설정해주세요.');
    }
    
    const form = FormApp.openById(FORM_ID);
    const participants = getParticipants();
    const allTeams = [...new Set(participants.map(p => p.team))];
    
    // 기존 항목 확인
    const items = form.getItems();
    
    // 각 팀별 채점 항목 추가
    allTeams.forEach(team => {
      // 팀별 섹션 추가
      const section = form.addSectionHeaderItem();
      section.setTitle(`${team} 평가`);
      section.setHelpText('각 항목을 1-10점 사이로 평가해주세요. (총 100점 만점)');
      
      // 11개 채점 항목 추가
      SCORING_ITEMS.forEach((item, index) => {
        const scaleItem = form.addScaleItem();
        scaleItem.setTitle(`${item}`);
        scaleItem.setBounds(1, 10);
        scaleItem.setRequired(true);
        scaleItem.setLabels('최저', '최고');
      });
      
      // MVP 점수 추가
      const mvpItem = form.addScaleItem();
      mvpItem.setTitle(`${team} - MVP 점수`);
      mvpItem.setBounds(1, 10);
      mvpItem.setRequired(true);
      mvpItem.setLabels('최저', '최고');
    });
    
    Logger.log('설문지 업데이트 완료');
    return '설문지가 성공적으로 업데이트되었습니다.';
  } catch (error) {
    Logger.log(`설문지 업데이트 오류: ${error.toString()}`);
    return `오류 발생: ${error.toString()}`;
  }
}

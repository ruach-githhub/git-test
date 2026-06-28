function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('2026 교사 일정, 회의록 및 예배순서')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 스프레드시트에서 데이터를 가져오는 함수
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appData = { meetings: [], schedules: [], worships: [] };
  
  // 1. 회의록 데이터 ('2026년' 탭)
  const meetingSheet = ss.getSheetByName("2026년");
  if (meetingSheet) {
    const mValues = meetingSheet.getDataRange().getValues();
    for (let i = 1; i < mValues.length; i++) {
      if (!mValues[i][0]) continue;
      appData.meetings.push({
        date: Utilities.formatDate(new Date(mValues[i][0]), "GMT+9", "yyyy-MM-dd"),
        agenda1: mValues[i][1] || '',
        agenda2: mValues[i][2] || '',
        agenda3: mValues[i][3] || '',
        notice: mValues[i][4] || ''
      });
    }
  }

  // 2. 일정 데이터 ('2026년 일정' 탭)
  const scheduleSheet = ss.getSheetByName("2026년 일정");
  if (scheduleSheet) {
    const sValues = scheduleSheet.getDataRange().getValues();
    for (let i = 1; i < sValues.length; i++) {
      if (!sValues[i][1]) continue; 
      appData.schedules.push({
        date: Utilities.formatDate(new Date(sValues[i][1]), "GMT+9", "yyyy-MM-dd"),
        count: Number(sValues[i][3]) || 0, 
        event1: sValues[i][4] || '',       
        event2: sValues[i][5] || '',       
        event3: sValues[i][6] || ''        
      });
    }
  }

  // 3. 예배 데이터 ('예배' 탭) - 열 추가됨
  const worshipSheet = ss.getSheetByName("예배");
  if (worshipSheet) {
    const wValues = worshipSheet.getDataRange().getValues();
    // 열 구조: YEAR(0), DAY(1), 찬양곡(2), 기도자(3), 설교 제목(4), 설교 본문(5), 
    // Connect(6), Insight(7), Real Life(8), Last week(9), This week(10), This Month(11)
    for (let i = 1; i < wValues.length; i++) {
      if (!wValues[i][1]) continue;
      appData.worships.push({
        date: Utilities.formatDate(new Date(wValues[i][1]), "GMT+9", "yyyy-MM-dd"),
        song: wValues[i][2] || '',
        prayer: wValues[i][3] || '',
        sermonTitle: wValues[i][4] || '',
        sermonText: wValues[i][5] || '',
        connect: wValues[i][6] || '',
        insight: wValues[i][7] || '',
        realLife: wValues[i][8] || '',
        lastWeek: wValues[i][9] || '',
        thisWeek: wValues[i][10] || '',
        thisMonth: wValues[i][11] || ''
      });
    }
  }

  return appData;
}

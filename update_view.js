function insertDataToSheet(data) {
  if (!scoreInputSheet) {
    SpreadsheetApp.getUi().alert('시트1을 찾을 수 없습니다.');
    return;
  }

  // 현재 날짜/시간을 가져와 "YYYY-MM-DD HH:mm" 형식으로 변환
  let now = new Date();
  let formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

  // 시트1의 마지막 행 찾기 (새로운 데이터가 추가될 위치)
  let lastRow = scoreInputSheet.getLastRow();
  let targetRow = lastRow + 1;

  // 경기 데이터 입력 (A열 ~ G열)
  scoreInputSheet.getRange(targetRow, 1).setValue(data.p1);       // 선수1
  scoreInputSheet.getRange(targetRow, 2).setValue(data.p2);       // 선수2
  scoreInputSheet.getRange(targetRow, 3).setValue('vs');          // "vs" 또는 빈칸
  scoreInputSheet.getRange(targetRow, 4).setValue(data.p3);       // 선수3
  scoreInputSheet.getRange(targetRow, 5).setValue(data.p4);       // 선수4
  scoreInputSheet.getRange(targetRow, 6).setValue(data.score);    // 점수 (예: "6:4")
  scoreInputSheet.getRange(targetRow, 7).setValue(formattedTime); // 경기 입력 시간
}


/**
 * 시트1에서 '초록색(#ccffcc)'으로 칠해진 칸들을 흰색(#ffffff)으로 되돌린다
 */
function revertSheet1Colors() {
  if (!scoreInputSheet) return;

  let lastRowSheet1 = scoreInputSheet.getLastRow();
  let lastColSheet1 = scoreInputSheet.getLastColumn();
  if (lastRowSheet1 < 2 || lastColSheet1 < 1) {
    return; // 실제 데이터가 없음
  }

  // 2행부터 마지막 행까지, 전체 열 범위를 가져옴
  let range = scoreInputSheet.getRange(2, 1, lastRowSheet1 - 1, lastColSheet1);
  let backgrounds = range.getBackgrounds();

  for (let r = 0; r < backgrounds.length; r++) {
    for (let c = 0; c < backgrounds[r].length; c++) {
      if (backgrounds[r][c] === '#ccffcc') {
        // 초록색 -> 흰색
        backgrounds[r][c] = '#ffffff';
      }
    }
  }

  // 수정된 색상 배열을 다시 Range에 반영
  range.setBackgrounds(backgrounds);
}

/**
 * (6) 시트2 업데이트 (기존 예시와 동일)
 */
function updateSheet2(sheet2, oldData, ratingMap) {
  for (let name in ratingMap) {
    let raw = ratingMap[name].NTR;
    // 1) 반올림
    raw = parseFloat(raw.toFixed(2));

    // 2) 범위(clamp)
    ratingMap[name].NTR = clamp(raw, 1.0, 16.5);
  }

  let existingNames = {};
  for (let i = 1; i < oldData.length; i++) {
    let name = oldData[i][0];
    if (name) {
      existingNames[name] = i;
    }
  }

  let newData = [];
  // 기존 행 업데이트
  for (let i = 1; i < oldData.length; i++) {
    let row = oldData[i];
    let name = row[0];
    if (!name) continue;
    if (ratingMap[name]) {
      row[1] = ratingMap[name].NTR;     // NTR
      row[2] = ratingMap[name].matches; // 경기수
    }
    newData.push(row);
  }

  // 신규 선수 추가
  for (let name in ratingMap) {
    if (existingNames[name] == null) {
      newData.push([
        name,
        ratingMap[name].NTR,
        ratingMap[name].matches
      ]);
    }
  }

  // ✅ NTR 기준으로 내림차순 정렬
  newData.sort((a, b) => b[1] - a[1]); // NTR 값(b[1])을 기준으로 내림차순 정렬

  // 시트2 초기화 (헤더 제외)
  sheet2.clearContents();
  sheet2.getRange(1,1).setValue('이름');
  sheet2.getRange(1,2).setValue('NTR');
  sheet2.getRange(1,3).setValue('경기수');

  if (newData.length > 0) {
    sheet2.getRange(2,1,newData.length,3).setValues(newData);
  }
}


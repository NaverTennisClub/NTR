
/**
 * 점수 차(diff)에 따라 margin_factor를 세분화 (예시)
 */
function calculateMarginFactor(scoreA, scoreB) {
  let diff = Math.abs(scoreA - scoreB);
  switch (diff) {
    case 0:
      return 1.0;
    case 1:
      return 0.2;
    case 2:
      return 0.4;
    case 3:
      return 0.6;
    case 4:
      return 0.8;
    case 5:
      return 0.9;
    case 6:
      return 1.0;
    default: // diff >= 7
      return 1.0;
  }
}

/**
 * (1) 새로 입력된 행만 NTR 계산
 *     - Script Property 이용, LAST_PROCESSED_ROW 까지 완료된 것으로 처리
 */
function calculateNTR() {
  // 1) 마지막으로 처리한 행 번호
  const props = PropertiesService.getScriptProperties();
  let lastProcessedRow = parseInt(props.getProperty('LAST_PROCESSED_ROW')) || 2;
  let lastRow = scoreInputSheet.getLastRow();

  // 새로 입력된 행이 없으면 종료
  if (lastRow <= lastProcessedRow) {
    SpreadsheetApp.getUi().alert('새로 추가된 데이터가 없습니다.');
    return;
  }

  // 2) 시트2 -> ratingMap 구성
  let ratingData = NTRScoreSheet.getDataRange().getValues();
  let ratingMap = {};
  for (let i = 1; i < ratingData.length; i++) {
    let row = ratingData[i];
    let name = row[0];
    let ntr = row[1];
    let matches = row[2] || 0;
    if (name) {
      ratingMap[name] = { NTR: ntr, matches: matches };
    }
  }

  // 3) 새로 추가된 경기 읽고, NTR 계산(온라인 업데이트 예시)
  let matches = [];
  for (let row = lastProcessedRow + 1; row <= lastRow; row++) {
    let rowData = scoreInputSheet.getRange(row, 1, 1, 6).getValues()[0];
    let p1 = rowData[0]; // 선수1
    let p2 = rowData[1]; // 선수2
    let p3 = rowData[3]; // 선수3
    let p4 = rowData[4]; // 선수4
    let scoreStr = rowData[5]; // ex) "4:6"

    // 입력 누락된 셀이 있으면 스킵
    if (!p1 || !p2 || !p3 || !p4 || !scoreStr) continue;

    let [scoreA, scoreB] = scoreStr.split(':').map(s => parseInt(s, 10));
    if (isNaN(scoreA) || isNaN(scoreB)) continue;

    // 신규 선수 초기값 세팅
    if (!ratingMap[p1]) ratingMap[p1] = { NTR: 3.0, matches: 0 };
    if (!ratingMap[p2]) ratingMap[p2] = { NTR: 3.0, matches: 0 };
    if (!ratingMap[p3]) ratingMap[p3] = { NTR: 3.0, matches: 0 };
    if (!ratingMap[p4]) ratingMap[p4] = { NTR: 3.0, matches: 0 };

    // 경기수 +1
    ratingMap[p1].matches++;
    ratingMap[p2].matches++;
    ratingMap[p3].matches++;
    ratingMap[p4].matches++;

    // (A) 실제 Elo 계산 (한 경기씩 즉시 반영)
    //   - runRatingAlgorithm()를 호출해도 되고,
    //   - 여기서 직접 계산 로직을 수행해도 됨
    //   - 여기서는 간단히 '한 경기만' 계산하는 함수를 호출하는 예시

    processOneMatch(ratingMap, {
      teamA: [p1, p2],
      teamB: [p3, p4],
      scoreA: scoreA,
      scoreB: scoreB
    });

    // (B) 처리 후, 시트1 해당 행을 초록색으로 칠함
    scoreInputSheet.getRange(row, 1, 1, 6).setBackground('#ccffcc');
  }

  // 4) 처리 완료 행 갱신
  props.setProperty('LAST_PROCESSED_ROW', lastRow.toString());

  // 5) 시트2에 반영
  updateSheet2(NTRScoreSheet, ratingData, ratingMap);

  SpreadsheetApp.getUi().alert('신규 입력분 NTR 계산 및 색상 표시 완료!');
}


/**
 * 한 경기(NTR) 처리 로직 (예시)
 */
function processOneMatch(ratingMap, game) {
  let ScaleFactor = 10;
  let K = 1.0;
  let level_diff_threshold = 2.0;
  let upset_factor = 1.3;

  // 팀 레이팅
  let teamA_rating = (ratingMap[game.teamA[0]].NTR + ratingMap[game.teamA[1]].NTR) / 2;
  let teamB_rating = (ratingMap[game.teamB[0]].NTR + ratingMap[game.teamB[1]].NTR) / 2;

  // 기대 승률
  let expectedA = 1 / (1 + Math.pow(10, (teamB_rating - teamA_rating) / ScaleFactor));

  // 실제 결과
  let actualA = 0.5;
  if (game.scoreA > game.scoreB) {
    actualA = 1.0;
  } else if (game.scoreA < game.scoreB) {
    actualA = 0.0;
  }
  let margin_factor = calculateMarginFactor(game.scoreA, game.scoreB);

  // 기본 diff
  let rating_diff_A = K * (actualA - expectedA) * margin_factor;
  let rating_diff_B = -rating_diff_A;

  // 업셋 보정
  let rating_gap = teamA_rating - teamB_rating;
  if (rating_gap >= level_diff_threshold && actualA === 0.0) {
    rating_diff_A *= upset_factor;
    rating_diff_B *= upset_factor;
  }
  if (rating_gap <= -level_diff_threshold && actualA === 1.0) {
    rating_diff_A *= upset_factor;
    rating_diff_B *= upset_factor;
  }

  // 업데이트 (클램핑)
  ratingMap[game.teamA[0]].NTR = clamp(ratingMap[game.teamA[0]].NTR + rating_diff_A, 1.0, 16.5);
  ratingMap[game.teamA[1]].NTR = clamp(ratingMap[game.teamA[1]].NTR + rating_diff_A, 1.0, 16.5);
  ratingMap[game.teamB[0]].NTR = clamp(ratingMap[game.teamB[0]].NTR + rating_diff_B, 1.0, 16.5);
  ratingMap[game.teamB[1]].NTR = clamp(ratingMap[game.teamB[1]].NTR + rating_diff_B, 1.0, 16.5);
}
/**
 * (2) 전체 재계산
 *   - 시트2의 기존 내용(헤더 제외)을 싹 지우고
 *   - LAST_PROCESSED_ROW를 1로 되돌린 뒤
 *   - 시트1의 "초록색" 부분을 흰색으로 원복
 *   - 시트1의 모든 데이터(2행~끝)를 다시 계산
 */
function recalcAllNTR() {
  // (A) 시트2 헤더(1행) 제외하고 지우기
  let lastRow = NTRScoreSheet.getLastRow();
  if (lastRow > 1) {
    NTRScoreSheet.getRange(2, 1, lastRow - 1, NTRScoreSheet.getLastColumn()).clearContent();
  }

  // (B) LAST_PROCESSED_ROW 리셋 → 다시 2행부터 처리하도록
  PropertiesService.getScriptProperties().setProperty('LAST_PROCESSED_ROW', '1');

  // (C) 시트1의 "초록색(#ccffcc)" 부분을 흰색(#ffffff)으로 원복
  revertSheet1Colors();

  // (D) 이제 calculateNTR() 호출 → 시트1의 2행 ~ 끝 전부 다시 계산
  calculateNTR();
  // (알림은 calculateNTR() 내부에서 이미 처리)
}



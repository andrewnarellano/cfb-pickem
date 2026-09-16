// CFB Pick'em — Google Apps Script Backend

const SHEET_PICKS  = 'Picks';
const SHEET_CONFIG = 'WeekConfig';
const SHEET_BONUS  = 'BonusQuestions';
const ADMIN_PASSWORD = 'aggies2026';

const PLAYERS = [
  'Michael Andres', 'Andrew Arellano', 'Jimbo Wilhite', 'Kendall Bicknell',
  'Andrew Johnson', 'Matt Countryman', 'Matt Recko', 'Mark Gold',
  'Ryce Recko', 'Adam Glenn', 'Tim Dickson', 'Jeff Estes', 'Jack Estes', 'Barrett Fischer'
];

const PICKS_HEADERS  = ['Timestamp','Year','Week','Username','GameId','AwayTeam','HomeTeam','SpreadDetail','Spread','PickedTeam','Confidence','PickType'];
const CONFIG_HEADERS = ['Year','Week','GameId','AwayTeam','HomeTeam','PickType','SpreadDetail','Spread','Total','DisplayOrder','GameDate'];
const BONUS_HEADERS  = ['Year','Week','QuestionId','QuestionText','OptionA','OptionB','CorrectAnswer'];

// Sheet colors
const C_HEADER = '#1f3864';
const C_TEXT   = '#ffffff';
const C_GREEN  = '#d9ead3';
const C_RED    = '#f4cccc';
const C_GRAY   = '#999999';

// Column layout: A=Away/Game, B=Home, C=Winner, D=%, E=spacer, then 3 cols per player
const LEFT_COLS = 5;
const P_WIDTH   = 3;
function pCol(i) { return LEFT_COLS + 1 + i * P_WIDTH; }

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getWeekConfig') return getWeekConfig(e.parameter.week, e.parameter.year);
  if (action === 'getPicks')      return getPicks(e.parameter.week, e.parameter.year);
  if (action === 'getAllPicks')   return getAllPicks();
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'submitPicks')        return submitPicks(data);
    if (data.action === 'saveWeekConfig')     return saveWeekConfig(data);
    if (data.action === 'saveBonusResults')   return saveBonusResults(data);
    if (data.action === 'generateScoreSheet') return generateScoreSheet(data);
    return jsonResponse({ error: 'Unknown action' });
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

// ── Week Config ───────────────────────────────────────────────────────────────

function getWeekConfig(week, year) {
  if (!week || !year) return jsonResponse({ games: [], bonusQuestions: [] });

  const configSheet = getOrCreateSheet(SHEET_CONFIG, CONFIG_HEADERS);
  const bonusSheet  = getOrCreateSheet(SHEET_BONUS, BONUS_HEADERS);

  const configRows = configSheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === String(year) && String(r[1]) === String(week))
    .sort((a, b) => a[9] - b[9])
    .map(r => ({
      year: r[0], week: r[1], gameId: r[2],
      awayTeam: r[3], homeTeam: r[4], pickType: r[5],
      spreadDetail: r[6], spread: r[7], total: r[8], displayOrder: r[9],
      gameDate: r[10] || '',
    }));

  const bonusRows = bonusSheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === String(year) && String(r[1]) === String(week))
    .map(r => ({
      year: r[0], week: r[1], questionId: r[2],
      questionText: r[3], optionA: r[4], optionB: r[5], correctAnswer: r[6],
    }));

  return jsonResponse({ games: configRows, bonusQuestions: bonusRows });
}

function saveWeekConfig(data) {
  if (data.password !== ADMIN_PASSWORD) return jsonResponse({ error: 'Unauthorized' });
  const { week, year, games, bonusQuestions } = data;
  if (!week || !year) return jsonResponse({ error: 'Missing week/year' });

  const configSheet = getOrCreateSheet(SHEET_CONFIG, CONFIG_HEADERS);
  const bonusSheet  = getOrCreateSheet(SHEET_BONUS, BONUS_HEADERS);

  deleteRowsWhere(configSheet, r => String(r[0]) === String(year) && String(r[1]) === String(week));
  deleteRowsWhere(bonusSheet,  r => String(r[0]) === String(year) && String(r[1]) === String(week));

  (games || []).forEach((g, i) => {
    configSheet.appendRow([year, week, g.gameId, g.awayTeam, g.homeTeam, g.pickType, g.spreadDetail || '', g.spread || '', g.total || '', i, g.gameDate || '']);
  });

  (bonusQuestions || []).forEach((q, i) => {
    bonusSheet.appendRow([year, week, `bonus_${i+1}`, q.questionText, q.optionA, q.optionB, '']);
  });

  return jsonResponse({ success: true });
}

function saveBonusResults(data) {
  if (data.password !== ADMIN_PASSWORD) return jsonResponse({ error: 'Unauthorized' });
  const { week, year, results } = data;
  if (!week || !year || !results) return jsonResponse({ error: 'Missing fields' });

  const bonusSheet = getOrCreateSheet(SHEET_BONUS, BONUS_HEADERS);
  const rows = bonusSheet.getDataRange().getValues();

  results.forEach(r => {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(year) && String(rows[i][1]) === String(week) && rows[i][2] === r.questionId) {
        bonusSheet.getRange(i + 1, 7).setValue(r.correctAnswer);
        break;
      }
    }
  });

  return jsonResponse({ success: true });
}

// ── Score Sheet Generation ────────────────────────────────────────────────────

function generateScoreSheet(data) {
  if (data.password !== ADMIN_PASSWORD) return jsonResponse({ error: 'Unauthorized' });
  const { week, year } = data;
  if (!week || !year) return jsonResponse({ error: 'Missing week/year' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Week ${week} - ${year}`;

  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  // Load game config
  const configSheet = getOrCreateSheet(SHEET_CONFIG, CONFIG_HEADERS);
  const bonusSheet  = getOrCreateSheet(SHEET_BONUS,  BONUS_HEADERS);

  const games = configSheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === String(year) && String(r[1]) === String(week))
    .sort((a, b) => a[9] - b[9])
    .map(r => ({
      gameId: String(r[2]), awayTeam: r[3], homeTeam: r[4],
      pickType: r[5], spreadDetail: r[6],
      spread: parseFloat(r[7]) || 0, total: parseFloat(r[8]) || 0
    }));

  const bonusQs = bonusSheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === String(year) && String(r[1]) === String(week))
    .map(r => ({
      questionId: String(r[2]), questionText: r[3],
      optionA: r[4], optionB: r[5], correctAnswer: r[6]
    }));

  // Load picks → pickMap[gameId][username]
  const picksSheet = getOrCreateSheet(SHEET_PICKS, PICKS_HEADERS);
  const allPicks = picksSheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === String(year) && String(r[2]) === String(week))
    .map(r => ({
      username: String(r[3]), gameId: String(r[4]),
      pickedTeam: String(r[9]), confidence: parseInt(r[10]) || 0, pickType: String(r[11] || 'ats')
    }));

  const pickMap = {};
  allPicks.forEach(p => {
    if (!pickMap[p.gameId]) pickMap[p.gameId] = {};
    pickMap[p.gameId][p.username] = p;
  });

  // Fetch ESPN scores → scoreMap[gameId]
  const scoreMap = {};
  try {
    let page = 1, totalPages = 1;
    do {
      const url = `https://site.api.espn.com/apis/site/v2/sports/football/college-football/scoreboard?seasontype=2&week=${week}&season=${year}&groups=80&limit=100&page=${page}`;
      const espn = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
      totalPages = espn.pageCount || 1;
      (espn.events || []).forEach(ev => {
        const comp = ev.competitions[0];
        const home = comp.competitors.find(c => c.homeAway === 'home');
        const away = comp.competitors.find(c => c.homeAway === 'away');
        const st = comp.status?.type;
        scoreMap[String(ev.id)] = {
          homeTeam: home?.team?.displayName || '', awayTeam: away?.team?.displayName || '',
          homeScore: parseInt(home?.score) || 0, awayScore: parseInt(away?.score) || 0,
          completed: st?.completed || false, state: st?.state || 'pre'
        };
      });
      page++;
    } while (page <= totalPages);
  } catch(e) {}

  // ── Helpers ────────────────────────────────────────────────────────────────

  function computeResult(pick, game, bonusQ) {
    if (pick.pickType === 'bonus') {
      if (!bonusQ || !bonusQ.correctAnswer) return null;
      const correct = bonusQ.correctAnswer === 'A' ? bonusQ.optionA : bonusQ.optionB;
      return pick.pickedTeam === correct ? 'WIN' : 'LOSS';
    }
    const sc = scoreMap[game.gameId];
    if (!sc || (!sc.completed && sc.state !== 'in')) return null;
    if (pick.pickType === 'ou') {
      if (!game.total) return null;
      const combined = sc.homeScore + sc.awayScore;
      if (Math.abs(combined - game.total) < 0.01) return 'PUSH';
      return pick.pickedTeam === 'Over' ? (combined > game.total ? 'WIN' : 'LOSS') : (combined < game.total ? 'WIN' : 'LOSS');
    }
    if (pick.pickType === 'su') {
      if (sc.homeScore === sc.awayScore) return 'PUSH';
      const winner = sc.homeScore > sc.awayScore ? sc.homeTeam : sc.awayTeam;
      return pick.pickedTeam === winner ? 'WIN' : 'LOSS';
    }
    // ATS
    const cover = (sc.homeScore - sc.awayScore) + game.spread;
    if (Math.abs(cover) < 0.01) return 'PUSH';
    return pick.pickedTeam === game.homeTeam ? (cover > 0 ? 'WIN' : 'LOSS') : (cover < 0 ? 'WIN' : 'LOSS');
  }

  function computePts(result, confidence, pickType) {
    if (pickType === 'bonus') return result === 'WIN' ? 5 : 0;
    if (result === 'WIN') return confidence;
    if (result === 'LOSS') return (pickType === 'ats' || pickType === 'ou') ? -confidence : 0;
    return 0;
  }

  function getActualWinner(game) {
    const sc = scoreMap[game.gameId];
    if (!sc || (!sc.completed && sc.state !== 'in')) return '';
    if (game.pickType === 'su') {
      if (sc.homeScore === sc.awayScore) return 'TIE';
      return sc.homeScore > sc.awayScore ? sc.homeTeam : sc.awayTeam;
    }
    if (game.pickType === 'ats') {
      const cover = (sc.homeScore - sc.awayScore) + game.spread;
      if (Math.abs(cover) < 0.01) return 'PUSH';
      const hSp = game.spread >= 0 ? `+${game.spread}` : `${game.spread}`;
      const aSp = (-game.spread) >= 0 ? `+${-game.spread}` : `${-game.spread}`;
      return cover > 0 ? `${game.homeTeam} (${hSp})` : `${game.awayTeam} (${aSp})`;
    }
    if (game.pickType === 'ou') {
      if (!game.total) return '';
      const combined = sc.homeScore + sc.awayScore;
      if (Math.abs(combined - game.total) < 0.01) return 'PUSH';
      return combined > game.total ? 'Over' : 'Under';
    }
    return '';
  }

  const playerTotals = {};
  PLAYERS.forEach(p => playerTotals[p] = 0);

  let row = 1;

  // ── Row 1: Title + player name headers ──────────────────────────────────────
  sheet.getRange(row, 1, 1, LEFT_COLS).merge()
    .setValue(`CFB Pick 'Em — Week ${week}`)
    .setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(13);
  PLAYERS.forEach((name, i) => {
    sheet.getRange(row, pCol(i), 1, P_WIDTH).merge()
      .setValue(name)
      .setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
  });
  row++;

  // ── Write one pick section ───────────────────────────────────────────────────

  function writeSection(sectionGames, label, pickType) {
    if (!sectionGames.length) return;

    // Compute per-player section totals for the header score
    const secTotals = {};
    PLAYERS.forEach(p => secTotals[p] = 0);
    sectionGames.forEach(game => {
      PLAYERS.forEach(player => {
        const pick = pickMap[game.gameId]?.[player];
        if (!pick) return;
        secTotals[player] += computePts(computeResult(pick, game, null), pick.confidence, pickType);
      });
    });

    // Section header row
    sheet.getRange(row, 1, 1, LEFT_COLS).merge()
      .setValue(label)
      .setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
    PLAYERS.forEach((player, i) => {
      const c = pCol(i);
      sheet.getRange(row, c).setValue(label).setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold');
      sheet.getRange(row, c + 1).setValue(secTotals[player]).setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold').setHorizontalAlignment('right');
      sheet.getRange(row, c + 2).setBackground(C_HEADER);
    });
    row++;

    // Column sub-headers
    if (pickType === 'ou') {
      sheet.getRange(row, 1).setValue('Game').setFontWeight('bold').setFontDecoration('underline');
    } else {
      sheet.getRange(row, 1).setValue('Away').setFontWeight('bold').setFontDecoration('underline');
      sheet.getRange(row, 2).setValue('Home').setFontWeight('bold').setFontDecoration('underline');
    }
    sheet.getRange(row, 3).setValue('Winner').setFontWeight('bold').setFontDecoration('underline');
    sheet.getRange(row, 4).setValue('%').setFontWeight('bold').setFontDecoration('underline');
    PLAYERS.forEach((_, i) => {
      const c = pCol(i);
      sheet.getRange(row, c).setValue('Winner').setFontWeight('bold');
      sheet.getRange(row, c + 1).setValue('Points').setFontWeight('bold');
    });
    row++;

    // One row per game
    sectionGames.forEach(game => {
      const winner = getActualWinner(game);
      const gamePicks = PLAYERS.map(p => pickMap[game.gameId]?.[p]).filter(Boolean);
      let correct = 0;
      if (winner) gamePicks.forEach(pick => { if (computeResult(pick, game, null) === 'WIN') correct++; });
      const pct = (gamePicks.length && winner) ? `${Math.round(correct / gamePicks.length * 100)}%` : '';

      if (pickType === 'ou') {
        sheet.getRange(row, 1).setValue(`${game.awayTeam} - ${game.homeTeam} (${game.total || '—'})`);
      } else if (pickType === 'ats') {
        const aSp = (-game.spread) >= 0 ? `+${-game.spread}` : `${-game.spread}`;
        const hSp = game.spread >= 0 ? `+${game.spread}` : `${game.spread}`;
        sheet.getRange(row, 1).setValue(`${game.awayTeam} (${aSp})`);
        sheet.getRange(row, 2).setValue(`${game.homeTeam} (${hSp})`);
      } else {
        sheet.getRange(row, 1).setValue(game.awayTeam);
        sheet.getRange(row, 2).setValue(game.homeTeam);
      }
      const wCell = sheet.getRange(row, 3);
      wCell.setValue(winner);
      if (winner) wCell.setFontWeight('bold');
      sheet.getRange(row, 4).setValue(pct).setFontColor(C_GRAY);

      PLAYERS.forEach((player, i) => {
        const pick = pickMap[game.gameId]?.[player];
        const c = pCol(i);
        if (!pick) return;
        const result = computeResult(pick, game, null);
        const pts = result !== null ? computePts(result, pick.confidence, pickType) : '';
        const bg = result === 'WIN' ? C_GREEN : result === 'LOSS' ? C_RED : null;
        sheet.getRange(row, c).setValue(pick.pickedTeam);
        sheet.getRange(row, c + 1).setValue(pick.confidence);
        if (pts !== '') sheet.getRange(row, c + 2).setValue(pts);
        if (bg) sheet.getRange(row, c, 1, P_WIDTH).setBackground(bg);
        playerTotals[player] += (typeof pts === 'number' ? pts : 0);
      });
      row++;
    });
  }

  // ── Write SU / ATS / O/U ─────────────────────────────────────────────────────
  writeSection(games.filter(g => g.pickType === 'su'),  "Pick 'Em Straight Up",        'su');
  writeSection(games.filter(g => g.pickType === 'ats'), "Pick 'Em Against the Spread", 'ats');
  writeSection(games.filter(g => g.pickType === 'ou'),  "Pick 'Em Over/Under",         'ou');

  // ── Bonus section ────────────────────────────────────────────────────────────
  if (bonusQs.length) {
    const bonusTotals = {};
    PLAYERS.forEach(p => bonusTotals[p] = 0);
    bonusQs.forEach(q => {
      PLAYERS.forEach(player => {
        const pick = pickMap[q.questionId]?.[player];
        if (!pick) return;
        bonusTotals[player] += computePts(computeResult(pick, null, q), pick.confidence, 'bonus');
      });
    });

    sheet.getRange(row, 1, 1, LEFT_COLS).merge()
      .setValue("Pick 'Em Bonus")
      .setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold').setHorizontalAlignment('center');
    PLAYERS.forEach((player, i) => {
      const c = pCol(i);
      sheet.getRange(row, c).setValue("Pick 'Em Bonus").setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold');
      sheet.getRange(row, c + 1).setValue(bonusTotals[player]).setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold').setHorizontalAlignment('right');
      sheet.getRange(row, c + 2).setBackground(C_HEADER);
    });
    row++;

    sheet.getRange(row, 1).setValue('Question').setFontWeight('bold').setFontDecoration('underline');
    sheet.getRange(row, 3).setValue('Answer').setFontWeight('bold').setFontDecoration('underline');
    sheet.getRange(row, 4).setValue('%').setFontWeight('bold').setFontDecoration('underline');
    PLAYERS.forEach((_, i) => {
      const c = pCol(i);
      sheet.getRange(row, c).setValue('Pick').setFontWeight('bold');
      sheet.getRange(row, c + 1).setValue('Pts (+5)').setFontWeight('bold');
    });
    row++;

    bonusQs.forEach(q => {
      const correct = q.correctAnswer ? (q.correctAnswer === 'A' ? q.optionA : q.optionB) : '';
      const qPicks = PLAYERS.map(p => pickMap[q.questionId]?.[p]).filter(Boolean);
      let qCorrect = 0;
      if (correct) qPicks.forEach(p => { if (computeResult(p, null, q) === 'WIN') qCorrect++; });
      const pct = (qPicks.length && correct) ? `${Math.round(qCorrect / qPicks.length * 100)}%` : '';

      sheet.getRange(row, 1).setValue(q.questionText);
      const aCell = sheet.getRange(row, 3);
      aCell.setValue(correct);
      if (correct) aCell.setFontWeight('bold');
      sheet.getRange(row, 4).setValue(pct).setFontColor(C_GRAY);

      PLAYERS.forEach((player, i) => {
        const pick = pickMap[q.questionId]?.[player];
        const c = pCol(i);
        if (!pick) return;
        const result = computeResult(pick, null, q);
        const pts = result !== null ? computePts(result, pick.confidence, 'bonus') : '';
        const bg = result === 'WIN' ? C_GREEN : result === 'LOSS' ? C_RED : null;
        sheet.getRange(row, c).setValue(pick.pickedTeam);
        sheet.getRange(row, c + 1).setValue(pts !== '' ? pts : '');
        if (bg) sheet.getRange(row, c, 1, P_WIDTH).setBackground(bg);
        playerTotals[player] += (typeof pts === 'number' ? pts : 0);
      });
      row++;
    });
  }

  // ── Correct Pick Bonus row ────────────────────────────────────────────────────
  const BONUS_PTS = 10;
  sheet.getRange(row, 1, 1, LEFT_COLS).merge()
    .setValue('Correct Pick Bonus (Picks from non-bonus categories)')
    .setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold');
  PLAYERS.forEach((player, i) => {
    const c = pCol(i);
    const submitted = allPicks.some(p => p.username === player);
    const bonus = submitted ? BONUS_PTS : 0;
    sheet.getRange(row, c).setValue('Correct Pick Bonus').setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold');
    sheet.getRange(row, c + 1).setValue(bonus).setBackground(C_HEADER).setFontColor(C_TEXT).setFontWeight('bold').setHorizontalAlignment('right');
    sheet.getRange(row, c + 2).setBackground(C_HEADER);
    playerTotals[player] += bonus;
  });
  row++;

  // ── Totals row ────────────────────────────────────────────────────────────────
  sheet.getRange(row, 1).setValue('TOTALS:').setFontWeight('bold');
  PLAYERS.forEach((player, i) => {
    sheet.getRange(row, pCol(i) + 1).setValue(playerTotals[player]).setFontWeight('bold');
  });

  // ── Column widths ─────────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 55);
  sheet.setColumnWidth(5, 8);
  PLAYERS.forEach((_, i) => {
    const c = pCol(i);
    sheet.setColumnWidth(c, 140);
    sheet.setColumnWidth(c + 1, 55);
    sheet.setColumnWidth(c + 2, 45);
  });

  return jsonResponse({ success: true, sheetName });
}

// ── Picks ─────────────────────────────────────────────────────────────────────

function submitPicks(data) {
  const { username, week, year, picks } = data;
  if (!username || !week || !year || !picks?.length) return jsonResponse({ error: 'Missing fields' });

  const sheet = getOrCreateSheet(SHEET_PICKS, PICKS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const gameIds = picks.map(p => String(p.gameId));

  const toDelete = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const row = rows[i];
    if (String(row[1]) === String(year) && String(row[2]) === String(week) &&
        row[3] === username && gameIds.includes(String(row[4]))) {
      toDelete.push(i + 1);
    }
  }
  toDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));

  const now = new Date();
  picks.forEach(p => {
    sheet.appendRow([
      now, year, week, username,
      p.gameId, p.awayTeam || '', p.homeTeam || '',
      p.spreadDetail || '', p.spread || '',
      p.pickedTeam, p.confidence, p.pickType || 'ats',
    ]);
  });

  return jsonResponse({ success: true, count: picks.length });
}

function getPicks(week, year) {
  if (!week || !year) return jsonResponse([]);
  const sheet = getOrCreateSheet(SHEET_PICKS, PICKS_HEADERS);
  const result = sheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === String(year) && String(r[2]) === String(week))
    .map(rowToPickObj);
  return jsonResponse(result);
}

function getAllPicks() {
  const sheet = getOrCreateSheet(SHEET_PICKS, PICKS_HEADERS);
  return jsonResponse(sheet.getDataRange().getValues().slice(1).map(rowToPickObj));
}

function rowToPickObj(r) {
  return {
    year: r[1], week: r[2], username: r[3], gameId: r[4],
    awayTeam: r[5], homeTeam: r[6], spreadDetail: r[7], spread: r[8],
    pickedTeam: r[9], confidence: r[10], pickType: r[11] || 'ats',
  };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function deleteRowsWhere(sheet, predicate) {
  const rows = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    if (predicate(rows[i])) toDelete.push(i + 1);
  }
  toDelete.forEach(r => sheet.deleteRow(r));
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

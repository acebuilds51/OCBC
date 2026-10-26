// ================================================================
//  OCBC Golf League — Google Apps Script v5
//  Paste into Extensions > Apps Script in your Google Sheet
//  Deploy > Manage deployments > edit the existing Web app > New version
//  (editing the existing deployment keeps the same URL the app uses)
//  v5: groups with fewer than 4 players no longer crash the Rounds sheet
// ================================================================

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'ping') return jsonResponse({ ok: true, message: 'OCBC Golf League connected!' });
  if (action === 'getAllRounds') return getAllRounds(e.parameter.year);
  if (action === 'getConfig') return getConfig(e.parameter.year);
  if (action === 'getLive') return getLive(e.parameter.year, e.parameter.week);
  if (action === 'testWrite') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Config');
    if (!sheet) return jsonResponse({ ok: false, error: 'No Config sheet' });
    setConfigValue(sheet, 'testWrite_ts', new Date().toISOString());
    return jsonResponse({ ok: true, message: 'Test write successful' });
  }
  return jsonResponse({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === 'saveRound') return saveRound(payload.data);
    if (payload.action === 'saveConfig') return saveConfig(payload.data);
    if (payload.action === 'saveLive') return saveLive(payload.data);
    if (payload.action === 'clearLive') return clearLiveSheet(payload.year, payload.week);
    if (payload.action === 'renamePlayer') return renamePlayerInSheet(payload.oldName, payload.newName, payload.year);
    return jsonResponse({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ================================================================
//  GET ALL ROUNDS — called on app startup for two-way sync
//  Returns all round data for a given year so every phone
//  gets the same leaderboard and history on load
// ================================================================
function getAllRounds(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  year = year || new Date().getFullYear().toString();

  // Always read config regardless of whether rounds exist
  var config = readConfig(ss, year);

  // Read raw round data
  var rounds = [];
  var rawSheet = ss.getSheetByName('RawRounds ' + year);
  if (rawSheet && rawSheet.getLastRow() >= 2) {
    var data = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 1).getValues();
    data.forEach(function(row) {
      try {
        var wd = JSON.parse(row[0]);
        rounds.push(wd);
      } catch(e) {}
    });
  }

  return jsonResponse({ ok: true, rounds: rounds, config: config });
}

// ================================================================
//  SAVE ROUND
// ================================================================
function saveRound(wd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = wd.year || new Date().getFullYear().toString();

  // 1. Save raw JSON blob for two-way sync
  saveRawRound(ss, wd, year);

  // 2. Save formatted rows to readable sheets
  saveToRoundsSheet(ss, wd, year);
  updatePlayerStats(ss, wd, year);
  writeWeekTab(ss, wd);

  // 3. Update config (advance week number)
  updateConfig(ss, wd, year);

  return jsonResponse({ ok: true });
}

// ── RAW ROUND STORAGE (for two-way sync) ─────────────────────
function saveRawRound(ss, wd, year) {
  var sheetName = 'RawRounds ' + year;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Round JSON (do not edit)', 'Week', 'Date', 'Tee Time']);
    sheet.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 800);
  }
  // Match by roundId (week+side+teeTime+sorted player names) so a different
  // group claiming the same tee time number creates a NEW row instead of
  // silently overwriting another group's already-submitted scores.
  var wdRoundId = wd.roundId || (wd.week + '_' + wd.side + '_' + wd.teeTime + '_' +
    (wd.players || []).map(function(p){ return p.name; }).sort().join('|'));
  var existing = sheet.getDataRange().getValues();
  for (var r = 1; r < existing.length; r++) {
    try {
      var ex = JSON.parse(existing[r][0]);
      var exRoundId = ex.roundId || (ex.week + '_' + ex.side + '_' + ex.teeTime + '_' +
        (ex.players || []).map(function(p){ return p.name; }).sort().join('|'));
      if (exRoundId === wdRoundId) {
        // Same group resubmitting (correction) — update existing row
        sheet.getRange(r+1, 1, 1, 4).setValues([[
          JSON.stringify(wd), wd.week, wd.date, wd.teeTime
        ]]);
        return;
      }
    } catch(e) {}
  }
  sheet.appendRow([JSON.stringify(wd), wd.week, wd.date, wd.teeTime]);
}

// ── CONFIG (week number, side, players) ──────────────────────
function readConfig(ss, year) {
  var sheet = ss.getSheetByName('Config');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var cfg = {};
  data.forEach(function(row) { if (row[0]) cfg[row[0]] = row[1]; });
  // Return ALL config keys so the app can read anything stored
  var result = {
    currentWeek: cfg['currentWeek_' + year] ? parseInt(cfg['currentWeek_' + year]) : null,
    currentSide: cfg['currentSide_' + year] || null,
    players: cfg['players'] ? cfg['players'] : null,
    subs: cfg['subs'] ? cfg['subs'] : null
  };
  // Pass through any dynamic keys (sideOverride, rainouts, etc.)
  Object.keys(cfg).forEach(function(k) {
    if (!result.hasOwnProperty(k)) result[k] = cfg[k];
  });
  return result;
}

function updateConfig(ss, wd, year) {
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value']);
    sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
  }
  var nextWeek = (wd.week || 1) + 1;
  var nextSide = wd.side === 'front' ? 'back' : 'front';
  setConfigValue(sheet, 'currentWeek_' + year, nextWeek);
  setConfigValue(sheet, 'currentSide_' + year, nextSide);
}

function saveConfig(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value', 'Description']);
    sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  Object.keys(data).forEach(function(k) {
    setConfigValue(sheet, k, data[k]);
  });
  // Also write a human-readable roster tab for easy viewing
  if (data.players) {
    writeRosterTab(ss, data.players, data.subs || '[]');
  }
  return jsonResponse({ ok: true });
}

function writeRosterTab(ss, playersJson, subsJson) {
  var sheet = ss.getSheetByName('Roster');
  if (!sheet) {
    sheet = ss.insertSheet('Roster');
  } else {
    sheet.clearContents();
  }
  sheet.appendRow(['Player Name', 'Type']);
  sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  try {
    var players = typeof playersJson === 'string' ? JSON.parse(playersJson) : playersJson;
    players.forEach(function(p) { sheet.appendRow([p, 'Regular']); });
  } catch(e) {}
  try {
    var subs = typeof subsJson === 'string' ? JSON.parse(subsJson) : subsJson;
    subs.forEach(function(p) { sheet.appendRow([p, 'Sub']); });
  } catch(e) {}
  sheet.autoResizeColumns(1, 2);
}

function setConfigValue(sheet, key, value) {
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (data[r][0] === key) {
      sheet.getRange(r+1, 2).setValue(typeof value === 'object' ? JSON.stringify(value) : value);
      return;
    }
  }
  sheet.appendRow([key, typeof value === 'object' ? JSON.stringify(value) : value]);
}

function getConfig(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return jsonResponse({ ok: true, config: readConfig(ss, year || new Date().getFullYear().toString()) });
}

// ── FORMATTED ROUNDS SHEET ────────────────────────────────────
function saveToRoundsSheet(ss, wd, year) {
  var sheetName = 'Rounds ' + year;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [
      'Week','Date','Side','Tee Time',
      'P1','P1 Sub','P2','P2 Sub','P3','P3 Sub','P4','P4 Sub',
      'P1 Score','P2 Score','P3 Score','P4 Score',
      'Team A Holes','Team B Holes','Match Winner',
      'Presses','Press Detail',
      'Skins','Skin Payout',
      'CTP Winners',
      'P1 Money','P2 Money','P3 Money','P4 Money'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length)
      .setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  var players = wd.players || [];
  // Groups can have 1-4 players; missing slots must not throw
  var pName = function(i) { return players[i] ? players[i].name : ''; };
  var pSub = function(i) { return players[i] ? (players[i].isSub ? 'Yes' : 'No') : ''; };
  var pMoney = function(i) { return players[i] ? ((wd.money || {})[players[i].name] || 0).toFixed(2) : ''; };
  var totals = [0,0,0,0];
  for (var i = 0; i < wd.scores.length; i++) {
    for (var j = 0; j < players.length; j++) {
      var s = wd.scores[i]['p'+j] || 0;
      if (s < 99) totals[j] += s;
    }
  }
  var skinCnt = Object.keys(wd.skins || {}).length;
  var skinPay = skinCnt > 0 ? (180/skinCnt).toFixed(2) : '0';
  var pressCnt = (wd.presses || []).length;
  var pressDetail = (wd.presses||[]).map(function(p,i){
    return 'P'+(i+1)+':H'+p.startHole+':'+p.calledBy+':'+(p.winner||'tie');
  }).join(', ');
  var ctpStr = Object.entries(wd.ctpWinners||{})
    .filter(function(e){ return e[1] && e[1] !== '__HIGHER_TEE__'; })
    .map(function(e){ return 'H'+e[0]+':'+e[1]; }).join(', ');

  // Check if row already exists for this round (same week, nine, tee time and group)
  var data = sheet.getDataRange().getValues();
  var existingRow = -1;
  for (var r = 1; r < data.length; r++) {
    if (data[r][0] == wd.week && data[r][2] == wd.side && data[r][3] == wd.teeTime && data[r][4] == pName(0)) { existingRow = r+1; break; }
  }

  var rowData = [
    wd.week, wd.date, wd.side, wd.teeTime,
    pName(0), pSub(0), pName(1), pSub(1), pName(2), pSub(2), pName(3), pSub(3),
    players[0] ? totals[0] : '', players[1] ? totals[1] : '', players[2] ? totals[2] : '', players[3] ? totals[3] : '',
    wd.awins, wd.bwins, wd.matchWinner||'Tie',
    pressCnt, pressDetail, skinCnt, skinPay, ctpStr,
    pMoney(0), pMoney(1), pMoney(2), pMoney(3)
  ];

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

// ── PLAYER STATS ─────────────────────────────────────────────
function updatePlayerStats(ss, wd, year) {
  // Rebuild stats from scratch using all raw rounds — more accurate
  var rawSheet = ss.getSheetByName('RawRounds ' + year);
  if (!rawSheet) return;

  var allRoundData = rawSheet.getRange(2, 1, Math.max(1, rawSheet.getLastRow()-1), 1).getValues();
  var allRounds = [];
  allRoundData.forEach(function(row) {
    try { allRounds.push(JSON.parse(row[0])); } catch(e) {}
  });

  // Build stats per player
  var stats = {};
  allRounds.forEach(function(r) {
    r.players.forEach(function(pl, idx) {
      var n = pl.name;
      if (!stats[n]) stats[n] = {isSub: pl.isSub, rounds:0, totalScore:0, money:0, wins:0, pressWins:0, skins:0, ctps:0, weekScores:[]};
      var score = r.scores.reduce(function(s,sc){ return s+(sc['p'+idx]||0); }, 0);
      stats[n].rounds++;
      stats[n].totalScore += score;
      stats[n].money += (r.money[n] || 0);
      stats[n].weekScores.push({week: r.week, score: score});
      var isWin = (r.matchWinner==='A'&&idx<2)||(r.matchWinner==='B'&&idx>=2);
      if (isWin) stats[n].wins++;
      (r.presses||[]).forEach(function(pr){
        var pw = (pr.winner==='A'&&idx<2)||(pr.winner==='B'&&idx>=2);
        if (pw) stats[n].pressWins++;
      });
    });
    Object.values(r.skins||{}).forEach(function(n){ if(stats[n]) stats[n].skins++; });
    Object.values(r.ctpWinners||{}).filter(function(v){ return v&&v!=='__HIGHER_TEE__'; })
      .forEach(function(n){ if(stats[n]) stats[n].ctps++; });
  });

  // Write Stats sheet
  var sheetName = 'Stats ' + year;
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet(sheetName);

  var headers = ['Player','Sub?','Rounds','Total Score','Season Avg','Last 4 Wk Avg','Season Money','Match Wins','Press Wins','Skin Wins','CTP Wins'];
  sheet.appendRow(headers);
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var rows = Object.keys(stats).sort(function(a,b){ return stats[b].money - stats[a].money; });
  rows.forEach(function(n) {
    var s = stats[n];
    var avg = s.rounds > 0 ? (s.totalScore/s.rounds).toFixed(1) : 0;
    var sorted = s.weekScores.slice().sort(function(a,b){ return b.week-a.week; });
    var last4 = sorted.slice(0,4);
    var r4avg = last4.length > 0 ? (last4.reduce(function(t,x){ return t+x.score; },0)/last4.length).toFixed(1) : avg;
    sheet.appendRow([n, s.isSub?'Yes':'No', s.rounds, s.totalScore, avg, r4avg,
      s.money.toFixed(2), s.wins, s.pressWins, s.skins, s.ctps]);
  });

  sheet.autoResizeColumns(1, headers.length);
}

// ── WEEK TAB ─────────────────────────────────────────────────
function writeWeekTab(ss, wd) {
  var year = wd.year || new Date().getFullYear().toString();
  var tabName = year + ' Wk' + wd.week + ' T' + wd.teeTime;
  var existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet(tabName);

  var holesData = wd.side === 'front'
    ? [{n:1,p:4},{n:2,p:3},{n:3,p:5},{n:4,p:4},{n:5,p:3},{n:6,p:4},{n:7,p:4},{n:8,p:4},{n:9,p:4}]
    : [{n:10,p:4},{n:11,p:4},{n:12,p:3},{n:13,p:5},{n:14,p:3},{n:15,p:4},{n:16,p:3},{n:17,p:5},{n:18,p:4}];

  sheet.getRange('A1').setValue(
    'OCBC Golf League — The Links at Crystal Lake — ' + year +
    ' Week ' + wd.week + ' (' + wd.side + ' 9) — Tee ' + wd.teeTime + ' — ' + wd.date
  ).setFontWeight('bold').setFontSize(11);
  sheet.getRange(1,1,1,13).merge();

  var pName = function(i) {
    return wd.players[i] ? wd.players[i].name + (wd.players[i].isSub?' (sub)':'') : '';
  };
  sheet.getRange('A3').setValue('TEAM A').setFontWeight('bold').setFontColor('#1D9E75');
  sheet.getRange('A4').setValue(pName(0));
  sheet.getRange('A5').setValue(pName(1));
  sheet.getRange('D3').setValue('TEAM B').setFontWeight('bold').setFontColor('#185FA5');
  sheet.getRange('D4').setValue(pName(2));
  sheet.getRange('D5').setValue(pName(3));

  var startRow = 7;
  var hdrArr = ['Player'].concat(holesData.map(function(h){ return 'H'+h.n; })).concat(['Total']);
  sheet.getRange(startRow,1,1,hdrArr.length).setValues([hdrArr])
    .setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');

  var parRow = ['Par'].concat(holesData.map(function(h){ return h.p; }))
    .concat([holesData.reduce(function(s,h){ return s+h.p; },0)]);
  sheet.getRange(startRow+1,1,1,parRow.length).setValues([parRow]).setBackground('#e8ecf5');

  wd.players.forEach(function(pl, idx) {
    var scores = wd.scores.map(function(sc){ return sc['p'+idx]||0; });
    var tot = scores.reduce(function(s,v){ return s+v; },0);
    sheet.getRange(startRow+2+idx,1,1,hdrArr.length).setValues(
      [[pl.name+(pl.isSub?' (sub)':'')].concat(scores).concat([tot])]
    );
  });

  var rr = startRow + 7;
  var teamA = [wd.players[0],wd.players[1]].filter(Boolean).map(function(p){return p.name;}).join('/');
  var teamB = [wd.players[2],wd.players[3]].filter(Boolean).map(function(p){return p.name;}).join('/');
  var matchTxt = wd.matchWinner==='A'
    ? teamA+' WIN (+$5/man)'
    : wd.matchWinner==='B'
    ? teamB+' WIN (+$5/man)'
    : 'TIE — no payout';
  sheet.getRange(rr,1).setValue('Match:').setFontWeight('bold');
  sheet.getRange(rr,2).setValue(matchTxt);

  if (wd.presses && wd.presses.length) {
    sheet.getRange(rr+1,1).setValue('Presses:').setFontWeight('bold');
    sheet.getRange(rr+1,2).setValue(wd.presses.map(function(pr,i){
      var w = pr.winner==='A'
        ? teamA+' WIN (+$10/man)'
        : pr.winner==='B'
        ? teamB+' WIN (+$10/man)'
        : 'TIED';
      return 'Press '+(i+1)+' (H'+pr.startHole+', Team '+pr.calledBy+'): '+w;
    }).join(' | '));
    rr++;
  }

  var skinEntries = Object.entries(wd.skins||{});
  var skinCnt = skinEntries.length;
  var skinPay = skinCnt>0 ? (180/skinCnt).toFixed(2) : 0;
  sheet.getRange(rr+1,1).setValue('Skins:').setFontWeight('bold');
  sheet.getRange(rr+1,2).setValue(skinCnt===0 ? 'No skins' :
    skinEntries.map(function(e){ return 'H'+e[0]+': '+e[1]+' (+$'+skinPay+')'; }).join(', '));

  var ctpEntries = Object.entries(wd.ctpWinners||{}).filter(function(e){ return e[1]&&e[1]!=='__HIGHER_TEE__'; });
  var par3Cnt = holesData.filter(function(h){ return h.p===3; }).length;
  var ctpPay = (36).toFixed(2);
  sheet.getRange(rr+2,1).setValue('CTP:').setFontWeight('bold');
  sheet.getRange(rr+2,2).setValue(
    ctpEntries.map(function(e){ return 'H'+e[0]+': '+e[1]+' (+$'+ctpPay+')'; }).join(', ') || 'None / Later tee time'
  );

  sheet.getRange(rr+3,1).setValue('Payouts:').setFontWeight('bold');
  var moneyObj = wd.money || {};
  sheet.getRange(rr+3,2).setValue(wd.players.map(function(pl){
    var m = moneyObj[pl.name]||0;
    return pl.name+': '+(m>=0?'+':'')+'$'+m.toFixed(2);
  }).join('   '));

  sheet.autoResizeColumns(1, hdrArr.length);
}

// ── LIVE SCORING ─────────────────────────────────────────────
function saveLive(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = data.year || new Date().getFullYear().toString();
  var sheetName = 'Live ' + year;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['Tee Time', 'Week', 'Hole Index', 'Players', 'Scores', 'A Wins', 'B Wins', 'Timestamp', 'JSON']);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  // Find existing row for this tee time + week, update it
  var lastRow = sheet.getLastRow();
  var found = false;
  if (lastRow > 1) {
    var teeCol = sheet.getRange(2, 1, lastRow-1, 2).getValues();
    for (var r = 0; r < teeCol.length; r++) {
      if (teeCol[r][0] == data.teeTime && teeCol[r][1] == data.week) {
        sheet.getRange(r+2, 1, 1, 9).setValues([[
          data.teeTime, data.week, data.holeIdx,
          data.players.map(function(p){return p.name;}).join(', '),
          JSON.stringify(data.scores),
          data.awins, data.bwins,
          new Date(data.timestamp).toLocaleTimeString(),
          JSON.stringify(data)
        ]]);
        found = true;
        break;
      }
    }
  }
  if (!found) {
    sheet.appendRow([
      data.teeTime, data.week, data.holeIdx,
      data.players.map(function(p){return p.name;}).join(', '),
      JSON.stringify(data.scores),
      data.awins, data.bwins,
      new Date(data.timestamp).toLocaleTimeString(),
      JSON.stringify(data)
    ]);
  }
  return jsonResponse({ ok: true });
}

function getLive(year, week) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  year = year || new Date().getFullYear().toString();
  var sheetName = 'Live ' + year;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ ok: true, live: [] });
  }
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 9).getValues();
  var live = [];
  data.forEach(function(row) {
    if (week && row[1] != week) return;
    try {
      var parsed = JSON.parse(row[8]);
      live.push(parsed);
    } catch(e) {}
  });
  return jsonResponse({ ok: true, live: live });
}

// ── HELPERS ───────────────────────────────────────────────────
function renamePlayerInSheet(oldName, newName, year) {
  if (!oldName || !newName) return jsonResponse({ ok: false, error: 'Missing names' });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var yr = year || new Date().getFullYear().toString();
  var count = 0;

  // Update RawRounds sheet
  var rawSheet = ss.getSheetByName('RawRounds ' + yr);
  if (rawSheet) {
    var rawData = rawSheet.getDataRange().getValues();
    for (var i = 1; i < rawData.length; i++) {
      var cell = rawData[i][0];
      if (!cell) continue;
      try {
        var parsed = typeof cell === 'string' ? JSON.parse(cell) : cell;
        var changed = false;
        if (parsed.players) {
          parsed.players.forEach(function(pl) {
            if (pl.name === oldName) { pl.name = newName; changed = true; }
          });
        }
        if (parsed.money && parsed.money[oldName] !== undefined) {
          parsed.money[newName] = parsed.money[oldName];
          delete parsed.money[oldName];
          changed = true;
        }
        if (parsed.skinsMoney && parsed.skinsMoney[oldName] !== undefined) {
          parsed.skinsMoney[newName] = parsed.skinsMoney[oldName];
          delete parsed.skinsMoney[oldName];
          changed = true;
        }
        if (parsed.ctpMoney && parsed.ctpMoney[oldName] !== undefined) {
          parsed.ctpMoney[newName] = parsed.ctpMoney[oldName];
          delete parsed.ctpMoney[oldName];
          changed = true;
        }
        if (parsed.ctpWinners) {
          Object.keys(parsed.ctpWinners).forEach(function(h) {
            if (parsed.ctpWinners[h] === oldName) { parsed.ctpWinners[h] = newName; changed = true; }
          });
        }
        if (parsed.leagueCTP) {
          Object.keys(parsed.leagueCTP).forEach(function(h) {
            if (parsed.leagueCTP[h] === oldName) { parsed.leagueCTP[h] = newName; changed = true; }
          });
        }
        if (parsed.skins) {
          Object.keys(parsed.skins).forEach(function(h) {
            if (parsed.skins[h] === oldName) { parsed.skins[h] = newName; changed = true; }
          });
        }
        if (changed) {
          rawSheet.getRange(i + 1, 1).setValue(JSON.stringify(parsed));
          count++;
        }
      } catch(e) {}
    }
  }

  // Update Config tab (weekLineup)
  var cfgSheet = ss.getSheetByName('Config');
  if (cfgSheet) {
    var cfgData = cfgSheet.getDataRange().getValues();
    for (var ci = 1; ci < cfgData.length; ci++) {
      var key = cfgData[ci][0];
      var val = cfgData[ci][1];
      if (key && key.toString().indexOf('weekLineup') !== -1 && val) {
        try {
          var lu = typeof val === 'string' ? JSON.parse(val) : val;
          var luChanged = false;
          for (var s = 1; s <= 10; s++) {
            if (lu[s]) {
              if (lu[s].captain === oldName) { lu[s].captain = newName; luChanged = true; }
              if (lu[s].prevCaptain === oldName) { lu[s].prevCaptain = newName; luChanged = true; }
              if (lu[s].members) {
                lu[s].members = lu[s].members.map(function(m) { return m === oldName ? newName : m; });
                luChanged = true;
              }
            }
          }
          if (lu.notPlaying) lu.notPlaying = lu.notPlaying.map(function(m) { if(m===oldName){luChanged=true;return newName;}return m; });
          if (lu.playersLooking) lu.playersLooking = lu.playersLooking.map(function(m) { if(m===oldName){luChanged=true;return newName;}return m; });
          if (lu.teamsLooking) lu.teamsLooking = lu.teamsLooking.map(function(t) {
            if(t.p1===oldName){t.p1=newName;luChanged=true;}
            if(t.p2===oldName){t.p2=newName;luChanged=true;}
            return t;
          });
          if (luChanged) {
            cfgSheet.getRange(ci + 1, 2).setValue(JSON.stringify(lu));
            count++;
          }
        } catch(e) {}
      }
    }
  }

  // Also update the players array in Config
  if (cfgSheet) {
    var cfgData2 = cfgSheet.getDataRange().getValues();
    for (var pi = 1; pi < cfgData2.length; pi++) {
      if (cfgData2[pi][0] === 'players') {
        try {
          var pList = typeof cfgData2[pi][1] === 'string' ? JSON.parse(cfgData2[pi][1]) : cfgData2[pi][1];
          if (Array.isArray(pList)) {
            var pIdx = pList.indexOf(oldName);
            if (pIdx >= 0) {
              pList[pIdx] = newName;
              cfgSheet.getRange(pi + 1, 2).setValue(JSON.stringify(pList));
              count++;
            }
          }
        } catch(e) {}
        break;
      }
    }
  }

  return jsonResponse({ ok: true, updated: count, oldName: oldName, newName: newName });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

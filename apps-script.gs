/**
 * MKCS Quiz Night — Scoring Backend
 * =================================
 *
 * SETUP (one time):
 *   1. Create a new Google Sheet. Name it whatever you like (e.g. "MKCS Quiz Scores").
 *   2. Open Extensions → Apps Script. Delete any boilerplate, paste this whole file in.
 *   3. Click Deploy → New deployment.
 *        - Type: Web app
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Copy the Web app URL. (Looks like https://script.google.com/macros/s/AKfy.../exec)
 *   4. Open submit.html and index.html, set the API_URL constant near the top to that URL.
 *   5. Share the Sheet with yourself only — players never touch it.
 *
 * The script auto-creates two tabs on first call: Teams and Scores.
 */

const SHEET_TEAMS     = 'Teams';
const SHEET_SCORES    = 'Scores';
const SHEET_ROUND_MAX = 'RoundMax';
const SHEET_REJECTED  = 'Rejected';
const SHEET_SETTINGS  = 'Settings';
const ROUNDS_MAX      = 12;   // absolute cap on number of rounds
const ROUNDS_DEFAULT  = 7;

function doGet(e) {
  const ss = SpreadsheetApp.getActive();
  ensureSheets_(ss);
  const teams    = readTeams_(ss);
  const scores   = readScores_(ss);
  const roundMax = readRoundMax_(ss);
  const rejected = readRejected_(ss);
  const rounds   = getRounds_(ss);
  return json_({ ok: true, teams, scores, roundMax, rejected, rounds, roundsMax: ROUNDS_MAX });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActive();
    ensureSheets_(ss);

    if (body.action === 'addTeam') {
      const name = cleanName_(body.name);
      if (!name) return json_({ ok: false, error: 'Team name required' });
      const teams = readTeams_(ss);
      if (teams.some(t => t.toLowerCase() === name.toLowerCase())) {
        return json_({ ok: true, teams, added: false });
      }
      ss.getSheetByName(SHEET_TEAMS).appendRow([name, new Date()]);
      return json_({ ok: true, teams: teams.concat(name), added: true });
    }

    if (body.action === 'setRounds') {
      const n = parseInt(body.rounds, 10);
      if (!(n >= 1 && n <= ROUNDS_MAX)) return json_({ ok: false, error: 'Rounds must be 1–' + ROUNDS_MAX });
      const sh = ss.getSheetByName(SHEET_SETTINGS);
      const data = sh.getDataRange().getValues();
      let rowIdx = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === 'rounds') { rowIdx = i; break; }
      }
      if (rowIdx >= 0) sh.getRange(rowIdx + 1, 1, 1, 2).setValues([['rounds', n]]);
      else sh.appendRow(['rounds', n]);
      return json_({ ok: true, rounds: n });
    }

    if (body.action === 'recordRejected') {
      const team = cleanName_(body.team);
      const round = parseInt(body.round, 10);
      if (!team) return json_({ ok: false, error: 'Team required' });
      if (!(round >= 1 && round <= ROUNDS_MAX)) return json_({ ok: false, error: 'Round must be 1–' + ROUNDS_MAX });
      addRejected_(ss, team, round);
      return json_({ ok: true, rejected: readRejected_(ss) });
    }

    if (body.action === 'setRoundMax') {
      const round = parseInt(body.round, 10);
      if (!(round >= 1 && round <= ROUNDS_MAX)) return json_({ ok: false, error: 'Round must be 1–' + ROUNDS_MAX });
      const sheet = ss.getSheetByName(SHEET_ROUND_MAX);
      const raw = body.max;
      const clear = raw === '' || raw === null || raw === undefined;
      const max = clear ? null : Number(raw);
      if (!clear && (!isFinite(max) || max < 0)) return json_({ ok: false, error: 'Max must be a positive number' });
      const data = sheet.getDataRange().getValues();
      let rowIdx = -1;
      for (let i = 1; i < data.length; i++) {
        if (Number(data[i][0]) === round) { rowIdx = i; break; }
      }
      if (clear) {
        if (rowIdx >= 0) sheet.deleteRow(rowIdx + 1);
      } else if (rowIdx >= 0) {
        sheet.getRange(rowIdx + 1, 1, 1, 2).setValues([[round, max]]);
      } else {
        sheet.appendRow([round, max]);
      }
      return json_({ ok: true, roundMax: readRoundMax_(ss) });
    }

    if (body.action === 'submitScore') {
      const team = cleanName_(body.team);
      const round = parseInt(body.round, 10);
      const score = Number(body.score);
      if (!team) return json_({ ok: false, error: 'Team required' });
      if (!(round >= 1 && round <= ROUNDS_MAX)) return json_({ ok: false, error: 'Round must be 1–' + ROUNDS_MAX });
      if (!isFinite(score)) return json_({ ok: false, error: 'Score must be a number' });
      const roundMax = readRoundMax_(ss);
      if (roundMax[round] !== undefined && score > roundMax[round]) {
        addRejected_(ss, team, round);
        return json_({ ok: false, error: 'Score ' + score + ' exceeds max of ' + roundMax[round] + ' for round ' + round, rejected: readRejected_(ss) });
      }
      if (score < 0) return json_({ ok: false, error: 'Score cannot be negative' });

      const sheet = ss.getSheetByName(SHEET_SCORES);
      // Upsert: replace any existing row for (team, round) so resubmits overwrite.
      const data = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim().toLowerCase() === team.toLowerCase() && Number(data[i][2]) === round) {
          sheet.getRange(i + 1, 1, 1, 4).setValues([[new Date(), team, round, score]]);
          updated = true;
          break;
        }
      }
      if (!updated) sheet.appendRow([new Date(), team, round, score]);
      removeRejected_(ss, team, round);
      return json_({ ok: true, team, round, score, updated });
    }

    if (body.action === 'renameTeam') {
      const oldName = cleanName_(body.oldName);
      const newName = cleanName_(body.newName);
      if (!oldName || !newName) return json_({ ok: false, error: 'Both names required' });
      if (oldName.toLowerCase() === newName.toLowerCase()) {
        // Allow case-only rename; otherwise no-op
      }
      // Reject collision with another existing team (case-insensitive, excluding the team being renamed)
      const teamsSheet = ss.getSheetByName(SHEET_TEAMS);
      const tData = teamsSheet.getDataRange().getValues();
      let row = -1;
      for (let i = 1; i < tData.length; i++) {
        const v = String(tData[i][0]).trim();
        if (v.toLowerCase() === oldName.toLowerCase()) row = i;
        else if (v.toLowerCase() === newName.toLowerCase()) return json_({ ok: false, error: 'Another team already uses that name' });
      }
      if (row < 0) return json_({ ok: false, error: 'Team not found' });
      teamsSheet.getRange(row + 1, 1).setValue(newName);
      // Cascade to scores
      const scoresSheet = ss.getSheetByName(SHEET_SCORES);
      const sData = scoresSheet.getDataRange().getValues();
      for (let i = 1; i < sData.length; i++) {
        if (String(sData[i][1]).trim().toLowerCase() === oldName.toLowerCase()) {
          scoresSheet.getRange(i + 1, 2).setValue(newName);
        }
      }
      // Cascade to rejected markers
      const rejSheet = ss.getSheetByName(SHEET_REJECTED);
      const rData = rejSheet.getDataRange().getValues();
      for (let i = 1; i < rData.length; i++) {
        if (String(rData[i][0]).trim().toLowerCase() === oldName.toLowerCase()) {
          rejSheet.getRange(i + 1, 1).setValue(newName);
        }
      }
      return json_({ ok: true });
    }

    if (body.action === 'deleteTeam') {
      const name = cleanName_(body.name);
      if (!name) return json_({ ok: false, error: 'Team required' });
      const teamsSheet = ss.getSheetByName(SHEET_TEAMS);
      const tData = teamsSheet.getDataRange().getValues();
      for (let i = tData.length - 1; i >= 1; i--) {
        if (String(tData[i][0]).trim().toLowerCase() === name.toLowerCase()) {
          teamsSheet.deleteRow(i + 1);
        }
      }
      const scoresSheet = ss.getSheetByName(SHEET_SCORES);
      const sData = scoresSheet.getDataRange().getValues();
      for (let i = sData.length - 1; i >= 1; i--) {
        if (String(sData[i][1]).trim().toLowerCase() === name.toLowerCase()) {
          scoresSheet.deleteRow(i + 1);
        }
      }
      removeRejectedTeam_(ss, name);
      return json_({ ok: true });
    }

    if (body.action === 'deleteScore') {
      const team = cleanName_(body.team);
      const round = parseInt(body.round, 10);
      if (!team) return json_({ ok: false, error: 'Team required' });
      if (!(round >= 1 && round <= ROUNDS_MAX)) return json_({ ok: false, error: 'Round must be 1-' + ROUNDS_MAX });
      const scoresSheet = ss.getSheetByName(SHEET_SCORES);
      const sData = scoresSheet.getDataRange().getValues();
      for (let i = sData.length - 1; i >= 1; i--) {
        if (String(sData[i][1]).trim().toLowerCase() === team.toLowerCase() && Number(sData[i][2]) === round) {
          scoresSheet.deleteRow(i + 1);
        }
      }
      removeRejected_(ss, team, round);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function ensureSheets_(ss) {
  let t = ss.getSheetByName(SHEET_TEAMS);
  if (!t) { t = ss.insertSheet(SHEET_TEAMS); t.appendRow(['Team Name', 'Added']); }
  let s = ss.getSheetByName(SHEET_SCORES);
  if (!s) { s = ss.insertSheet(SHEET_SCORES); s.appendRow(['Timestamp', 'Team', 'Round', 'Score']); }
  let m = ss.getSheetByName(SHEET_ROUND_MAX);
  if (!m) { m = ss.insertSheet(SHEET_ROUND_MAX); m.appendRow(['Round', 'Max']); }
  let r = ss.getSheetByName(SHEET_REJECTED);
  if (!r) { r = ss.insertSheet(SHEET_REJECTED); r.appendRow(['Team', 'Round']); }
  let g = ss.getSheetByName(SHEET_SETTINGS);
  if (!g) { g = ss.insertSheet(SHEET_SETTINGS); g.appendRow(['Key', 'Value']); g.appendRow(['rounds', ROUNDS_DEFAULT]); }
}

function getRounds_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === 'rounds') {
        const n = parseInt(data[i][1], 10);
        if (n >= 1 && n <= ROUNDS_MAX) return n;
      }
    }
  }
  return ROUNDS_DEFAULT;
}

function readRejected_(ss) {
  const sh = ss.getSheetByName(SHEET_REJECTED);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues().map(r => ({
    team: String(r[0]).trim(),
    round: Number(r[1]),
  })).filter(x => x.team && x.round);
}

function addRejected_(ss, team, round) {
  const sh = ss.getSheetByName(SHEET_REJECTED);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === team.toLowerCase() && Number(data[i][1]) === round) return;
  }
  sh.appendRow([team, round]);
}

function removeRejected_(ss, team, round) {
  const sh = ss.getSheetByName(SHEET_REJECTED);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toLowerCase() === team.toLowerCase() && Number(data[i][1]) === round) sh.deleteRow(i + 1);
  }
}

function removeRejectedTeam_(ss, team) {
  const sh = ss.getSheetByName(SHEET_REJECTED);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toLowerCase() === team.toLowerCase()) sh.deleteRow(i + 1);
  }
}

function readRoundMax_(ss) {
  const sh = ss.getSheetByName(SHEET_ROUND_MAX);
  const last = sh.getLastRow();
  const out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    const round = Number(r[0]);
    const max = Number(r[1]);
    if (round >= 1 && round <= ROUNDS_MAX && isFinite(max)) out[round] = max;
  });
  return out;
}

function readTeams_(ss) {
  const sh = ss.getSheetByName(SHEET_TEAMS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}

function readScores_(ss) {
  const sh = ss.getSheetByName(SHEET_SCORES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues().map(r => ({
    team: String(r[1]).trim(),
    round: Number(r[2]),
    score: Number(r[3]),
  })).filter(s => s.team && s.round);
}

function cleanName_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().slice(0, 60);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

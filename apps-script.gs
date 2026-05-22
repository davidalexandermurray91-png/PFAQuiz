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

const SHEET_TEAMS  = 'Teams';
const SHEET_SCORES = 'Scores';
const ROUNDS_MAX   = 7;

function doGet(e) {
  const ss = SpreadsheetApp.getActive();
  ensureSheets_(ss);
  const teams  = readTeams_(ss);
  const scores = readScores_(ss);
  return json_({ ok: true, teams, scores, rounds: ROUNDS_MAX });
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

    if (body.action === 'submitScore') {
      const team = cleanName_(body.team);
      const round = parseInt(body.round, 10);
      const score = Number(body.score);
      if (!team) return json_({ ok: false, error: 'Team required' });
      if (!(round >= 1 && round <= ROUNDS_MAX)) return json_({ ok: false, error: 'Round must be 1–' + ROUNDS_MAX });
      if (!isFinite(score)) return json_({ ok: false, error: 'Score must be a number' });

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
      return json_({ ok: true, team, round, score, updated });
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

/**
 * WARM PIXEL
 * Google Apps Script backend
 * Game: original pixel platform-fighter inspired by classic arcade platformers.
 */

const DB_SHEET = 'WarmPixel_Data';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Warm Pixel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function saveProfile(profile) {
  const clean = sanitizeProfile_(profile);
  if (!clean.name || !clean.birthYear || !['female', 'male'].includes(clean.character)) {
    throw new Error('Invalid profile.');
  }

  const sheet = getDataSheet_();
  const now = new Date();
  const id = Utilities.getUuid();

  sheet.appendRow([
    id,
    now,
    clean.name,
    clean.birthYear,
    clean.character,
    0,
    0
  ]);

  return {
    id,
    name: clean.name,
    birthYear: clean.birthYear,
    character: clean.character
  };
}

function saveScore(payload) {
  const p = payload || {};
  const id = String(p.id || '').trim();
  const distance = Math.max(0, Math.floor(Number(p.distance) || 0));
  const score = Math.max(0, Math.floor(Number(p.score) || 0));

  if (!id) throw new Error('Missing player id.');

  const sheet = getDataSheet_();
  const values = sheet.getDataRange().getValues();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]) === id) {
      const bestDistance = Math.max(Number(values[r][5]) || 0, distance);
      const bestScore = Math.max(Number(values[r][6]) || 0, score);

      sheet.getRange(r + 1, 6, 1, 2).setValues([[bestDistance, bestScore]]);
      return { distance: bestDistance, score: bestScore };
    }
  }

  throw new Error('Player not found.');
}

function getLeaderboard() {
  const sheet = getDataSheet_();
  const values = sheet.getDataRange().getValues();

  const rows = values.slice(1)
    .filter(r => r[2])
    .map(r => ({
      name: String(r[2]),
      character: String(r[4] || 'male'),
      distance: Number(r[5]) || 0,
      score: Number(r[6]) || 0
    }))
    .sort((a, b) => b.distance - a.distance || b.score - a.score)
    .slice(0, 10);

  return rows;
}

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DB_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(DB_SHEET);
    sheet.appendRow([
      'ID',
      'Created At',
      'Name',
      'Birth Year',
      'Character',
      'Best Distance',
      'Best Score'
    ]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function sanitizeProfile_(profile) {
  const rawName = String(profile && profile.name || '').trim();
  const name = rawName.replace(/[<>]/g, '').slice(0, 18);
  const birthYear = Math.floor(Number(profile && profile.birthYear));

  const currentYear = new Date().getFullYear();
  const safeYear = birthYear >= 1900 && birthYear <= currentYear ? birthYear : 0;

  return {
    name,
    birthYear: safeYear,
    character: String(profile && profile.character || '').toLowerCase()
  };
}

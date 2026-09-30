/**
 * WARM PIXEL - ONE TIME SETUP
 *
 * Run setupWarmPixel() once from the Apps Script editor.
 * It creates the Google Sheet used for player profiles and scores.
 */

function setupWarmPixel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('WarmPixel_Data');

  if (!sheet) {
    sheet = ss.insertSheet('WarmPixel_Data');
  }

  const headers = [
    'ID',
    'Created At',
    'Name',
    'Birth Year',
    'Character',
    'Best Distance',
    'Best Score'
  ];

  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  if (firstRow.join('').trim() === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert(
    'Warm Pixel setup complete.\\n\\n' +
    'Next: Deploy > New deployment > Web app.\\n' +
    'Execute as: Me\\n' +
    'Who has access: Anyone'
  );
}

import { google } from "googleapis";
import fs from "fs";
import path from "path";
require("dotenv").config();

const CREDENTIALS_FILEPATH = process.env.GOOGLE_CREDENTIALS_PATH;
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];

function getAuthClient() {
  const credentials = JSON.parse(
    fs.readFileSync(path.resolve(CREDENTIALS_FILEPATH))
  );

  return google.auth.JWT(
    credentials.client_email,
    null,
    credentials.private_key,
    SCOPES
  );
}

async function getSheets() {
  const auth = getAuthClient();
  await auth.authorize();

  return google.sheets({ version: "v4", auth });
}

module.exports = {
  async getAllRows() {
    const sheets = await getSheets();
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const sheetName = process.env.SHEET_NAME;
    const range = `${sheetName}!A:Z`;
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range,
    });
    return res.data.values || [];
  },

  async appendRow(values) {
    const sheets = await getSheets();
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const sheetName = process.env.SHEET_NAME;
    const res = await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: `${sheetName}!A:Z`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [values] },
    });
    return res.data;
  },

  async updateRowById(id, updates) {
    const sheets = await getSheets();
    const sheetId = process.env.GOOGLE_SHEET_ID;
    const sheetName = process.env.SHEET_NAME;

    const range = `${sheetName}!A:Z`;
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range,
    });
    const rows = res.data.values || [];
    if (rows.length === 0) return null;

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] && String(rows[i][0]) === String(id)) {
        const rowIndexInSheet = i + 1;

        const requests = [];
        for (const [colIndexStr, value] of Object.entries(updates)) {
          const colIndex = parseInt(colIndexStr, 10);
          const colLetter = String.fromCharCode("A".charCodeAt(0) + colIndex);
          const cell = `${sheetName}!${colLetter}${rowIndexInSheet}`;
          await sheets.spreadsheets.values.update({
            spreadsheetId: sheetId,
            range: cell,
            valueInputOption: "USER_ENTERED",
            requestBody: { values: [[value]] },
          });
        }
        return true;
      }
    }
    return false;
  },
};

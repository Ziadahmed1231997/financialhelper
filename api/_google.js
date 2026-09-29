import { google } from "googleapis";

export const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;

export async function getSheets() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is missing in Vercel Environment Variables");
  }
  if (!SPREADSHEET_ID) {
    throw new Error("GOOGLE_SHEET_ID is missing in Vercel Environment Variables");
  }

  let credentials;
  try {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } catch {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON");
  }

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });

  return google.sheets({ version: "v4", auth });
}
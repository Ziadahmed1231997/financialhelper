import { getSheets, SPREADSHEET_ID } from "./_google.js";

export default async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).json({ ok: false, error: "Method not allowed" });
  try {
    const sheets = await getSheets();
    const result = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      fields: "properties.title,sheets.properties.title"
    });
    return res.status(200).json({
      ok: true,
      database: result.data.properties?.title || "FinanceControl Database",
      sheets: (result.data.sheets || []).map(s => s.properties?.title).filter(Boolean)
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
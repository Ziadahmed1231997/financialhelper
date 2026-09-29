import { getSheets, SPREADSHEET_ID } from "./_google.js";

const DB_NAME = "FinanceControl Database";
const RANGE_SPECS = {
  AppConfig: "AppConfig!A2:D",
  Categories: "Categories!A2:J",
  Accounts: "Accounts!A2:L",
  Transactions: "Transactions!A2:R",
  Recurring: "Recurring!A2:S",
  RecurringChanges: "RecurringChanges!A2:J",
  FXRates: "FXRates!A2:J",
  MonthlyPlans: "MonthlyPlans!A2:N",
  Goals: "Goals!A2:O",
  ForecastEvents: "ForecastEvents!A2:O",
  Scenarios: "Scenarios!A2:M",
  ScenarioItems: "ScenarioItems!A2:O"
};

const now = () => new Date().toISOString();
const bool = v => v === true || String(v).toLowerCase() === "true" || String(v) === "1" || String(v).toLowerCase() === "yes";
const num = v => Number(v) || 0;
const month = v => String(v || "").slice(0, 7);
const dateFromMonth = v => v ? `${String(v).slice(0, 7)}-01` : "";
const cleanRows = rows => (rows || []).filter(r => Array.isArray(r) && r.some(v => String(v ?? "").trim() !== ""));

function categoryMaps(rows) {
  const idToName = new Map();
  const nameToId = new Map();
  for (const r of cleanRows(rows)) {
    const id = String(r[0] || "").trim();
    const name = String(r[1] || "").trim();
    if (id && name) { idToName.set(id, name); nameToId.set(name.toLowerCase(), id); }
  }
  return { idToName, nameToId };
}
function categoryName(v, map) { const x=String(v||""); return map.idToName.get(x) || x || "Other"; }
function categoryId(v, map) { const x=String(v||"Other").trim(); return map.nameToId.get(x.toLowerCase()) || x; }

async function readState() {
  const sheets = await getSheets();
  const ranges = Object.values(RANGE_SPECS);
  const result = await sheets.spreadsheets.values.batchGet({ spreadsheetId: SPREADSHEET_ID, ranges });
  const values = Object.fromEntries(Object.keys(RANGE_SPECS).map((k, i) => [k, cleanRows(result.data.valueRanges?.[i]?.values || [])]));
  const cats = categoryMaps(values.Categories);
  const config = Object.fromEntries(values.AppConfig.map(r => [String(r[0]||""), r[1]]));

  const state = {
    version: 5,
    settings: {
      defaultFx: num(config.default_fx_egp_per_sar) || 13,
      reportCurrency: config.reporting_currency || "SAR",
      projectStartDate: config.project_start_date || "2026-10-01"
    },
    accounts: values.Accounts.filter(r => !r[11]).map(r => ({
      id: String(r[0]||""), name: String(r[1]||""), type: String(r[2]||"checking").toLowerCase(),
      currency: String(r[3]||"SAR").toUpperCase(), balance: num(r[5] !== "" && r[5] != null ? r[5] : r[4]),
      includeNetWorth: r[6] === "" || r[6] == null ? true : bool(r[6]), notes: String(r[8]||"")
    })),
    transactions: values.Transactions.filter(r => !r[16]).map(r => ({
      id: String(r[0]||""), date: String(r[1]||"").slice(0,10), type: String(r[2]||"expense").toLowerCase(),
      name: String(r[3]||""), category: categoryName(r[4], cats), accountId: String(r[5]||""),
      amount: num(r[6]), currency: String(r[7]||"SAR").toUpperCase(), fxRate: num(r[8]) || 1,
      sarAmount: num(r[9]), recurringId: String(r[10]||""), status: String(r[11]||"POSTED"),
      notes: String(r[12]||""), source: String(r[13]||"WEB")
    })),
    recurring: values.Recurring.filter(r => !r[17]).map(r => ({
      id: String(r[0]||""), name: String(r[1]||""), type: String(r[2]||"expense").toLowerCase(),
      category: categoryName(r[3], cats), accountId: String(r[4]||""), amount: num(r[5]),
      currency: String(r[6]||"SAR").toUpperCase(), frequency: String(r[7]||"monthly").toLowerCase(),
      startMonth: month(r[8]), endMonth: month(r[9]), fixed: bool(r[10]), active: r[11] === "" || r[11] == null ? true : bool(r[11]),
      day: num(r[13]) || 1, notes: String(r[14]||"")
    })),
    changes: values.RecurringChanges.filter(r => !r[9]).map(r => ({
      id: String(r[0]||""), recurringId: String(r[1]||""), effectiveMonth: month(r[2]),
      oldAmount: num(r[3]), newAmount: num(r[4]), reason: String(r[6]||"")
    })),
    fxRates: values.FXRates.filter(r => !r[9] && String(r[2]||"").toUpperCase() === "EGP" && String(r[3]||"").toUpperCase() === "SAR").map(r => ({
      id: String(r[0]||""), month: month(r[1]), rate: num(r[4]), note: String(r[5]||"")
    })),
    budgets: values.MonthlyPlans.filter(r => !r[10]).map(r => ({
      id: String(r[0]||""), startMonth: month(r[1]), category: categoryName(r[2], cats), amount: num(r[3]),
      currency: String(r[4]||"SAR").toUpperCase(), bucket: String(r[5]||"flex").toLowerCase(), rollover: bool(r[6]), notes: String(r[7]||""), endMonth: month(r[12]), scheduleMode: String(r[13]||"ongoing").toLowerCase()
    })),
    goals: values.Goals.filter(r => !r[13]).map(r => ({
      id: String(r[0]||""), name: String(r[1]||""), target: num(r[2]), currency: String(r[3]||"SAR").toUpperCase(),
      dueMonth: month(r[4]), current: num(r[5]), monthlyContribution: num(r[6]), priority: String(r[7]||""),
      status: String(r[8]||"ACTIVE"), accountId: String(r[9]||""), notes: String(r[10]||"")
    })),
    events: values.ForecastEvents.filter(r => !r[13]).map(r => ({
      id: String(r[0]||""), name: String(r[1]||""), date: String(r[2]||"").slice(0,10), type: String(r[3]||"expense").toLowerCase(),
      category: categoryName(r[4], cats), amount: num(r[5]), currency: String(r[6]||"SAR").toUpperCase(),
      probability: num(r[7]) || 100, scenarioId: String(r[8]||""), status: String(r[9]||"PLANNED"), notes: String(r[10]||""), accountId: ""
    })),
    spendScenarios: values.Scenarios.filter(r => !r[9] && String(r[10]||"").toUpperCase() === "SPEND").map(r => ({
      id: String(r[0]||""), name: String(r[1]||""), notes: String(r[6]||""),
      status: String(r[11]||"DRAFT").toUpperCase(), approvedAt: String(r[12]||""),
      createdAt: String(r[7]||"")
    })),
    scenarioItems: values.ScenarioItems.filter(r => !r[13]).map(r => ({
      id: String(r[0]||""), scenarioId: String(r[1]||""), name: String(r[2]||""),
      category: categoryName(r[3], cats), amount: num(r[4]), currency: String(r[5]||"SAR").toUpperCase(),
      timing: String(r[6]||"ONE_TIME").toUpperCase(), startMonth: month(r[7]), endMonth: month(r[8]),
      bucket: String(r[9]||"nonmonthly").toLowerCase(), notes: String(r[10]||""), createdAt: String(r[11]||"")
    }))
  };
  return { state, cats };
}

async function writeState(state) {
  const sheets = await getSheets();
  const current = await readState();
  const cats = current.cats;
  const ts = now();
  const s = state || {};
  const settings = s.settings || {};
  const arrays = k => Array.isArray(s[k]) ? s[k] : [];
  const projectStartDate = settings.projectStartDate || current.state?.settings?.projectStartDate || "2026-10-01";
  const projectStartMonth = month(projectStartDate);
  const beforeStartDate = v => v && String(v).slice(0,10) < projectStartDate;
  const beforeStartMonth = v => v && month(v) < projectStartMonth;
  const invalid = [];
  arrays("transactions").forEach(x => { if (beforeStartDate(x.date)) invalid.push(`Transaction "${x.name||x.id}" is before project start`); });
  arrays("recurring").forEach(x => { if (beforeStartMonth(x.startMonth)) invalid.push(`Recurring item "${x.name||x.id}" starts before project start`); });
  arrays("changes").forEach(x => { if (beforeStartMonth(x.effectiveMonth)) invalid.push(`Recurring change "${x.id}" is before project start`); });
  arrays("fxRates").forEach(x => { if (beforeStartMonth(x.month)) invalid.push(`FX rate "${x.id}" is before project start`); });
  arrays("budgets").forEach(x => { if (beforeStartMonth(x.startMonth)) invalid.push(`Monthly plan "${x.category||x.id}" starts before project start`); });
  arrays("events").forEach(x => { if (beforeStartDate(x.date)) invalid.push(`Forecast event "${x.name||x.id}" is before project start`); });
  arrays("scenarioItems").forEach(x => { if (beforeStartMonth(x.startMonth)) invalid.push(`Scenario item "${x.name||x.id}" starts before project start`); });
  if (invalid.length) throw new Error(`Project starts on ${projectStartDate}. ${invalid.slice(0,5).join("; ")}`);

  const appConfig = [
    ["schema_version","1.2","FinanceControl Google Sheets database schema",ts],
    ["app_name","FinanceControl","Application name",ts],
    ["reporting_currency",settings.reportCurrency || "SAR","Primary dashboard/reporting currency",ts],
    ["supported_currencies","SAR,EGP","Currencies accepted by the app",ts],
    ["default_forecast_months","12","Default forecast horizon",ts],
    ["timezone","Asia/Riyadh","Application timezone",ts],
    ["owner_email","ziadrehiem@gmail.com","Database owner",ts],
    ["database_type","GOOGLE_SHEETS","Backend storage type",ts],
    ["api_version","v1","Expected Vercel API contract version",ts],
    ["default_fx_egp_per_sar",num(settings.defaultFx)||13,"Fallback EGP per 1 SAR",ts],
    ["project_start_date",projectStartDate,"FinanceControl project start date; operational entries before this date are not allowed",ts]
  ];

  const data = {
    AppConfig: appConfig,
    Accounts: arrays("accounts").map(a => [a.id,a.name,a.type,a.currency,num(a.balance),num(a.balance),a.includeNetWorth!==false,true,a.notes||"",ts,ts,""]),
    Transactions: arrays("transactions").map(t => [t.id,t.date,t.type,t.name,categoryId(t.category,cats),t.accountId||"",num(t.amount),t.currency,num(t.fxRate)||1,num(t.sarAmount),t.recurringId||"",t.status||"POSTED",t.notes||"",t.source||"WEB",ts,ts,"",1]),
    Recurring: arrays("recurring").map(r => [r.id,r.name,r.type,categoryId(r.category,cats),r.accountId||"",num(r.amount),r.currency,r.frequency||"monthly",dateFromMonth(r.startMonth),dateFromMonth(r.endMonth),r.fixed!==false,r.active!==false,"",num(r.day)||1,r.notes||"",ts,ts,"",1]),
    RecurringChanges: arrays("changes").map(c => [c.id,c.recurringId,dateFromMonth(c.effectiveMonth),num(c.oldAmount),num(c.newAmount),"",c.reason||"",ts,ts,""]),
    FXRates: arrays("fxRates").map(f => [f.id,dateFromMonth(f.month),"EGP","SAR",num(f.rate),f.note||"Manual","true",ts,ts,""]),
    MonthlyPlans: arrays("budgets").map(b => [b.id,b.startMonth,categoryId(b.category,cats),num(b.amount),b.currency,b.bucket||"flex",!!b.rollover,b.notes||"",ts,ts,"",1,b.endMonth||"",b.scheduleMode||"ongoing"]),
    Goals: arrays("goals").map(g => [g.id,g.name,num(g.target),g.currency,dateFromMonth(g.dueMonth),num(g.current),num(g.monthlyContribution),g.priority||"",g.status||"ACTIVE",g.accountId||"",g.notes||"",ts,ts,"",1]),
    ForecastEvents: arrays("events").map(e => [e.id,e.name,e.date,e.type,categoryId(e.category,cats),num(e.amount),e.currency,num(e.probability)||100,e.scenarioId||"",e.status||"PLANNED",e.notes||"",ts,ts,"",1]),
    Scenarios: [
      ["scn_base","Base",1,1,"",true,"Normal recurring income/expense assumptions",ts,ts,"","STRESS","ACTIVE",""],
      ["scn_conservative","Conservative",0.95,1.1,"",true,"5% lower income and 10% higher expenses",ts,ts,"","STRESS","ACTIVE",""],
      ["scn_lean","Lean",1,0.9,"",true,"10% lower expenses",ts,ts,"","STRESS","ACTIVE",""],
      ["scn_custom","Custom",1,1,"",true,"User-controlled stress scenario",ts,ts,"","STRESS","ACTIVE",""],
      ...arrays("spendScenarios").map(x => [x.id,x.name,1,1,"",true,x.notes||"",x.createdAt||ts,ts,"","SPEND",x.status||"DRAFT",x.approvedAt||""])
    ],
    ScenarioItems: arrays("scenarioItems").map(x => [x.id,x.scenarioId,x.name,categoryId(x.category,cats),num(x.amount),x.currency||"SAR",x.timing||"ONE_TIME",x.startMonth||"",x.endMonth||"",x.bucket||"nonmonthly",x.notes||"",x.createdAt||ts,ts,"",1])
  };

  const clearRanges = Object.keys(data).map(k => RANGE_SPECS[k]);
  await sheets.spreadsheets.values.batchClear({ spreadsheetId: SPREADSHEET_ID, requestBody: { ranges: clearRanges } });

  const updates = Object.entries(data).filter(([, values]) => values.length).map(([name, values]) => ({
    range: RANGE_SPECS[name].replace(/A2:.+$/, "A2"),
    majorDimension: "ROWS",
    values
  }));
  if (updates.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: "RAW", data: updates }
    });
  }
  return { ok: true, database: DB_NAME, counts: Object.fromEntries(Object.entries(data).map(([k,v]) => [k,v.length])) };
}

export default async function handler(req, res) {
  try {
    if (req.method === "GET") {
      const { state } = await readState();
      return res.status(200).json({ ok: true, database: DB_NAME, state });
    }
    if (req.method === "PUT") {
      const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
      const result = await writeState(body.state || body);
      return res.status(200).json(result);
    }
    return res.status(405).json({ ok: false, error: "Method not allowed" });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
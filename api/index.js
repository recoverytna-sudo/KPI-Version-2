const { google } = require('googleapis');

const jsonResponse = (res, statusCode, data) => {
  res.status(statusCode).json(data);
};

// Google Sheet ထဲရှိ Tab နာမည်ကို အလိုအလျောက် ရှာဖွေပေးသည့် Helper Function
async function getTabTitle(sheets, spreadsheetId, preferredName) {
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetList = meta.data.sheets || [];
    
    const matched = sheetList.find(s => 
      s.properties.title.trim().toLowerCase() === preferredName.trim().toLowerCase()
    );
    if (matched) return matched.properties.title;
    if (sheetList.length > 0) return sheetList[0].properties.title;
  } catch (e) {}
  return preferredName;
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    let privateKey = process.env.PRIVATE_KEY || '';
    if (privateKey.startsWith('"') && privateKey.endsWith('"')) {
      privateKey = privateKey.slice(1, -1);
    }
    privateKey = privateKey.replace(/\\n/g, '\n');

    const auth = new google.auth.JWT(
      process.env.CLIENT_EMAIL,
      null,
      privateKey,
      ['https://www.googleapis.com/auth/spreadsheets']
    );

    const sheets = google.sheets({ version: 'v4', auth });
    
    const SHEET_ID_KPI = process.env.SHEET_ID_KPI;
    const SHEET_ID_EMP = process.env.SHEET_ID_EMP || process.env.SHEET_ID_KPI;

    // =========================================================================
    // 🌟 GET REQUEST (Employee ID Lookup & Raw Data Fetches)
    // =========================================================================
    if (req.method === 'GET') {
      const { action, id } = req.query;

      if (action === 'getRawData') {
        try {
          const rawTab = await getTabTitle(sheets, SHEET_ID_KPI, "Raw Data");
          const rawRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SHEET_ID_KPI,
            range: `'${rawTab}'!A2:C1000`,
            valueRenderOption: 'FORMATTED_VALUE'
          });

          const rows = rawRes.data.values || [];
          const srNames = [];
          const resolutions = [];
          const dupReasons = [];

          rows.forEach(r => {
            if (r[0] && r[0].trim()) srNames.push(r[0].trim());
            if (r[1] && r[1].trim()) resolutions.push(r[1].trim());
            if (r[2] && r[2].trim()) dupReasons.push(r[2].trim());
          });

          return jsonResponse(res, 200, {
            status: "success",
            srNames: [...new Set(srNames)],
            resolutions: [...new Set(resolutions)],
            duplicateReasons: [...new Set(dupReasons)]
          });
        } catch (e) {
          return jsonResponse(res, 200, { status: "error", message: "Raw Data ဖတ်မရပါ: " + e.message });
        }
      }

      if (action === 'getEmployee' || id) {
        const searchId = id ? String(id).trim() : "";
        if (!searchId) {
          return jsonResponse(res, 200, { status: "error", message: "ID လိုအပ်ပါသည်။" });
        }

        try {
          const empTab = await getTabTitle(sheets, SHEET_ID_EMP, "Mgr Fb-30%");
          const empRes = await sheets.spreadsheets.values.get({
            spreadsheetId: SHEET_ID_EMP,
            range: `'${empTab}'!A2:B3000`,
            valueRenderOption: 'FORMATTED_VALUE'
          });

          const rows = empRes.data.values || [];
          const cleanSearch = searchId.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

          for (let i = 0; i < rows.length; i++) {
            const rawName = rows[i][0] ? String(rows[i][0]).trim() : "";
            const rawId = rows[i][1] ? String(rows[i][1]).trim() : "";

            const cleanIdCol = rawId.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();
            const cleanNameCol = rawName.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

            if (cleanIdCol === cleanSearch && cleanSearch !== "") {
              return jsonResponse(res, 200, { status: "success", id: searchId, name: rawName });
            }
            if (cleanNameCol === cleanSearch && cleanSearch !== "") {
              return jsonResponse(res, 200, { status: "success", id: searchId, name: rawId });
            }
          }
          
          return jsonResponse(res, 200, { status: "error", message: "Employee ID မရှိသေးပါ" });

        } catch (e) {
          return jsonResponse(res, 200, { status: "error", message: "Employee Sheet Error: " + e.message });
        }
      }

      return jsonResponse(res, 200, { status: "running", message: "KPI API is active!" });
    }

    // =========================================================================
    // 🌟 POST REQUEST (Validation & Double Sheet Scan Logic)
    // =========================================================================
    if (req.method === 'POST') {
      let bodyData = req.body;
      if (typeof bodyData === 'string') {
        try { bodyData = JSON.parse(bodyData); } catch (e) {}
      }

      const {
        date,
        srName,
        serviceId,
        customerName,
        financeId,
        financeName,
        reason,
        hasTicket,
        ticketInvoiceNo,
        resolutionDetail,
        otherResolutionReason,
        duplicateReason,
        otherDuplicateReason
      } = bodyData || {};

      if (!date || !srName || !serviceId || !financeName) {
        return jsonResponse(res, 200, { status: "error", message: "လိုအပ်သော Parameter များ မပါဝင်ပါ။" });
      }

      const kpiTab = await getTabTitle(sheets, SHEET_ID_KPI, "KPI Calculation");
      const dupTab = await getTabTitle(sheets, SHEET_ID_KPI, "KPI Duplicate");

      // 🔍 Sheet ၂ ခုစလုံး ('KPI Calculation' & 'Duplicate') မှ Data များကို ပြိုင်တူ ဆွဲထုတ်ခြင်း
      const [kpiRes, dupRes] = await Promise.all([
        sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID_KPI,
          range: `'${kpiTab}'!A:J`
        }).catch(() => ({ data: { values: [] } })),
        sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID_KPI,
          range: `'${dupTab}'!A:J`
        }).catch(() => ({ data: { values: [] } }))
      ]);

      const kpiRows = kpiRes.data.values || [];
      const dupRows = dupRes.data.values || [];

      // Sheet (၂) ခုလုံးမှ တင်ထားသမျှ Data များကို ပေါင်းစည်းခြင်း
      const combinedRows = [];
      for (let i = 1; i < kpiRows.length; i++) combinedRows.push(kpiRows[i]);
      for (let j = 1; j < dupRows.length; j++) combinedRows.push(dupRows[j]);

      // MM Time (UTC+6:30) Timestamp
      const now = new Date();
      const mmTime = new Date(now.getTime() + (6.5 * 60 * 60 * 1000));
      const mmMonth = String(mmTime.getUTCMonth() + 1).padStart(2, '0');
      const mmDay = String(mmTime.getUTCDate()).padStart(2, '0');
      const mmYear = mmTime.getUTCFullYear();
      const hours = String(mmTime.getUTCHours()).padStart(2, '0');
      const minutes = String(mmTime.getUTCMinutes()).padStart(2, '0');
      const seconds = String(mmTime.getUTCSeconds()).padStart(2, '0');
      const timestamp = `${mmMonth}/${mmDay}/${mmYear} ${hours}:${minutes}:${seconds}`;

      // Date Format: MM/DD/YYYY (e.g., 09/21/2026)
      let formattedDate = date;
      const dateParts = String(date).trim().split('-');
      if (dateParts.length === 3) {
        formattedDate = `${dateParts[1]}/${dateParts[2]}/${dateParts[0]}`;
      }
      const targetDate = new Date(formattedDate);

      let sevenDaysCount = 0;
      let sameDayRecords = [];

      const cleanCurrentId = String(serviceId).trim().toLowerCase();

      // Sheet ၂ ခုပေါင်းထားသော combinedRows ထဲတွင် ၇ရက်အတွင်း/၁ရက်အတွင်း Duplicate Scan ဖတ်ခြင်း
      for (let i = 0; i < combinedRows.length; i++) {
        const row = combinedRows[i];
        const rowDateStr = row[1] ? String(row[1]).replace("'", "").trim() : "";
        const rowServiceId = row[3] ? String(row[3]).trim().toLowerCase() : "";

        if (rowServiceId === cleanCurrentId && rowDateStr) {
          const rDate = new Date(rowDateStr);
          if (!isNaN(rDate.getTime())) {
            const diffTime = Math.abs(targetDate - rDate);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            if (diffDays <= 7) {
              sevenDaysCount++;
            }
            if (rDate.toDateString() === targetDate.toDateString()) {
              sameDayRecords.push({
                financeName: row[5] ? String(row[5]).trim() : "",
                financeId: row[6] ? String(row[6]).trim() : ""
              });
            }
          }
        }
      }

      // 🔴 Check 1: ၇ ရက်အတွင်း ၃ ကြိမ် ရောက်ပြီး ထပ်တင်ပါက (၄ ကြိမ်မြောက်ဖြစ်လျှင်) တင်ခွင့်မပြုပါ
      if (sevenDaysCount >= 3) {
        return jsonResponse(res, 200, {
          status: "blocked",
          message: "ဤ Customer သည် 7ရက်အတွင်း ၃ကြိမ် KPI လင်ခ် တင်ပြီးပြီဖြစ်သောကြောင့် ထပ်မံတင်ခွင့်မရှိတော့ပါ။"
        });
      }

      // Final Value Selection Logic
      let finalResolution = "";
      if (String(hasTicket) === "true" || String(hasTicket) === "yes") {
        finalResolution = ticketInvoiceNo ? `TKT No/Invoice No: ${ticketInvoiceNo}` : "TKT/Invoice Exists";
      } else {
        finalResolution = (resolutionDetail === "Others") ? otherResolutionReason : resolutionDetail;
      }

      let finalDupReason = (duplicateReason === "Others") ? otherDuplicateReason : duplicateReason;
      const sheetDateValue = "'" + formattedDate;

      // 🟡 Check 2: Sheet ၂ ခုလုံးမှ Scan ဖတ်၍ ယနေ့တွင် Duplicate ရှိနေပါက ဆုံးဖြတ်ခြင်း
      if (sameDayRecords.length > 0 && !duplicateReason) {
        const currentFinanceId = String(financeId).trim();
        const sameStaffRecord = sameDayRecords.find(r => r.financeId === currentFinanceId);

        if (sameStaffRecord) {
          return jsonResponse(res, 200, {
            status: "need_duplicate_reason",
            message: `ဤ Customer အား ယနေ့အတွင်း သင်ကိုယ်တိုင် (${sameDayRecords.length}) ကြိမ်တင်ပြီးပါပြီ။ Connection မကောင်းလို့ ထပ်တင်တာဆိုရင် "မှားတင်မိတာပါ" ကိုနှိပ်ပေးပါ။ နှစ်ခါရှင်းလိုက်ရတာဆိုရင် Duplicate Reason ထည့်ပေးပါ။`
          });
        } else {
          const otherStaff = sameDayRecords[0];
          return jsonResponse(res, 200, {
            status: "need_duplicate_reason",
            message: `ဤ Customer အား ယနေ့အတွင်း ${otherStaff.financeName} (ID - ${otherStaff.financeId}) မှ တင်ထားပြီးဖြစ်ပါသည်။ ထပ်ရှင်းပြပေးရပါက Duplicate Reason ထည့်ပေးပါ။`
          });
        }
      }

      // 🟢 Duplicate Reason ပါလာလျှင် Duplicate Sheet သို့ သိမ်းဆည်းခြင်း (Finance Name က အရင်၊ Finance ID က အနောက်)
      if (duplicateReason) {
        const dupRecord = [[
          timestamp,
          sheetDateValue,
          srName,
          serviceId,
          customerName,
          financeName,
          financeId,
          reason,
          finalResolution || "",
          finalDupReason || ""
        ]];

        await sheets.spreadsheets.values.append({
          spreadsheetId: SHEET_ID_KPI,
          range: `'${dupTab}'!A1`,
          valueInputOption: 'USER_ENTERED',
          resource: { values: dupRecord }
        });

        return jsonResponse(res, 200, {
          status: "success",
          message: "အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။"
        });
      }

      // 🟢 Normal Insert: KPI Calculation Sheet သို့ သိမ်းဆည်းခြင်း (Finance Name က အရင်၊ Finance ID က အနောက်)
      const kpiRecord = [[
        timestamp,
        sheetDateValue,
        srName,
        serviceId,
        customerName,
        financeName,
        financeId,
        reason,
        finalResolution || ""
      ]];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID_KPI,
        range: `'${kpiTab}'!A1`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: kpiRecord }
      });

      return jsonResponse(res, 200, { status: "success", message: "အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီ။" });
    }

  } catch (error) {
    return jsonResponse(res, 500, { status: "error", message: error.message || String(error) });
  }
};

// =========================
// CONFIG
// =========================
const RIS_SHEET = "RIS - Appointment Procedure Sum";
const RIS_HEADER_ROW = 8; // 1-based
const BILL_SHEET = "Report_ChargeTransactionDetail";
const BILL_HEADER_ROW = 10; // 1-based

// RIS headers A–U
const RIS_HEADERS = [
  "Modality","Location","Resource Name","Technologist","Radiologist",
  "Date of Service","Appointment ID","Accession Number","Dictation Status",
  "MRN","Patient Name","Patient DOB","Procedure","Exam Code","CPT Code",
  "Appointment Status","Order Priority","Referring Provider",
  "Payer Attached to Procedure","Primary Payer","Secondary Payer"
];

// Billing headers A–N
const BILL_HEADERS = [
  "Patient","Location","DOS","Charge Post","Procedure","ASA Code",
  "Charge Amt","Total Payment","Max Pay Date","Max Pay Post",
  "Primary Ins","Secondary Ins","Tertiary Ins","Order Num"
];

// =========================
// UTILITIES
// =========================
function normalize(v) {
  return v == null ? "" : String(v).trim();
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function extractRows(ws, headerRowIndex) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const rows = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? cell.v : "");
    }
    rows.push(row);
  }

  const header = rows[headerRowIndex - 1] || [];
  const data = rows.slice(headerRowIndex);

  return { header, data };
}

function buildIndex(header, data, keyName) {
  const col = header.findIndex(h => normalize(h) === keyName);
  const map = new Map();

  for (const row of data) {
    const key = normalize(row[col]);
    if (key && !map.has(key)) map.set(key, row);
  }

  return { map, col };
}

// =========================
// MAIN RECONCILIATION
// =========================
async function runReconciliation() {
  const risFile = document.getElementById("risFile").files[0];
  const billFile = document.getElementById("billingFile").files[0];
  const summary = document.getElementById("summary");

  if (!risFile || !billFile) {
    summary.textContent = "Select both RIS and Billing files.";
    return;
  }

  try {
    const [risWb, billWb] = await Promise.all([
      readWorkbook(risFile),
      readWorkbook(billFile)
    ]);

    const risWs = risWb.Sheets[RIS_SHEET];
    const billWs = billWb.Sheets[BILL_SHEET];

    if (!risWs) throw new Error(`Missing sheet: ${RIS_SHEET}`);
    if (!billWs) throw new Error(`Missing sheet: ${BILL_SHEET}`);

    // Extract RIS rows
    const { header: risHeader, data: risData } =
      extractRows(risWs, RIS_HEADER_ROW);

    // Extract Billing rows
    const { header: billHeader, data: billData } =
      extractRows(billWs, BILL_HEADER_ROW);

    // Build Billing index on Order Num
    const { map: billIndex } =
      buildIndex(billHeader, billData, "Order Num");

    // Find RIS Accession Number column
    const risAccCol = risHeader.findIndex(h => normalize(h) === "Accession Number");

    // Output arrays
    const MATCH = [];
    const NOMATCH = [];

    const outputHeader = [...RIS_HEADERS, "Reconcile", ...BILL_HEADERS];

    let matchCount = 0;
    let noMatchCount = 0;

    for (const row of risData) {
      if (!row.some(v => v !== "")) continue;

      const acc = normalize(row[risAccCol]);
      const risOut = RIS_HEADERS.map((h, i) => row[i] ?? "");

      if (billIndex.has(acc)) {
        const billRow = billIndex.get(acc);
        const billOut = BILL_HEADERS.map((h, i) => billRow[i] ?? "");
        MATCH.push([...risOut, "MATCH", ...billOut]);
        matchCount++;
      } else {
        NOMATCH.push([...risOut, "NO MATCH", ...Array(BILL_HEADERS.length).fill("")]);
        noMatchCount++;
      }
    }

    const total = matchCount + noMatchCount;

    // Build workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([outputHeader, ...MATCH]), "MATCH");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([outputHeader, ...NOMATCH]), "NO MATCH");

    const summarySheet = XLSX.utils.aoa_to_sheet([
      ["Reconcile", "Count", "Percent"],
      ["MATCH", matchCount, matchCount / total],
      ["NO MATCH", noMatchCount, noMatchCount / total],
      ["Grand Total", total, 1]
    ]);
    XLSX.utils.book_append_sheet(wb, summarySheet, "SUMMARY");

    XLSX.writeFile(wb, "RIS_Billing_Reconciliation.xlsx");

    summary.textContent =
      `MATCH: ${matchCount}\nNO MATCH: ${noMatchCount}\nTOTAL: ${total}`;

  } catch (err) {
    summary.textContent = "ERROR: " + err.message;
  }
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("runBtn").onclick = runReconciliation;
});

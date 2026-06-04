// =========================
// Utility Functions
// =========================

function normalizeAccession(value) {
  return String(value || "").trim();
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

  const header = rows[headerRowIndex];
  const data = rows.slice(headerRowIndex + 1);
  return { header, data };
}

// =========================
// DATE FORMATTING (FINAL FIX)
// =========================

const RIS_DATE_COLS = [4, 10];     // Date of Service, DOB
const BILL_DATE_COLS = [2, 3, 8, 9]; // DOS, Charge Post, Max Pay Date, Max Pay Post

function applyDateFormatting(ws, dateCols, startRow = 1) {
  const range = XLSX.utils.decode_range(ws["!ref"]);

  for (let r = startRow; r <= range.e.r; r++) {
    for (const c of dateCols) {
      const cellAddr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellAddr];
      if (!cell || cell.v === undefined || cell.v === null) continue;

      let parsed;

      // Case 1: Excel serial number (e.g., 46024)
      if (typeof cell.v === "number") {
        parsed = new Date(Date.UTC(1899, 11, 30) + cell.v * 86400000);
      }
      // Case 2: String date
      else {
        parsed = new Date(cell.v);
      }

      if (isNaN(parsed.getTime())) continue;

      // Convert JS Date → Excel serial number
      const excelSerial =
        (parsed - new Date(Date.UTC(1899, 11, 30))) / 86400000;

      cell.v = excelSerial;
      cell.t = "n";
      cell.z = "00/00/0000";
    }
  }
}

// =========================
// Main Reconciliation
// =========================

async function runReconciliation() {
  const summary = document.getElementById("summary");
  summary.textContent = "Processing…";

  await new Promise(r => setTimeout(r, 50));

  try {
    // -------------------------
    // Load Billing File
    // -------------------------
    const billingFile = document.getElementById("billingFile").files[0];
    if (!billingFile) {
      summary.textContent = "ERROR: Please upload the Billing file.";
      return;
    }

    const billingData = await billingFile.arrayBuffer();
    const billingWb = XLSX.read(billingData);
    const billingWs = billingWb.Sheets[billingWb.SheetNames[0]];

    const BILL_HEADER_ROW = 9;
    const BILL_KEY_COL = 13;

    const { header: billHeader, data: billData } =
      extractRows(billingWs, BILL_HEADER_ROW);

    const billDict = new Map();

    for (let i = 0; i < billData.length; i++) {
      const row = billData[i];
      const key = normalizeAccession(row[BILL_KEY_COL]);
      if (key.length > 0 && !billDict.has(key)) {
        billDict.set(key, i);
      }
    }

    // -------------------------
    // Load RIS File
    // -------------------------
    const risFile = document.getElementById("risFile").files[0];
    if (!risFile) {
      summary.textContent = "ERROR: Please upload the RIS file.";
      return;
    }

    const risDataBuf = await risFile.arrayBuffer();
    const risWb = XLSX.read(risDataBuf);
    const risWs = risWb.Sheets[risWb.SheetNames[0]];

    const RIS_HEADER_ROW = 7;
    const RIS_KEY_COL = 7;

    const { header: risHeader, data: risData } =
      extractRows(risWs, RIS_HEADER_ROW);

    // -------------------------
    // Reconciliation Logic
    // -------------------------
    const MATCH = [];
    const NOMATCH = [];

    let matchCount = 0;
    let noMatchCount = 0;

    const billColumnCount = billHeader.length;
    const emptyBillRow = Array(billColumnCount).fill("");

    for (let i = 0; i < risData.length; i++) {
      const risRow = risData[i];
      const accession = normalizeAccession(risRow[RIS_KEY_COL]);

      if (!accession) continue;

      if (billDict.has(accession)) {
        const billIndex = billDict.get(accession);
        const billRow = billData[billIndex];
        MATCH.push([...risRow, "MATCH", ...billRow]);
        matchCount++;
      } else {
        NOMATCH.push([...risRow, "NO MATCH", ...emptyBillRow]);
        noMatchCount++;
      }
    }

    // -------------------------
    // Build Output Workbook
    // -------------------------
    const outWb = XLSX.utils.book_new();

    // MATCH sheet
    const matchSheet = XLSX.utils.aoa_to_sheet([
      [...risHeader, "Status", ...billHeader],
      ...MATCH
    ]);
    XLSX.utils.book_append_sheet(outWb, matchSheet, "MATCH");
    applyDateFormatting(outWb.Sheets["MATCH"], [...RIS_DATE_COLS, ...BILL_DATE_COLS], 1);

    // NO MATCH sheet
    const noMatchSheet = XLSX.utils.aoa_to_sheet([
      [...risHeader, "Status", ...billHeader],
      ...NOMATCH
    ]);
    XLSX.utils.book_append_sheet(outWb, noMatchSheet, "NO MATCH");
    applyDateFormatting(outWb.Sheets["NO MATCH"], [...RIS_DATE_COLS, ...BILL_DATE_COLS], 1);

    // SUMMARY sheet
    XLSX.utils.book_append_sheet(
      outWb,
      XLSX.utils.aoa_to_sheet([
        ["MATCH", matchCount],
        ["NO MATCH", noMatchCount],
        ["TOTAL ACCESSIONS", matchCount + noMatchCount]
      ]),
      "SUMMARY"
    );

    XLSX.writeFile(outWb, "Reconciliation_Output.xlsx");

    summary.textContent =
      `MATCH: ${matchCount}\n` +
      `NO MATCH: ${noMatchCount}\n` +
      `TOTAL ACCESSIONS: ${matchCount + noMatchCount}`;

  } catch (err) {
    summary.textContent = "ERROR: " + err.message;
  }
}

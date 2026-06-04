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
// Convert ANY Excel date to readable string
// =========================

function fixDate(v) {
  // Excel serial number
  if (typeof v === "number") {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return d.toLocaleDateString("en-US");
  }

  // Text date
  const d = new Date(v);
  if (!isNaN(d)) {
    return d.toLocaleDateString("en-US");
  }

  return v; // Not a date
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

      // Fix RIS dates
      const risDOS = fixDate(risRow[4]);
      const risDOB = fixDate(risRow[10]);

      if (billDict.has(accession)) {
        const billIndex = billDict.get(accession);
        const billRow = billData[billIndex];

        // Fix Billing dates
        const billDOS = fixDate(billRow[2]);
        const billPost = fixDate(billRow[3]);
        const billMaxPay = fixDate(billRow[8]);
        const billMaxPost = fixDate(billRow[9]);

        MATCH.push([
          ...risRow.slice(0, 4),
          risDOS,
          ...risRow.slice(5, 10),
          risDOB,
          ...risRow.slice(11),
          "MATCH",
          ...billRow.slice(0, 2),
          billDOS,
          billPost,
          ...billRow.slice(4, 8),
          billMaxPay,
          billMaxPost,
          ...billRow.slice(10)
        ]);

        matchCount++;
      } else {
        NOMATCH.push([
          ...risRow.slice(0, 4),
          risDOS,
          ...risRow.slice(5, 10),
          risDOB,
          ...risRow.slice(11),
          "NO MATCH",
          ...emptyBillRow
        ]);

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

    // NO MATCH sheet
    const noMatchSheet = XLSX.utils.aoa_to_sheet([
      [...risHeader, "Status", ...billHeader],
      ...NOMATCH
    ]);
    XLSX.utils.book_append_sheet(outWb, noMatchSheet, "NO MATCH");

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

// =========================
// Utility Functions
// =========================

function normalizeAccession(value) {
  // EXACT VBA behavior: Trim only, no stripping
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
// Main Reconciliation
// =========================

async function runReconciliation() {
  const summary = document.getElementById("summary");
  summary.textContent = "Processing…";

  // Force UI update
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

    // VBA starts at row 11 → header row is 10 (0‑based)
    const BILL_HEADER_ROW = 10;
    const BILL_KEY_COL = 13; // Column N = index 13

    const { header: billHeader, data: billData } = extractRows(billingWs, BILL_HEADER_ROW);

    // Build dictionary EXACTLY like VBA
    const billDict = new Map();

    for (let i = 0; i < billData.length; i++) {
      const row = billData[i];
      const key = normalizeAccession(row[BILL_KEY_COL]);

      if (key.length > 0) {
        if (!billDict.has(key)) {
          billDict.set(key, i); // store FIRST matching row only
        }
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

    // VBA starts at row 9 → header row is 8 (0‑based)
    const RIS_HEADER_ROW = 8;
    const RIS_KEY_COL = 7; // Column H = index 7

    const { header: risHeader, data: risData } = extractRows(risWs, RIS_HEADER_ROW);

    // -------------------------
    // Reconciliation Logic
    // -------------------------
    const MATCH = [];
    const NOMATCH = [];

    let matchCount = 0;
    let noMatchCount = 0;

    const emptyBillRow = Array(billHeader.length).fill("");

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

    XLSX.utils.book_append_sheet(
      outWb,
      XLSX.utils.aoa_to_sheet([[...risHeader, "Status", ...billHeader], ...MATCH]),
      "MATCH"
    );

    XLSX.utils.book_append_sheet(
      outWb,
      XLSX.utils.aoa_to_sheet([[...risHeader, "Status", ...billHeader], ...NOMATCH]),
      "NO MATCH"
    );

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



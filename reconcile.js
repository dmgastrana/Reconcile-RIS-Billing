// =========================
// Utility Functions
// =========================

function normalize(value) {
  return String(value || "").trim().toLowerCase();
}

// VBA-style normalization: strip all non-digits so JS matches like Excel/VBA
function normalizeAccession(value) {
  return String(value || "")
    .trim()
    .replace(/\D/g, "");
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

function buildIndex(header, data, keyName) {
  const keyCol = header.findIndex(h => normalize(h) === normalize(keyName));
  const index = new Map();

  if (keyCol === -1) return index;

  for (const row of data) {
    const key = normalizeAccession(row[keyCol]);
    if (key) index.set(key, row);
  }

  return index;
}

// =========================
// Main Reconciliation
// =========================

async function runReconciliation() {
  const summary = document.getElementById("summary");
  summary.textContent = "Processing…";

  try {
    // -------------------------
    // Load Billing File
    // -------------------------
    const billingFile = document.getElementById("billingFile").files[0];
    const billingData = await billingFile.arrayBuffer();
    const billingWb = XLSX.read(billingData);
    const billingWs = billingWb.Sheets[billingWb.SheetNames[0]];

    // Billing header row is Excel row 10 → index 9
    const { header: billHeader, data: billData } = extractRows(billingWs, 9);

    // Build index on the column that matches RIS Accession Number
    const billIndex = buildIndex(billHeader, billData, "Order Num");

    // -------------------------
    // Load RIS File
    // -------------------------
    const risFile = document.getElementById("risFile").files[0];
    const risDataBuf = await risFile.arrayBuffer();
    const risWb = XLSX.read(risDataBuf);
    const risWs = risWb.Sheets[risWb.SheetNames[0]];

    // RIS header row is Excel row 8 → index 7
    const { header: risHeader, data: risData } = extractRows(risWs, 7);

    // Find Accession Number column
    const risAccCol = risHeader.findIndex(
      h => normalize(h) === "accession number"
    );

    if (risAccCol === -1) {
      summary.textContent = "ERROR: RIS file missing 'Accession Number' column.";
      return;
    }

    // -------------------------
    // Reconciliation Logic (optimized)
// -------------------------
    const MATCH = [];
    const NOMATCH = [];

    let matchCount = 0;
    let noMatchCount = 0;

    // Prebuild RIS rows (fast: no map inside main loop)
    const risRows = risData.map(row => risHeader.map((h, i) => row[i] ?? ""));

    // Prebuild Billing rows (fast lookup by normalized accession)
    const billRows = {};
    for (const [key, row] of billIndex.entries()) {
      billRows[key] = billHeader.map((h, i) => row[i] ?? "");
    }

    // Prebuild empty billing row for NO MATCH
    const emptyBillRow = Array(billHeader.length).fill("");

    // Single fast loop over RIS
    for (let i = 0; i < risData.length; i++) {
      const acc = normalizeAccession(risData[i][risAccCol]);
      if (!acc) continue;

      const risOut = risRows[i];

      if (billRows[acc]) {
        MATCH.push([...risOut, "MATCH", ...billRows[acc]]);
        matchCount++;
      } else {
        NOMATCH.push([...risOut, "NO MATCH", ...emptyBillRow]);
        noMatchCount++;
      }
    }

    // -------------------------
    // Build Output Workbook
    // -------------------------
    const outWb = XLSX.utils.book_new();

    const matchSheet = XLSX.utils.aoa_to_sheet([
      [...risHeader, "Status", ...billHeader],
      ...MATCH
    ]);
    XLSX.utils.book_append_sheet(outWb, matchSheet, "MATCH");

    const noMatchSheet = XLSX.utils.aoa_to_sheet([
      [...risHeader, "Status", ...billHeader],
      ...NOMATCH
    ]);
    XLSX.utils.book_append_sheet(outWb, noMatchSheet, "NO MATCH");

    const summarySheet = XLSX.utils.aoa_to_sheet([
      ["MATCH", matchCount],
      ["NO MATCH", noMatchCount],
      ["TOTAL ACCESSIONS", matchCount + noMatchCount]
    ]);
    XLSX.utils.book_append_sheet(outWb, summarySheet, "SUMMARY");

    XLSX.writeFile(outWb, "Reconciliation_Output.xlsx");

    // -------------------------
    // UI Summary
    // -------------------------
    summary.textContent =
      `MATCH: ${matchCount}\n` +
      `NO MATCH: ${noMatchCount}\n` +
      `TOTAL ACCESSIONS: ${matchCount + noMatchCount}`;

  } catch (err) {
    summary.textContent = "ERROR: " + err.message;
  }
}


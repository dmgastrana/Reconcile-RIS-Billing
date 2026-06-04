// =========================
// Utility Functions
// =========================

function normalize(value) {
  return String(value || "").trim().toLowerCase();
}

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

  // Force UI update
  await Promise.resolve();

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

    const { header: billHeader, data: billData } = extractRows(billingWs, 9);
    const billIndex = buildIndex(billHeader, billData, "Order Num");

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

    const { header: risHeader, data: risData } = extractRows(risWs, 7);

    const risAccCol = risHeader.findIndex(
      h => normalize(h) === "accession number"
    );

    if (risAccCol === -1) {
      summary.textContent = "ERROR: RIS file missing 'Accession Number' column.";
      return;
    }

    // -------------------------
    // Reconciliation Logic
    // -------------------------
    const MATCH = [];
    const NOMATCH = [];

    let matchCount = 0;
    let noMatchCount = 0;

    const risRows = risData.map(row => risHeader.map((h, i) => row[i] ?? ""));

    const billRows = {};
    for (const [key, row] of billIndex.entries()) {
      billRows[key] = billHeader.map((h, i) => row[i] ?? "");
    }

    const emptyBillRow = Array(billHeader.length).fill("");

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


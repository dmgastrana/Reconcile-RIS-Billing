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

// Convert Excel serial or text date → "MM/DD/YYYY"
function fixDate(v) {
  let d;

  if (typeof v === "number") {
    d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  } else {
    d = new Date(v);
  }

  if (isNaN(d)) return v;

  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const yyyy = d.getUTCFullYear();
  return `${mm}/${dd}/${yyyy}`;
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

    const { header: billHeader, data: billData } =
      extractRows(billingWs, BILL_HEADER_ROW);

    // ⭐ FIXED: detect "Order Num" instead of hardcoding column 13
    const BILL_KEY_COL = billHeader.indexOf("Order Num");
    if (BILL_KEY_COL === -1) {
      summary.textContent = "ERROR: Billing file missing 'Order Num' column.";
      return;
    }

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

    const dosIndex = risHeader.indexOf("Date of Service");
    const dobIndex = risHeader.indexOf("Patient DOB");

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

      const fixedRIS = [...risRow];

      // ⭐ FIX DOS
      fixedRIS[dosIndex] = fixDate(risRow[dosIndex]);

      // ⭐ FIX DOB
      if (dobIndex !== -1) {
        fixedRIS[dobIndex] = fixDate(risRow[dobIndex]);
      }

      if (billDict.has(accession)) {
        const billIndex = billDict.get(accession);
        const billRow = billData[billIndex];

        MATCH.push([
          ...fixedRIS,
          "MATCH",
          ...billRow
        ]);

        matchCount++;
      } else {
        NOMATCH.push([
          ...fixedRIS,
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

    XLSX.utils.book_append_sheet(
      outWb,
      XLSX.utils.aoa_to_sheet([
        [...risHeader, "Reconcile", ...billHeader],
        ...MATCH
      ]),
      "MATCH"
    );

    XLSX.utils.book_append_sheet(
      outWb,
      XLSX.utils.aoa_to_sheet([
        [...risHeader, "Reconcile", ...billHeader],
        ...NOMATCH
      ]),
      "NO MATCH"
    );

    const total = matchCount + noMatchCount;
    const matchPct = ((matchCount / total) * 100).toFixed(2) + "%";
    const noMatchPct = ((noMatchCount / total) * 100).toFixed(2) + "%";

    XLSX.utils.book_append_sheet(
      outWb,
      XLSX.utils.aoa_to_sheet([
        ["Reconcile", "Count", "Percent"],
        ["MATCH", matchCount, matchPct],
        ["NO MATCH", noMatchCount, noMatchPct],
        ["TOTAL ACCESSION", total, "100%"]
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

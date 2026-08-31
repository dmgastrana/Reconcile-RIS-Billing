// =========================
// Utility Functions
// =========================

function normalizeAccession(value) { return String(value || "").trim(); }

function normalizeName(raw) {
  let s = String(raw || "");
  if (s.includes("-")) {
    const parts = s.split("-");
    s = parts[parts.length - 1].trim();
  }
  s = s.replace(/\u00A0|\u3000|\u200B|\uFEFF|\u202F/g, " ");
  s = s.replace(/[,.;]/g, " ").replace(/\s\s+/g, " ").trim().toUpperCase();
  if (!s) return "";
  const parts = s.split(" ");
  if (String(raw).includes(",")) {
    const last = parts[0];
    return (parts.slice(1).join(" ") + " " + last).trim();
  }
  return (parts.slice(0, -1).join(" ") + " " + parts[parts.length - 1]).trim();
}

function cleanCPT(rawCPT) {
  const s = String(rawCPT || "");
  let out = "";
  for (let i = 0; i < s.length; i++) if (/[0-9]/.test(s[i])) out += s[i];
  return out;
}

function fixDate(v) {
  if (!v) return v;
  let d;
  if (typeof v === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    d = new Date(excelEpoch.getTime() + v * 86400000);
  } else d = new Date(v);
  if (isNaN(d)) return v;
  return `${String(d.getUTCMonth() + 1).padStart(2,"0")}/${String(d.getUTCDate()).padStart(2,"0")}/${d.getUTCFullYear()}`;
}

// =========================
// PASS 1 — Order Num match
// =========================

function firstPass_OrderNum(risAOA, billAOA, colMap, startReconCol) {
  const dictOrder = {};
  const orderCol = colMap["Order Num"];
  for (let r = 10; r < billAOA.length; r++) {
    const key = String(billAOA[r][orderCol] || "").trim();
    if (key) {
      if (!dictOrder[key]) dictOrder[key] = [];
      dictOrder[key].push(r);
    }
  }
  for (let r = 8; r < risAOA.length; r++) {
    const accession = normalizeAccession(risAOA[r][7]);
    if (!accession) continue;
    if (dictOrder[accession]) {
      const bestRow = dictOrder[accession][0];
      const billRow = billAOA[bestRow];
      risAOA[r][startReconCol] = "MATCH";
      risAOA[r][startReconCol+1] = billRow[colMap["Patient"]];
      risAOA[r][startReconCol+2] = billRow[colMap["Location"]];
      risAOA[r][startReconCol+3] = fixDate(billRow[colMap["DOS"]]);
      risAOA[r][startReconCol+4] = fixDate(billRow[colMap["Charge Post"]]);
      risAOA[r][startReconCol+5] = billRow[colMap["Procedure"]];
      risAOA[r][startReconCol+6] = billRow[colMap["ASA Code"]];
      risAOA[r][startReconCol+7] = billRow[colMap["Charge Amt"]];
      risAOA[r][startReconCol+8] = billRow[colMap["Total Payment"]];
      risAOA[r][startReconCol+9] = fixDate(billRow[colMap["Max Pay Date"]]);
      risAOA[r][startReconCol+10] = fixDate(billRow[colMap["Max Pay Post"]]);
      risAOA[r][startReconCol+11] = billRow[colMap["Primary Ins"]];
      risAOA[r][startReconCol+12] = billRow[colMap["Secondary Ins"]];
      risAOA[r][startReconCol+13] = billRow[colMap["Tertiary Ins"]];
      risAOA[r][startReconCol+14] = billRow[colMap["Order Num"]];
    }
  }
  return risAOA;
}

// =========================
// PASS 2 — Name + CPT match
// =========================

function secondPass_NameCPT(risAOA, billAOA, colMap, startReconCol) {
  const dictNameCPT = {};
  const patientCol = colMap["Patient"];
  const procCol = colMap["Procedure"];
  let firstChargeRow = 0;

  for (let r = 0; r < billAOA.length; r++) {
    if (String(billAOA[r][patientCol] || "").includes("-")) {
      firstChargeRow = r;
      break;
    }
  }

  if (firstChargeRow > 0) {
    for (let r = firstChargeRow; r < billAOA.length; r++) {
      const rptName = normalizeName(billAOA[r][patientCol]);
      const rptCPT = cleanCPT(billAOA[r][procCol]);
      if (rptName && rptCPT) {
        const key = rptName + "|" + rptCPT;
        if (!dictNameCPT[key]) dictNameCPT[key] = r;
      }
    }

    for (let r = 8; r < risAOA.length; r++) {
      if (String(risAOA[r][startReconCol] || "").trim().toUpperCase() === "MATCH") continue;
      if (normalizeAccession(risAOA[r][7])) continue;

      const risName = normalizeName(risAOA[r][11]);
      const cpt = cleanCPT(risAOA[r][19]);
      if (risName && cpt) {
        const key = risName + "|" + cpt;
        if (dictNameCPT[key] != null) {
          const bestRow = dictNameCPT[key];
          const billRow = billAOA[bestRow];
          risAOA[r][startReconCol] = "MATCH";
          risAOA[r][startReconCol+1] = billRow[colMap["Patient"]];
          risAOA[r][startReconCol+2] = billRow[colMap["Location"]];
          risAOA[r][startReconCol+3] = fixDate(billRow[colMap["DOS"]]);
          risAOA[r][startReconCol+4] = fixDate(billRow[colMap["Charge Post"]]);
          risAOA[r][startReconCol+5] = billRow[colMap["Procedure"]];
          risAOA[r][startReconCol+6] = billRow[colMap["ASA Code"]];
          risAOA[r][startReconCol+7] = billRow[colMap["Charge Amt"]];
          risAOA[r][startReconCol+8] = billRow[colMap["Total Payment"]];
          risAOA[r][startReconCol+9] = fixDate(billRow[colMap["Max Pay Date"]]);
          risAOA[r][startReconCol+10] = fixDate(billRow[colMap["Max Pay Post"]]);
          risAOA[r][startReconCol+11] = billRow[colMap["Primary Ins"]];
          risAOA[r][startReconCol+12] = billRow[colMap["Secondary Ins"]];
          risAOA[r][startReconCol+13] = billRow[colMap["Tertiary Ins"]];
          risAOA[r][startReconCol+14] = billRow[colMap["Order Num"]];
        }
      }
    }
  }
  return risAOA;
}

// =========================
// PASS 3 — duplicate-no accession
// =========================

function flagDuplicateNoAccession(risAOA, startReconCol) {
  const dict = {};
  for (let i = 8; i < risAOA.length; i++) {
    const accession = String(risAOA[i][7] || "").trim();
    if (accession) {
      const key = `${risAOA[i][11]}|${risAOA[i][5]}|${risAOA[i][10]}|${risAOA[i][6]}`;
      dict[key] = true;
    }
  }
  for (let i = 8; i < risAOA.length; i++) {
    const accession = String(risAOA[i][7] || "").trim();
    const reconcile = String(risAOA[i][startReconCol] || "").trim();
    const radiologist = String(risAOA[i][4] || "").trim();
    if (!accession && !reconcile && !radiologist) {
      const key = `${risAOA[i][11]}|${risAOA[i][5]}|${risAOA[i][10]}|${risAOA[i][6]}`;
      if (dict[key]) risAOA[i][startReconCol] = "duplicate-no accession";
    }
  }
  return risAOA;
}

// =========================
// PASS 4 — duplicate-different accession & no radiology
// =========================

function flagDifferentAccessionNoRadiology(risAOA, startReconCol) {
  const dict = {};
  for (let i = 8; i < risAOA.length; i++) {
    const key = `${risAOA[i][11]}|${risAOA[i][5]}|${risAOA[i][10]}|${risAOA[i][6]}`;
    const accession = String(risAOA[i][7] || "").trim();
    if (!dict[key]) dict[key] = accession;
    else if (accession && dict[key] !== accession) dict[key] = "MULTI";
  }
  for (let i = 8; i < risAOA.length; i++) {
    const key = `${risAOA[i][11]}|${risAOA[i][5]}|${risAOA[i][10]}|${risAOA[i][6]}`;
    const accession = String(risAOA[i][7] || "").trim();
    const radiologist = String(risAOA[i][4] || "").trim();
    const reconcile = String(risAOA[i][startReconCol] || "").trim();
    if (dict[key] === "MULTI" && accession && !radiologist && !reconcile) {
      risAOA[i][startReconCol] = "duplicate-different accession & no radiology";
    }
  }
  return risAOA;
}

// =========================
// PASS 5 — NO MATCH for Completed WO Report / Reported
// =========================

function flagNoMatchCompletedOrReported(risAOA, startReconCol) {
  for (let r = 8; r < risAOA.length; r++) {
    const markVal = String(risAOA[r][startReconCol] || "").trim();
    const statusVal = String(risAOA[r][24] || "").trim();
    if (!markVal && (statusVal === "Completed WO Report" || statusVal === "Reported")) {
      risAOA[r][startReconCol] = "NO MATCH";
    }
  }
  return risAOA;
}

// =========================
// PASS 6 — FINAL CLEANUP
// =========================

function finalizeReconcileColumn(risAOA, startReconCol) {
  for (let r = 8; r < risAOA.length; r++) {
    const val = String(risAOA[r][startReconCol] || "").trim();
    if (!["MATCH","NO MATCH","duplicate-no accession","duplicate-different accession & no radiology"].includes(val)) {
      risAOA[r][startReconCol] = "NO MATCH";
    }
  }
  return risAOA;
}

// =========================
// MAIN RECONCILIATION
// =========================

async function runReconciliation() {
  const summary = document.getElementById("summary");
  summary.textContent = "Processing…";
  await new Promise(r => setTimeout(r, 50));

  try {
    const billingFile = document.getElementById("billingFile").files[0];
    if (!billingFile) { summary.textContent = "ERROR: Please upload the Billing file."; return; }

    const billingData = await billingFile.arrayBuffer();
    const billingWb = XLSX.read(billingData);
    const wsRpt = billingWb.Sheets[billingWb.SheetNames[0]];
    const billAOA = XLSX.utils.sheet_to_json(wsRpt, { header: 1 });

    const risFile = document.getElementById("risFile").files[0];
    if (!risFile) { summary.textContent = "ERROR: Please upload the RIS file."; return; }

    const risDataBuf = await risFile.arrayBuffer();
    const risWb = XLSX.read(risDataBuf);
    const wsRIS = risWb.Sheets[risWb.SheetNames[0]];
    let risAOA = XLSX.utils.sheet_to_json(wsRIS, { header: 1 });

    // ⭐ Extract DOS range (Row 5, Column A)
    let dosRange = "";
    const headerLine = String(risAOA[4][0] || "").trim();
    if (headerLine.startsWith("Report ran for the period")) {
      dosRange = headerLine.replace("Report ran for the period:", "").trim();
    }
    window.reconDOS = dosRange;

    // ⭐ Fix RIS DOS column (index 5)
    for (let r = 8; r < risAOA.length; r++) risAOA[r][5] = fixDate(risAOA[r][5]);

    // Remove footer
    risAOA = risAOA.filter(row => !String(row[0] || "").trim().startsWith("Confidential and Proprietary"));

    const billHeaderRow = billAOA[9] || [];
    const colMap = {};
    for (let c = 0; c < billHeaderRow.length; c++) {
      const hdr = String(billHeaderRow[c] || "").trim();
      if (hdr && !colMap[hdr]) colMap[hdr] = c;
    }

    const headerRow = 7;
    const startReconCol = (risAOA[headerRow] || []).length;
    const newHeaders = [
      "Reconcile","Patient","Location","DOS","Charge Post","Procedure","ASA Code",
      "Charge Amt","Total Payment","Max Pay Date","Max Pay Post","Primary Ins",
      "Secondary Ins","Tertiary Ins","Order Num"
    ];

    if (!risAOA[headerRow]) risAOA[headerRow] = [];
    for (let i = 0; i < newHeaders.length; i++) {
      risAOA[headerRow][startReconCol + i] = newHeaders[i];
    }

    risAOA = firstPass_OrderNum(risAOA, billAOA, colMap, startReconCol);
    risAOA = secondPass_NameCPT(risAOA, billAOA, colMap, startReconCol);
    risAOA = flagDuplicateNoAccession(risAOA, startReconCol);
    risAOA = flagDifferentAccessionNoRadiology(risAOA, startReconCol);
    risAOA = flagNoMatchCompletedOrReported(risAOA, startReconCol);
    risAOA = finalizeReconcileColumn(risAOA, startReconCol);

    // =========================
    // SUMMARY CALCULATIONS
    // =========================

    const totalAppt = risAOA.length - 8;
    let matchCount = 0, noMatchCount = 0;

    for (let r = 8; r < risAOA.length; r++) {
      const val = String(risAOA[r][startReconCol] || "").trim().toUpperCase();
      if (val === "MATCH") matchCount++;
      if (val === "NO MATCH") noMatchCount++;
    }

    const pctNoMatch = ((noMatchCount / matchCount) * 100).toFixed(2);

    window.totalAppt = totalAppt;
    window.matchCount = matchCount;
    window.noMatchCount = noMatchCount;
    window.pctNoMatch = pctNoMatch;




// =========================
// GROUP BY RECONCILE (Pivot-style)
// =========================

const groupCounts = {
  "duplicate-different accession & no radiology": 0,
  "duplicate-no accession": 0,
  "MATCH": 0,
  "NO MATCH": 0
};

for (let r = 8; r < risAOA.length; r++) {
  const val = String(risAOA[r][startReconCol] || "").trim();
  if (groupCounts[val] != null) {
    groupCounts[val]++;
  }
}

groupCounts["Grand Total"] = risAOA.length - 8;

// expose to HTML
window.groupCounts = groupCounts;


    














    

    // =========================
    // WRITE OUTPUT FILE
    // =========================

    const outSheet = XLSX.utils.aoa_to_sheet(risAOA);
    const outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, outSheet, "RIS - Appointment Procedure Sum");
    XLSX.writeFile(outWb, "Reconciliation_Output.xlsx");

    summary.textContent = "Reconciliation complete.\nOutput file downloaded.";

  } catch (err) {
    summary.textContent = "ERROR: " + err.message;
  }
}

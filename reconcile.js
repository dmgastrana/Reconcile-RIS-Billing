async function runReconciliation() {

    // 1) Load Billing + RIS
    // 2) Build column map
    // 3) Write headers V–AJ

    // ---------------------------------------------------------
    // PASS 1 — Order Num match (VBA: Append_ChargeDetail_To_RIS)
    // ---------------------------------------------------------
    risAOA = firstPass_OrderNum(risAOA, billAOA, colMap);

    // ---------------------------------------------------------
    // PASS 2 — Name + CPT match (VBA: SecondMatch_Name_CPT_Only)
    // ---------------------------------------------------------
    risAOA = secondPass_NameCPT(risAOA, billAOA, colMap);

    // ---------------------------------------------------------
    // PASS 3 — duplicate-no-accession
    // (VBA: findduplicatenoaccession)
    // ---------------------------------------------------------
    risAOA = flagDuplicateNoAccession(risAOA);

    // ---------------------------------------------------------
    // PASS 4 — duplicate-different accession & no radiology
    // (VBA: flagDifferentAccessionNoRadiology)
    // ---------------------------------------------------------
    risAOA = flagDifferentAccessionNoRadiology(risAOA);

    // ---------------------------------------------------------
    // PASS 5 — NO MATCH for Completed WO Report / Reported
    // (VBA: write_NoMatch_Completed_WO_Report_and_Reported)
    // ---------------------------------------------------------
    risAOA = flagNoMatchCompletedOrReported(risAOA);

    // ---------------------------------------------------------
    // Write workbook
    // ---------------------------------------------------------
    const outSheet = XLSX.utils.aoa_to_sheet(risAOA);
    const outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, outSheet, "RIS - Appointment Procedure Sum");
    XLSX.writeFile(outWb, "Reconciliation_Output.xlsx");

    // Update webpage summary
    window.matchCount = matchCount;
    window.noMatchCount = noMatchCount;
}



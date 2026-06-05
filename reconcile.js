<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Reconciliation Tool</title>

  <!-- SheetJS -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

  <!-- Reconciliation Logic -->
  <script src="reconcile.js"></script>

  <style>
    body {
      font-family: Arial, sans-serif;
      background: #f4f6f9;
      margin: 0;
      padding: 40px;
      display: flex;
      justify-content: center;
    }

    .container {
      background: white;
      padding: 30px;
      width: 650px;
      border-radius: 12px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    }

    h2 {
      text-align: center;
      margin-bottom: 20px;
      color: #333;
    }

    .upload-btn {
      width: 100%;
      padding: 12px;
      background: #6b7280;
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      cursor: pointer;
      margin-bottom: 5px;
    }

    .upload-btn:hover {
      background: #4b5563;
    }

    input[type="file"] {
      display: none;
    }

    button {
      width: 100%;
      margin-top: 10px;
      padding: 14px;
      background: #2563eb;
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      cursor: pointer;
      transition: 0.2s;
    }

    button:hover {
      background: #1e40af;
    }

    .upload-status {
      margin: 12px 0 28px 0;
      color: green;
      font-size: 16px;
      font-weight: 500;
    }

    #summary {
      margin-top: 20px;
      white-space: pre;
      font-weight: bold;
      text-align: left;
      font-family: monospace;
      font-size: 16px;
      line-height: 1.6;
    }

    #passwordPage { display: block; }
    #reconPage { display: none; }

    #pwError {
      color: red;
      text-align: center;
      margin-top: 10px;
      font-weight: bold;
    }
  </style>
</head>

<body>

<div class="container">

  <!-- PASSWORD PAGE -->
  <div id="passwordPage">
    <h2>Enter Password</h2>
    <input id="pwInput" type="password" placeholder="Enter password"
           style="width:100%;padding:12px;font-size:16px;border-radius:8px;border:1px solid #ccc;">
    <button onclick="checkPassword()">Submit</button>
    <div id="pwError"></div>
  </div>

  <!-- RECON PAGE -->
  <div id="reconPage">

    <h2>Reconciliation</h2>

    <!-- BILLING FILE -->
    <label class="upload-btn" for="billingFile">Upload Billing Charge Report</label>
    <input type="file" id="billingFile" onchange="showFileName('billingFile','billingStatus')">
    <div id="billingStatus" class="upload-status"></div>

    <!-- RIS FILE -->
    <label class="upload-btn" for="risFile">Upload Abbadox RIS Procedure Report</label>
    <input type="file" id="risFile" onchange="showFileName('risFile','risStatus')">
    <div id="risStatus" class="upload-status"></div>

    <!-- RECONCILE BUTTON -->
    <button id="runBtn">Reconcile</button>

    <!-- SUMMARY OUTPUT -->
    <div id="summary"></div>

  </div>

</div>

<script>
function checkPassword() {
  const pw = document.getElementById("pwInput").value;
  if (pw === "dmg1129") {
    document.getElementById("passwordPage").style.display = "none";
    document.getElementById("reconPage").style.display = "block";
  } else {
    document.getElementById("pwError").textContent = "Incorrect password";
  }
}

function showFileName(inputId, statusId) {
  const fileInput = document.getElementById(inputId);
  const statusDiv = document.getElementById(statusId);
  if (fileInput.files.length > 0) {
    statusDiv.textContent = "Uploaded: " + fileInput.files[0].name;
  } else {
    statusDiv.textContent = "";
  }
}

document.getElementById("runBtn").addEventListener("click", async () => {
  await runReconciliation();

  const total = window.matchCount + window.noMatchCount;

  const matchPct = ((window.matchCount / total) * 100).toFixed(2);
  const noMatchPct = ((window.noMatchCount / total) * 100).toFixed(2);

  document.getElementById("summary").textContent =
    "Reconcile            Count       Percent\n" +
    "MATCH                " + window.matchCount.toLocaleString() + "       " + matchPct + "%\n" +
    "NO MATCH             " + window.noMatchCount.toLocaleString() + "       " + noMatchPct + "%\n" +
    "TOTAL ACCESSION      " + total.toLocaleString() + "       100%";
});
</script>

</body>
</html> 

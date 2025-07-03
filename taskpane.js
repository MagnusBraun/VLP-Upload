Office.onReady(() => {
  document.getElementById("fileInput").onchange = uploadPDF;
});

async function uploadPDF() {
  const input = document.getElementById("fileInput");
  const files = input.files;
  if (files.length === 0) {
    showError("Bitte wähle mindestens eine PDF-Datei aus.");
    return;
  }

  const preview = document.getElementById("preview");
  preview.innerHTML = "<p><em>PDFs werden verarbeitet...</em></p>";

  const allResults = [];

  for (const file of files) {
    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch("https://pmfusion-api.onrender.com/process", {
        method: "POST",
        body: formData
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Fehler bei der Verarbeitung");
      }

      const data = await res.json();
      allResults.push(data);
    } catch (err) {
      console.error("Fehler:", err.message);
      showError("Verarbeitung fehlgeschlagen: " + err.message);
      return;
    }
  }

  // Alle PDF-Ergebnisse kombinieren
  const combined = {};
  for (const data of allResults) {
    for (const key in data) {
      combined[key] = (combined[key] || []).concat(data[key]);
    }
  }

function previewInTable(mapped) {
  const preview = document.getElementById("preview");
  preview.innerHTML = "";

  const headers = Object.keys(mapped);
  const maxLength = Math.max(...headers.map(k => mapped[k].length));

  const table = document.createElement("table");
  table.border = "1";

  const thead = table.createTHead();
  const headRow = thead.insertRow();
  headers.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    headRow.appendChild(th);
  });

  const tbody = table.createTBody();
  for (let i = 0; i < maxLength; i++) {
    const row = tbody.insertRow();
    headers.forEach(h => {
      const cell = row.insertCell();
      cell.textContent = mapped[h][i] || "";
    });
  }

  preview.appendChild(table);

  const insertBtn = document.createElement("button");
  insertBtn.textContent = "In Excel einfügen";
  insertBtn.onclick = () => insertToExcel(mapped);
  preview.appendChild(insertBtn);
}

async function insertToExcel(mapped) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headerRange = sheet.getRange("A1:Z1");
    headerRange.load("values");
    await context.sync();
    
    const excelHeaders = headerRange.values?.[0] || [];
    if (excelHeaders.length === 0) {
      console.log("Keine Spaltenüberschriften in Excel gefunden.");
      return;
    }

    const colCount = excelHeaders.length;
    const maxRows = Math.max(...Object.values(mapped).map(col => col.length));

    const usedRange = sheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    const startRow = usedRange.rowCount;

    // 🔄 Normiere mapped keys (klein, getrimmt)
    const normalizedMap = {};
    for (const key in mapped) {
      const norm = key.toLowerCase().trim();
      normalizedMap[norm] = mapped[key];
    }

    const dataRows = [];
    for (let i = 0; i < maxRows; i++) {
      const row = [];
      for (let h = 0; h < colCount; h++) {
        const header = excelHeaders[h];
        const normHeader = header ? header.toString().toLowerCase().trim() : "";
        const colData = normalizedMap[normHeader] || [];
        row.push(colData[i] || "");
      }
      row.some(cell => cell !== "") && dataRows.push(row); // füge nur nicht-leere Zeilen ein
    }

    if (dataRows.length === 0) {
      console.log("Keine passenden Datenzeilen gefunden");
      return;
    }

    const range = sheet.getRangeByIndexes(startRow, 0, dataRows.length, colCount);
    range.values = dataRows;
    range.format.font.name = "Calibri";
    range.format.font.size = 11;
    range.format.horizontalAlignment = "Left";


    await context.sync();
  });
}


function showError(msg) {
  const preview = document.getElementById("preview");
  preview.innerHTML = `<div style="color:red;font-weight:bold">${msg}</div>`;
}

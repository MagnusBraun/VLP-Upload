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
  const errors = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const formData = new FormData();
    formData.append("file", file);

    preview.innerHTML = `<p><em>Verarbeite Datei ${i + 1} von ${files.length}: ${file.name}</em></p>`;

    try {
      const res = await fetch("https://vlp-upload.onrender.com/process", {
        method: "POST",
        body: formData
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || "Serverfehler");
      }

      const data = await res.json();
      allResults.push(data);
    } catch (err) {
      errors.push(`${file.name}: ${err.message}`);
      continue;
    }
  }

  input.value = "";

  if (allResults.length === 0) {
    showError("Keine gültigen PDF-Dateien verarbeitet.");
    return;
  }

  const combined = {};
  for (const data of allResults) {
    for (const key in data) {
      combined[key] = (combined[key] || []).concat(data[key]);
    }
  }

  previewInTable(combined);

  if (errors.length > 0) {
    const errorDiv = document.createElement("div");
    errorDiv.style.color = "orangered";
    errorDiv.style.marginTop = "1em";
    errorDiv.innerHTML = "<strong>Folgende Dateien konnten nicht verarbeitet werden:</strong><br>" +
                         errors.map(e => `• ${e}`).join("<br>");
    preview.appendChild(errorDiv);
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

// normalize string: lowercase, no spaces/symbols
function normalizeLabel(label) {
  return label.toLowerCase().replace(/[^a-z0-9]/gi, "");
}

// map excel headers to closest pdf keys
function createHeaderMap(excelHeaders, mappedKeys) {
  const excelMap = {};
  const normalizedExcel = excelHeaders.map(h => normalizeLabel(h));
  const normalizedPDF = mappedKeys.map(k => normalizeLabel(k));

  for (let i = 0; i < excelHeaders.length; i++) {
    const excelHeaderNorm = normalizedExcel[i];
    let bestMatch = null;
    for (let j = 0; j < mappedKeys.length; j++) {
      const pdfNorm = normalizedPDF[j];
      if (excelHeaderNorm === pdfNorm ||
          pdfNorm.includes(excelHeaderNorm) ||
          excelHeaderNorm.includes(pdfNorm)) {
        bestMatch = mappedKeys[j];
        break;
      }
    }
    excelMap[excelHeaders[i]] = bestMatch;
  }

  return excelMap;
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

    const headerMap = createHeaderMap(excelHeaders, Object.keys(mapped));

    const dataRows = [];
    for (let i = 0; i < maxRows; i++) {
      const row = [];
      for (let h = 0; h < colCount; h++) {
        const excelHeader = excelHeaders[h];
        const pdfKey = headerMap[excelHeader];
        const colData = pdfKey ? mapped[pdfKey] : [];
        row.push(colData[i] || "");
      }
      if (row.some(cell => cell !== "")) {
        dataRows.push(row);
      }
    }

    if (dataRows.length === 0) {
      console.log("Keine passenden Datenzeilen gefunden.");
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


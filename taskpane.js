Office.onReady(() => {
  const input = document.getElementById("fileInput");
  if (input) {
    input.onchange = uploadPDF;
  }
});

const apiUrl = "https://vlp-upload.onrender.com/process";
const storageKey = "pmfusion-column-mapping";

function normalizeLabel(label) {
  return label.toLowerCase().replace(/[^a-z0-9]/gi, "");
}

const columnAliases = {
  "Kabelnummer": ["kabelnummer", "kabel-nr", "kabelnr", "knr", "kabnr"],
  "Kabeltyp": ["typ", "kabel-typ", "kabeltype"],
  "Trommelnummer": ["trommelnummer", "trommel-nr", "trommel-nummer"],
  "Durchmesser": ["durchmesser", "ø", "dm", "ømm", "Ømm", "Ø"],
  "von Ort": ["von ort"],
  "bis Ort": ["bis ort"],
  "von km": ["von km", "von kilometer"],
  "bis km": ["bis km", "bis kilometer"],
  "Metr. (von)": ["metr. von"],
  "Metr. (bis)": ["metr. bis"],
  "SOLL": ["soll"],
  "IST": ["ist"],
  "Verlegeart": ["verlegeart"],
  "Bemerkung": ["bemerkung", "bemerkungen"]
};

function loadSavedMappings() {
  const json = localStorage.getItem(storageKey);
  return json ? JSON.parse(json) : {};
}

function saveMappings(headerMap) {
  localStorage.setItem(storageKey, JSON.stringify(headerMap));
}

function resetMappings() {
  localStorage.removeItem(storageKey);
  alert("Gespeicherte Zuordnungen wurden zurückgesetzt.");
}

function createHeaderMapWithAliases(excelHeaders, mappedKeys, aliases) {
  const excelMap = {};
  const normMapped = {};
  mappedKeys.forEach(k => {
    normMapped[normalizeLabel(k)] = k;
  });

  for (const excelHeader of excelHeaders) {
    const cleaned = excelHeader?.trim();
    if (!cleaned) continue;

    const aliasList = aliases[cleaned] || [];
    let match = null;
    for (const alias of aliasList) {
      const normAlias = normalizeLabel(alias);
      if (normMapped[normAlias]) {
        match = normMapped[normAlias];
        break;
      }
    }

    excelMap[excelHeader] = match || null;
  }

  return excelMap;
}

async function resolveMissingMappings(headerMap, mappedKeys) {
  return new Promise((resolve) => {
    const missing = Object.entries(headerMap).filter(([k, v]) => k.trim() !== "" && v === null);
    if (missing.length === 0) return resolve(headerMap);

    const overlay = document.createElement("div");
    overlay.style.position = "fixed";
    overlay.style.top = "0";
    overlay.style.left = "0";
    overlay.style.width = "100%";
    overlay.style.height = "100%";
    overlay.style.backgroundColor = "rgba(0,0,0,0.4)";
    overlay.style.zIndex = "9999";
    overlay.style.padding = "2em";
    overlay.style.overflow = "auto";

    const box = document.createElement("div");
    box.style.background = "white";
    box.style.padding = "1em";
    box.style.borderRadius = "8px";
    box.style.maxWidth = "500px";
    box.style.margin = "auto";

    const title = document.createElement("h3");
    title.textContent = "Manuelle Spaltenzuordnung erforderlich:";
    box.appendChild(title);

    missing.forEach(([excelCol]) => {
      const label = document.createElement("label");
      label.textContent = `Excel: ${excelCol}`;
      label.style.display = "block";
      label.style.marginTop = "10px";

      const select = document.createElement("select");
      select.dataset.excelCol = excelCol;

      const none = document.createElement("option");
      none.value = "";
      none.textContent = "Keine Zuordnung";
      select.appendChild(none);

      mappedKeys.forEach(key => {
        const option = document.createElement("option");
        option.value = key;
        option.textContent = key;
        select.appendChild(option);
      });

      box.appendChild(label);
      box.appendChild(select);
    });

    const button = document.createElement("button");
    button.textContent = "Zuordnung übernehmen";
    button.style.marginTop = "1em";
    button.onclick = () => {
      const selects = box.querySelectorAll("select");
      selects.forEach(select => {
        const col = select.dataset.excelCol;
        const val = select.value;
        if (val) headerMap[col] = val;
      });
      overlay.remove();
      resolve(headerMap);
    };

    box.appendChild(button);
    overlay.appendChild(box);
    document.body.appendChild(overlay);
  });
}

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
      const res = await fetch(apiUrl, {
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

  const resetBtn = document.createElement("button");
  resetBtn.textContent = "Zuordnungen zurücksetzen";
  resetBtn.style.marginLeft = "1em";
  resetBtn.onclick = resetMappings;
  preview.appendChild(resetBtn);
}

async function insertToExcel(mapped) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headerRange = sheet.getRange("A1:Z1");
    headerRange.load("values");
    await context.sync();

    const excelHeaders = headerRange.values?.[0] || [];
    if (excelHeaders.length === 0) return;

    const colCount = excelHeaders.length;
    const maxRows = Math.max(...Object.values(mapped).map(col => col.length));

    const usedRange = sheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    const startRow = usedRange.rowCount;

    const saved = loadSavedMappings();
    let headerMap = createHeaderMapWithAliases(excelHeaders, Object.keys(mapped), columnAliases);
    for (const key in saved) {
      if (headerMap[key] === null && saved[key]) {
        headerMap[key] = saved[key];
      }
    }

    headerMap = await resolveMissingMappings(headerMap, Object.keys(mapped));
    saveMappings(headerMap);

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

    if (dataRows.length === 0) return;

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

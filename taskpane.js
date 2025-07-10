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
function extractVLPNumber(filename) {
  const match = filename.match(/VLP[\s\-]*(\d+)/i);
  return match ? `VLP ${match[1]}` : "";
}

const columnAliases = {
  "Kabelnummer": ["kabelnummer", "kabel-nr", "kabelnr"],
  "Kabeltyp": ["typ", "kabel-typ", "kabeltype"],
  "Trommelnummer": ["trommelnummer", "trommel-nr", "trommel-nummer"],
  "Durchmesser": ["durchmesser", "ø", "ømm", "Ømm", "Ø"],
  "von Ort": ["von ort"],
  "bis Ort": ["bis ort"],
  "von km": ["von km", "von kilometer"],
  "bis km": ["bis km", "bis kilometer"],
  "Metr. (von)": ["metr. von"],
  "Metr. (bis)": ["metr. bis"],
  "SOLL": ["soll"],
  "IST": ["ist"],
  "Verlegeart": ["verlegeart"],
  "Bemerkung": ["bemerkung", "bemerkungen"],
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

    let match = null;
    const normExcel = normalizeLabel(cleaned);

    // Direkt 1:1 Match prüfen
    if (normMapped[normExcel]) {
      match = normMapped[normExcel];
    } else {
      // Alias-Check: durchsuche alle Aliase
      for (const [stdLabel, aliasList] of Object.entries(aliases)) {
        for (const alias of aliasList) {
          if (normalizeLabel(alias) === normExcel && normMapped[normExcel]) {
            match = normMapped[normExcel];
            break;
          }
        }
        if (match) break;
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

      let data = await res.json();
      const vlpNumber = extractVLPNumber(file.name);
      const keys = Object.keys(data);
      const rowCount = Object.values(data)[0]?.length || 0;
      
      // Leere Daten verhindern, auch wenn VLP existiert
      const filteredData = {};
      for (const key of keys) {
        filteredData[key] = [];
      }
      filteredData["VLP"] = [];
      
      for (let i = 0; i < rowCount; i++) {
        let hasRealContent = false;

        for (const key of keys) {
          const value = data[key]?.[i];
          const str = value?.toString().trim();
          if (str && str !== "0") {
            hasRealContent = true;
            break;
          }
        }
      
        if (hasRealContent) {
          for (const key of keys) {
            filteredData[key].push(data[key]?.[i] ?? "");
          }
          filteredData["VLP"].push(vlpNumber);
        }
      }
      
      // Wenn nach dem Filtern KEINE Zeile übrig bleibt
      const filteredRowCount = filteredData["VLP"].length;
      if (filteredRowCount === 0) {
        throw new Error("Keine gültigen Datenzeilen in dieser PDF.");
      }
      
      data = filteredData;
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
  // Stelle sicher, dass VLP ganz am Ende steht
  if (!headers.includes("VLP") && mapped["VLP"]) {
    headers.push("VLP");
  }
  
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
    usedRange.load(["values", "rowCount"]);
    await context.sync();

    const startRow = usedRange.rowCount;
    const insertedRowNumbers = [];
    const insertedKeys = new Set();

    const saved = loadSavedMappings();
    let headerMap = createHeaderMapWithAliases(excelHeaders, Object.keys(mapped), columnAliases);
    for (const key in saved) {
      if (headerMap[key] === null && saved[key]) {
        headerMap[key] = saved[key];
      }
    }
    headerMap = await resolveMissingMappings(headerMap, Object.keys(mapped));
    saveMappings(headerMap);

    const existingRows = usedRange.values.slice(1);
    const keyCols = ["Kabelnummer", "von Ort", "von km", "bis Ort", "bis km"];
    const keyIndexes = keyCols
      .map(key => excelHeaders.findIndex(h => normalizeLabel(h) === normalizeLabel(key)))
      .filter(i => i !== -1);

    const existingKeys = new Set(
      existingRows.map(row =>
        keyIndexes.map(i => (row[i] || "").toString().trim().toLowerCase()).join("|")
      )
    );

    const dataRows = [];

    for (let i = 0; i < maxRows; i++) {
      const row = [];
      const keyParts = [];
      for (let h = 0; h < colCount; h++) {
        const excelHeader = excelHeaders[h];
        const pdfKey = headerMap[excelHeader];
        let colData = [];
        if (mapped.hasOwnProperty(excelHeader)) {
          colData = mapped[excelHeader];
        } else if (pdfKey && mapped.hasOwnProperty(pdfKey)) {
          colData = mapped[pdfKey];
        }
        const val = colData[i] || "";
        row.push(val);
        if (keyIndexes.includes(h)) {
          keyParts.push(val.toString().trim().toLowerCase());
        }
      }
      const keyString = keyParts.join("|");
      if (!existingKeys.has(keyString)) {
        existingKeys.add(keyString);
        const newRowNum = startRow + dataRows.length + 1;
        insertedRowNumbers.push(newRowNum);
        insertedKeys.add(keyString);
        dataRows.push(row);
      }
    }

    if (dataRows.length > 0) {
      const range = sheet.getRangeByIndexes(startRow, 0, dataRows.length, colCount);
      range.values = dataRows;
      range.format.font.name = "Calibri";
      range.format.font.size = 11;
      range.format.horizontalAlignment = "Left";
      await context.sync();
    }
    // 🧹 Nach dem Einfügen: fehlerhafte (fast leere) Zeilen entfernen
    const cleanupCols = ["Kabelnummer", "Kabeltyp", "von Ort", "bis Ort"];
    const cleanupIndexes = cleanupCols.map(col =>
      excelHeaders.findIndex(h => normalizeLabel(h) === normalizeLabel(col))
    ).filter(i => i !== -1);
    
    const invalidRows = [];
    
    const rowRanges = insertedRowNumbers.map(rowNum => {
      const range = sheet.getRangeByIndexes(rowNum - 1, 0, 1, colCount);
      range.load("values");
      return { rowNum, range };
    });
    await context.sync();
    
    for (const { rowNum, range } of rowRanges) {
      const values = range.values[0];
      const allRelevantEmpty = cleanupIndexes.every(i =>
        !values[i] || values[i].toString().trim() === ""
      );
    
      if (allRelevantEmpty) {
        invalidRows.push(rowNum);
      }
    }
    
    // Jetzt löschen – von unten nach oben!
    for (const row of invalidRows.sort((a, b) => b - a)) {
      sheet.getRangeByIndexes(row - 1, 0, 1, colCount).delete(Excel.DeleteShiftDirection.up);
    }
    
    await context.sync();

    // ✅ Jetzt Duplikate prüfen, vor Sortierung!
    await detectAndHandleDuplicates(context, sheet, excelHeaders, insertedRowNumbers, insertedKeys);

    // 📊 Jetzt sortieren nach Kabelnummer
    const updatedRange = sheet.getUsedRange();
    updatedRange.load("rowCount");
    await context.sync();
    const kabelIndex = excelHeaders.findIndex(h => normalizeLabel(h) === normalizeLabel("Kabelnummer"));
    const vonKmIndex = excelHeaders.findIndex(h => normalizeLabel(h) === normalizeLabel("von km"));
    
    if (kabelIndex !== -1 && vonKmIndex !== -1) {
      const sortRange = sheet.getRangeByIndexes(1, 0, updatedRange.rowCount - 1, colCount);
      sortRange.sort.apply([
        { key: kabelIndex, ascending: true },
        { key: vonKmIndex, ascending: true }
      ]);
      await context.sync();
    }
    await applyDuplicateBoxHighlightingAfterSort(context, sheet);
    await applyKabelnummerGroupingBorders(context, sheet, excelHeaders);
    // 🧹 Leere Zeilen entfernen
    const fullRange = sheet.getUsedRange();
    fullRange.load(["values", "rowCount"]);
    await context.sync();

    const emptyRows = fullRange.values.map((row, idx) => ({
      isEmpty: row.every(cell => cell === "" || cell === null),
      idx
    })).filter(r => r.isEmpty).map(r => r.idx + 1).sort((a, b) => b - a);

    for (const row of emptyRows) {
      sheet.getRange(`A${row}:Z${row}`).delete(Excel.DeleteShiftDirection.up);
    }
    await context.sync();
  });
}


async function removeEmptyRows(context, sheet) {
  const usedRange = sheet.getUsedRange();
  usedRange.load(["values", "rowCount", "columnCount"]);
  await context.sync();

  const rows = usedRange.values;
  const rowCount = usedRange.rowCount;
  const colCount = usedRange.columnCount;

  const rowsToDelete = [];

  for (let i = 1; i < rowCount; i++) { // Zeile 0 = Header
    const isEmpty = rows[i].every(cell => !cell || cell.toString().trim() === "");
    if (isEmpty) rowsToDelete.push(i + 1); // Excel ist 1-basiert
  }

  for (const r of rowsToDelete.reverse()) {
    sheet.getRange(`A${r}:Z${r}`).delete(Excel.DeleteShiftDirection.up);
  }

  await context.sync();
}

function showDuplicateChoiceDialog(message, onSkipNew, onReplaceOld, onKeepAllMarked) {
  const overlay = document.createElement("div");
  overlay.style.position = "fixed";
  overlay.style.top = "0";
  overlay.style.left = "0";
  overlay.style.width = "100%";
  overlay.style.height = "100%";
  overlay.style.backgroundColor = "rgba(0,0,0,0.4)";
  overlay.style.zIndex = "9999";
  overlay.style.display = "flex";
  overlay.style.alignItems = "center";
  overlay.style.justifyContent = "center";

  const dialog = document.createElement("div");
  dialog.style.background = "white";
  dialog.style.padding = "1.5em";
  dialog.style.borderRadius = "8px";
  dialog.style.maxWidth = "450px";
  dialog.style.textAlign = "center";
  dialog.style.boxShadow = "0 0 10px rgba(0,0,0,0.3)";

  const msg = document.createElement("p");
  msg.textContent = message;
  msg.style.whiteSpace = "pre-line";
  dialog.appendChild(msg);

  const buttons = document.createElement("div");
  buttons.style.marginTop = "1em";

  const btn1 = document.createElement("button");
  btn1.textContent = "1: Duplikate nicht hinzufügen";
  btn1.style.margin = "0.5em";
  btn1.onclick = () => {
    overlay.remove();
    onSkipNew();
  };

  const btn2 = document.createElement("button");
  btn2.textContent = "2: Alte Zeilen ersetzen";
  btn2.style.margin = "0.5em";
  btn2.onclick = () => {
    overlay.remove();
    onReplaceOld();
  };

  const btn3 = document.createElement("button");
  btn3.textContent = "3: Duplikate behalten & markieren";
  btn3.style.margin = "0.5em";
  btn3.onclick = () => {
    overlay.remove();
    onKeepAllMarked();
  };

  buttons.appendChild(btn1);
  buttons.appendChild(btn2);
  buttons.appendChild(btn3);
  dialog.appendChild(buttons);
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);
}


async function detectAndHandleDuplicates(context, sheet, headers, insertedRowNumbers = []) {
  const keyCols = ["Kabelnummer", "von Ort", "von km", "bis Ort", "bis km"];
  const keyIndexes = keyCols.map(k =>
    headers.findIndex(h => normalizeLabel(h) === normalizeLabel(k))
  ).filter(i => i !== -1);

  if (keyIndexes.length < 2) return;

  const usedRange = sheet.getUsedRange();
  usedRange.load(["values", "rowCount"]);
  await context.sync();

  const allRows = usedRange.values.slice(1);
  const newRowSet = new Set(insertedRowNumbers);
  const existingRows = allRows.filter((_, idx) => !newRowSet.has(idx + 2)); // +2 = Header + 1-based

  const existingKeyMap = new Map();
  existingRows.forEach((row, idx) => {
    const key = keyIndexes.map(i => (row[i] || "").toString().trim().toLowerCase()).join("|");
    const excelRowNum = idx + 2;
    if (!existingKeyMap.has(key)) existingKeyMap.set(key, []);
    existingKeyMap.get(key).push(excelRowNum);
  });

  const dupeNewRows = [];
  const dupeOldRows = new Set();
  const duplicateKeys = new Set();

  const startCol = headers.findIndex(h => normalizeLabel(h) === "kabelnummer");
  const endCol = headers.findIndex(h => normalizeLabel(h) === "vlp");
  const colCount = endCol >= startCol ? endCol - startCol + 1 : 1;

  for (const rowNum of insertedRowNumbers) {
    const range = sheet.getRangeByIndexes(rowNum - 1, 0, 1, headers.length);
    range.load("values");
    await context.sync();

    const row = range.values[0];
    const key = keyIndexes.map(i => (row[i] || "").toString().trim().toLowerCase()).join("|");
    const dupOlds = existingKeyMap.get(key) || [];

    if (dupOlds.length > 0) {
      const newRange = sheet.getRangeByIndexes(rowNum - 1, startCol, 1, colCount);
      newRange.format.fill.color = "#FFD966"; // dunkelgelb
      dupeNewRows.push(rowNum);

      for (const dup of dupOlds) {
        const dupRange = sheet.getRangeByIndexes(dup - 1, startCol, 1, colCount);
        dupRange.format.fill.color = "#FFF2CC"; // hellgelb
        dupeOldRows.add(dup);
        duplicateKeys.add(key);
      }
    }
  }

  if (dupeNewRows.length === 0) return;

  await context.sync();

  return new Promise(resolve => {
    showDuplicateChoiceDialog(
      `${dupeNewRows.length} Duplikate erkannt. Wie möchtest du fortfahren?`,
      async () => {
        // 1. NICHT hinzufügen
        for (const row of dupeNewRows.sort((a, b) => b - a)) {
          sheet.getRangeByIndexes(row - 1, startCol, 1, colCount).delete(Excel.DeleteShiftDirection.up);
        }
        for (const row of dupeOldRows) {
          sheet.getRangeByIndexes(row - 1, startCol, 1, colCount).format.fill.clear();
        }
        await context.sync();
        resolve();
      },
      async () => {
        // 2. ALTE ZEILEN ERSETZEN (nach dem Löschen: nochmal clear auf die neuen Reihenfolge!)
        for (const row of [...dupeOldRows].sort((a, b) => b - a)) {
          sheet.getRangeByIndexes(row - 1, 0, 1, headers.length).delete(Excel.DeleteShiftDirection.up);
        }
        await context.sync();

        const updatedRange = sheet.getUsedRange();
        updatedRange.load(["rowCount"]);
        await context.sync();

        const remainingStartRow = updatedRange.rowCount - insertedRowNumbers.length + 1;

        for (let i = 0; i < insertedRowNumbers.length; i++) {
          const rowIndex = remainingStartRow + i - 1;
          const clearRange = sheet.getRangeByIndexes(rowIndex, startCol, 1, colCount);
          clearRange.format.fill.clear();
        }

        await context.sync();
        resolve();
      },
      async () => {
        // 3. BEHALTEN & später umrahmen
        for (const row of [...dupeOldRows, ...dupeNewRows]) {
          sheet.getRangeByIndexes(row - 1, startCol, 1, colCount).format.fill.clear();
        }

        // In Workbook-Einstellungen speichern (statt in NamedRange!)
        context.workbook.settings.add("DuplikatKeys", JSON.stringify({
          keys: [...duplicateKeys],
          startCol,
          colCount,
          keyCols: keyCols.map(k => normalizeLabel(k))
        }));
        await context.sync();
        resolve();
      }
    );
  });
}


async function applyDuplicateBoxHighlightingAfterSort(context, sheet) {
  const setting = context.workbook.settings.getItemOrNullObject("DuplikatKeys");
  setting.load("value"); // <-- wichtig!
  await context.sync();

  // Jetzt sicher prüfen und verwenden
  if (setting.isNullObject || !setting.value) return;

  const raw = setting.value;
  setting.delete(); // optional
  await context.sync();

  const { keys, startCol, colCount, keyCols } = JSON.parse(raw);

  if (!keys || !Array.isArray(keys) || keys.length === 0) return;
  if (startCol < 0 || colCount <= 0) return;

  const usedRange = sheet.getUsedRange();
  usedRange.load(["values", "rowCount", "columnCount"]);
  await context.sync();

  const headerRange = sheet.getRange("A1:Z1");
  headerRange.load("values");
  await context.sync();

  const headerRow = headerRange.values[0];
  const keyIndexes = keyCols.map(k =>
    headerRow.findIndex(h => normalizeLabel(h) === k)
  ).filter(i => i !== -1);

  if (keyIndexes.length < 2) return;

  const values = usedRange.values.slice(1); // ohne Header
  const matchedKeys = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const key = keyIndexes.map(j => (row[j] || "").toString().trim().toLowerCase()).join("|");
    if (keys.includes(key)) {
      if (!matchedKeys.has(key)) matchedKeys.set(key, []);
      matchedKeys.get(key).push(i + 1); // Excel-Zeile (1-basiert, +1 für Header)
    }
  }

 for (const rows of matchedKeys.values()) {
  for (const row of rows) {
    const cellRange = sheet.getRangeByIndexes(row, startCol, 1, colCount);
    cellRange.format.font.color = "#B8860B"; // Dunkelgelb
  }
}


  await context.sync();
}
async function applyKabelnummerGroupingBorders(context, sheet, headers) {
  const usedRange = sheet.getUsedRange();
  usedRange.load("values");
  await context.sync();

  const rows = usedRange.values;
  const kabelIndex = headers.findIndex(h => normalizeLabel(h) === "kabelnummer");
  const startCol = headers.findIndex(h => normalizeLabel(h) === "kabelnummer");
  const endCol = headers.findIndex(h => normalizeLabel(h) === "vlp");
  const colCount = endCol >= startCol ? endCol - startCol + 1 : 1;

  if (kabelIndex === -1 || startCol === -1 || endCol === -1) return;

  let currentGroup = [];
  let currentKabel = null;

  for (let i = 1; i < rows.length; i++) { // Zeile 1 = Index 0 = Header
    const row = rows[i];
    const kabelVal = row[kabelIndex]?.toString().trim() ?? "";

    if (kabelVal !== currentKabel) {
      if (currentGroup.length > 0) {
        await drawThinBorder(sheet, currentGroup[0] + 1, currentGroup[currentGroup.length - 1] + 1, startCol, colCount);
      }
      currentGroup = [i];
      currentKabel = kabelVal;
    } else {
      currentGroup.push(i);
    }
  }

  if (currentGroup.length > 0) {
    await drawThinBorder(sheet, currentGroup[0] + 1, currentGroup[currentGroup.length - 1] + 1, startCol, colCount);
  }

  await context.sync();
}

async function drawThinBorder(sheet, startRow, endRow, startCol, colCount) {
  const range = sheet.getRangeByIndexes(startRow, startCol, endRow - startRow + 1, colCount);
  const borders = range.format.borders;

  ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(edge => {
    const border = borders.getItem(edge);
    border.style = "Continuous";
    border.weight = "Thin";
    border.color = "black";
  });
}

function showError(msg) {
  const preview = document.getElementById("preview");
  preview.innerHTML = `<div style="color:red;font-weight:bold">${msg}</div>`;
}

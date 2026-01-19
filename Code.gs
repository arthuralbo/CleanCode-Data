/**
 * @OnlyCurrentDoc
 * CleanCode Data Engine
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CleanCode Data')
      .addItem('Launch CleanCode Data', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate().setTitle('CleanCode Data');
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialSchema() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const firstRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  return headers.map((name, i) => {
    let type = "Categorical";
    const sample = firstRow[i];
    if (sample instanceof Date) type = "Date";
    else if (typeof sample === "number") type = "Numeric";
    return { name: name || `Col ${i+1}`, index: i + 1, suggestedType: type };
  });
}

function performDiagnostic(selections) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // --- PHASE 1: STRICT VALIDATION GATEKEEPER ---
  for (let sel of selections) {
    const colIdx = sel.index - 1;
    const colName = headers[colIdx] || `Colonne ${sel.index}`;

    for (let r = 0; r < rows.length; r++) {
      const cell = rows[r][colIdx];
      
      // On ignore les cellules vides pour la validation de type
      if (cell === "" || cell === null || cell === undefined) continue; 

      if (sel.type === "Numeric") {
        // Vérifie si la valeur est un nombre ou peut être convertie proprement
        const isInvalid = isNaN(parseFloat(cell)) || (typeof cell === 'string' && /[a-zA-Z]/.test(cell));
        if (isInvalid) {
          sheet.getRange(r + 2, sel.index).activate(); // Sélectionne la cellule fautive
          return { 
            error: `Numeric Error: The value "${cell}" in [${colName}] at Row ${r + 2} is not a number.` 
          };
        }
      }

      if (sel.type === "Date") {
        let d = (cell instanceof Date) ? cell : new Date(cell);
        if (isNaN(d.getTime())) {
          sheet.getRange(r + 2, sel.index).activate();
          return { 
            error: `Erreur Date : La valeur "${cell}" dans [${colName}] à la ligne ${r + 2} n'est pas une date valide.` 
          };
        }
      }
    }
  }

  // --- PHASE 2: REPORT GENERATION ---
  // If we reach here, the data types are valid. Now we count for the Badges.
  let report = {
    hasMissing: false, missingCount: 0, missingDetails: {},
    hasDates: false, dateCount: 0, dateDetails: {},
    hasOutliers: false, outlierCount: 0, outlierDetails: {},
    hasStrings: false, stringCount: 0, stringDetails: {}, 
    hasCleanup: false, cleanupCount: 0, cleanupDetails: {},
    totalIssues: 0
  };

  selections.forEach(sel => {
    const colIdx = sel.index - 1;
    const colName = headers[colIdx];
    let colNumericValues = [];

    rows.forEach(row => {
      const cell = row[colIdx];
      
      // 1. Missing Values
      if (cell === "" || cell === null) {
        report.hasMissing = true; 
        report.missingCount++;
        report.missingDetails[colName] = (report.missingDetails[colName] || 0) + 1;
        return;
      }

      // 2. Categorical (Cleanup & Encoding)
      if (sel.type === "Categorical") {
        report.hasStrings = true;
        report.stringDetails[colName] = (report.stringDetails[colName] || 0) + 1;
        
        // Detect if cleanup is needed
        if (typeof cell === 'string' && (cell !== cell.trim() || cell !== cell.toLowerCase())) {
          report.hasCleanup = true; 
          report.cleanupCount++;
          report.cleanupDetails[colName] = (report.cleanupDetails[colName] || 0) + 1;
        }
      }

      // 3. Dates
      if (sel.type === "Date") {
        report.hasDates = true;
        // This ensures the Date Badge pill gets a value
        report.dateDetails[colName] = (report.dateDetails[colName] || 0) + 1;
      }

      // 4. Numeric (for Outliers)
      if (sel.type === "Numeric") {
        colNumericValues.push(parseFloat(cell));
      }
    });

    // 5. Outlier Calculation
    if (sel.type === "Numeric" && colNumericValues.length > 2) {
      const n = colNumericValues.length;
      const mean = colNumericValues.reduce((a, b) => a + b) / n;
      const stdDev = Math.sqrt(colNumericValues.reduce((s, x) => s + Math.pow(x - mean, 2), 0) / n);
      const outliers = colNumericValues.filter(x => Math.abs(x - mean) > (3 * stdDev));
      if (outliers.length > 0) {
        report.hasOutliers = true; 
        report.outlierCount += outliers.length;
        report.outlierDetails[colName] = (report.outlierDetails[colName] || 0) + outliers.length;
      }
    }
  });

  report.totalIssues = report.missingCount + report.outlierCount + report.cleanupCount;
  return report;
}

/**
 * Transforms messy date columns based on user-selected Locale and Format.
 */
function applyDateTransformation(dateConfigs) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const timezone = Session.getScriptTimeZone();
  
  // Sort descending to handle column shifts
  const sortedConfigs = dateConfigs.sort((a, b) => b.index - a.index);

  sortedConfigs.forEach(config => {
    const colIdx = config.index;
    const colName = headers[colIdx - 1];
    const raw = data.slice(1).map(r => r[colIdx - 1]);
    
    const processed = raw.map(val => {
      if (val === "" || val === null) return [""];
      let d;
      
      // Better Date Parsing for US vs EU
      if (typeof val === 'string' && val.includes('/')) {
        let p = val.split(/[\/\-\. ]/);
        let day = (config.locale === "US") ? (p[1] || 1) : p[0];
        let month = (config.locale === "US") ? p[0] : (p[1] || 1);
        let year = p[2] || new Date().getFullYear();
        if (year.toString().length === 2) year = "20" + year;
        d = new Date(year, month - 1, day);
      } else {
        d = (val instanceof Date) ? val : new Date(val);
      }

      if (d && !isNaN(d.getTime())) {
        if (config.format === "UNIX") return [Math.floor(d.getTime() / 1000).toString()];
        if (config.format === "US_LONG") return [Utilities.formatDate(d, timezone, "MMMM dd, yyyy")];
        return [Utilities.formatDate(d, timezone, "yyyy-MM-dd")];
      }
      return [val];
    });

    sheet.insertColumnAfter(colIdx);
    const targetRange = sheet.getRange(2, colIdx + 1, processed.length, 1);
    sheet.getRange(1, colIdx + 1).setValue(colName + "_cleaned").setFontWeight("bold");
    
    // FORCE PLAIN TEXT to prevent Sheets from auto-changing "March 02" back to "3/2/2026"
    targetRange.setNumberFormat("@"); 
    targetRange.setValues(processed);
  });
  return "Date transformation successful. Created new columns.";
}

function applyScalingTransformation(scalingConfigs) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const fullData = sheet.getDataRange().getValues();
  const headers = fullData[0];
  const rowsToDelete = new Set(); // Using a Set to avoid duplicate row deletions

  // Sort descending by index to handle column additions correctly
  const sortedConfigs = scalingConfigs.sort((a, b) => b.index - a.index);

  sortedConfigs.forEach(config => {
    const colIdx = config.index;
    const rawValues = sheet.getRange(2, colIdx, sheet.getLastRow() - 1, 1).getValues();
    const numericValues = rawValues.map(r => parseFloat(r[0])).filter(v => !isNaN(v));

    if (numericValues.length < 2) return;

    // Calculate Stats for Thresholds
    const n = numericValues.length;
    const mean = numericValues.reduce((a, b) => a + b, 0) / n;
    const stdDev = Math.sqrt(numericValues.reduce((s, x) => s + Math.pow(x - mean, 2), 0) / n);
    const upper = mean + (3 * stdDev);
    const lower = mean - (3 * stdDev);

    if (config.method === "DROP") {
      // 1. COLLECT ROWS FOR DELETION
      rawValues.forEach((row, i) => {
        const val = parseFloat(row[0]);
        if (!isNaN(val) && (val > upper || val < lower)) {
          rowsToDelete.add(i + 2); // +2 for header and 1-based indexing
        }
      });
    } else {
      // 2. STANDARD SCALING (Create New Column)
      const newHeader = `${headers[colIdx - 1]}_${config.method.toLowerCase()}`;
      sheet.insertColumnAfter(colIdx);
      sheet.getRange(1, colIdx + 1).setValue(newHeader).setFontWeight("bold");

      const transformed = rawValues.map(row => {
        let x = parseFloat(row[0]);
        if (isNaN(x)) return [row[0]];
        let res;
        if (config.method === "Z-SCORE") res = (stdDev === 0) ? 0 : (x - mean) / stdDev;
        else if (config.method === "MIN-MAX") {
          const min = Math.min(...numericValues);
          const max = Math.max(...numericValues);
          res = (max === min) ? 0 : (x - min) / (max - min);
        }
        else if (config.method === "LOG") res = x < 0 ? 0 : Math.log1p(x);
        else if (config.method === "WINSORIZE") res = x > upper ? upper : (x < lower ? lower : x);
        return [res];
      });
      sheet.getRange(2, colIdx + 1, transformed.length, 1).setValues(transformed).setNumberFormat("0.0000");
    }
  });

  // 3. PHYSICAL DELETION (Process backwards to keep indices valid)
  if (rowsToDelete.size > 0) {
    const sortedRows = Array.from(rowsToDelete).sort((a, b) => b - a);
    sortedRows.forEach(rowIdx => sheet.deleteRow(rowIdx));
    return `Cleaned! Deleted ${rowsToDelete.size} rows and applied scaling.`;
  }

  return "Scaling applied successfully.";
}


function applyMissingTransformation(configs) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowsToDelete = new Set();
  
  const sortedConfigs = configs.sort((a, b) => b.index - a.index);

  sortedConfigs.forEach(config => {
    const colIdx = config.index;
    const raw = data.slice(1).map(r => r[colIdx - 1]);

    if (config.method === "DROP") {
      raw.forEach((val, i) => { if (val === "" || val === null) rowsToDelete.add(i + 2); });
    } else {
      let lastKnownValue = null; // Memory for Forward Fill
      
      // Calculate Stats (only for non-fill methods)
      const nums = raw.map(v => parseFloat(v)).filter(v => !isNaN(v));
      const mean = nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : 0;
      
      const processed = raw.map((val, i) => {
        if (val !== "" && val !== null) {
          lastKnownValue = val;
          return [val];
        }
        
        switch (config.method) {
          case "MEAN": return [mean];
          case "ZERO": return [0];
          case "FORWARD": return [lastKnownValue || ""]; // Carrying forward memory
          case "CUSTOM": return [config.customVal];
          default: return [""];
        }
      });

      sheet.insertColumnAfter(colIdx);
      sheet.getRange(1, colIdx + 1).setValue(headers[colIdx-1] + "_imputed").setFontWeight("bold");
      sheet.getRange(2, colIdx + 1, processed.length, 1).setValues(processed);
    }
  });

  if (rowsToDelete.size > 0) {
    const sortedRows = Array.from(rowsToDelete).sort((a, b) => b - a);
    sortedRows.forEach(rowIdx => sheet.deleteRow(rowIdx));
    return `Cleaned! Deleted ${rowsToDelete.size} rows.`;
  }
  return "Missing values imputed successfully.";
}

/**
 * Applies multiple structural fixes to selected columns.
 */
function applyStructuralCleanup(configs) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Sort descending to handle column insertion correctly
  const sortedConfigs = configs.sort((a, b) => b.index - a.index);

  sortedConfigs.forEach(config => {
    const colIdx = config.index;
    const colName = headers[colIdx - 1];
    const newHeader = `${colName}_cleaned`;
    
    sheet.insertColumnAfter(colIdx);
    sheet.getRange(1, colIdx + 1).setValue(newHeader).setFontWeight("bold");

    const rawData = sheet.getRange(2, colIdx, sheet.getLastRow() - 1, 1).getValues();
    
    const cleaned = rawData.map(row => {
      let val = row[0].toString();
      if (!val) return [""];

      // Apply transformations based on selected checkboxes
      if (config.doTrim) val = val.trim();
      if (config.doLower) val = val.toLowerCase();
      if (config.doUpper) val = val.toUpperCase();
      if (config.doAlphaNum) val = val.replace(/[^a-z0-9 ]/gi, '');
      
      return [val];
    });

    sheet.getRange(2, colIdx + 1, cleaned.length, 1).setValues(cleaned);
  });

  return "Structural cleanup finished. New columns created.";
}

/**
 * Handles Categorical Encoding Transformations
 */
/**
 * Handles Categorical Encoding Transformations while preserving nulls
 */
function applyEncodingTransformation(configs) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const sortedConfigs = configs.sort((a, b) => b.index - a.index);

  sortedConfigs.forEach(config => {
    const colIdx = config.index;
    const colName = headers[colIdx - 1];
    const rawData = sheet.getRange(2, colIdx, sheet.getLastRow() - 1, 1).getValues().map(r => r[0].toString());

    if (config.method === "ONE-HOT") {
      // Get unique values but exclude the empty string from becoming its own column
      const uniqueValues = [...new Set(rawData)].filter(v => v !== "" && v !== "null" && v !== "undefined");
      
      uniqueValues.forEach((val, i) => {
        const newColIdx = colIdx + i;
        sheet.insertColumnAfter(newColIdx);
        sheet.getRange(1, newColIdx + 1).setValue(`${colName}_${val}`).setFontWeight("bold");
        
        const dummyData = rawData.map(r => {
          if (r === "" || r === "null") return [""]; // KEEP EMPTY
          return [r === val ? 1 : 0];
        });
        sheet.getRange(2, newColIdx + 1, dummyData.length, 1).setValues(dummyData);
      });
    } 
    else if (config.method === "LABEL") {
      const map = {};
      config.order.forEach((val, i) => map[val] = i);

      sheet.insertColumnAfter(colIdx);
      sheet.getRange(1, colIdx + 1).setValue(`${colName}_encoded`).setFontWeight("bold");

      const encodedData = rawData.map(r => {
        // If the original cell is empty, return empty instead of -1
        if (r === "" || r === "null") return [""]; 
        return [map[r] !== undefined ? map[r] : ""]; 
      });
      sheet.getRange(2, colIdx + 1, encodedData.length, 1).setValues(encodedData);
    }
  });
  return "Encoding complete. New numeric features added.";
}

function getUniqueValues(colIdx) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getRange(2, colIdx, sheet.getLastRow() - 1, 1).getValues();
  const unique = [...new Set(data.map(r => r[0].toString()))].filter(v => v !== "");
  return unique;
}

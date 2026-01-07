/**
 * DATA PREP ENGINE - CORE
 */

function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Launch Data Prep', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate().setTitle('Data Prep Engine');
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
  const range = sheet.getDataRange();
  const data = range.getValues();
  const headers = data[0];
  const rows = data.slice(1);

  for (let sel of selections) {
    const colIdx = sel.index - 1;
    for (let r = 0; r < rows.length; r++) {
      const cell = rows[r][colIdx];
      if (cell === "" || cell === null) continue;
      if (sel.type === "Numeric" && isNaN(parseFloat(cell))) {
        sheet.getRange(r + 2, sel.index).activate();
        return { error: `Type Mismatch in [${headers[colIdx]}] at Row ${r + 2}: Found "${cell}" (cannot be Numeric).` };
      }
    }
    const colRange = sheet.getRange(2, sel.index, sheet.getLastRow() - 1, 1);
    if (sel.type === "Numeric") colRange.setNumberFormat("#.####################");
    else if (sel.type === "Categorical") colRange.setNumberFormat("@");
  }

  let report = {
    hasMissing: false, missingCount: 0, missingDetails: {},
    hasDates: false, dateCount: 0, dateDetails: {},
    hasOutliers: false, outlierCount: 0, outlierDetails: {},
    hasStrings: false, stringCount: 0, stringDetails: {},
    hasCleanup: false, cleanupCount: 0, cleanupDetails: {},
    totalChecked: selections.length, totalIssues: 0
  };

  selections.forEach(sel => {
    const colIdx = sel.index - 1;
    const colName = headers[colIdx];
    let colValues = [];
    rows.forEach(row => {
      const cell = row[colIdx];
      if (cell === "" || cell === null) {
        report.hasMissing = true; report.missingCount++;
        report.missingDetails[colName] = (report.missingDetails[colName] || 0) + 1;
        return; 
      }
      if (typeof cell === 'string') {
        report.hasStrings = true; report.stringCount++;
        report.stringDetails[colName] = (report.stringDetails[colName] || 0) + 1;
        if (cell !== cell.trim() || cell !== cell.toLowerCase()) {
          report.hasCleanup = true; report.cleanupCount++;
          report.cleanupDetails[colName] = (report.cleanupDetails[colName] || 0) + 1;
        }
      }
      if (cell instanceof Date || sel.type === "Date") {
        report.hasDates = true; report.dateCount++;
        report.dateDetails[colName] = (report.dateDetails[colName] || 0) + 1;
      }
      if (sel.type === "Numeric" && !isNaN(parseFloat(cell))) colValues.push(parseFloat(cell));
    });

    if (sel.type === "Numeric" && colValues.length > 2) {
      const n = colValues.length;
      const mean = colValues.reduce((a, b) => a + b) / n;
      const stdDev = Math.sqrt(colValues.reduce((s, x) => s + Math.pow(x - mean, 2), 0) / n);
      if (stdDev > 0) {
        const outliers = colValues.filter(x => Math.abs(x - mean) > (3 * stdDev));
        if (outliers.length > 0) {
          report.hasOutliers = true; report.outlierCount += outliers.length;
          report.outlierDetails[colName] = outliers.length;
        }
      }
    }
  });
  report.totalIssues = report.missingCount + report.outlierCount + report.stringCount + report.cleanupCount;
  return report;
}

function applyDateTransformation(dateConfigs) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timezone = Session.getScriptTimeZone();
  const sortedConfigs = dateConfigs.sort((a, b) => b.index - a.index);

  sortedConfigs.forEach(config => {
    const colIdx = config.index;
    sheet.insertColumnAfter(colIdx);
    sheet.getRange(1, colIdx + 1).setValue(headers[colIdx - 1] + "_cleaned").setFontWeight("bold");

    const raw = sheet.getRange(2, colIdx, sheet.getLastRow() - 1, 1).getValues();
    const processed = raw.map(row => {
      let val = row[0];
      if (!val) return [""];
      let d;
      if (typeof val === 'string') {
        let p = val.split(/[\/\-\. ]/);
        if (p.length >= 2) {
          let day = (config.locale === "US") ? (p[1] || 1) : p[0];
          let month = (config.locale === "US") ? p[0] : (p[1] || 1);
          let year = p[2] || new Date().getFullYear();
          if (year.toString().length === 2) year = "20" + year;
          d = new Date(year, month - 1, day);
        }
      } else if (val instanceof Date) { d = val; }

      if (d && !isNaN(d.getTime())) {
        return (config.format === "UNIX") ? [Math.floor(d.getTime()/1000)] : [Utilities.formatDate(d, timezone, "yyyy-MM-dd HH:mm:ss")];
      }
      return [val];
    });
    sheet.getRange(2, colIdx + 1, processed.length, 1).setValues(processed);
  });
  return "Date transformation successful.";
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
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowsToDelete = new Set();
  
  // Sort descending by index to handle column additions safely
  const sortedConfigs = configs.sort((a, b) => b.index - a.index);

  sortedConfigs.forEach(config => {
    const colIdx = config.index;
    const raw = sheet.getRange(2, colIdx, sheet.getLastRow() - 1, 1).getValues();

    if (config.method === "DROP") {
      // COLLECT ROWS FOR DELETION
      raw.forEach((row, i) => {
        if (row[0] === "" || row[0] === null) {
          rowsToDelete.add(i + 2); // +2 for header offset
        }
      });
    } else {
      // CALCULATE STATS FOR IMPUTATION
      const nums = raw.map(r => parseFloat(r[0])).filter(v => !isNaN(v));
      const mean = nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : 0;
      const median = nums.length ? nums.sort((a,b)=>a-b)[Math.floor(nums.length/2)] : 0;
      
      const modeMap = {};
      raw.forEach(r => { if(r[0]) modeMap[r[0]] = (modeMap[r[0]] || 0) + 1; });
      const mode = Object.keys(modeMap).reduce((a, b) => modeMap[a] > modeMap[b] ? a : b, "");

      // Prepare New Column
      sheet.insertColumnAfter(colIdx);
      sheet.getRange(1, colIdx + 1).setValue(headers[colIdx-1] + "_imputed").setFontWeight("bold");

      const processed = raw.map((row, i) => {
        let val = row[0];
        if (val !== "" && val !== null) return [val];
        
        switch (config.method) {
          case "MEAN": return [mean];
          case "MEDIAN": return [median];
          case "ZERO": return [0];
          case "MODE": return [mode];
          case "LABEL": return ["Unknown"];
          case "CUSTOM": return [config.customVal];
          case "FORWARD": return i > 0 ? [raw[i-1][0]] : [""];
          default: return [""];
        }
      });
      sheet.getRange(2, colIdx + 1, processed.length, 1).setValues(processed);
    }
  });

  // EXECUTE ROW DELETIONS
  if (rowsToDelete.size > 0) {
    const sortedRows = Array.from(rowsToDelete).sort((a, b) => b - a);
    sortedRows.forEach(rowIdx => sheet.deleteRow(rowIdx));
    return `Cleaned! Deleted ${rowsToDelete.size} rows with missing values and imputed others.`;
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

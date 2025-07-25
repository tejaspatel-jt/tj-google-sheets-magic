/**
 * This script provides various tools for managing and merging leads data in Google Sheets.
 * It includes functionalities for:
 *    1. Freezing the first row in all sheets
 *    2. Merging sheets from external spreadsheets with header union (both dynamic and static)
 *    3. Deduplicating data based on a specified key
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ Tejas External Data Tools 🚀')

    // CLEAN-UP PART
    .addItem('1️⃣ Freeze First Row in All Sheets ✅', 'freezeFirstRowInAllSheets')

    // MERGING PART
    .addItem('🔗 ➕ Merge Sheets from External Spreadsheet Tabs - Dynamic 📂 ✅', 'merge_External_SpreadSheet_Tabs_Header_Union_Dynamically')
    .addItem('🔗 ➕ Merge Sheets from External Spreadsheet Tabs - Static 📂 ❌', 'merge_External_SpreadSheet_Tabs_Header_Union_Static')
    

    .addToUi();
    
}

// 1. Freeze first row in all sheets
function freezeFirstRowInAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    sheets[i].setFrozenRows(1);
  }
}


/**
 * 🔗 Merge External Spreadsheets using 'MergeSheet_Config' Sheet for Source Info with Header Union
 * 
 * 🧠 Reads configuration sheet listing spreadsheets and tabs to merge,
 * using columns: 'SpreadSheet URL', 'SpreadSheet Name' (optional), and 'Sheet Name' (tab).
 * 
 * - Merges all data by union of headers (case-insensitive optional).
 * - Optionally deduplicates.
 * - Supports overwrite or append mode controlled by config.
 * - Don't append duplicate records based on deduplication key (e.g. Email) while merging.
 * 
 * 🔁 Safe for repeated runs.
 */
function merge_External_SpreadSheet_Tabs_Header_Union_Dynamically() {
  const config = {
    configSheetName: 'MergeSheet_Config', // Name of the config tab holding source sheet info
    outputSheet: 'Leads_MasterData',
    deduplication: {
      allowDuplicates: false, // Keep it false always to deduplicate, Option is here for misuse only 🙂
      key: 'Email'
    },
    options: {
      // Treat 'Region' and 'region' as same? Set to true if you want case-sensitive headers : KEEP IT FALSE only for similar Columns Overheads
      caseSensitiveHeaders: false   // Keeping this `False` will Treat 'Region' and 'region' as same. Set to true if you want case-sensitive headers
    },
    // KEEP IT FALSE only to Process New Rows Only, Ideal in Case if You have added Other Columns for Tracking purpose
    overwriteOutputSheet: false // ALWAYS keep false as per your ask, never overwrite, only append unique rows
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- Step 0: Read source file configs from config sheet ---
  let configSheet = ss.getSheetByName(config.configSheetName);
  if (!configSheet) {
    ui.alert(
      "⚠️ Missing Config Sheet ⚠️",

      `❌ Config sheet '${config.configSheetName}' not found!

        Let Me create it for you.
        
        Just click on 'OK' and I will create it for you.`,
      ui.ButtonSet.OK
    );

    // Create the config sheet with required headers
    configSheet = ss.insertSheet(config.configSheetName);
    const headers = ["SpreadSheet Name", "Sheet Name", "SpreadSheet URL"];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    ui.alert(
      "✅ Configuration Sheet Created ✅",
      `A new sheet named '${config.configSheetName}' was created with the required columns
            - 'SpreadSheet Name' (optional, for info only)
            - 'Sheet Name' (mandatory)
            - 'SpreadSheet URL' (mandatory)
    
    Please fill it with your source configuration and rerun the script.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const configData = configSheet.getDataRange().getValues();
  if (configData.length < 2) {
    ui.alert(
        "⚠️ Missing Required Configuration ⚠️",
        `❌ Config sheet '${config.configSheetName}' is empty\n
        It must have below columns:
            - 'SpreadSheet Name' (optional, for info only)`
            - 'Sheet Name' (mandatory)
            - 'SpreadSheet URL' (mandatory),
          ui.ButtonSet.OK
    );
    return;
  }

  const headerRow = configData[0];

  // Find indexes of key columns by exact header name match (case-insensitive, trimmed)
  function findHeaderIndex(header) {
    return headerRow.findIndex(h => h && h.toString().trim().toLowerCase() === header.toLowerCase());
  }

  const urlColIndex = findHeaderIndex('SpreadSheet URL');
  const spreadSheetNameColIndex = findHeaderIndex('SpreadSheet Name'); // Optional
  const sheetNameColIndex = findHeaderIndex('Sheet Name');

  if (urlColIndex === -1) {
    ui.alert(`❌ Config sheet '${config.configSheetName}' must have a column header named exactly 'SpreadSheet URL'.`);
    return;
  }
  if (sheetNameColIndex === -1) {
    ui.alert(`❌ Config sheet '${config.configSheetName}' must have a column header named exactly 'Sheet Name'.`);
    return;
  }

  // Parse sourceFiles from config sheet rows
  const sourceFiles = [];
  for (let i = 1; i < configData.length; i++) {
    const row = configData[i];
    const url = (row[urlColIndex] || '').toString().trim();
    if (!url) continue;

    const spreadSheetName = (spreadSheetNameColIndex >= 0) ? (row[spreadSheetNameColIndex] || '').toString().trim() : '';
    const sheetName = (row[sheetNameColIndex] || '').toString().trim();
    if (!sheetName) {
      ui.alert(`❌ Row ${i + 1} in config sheet missing required 'Sheet Name' value.`);
      return;
    }

    // Extract fileId from URL
    const matchId = url.match(/\/d\/([a-zA-Z0-9-_]+)(\/|$)/);
    if (!matchId) {
      ui.alert(`❌ Invalid spreadsheet URL in config sheet row ${i + 1}: ${url}`);
      return;
    }

    sourceFiles.push({
      fileId: matchId[1],
      spreadSheetName: spreadSheetName, // Optional, not used in current logic — just for info or future logging
      sheetName: sheetName
    });
  }

  if (sourceFiles.length === 0) {
    ui.alert(`❌ No valid source files found in config sheet '${config.configSheetName}'.`);
    return;
  }

  // Inform user (modal dialog) — optional UX enhancement
  // ui.showModalDialog(HtmlService.createHtmlOutput('<p>Starting merge operation... please wait 👨‍💻</p>'), 'Merging external sources');
  ss.toast("🔄 Merging & processing external sources...", "Processing", -1);

  // --- Step 1: Load headers + data, build global union of headers ---
  const allHeadersMap = new Map();
  const sheetData = [];

  sourceFiles.forEach(({ fileId, sheetName, spreadSheetName }) => {
    let sourceSS, sheet;
    try {
      sourceSS = SpreadsheetApp.openById(fileId);
      sheet = sourceSS.getSheetByName(sheetName);
      if (!sheet) throw new Error(`Sheet '${sheetName}' not found in spreadsheet with id '${fileId}'`);
    } catch(e) {
      throw new Error(`Failed to open spreadsheet or sheet: ${e.message}`);
    }

    const values = sheet.getDataRange().getValues();
    if (values.length < 1) return;

    const headers = values[0];
    const rows = values.slice(1);

    const normalizedHeaders = headers.map(h =>
      config.options.caseSensitiveHeaders ? h.trim() : h.toString().trim().toLowerCase()
    );

    normalizedHeaders.forEach((norm, i) => {
      if (!allHeadersMap.has(norm)) {
        allHeadersMap.set(norm, headers[i].toString().trim());
      }
    });

    sheetData.push({ normalizedHeaders, originalHeaders: headers, rows });
  });

  const finalHeaders = Array.from(allHeadersMap.values());
  const finalData = [];

  // --- Step 2: Align all rows against header union ---
  sheetData.forEach(({ normalizedHeaders, rows }) => {
    const headerIndex = normalizedHeaders.reduce((map, h, i) => {
      map[h] = i;
      return map;
    }, {});

    rows.forEach(row => {
      const newRow = finalHeaders.map(h => {
        const norm = config.options.caseSensitiveHeaders ? h.trim() : h.toLowerCase();
        const index = headerIndex[norm];
        return index !== undefined ? row[index] : '';
      });
      finalData.push(newRow);
    });
  });

  // --- Step 3: Deduplicate if needed (among imported data only, not vs master sheet yet) ---
  let cleanedData = finalData;
  let removedCount = 0;

  if (!config.deduplication.allowDuplicates) {
    const dedupKeyNorm = config.options.caseSensitiveHeaders ? config.deduplication.key : config.deduplication.key.toLowerCase();
    const keyIndex = finalHeaders.findIndex(h =>
      (config.options.caseSensitiveHeaders ? h : h.toLowerCase()) === dedupKeyNorm
    );
    if (keyIndex === -1) throw new Error(`Deduplication key "${config.deduplication.key}" not found in final headers.`);

    const recordMap = new Map();

    for (const row of finalData) {
      const rawKey = String(row[keyIndex] || '').trim();
      const normKey = config.options.caseSensitiveHeaders ? rawKey : rawKey.toLowerCase();
      if (!normKey) continue;

      const nonEmptyCount = row.filter(cell => cell !== '' && cell !== null).length;
      const existing = recordMap.get(normKey);

      if (!existing || nonEmptyCount > existing.nonEmptyCount) {
        recordMap.set(normKey, { row, nonEmptyCount });
      } else {
        removedCount++;
      }
    }
    cleanedData = Array.from(recordMap.values()).map(v => v.row);
  }

  // --- Step 4: Write output 
  let targetSheet = ss.getSheetByName(config.outputSheet);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(config.outputSheet);
    targetSheet.clear();
  } else {
    // NEVER clear, only append
  }

  // Write headers only if sheet is new/empty
  if (targetSheet.getLastRow() === 0) {
    targetSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  }

  // Declare toAppend in higher scope to access later in toast as well
  let toAppend = [];

  if (cleanedData.length > 0) {
    // -- APPEND MODE LOGIC ONLY --
    // Only append those rows where deduplication key (e.g. Email) is not present in Lead_MasterData
    const existingDataRange = targetSheet.getDataRange();
    const existingData = existingDataRange.getValues();
    // Find index of deduplication key/email in the master sheet header
    let existingKeyIndex = -1;
    if (existingData.length > 0 && existingData[0]) {
      existingKeyIndex = existingData[0].findIndex(h =>
        (config.options.caseSensitiveHeaders ? h : String(h).toLowerCase()) === (config.options.caseSensitiveHeaders ? config.deduplication.key : config.deduplication.key.toLowerCase())
      );
    }
    // If header not found, fallback to 0 (risk: broken append)
    if (existingKeyIndex === -1) existingKeyIndex = 0;

    // Build set of emails already in Leads_MasterData
    const existingSet = new Set();
    for (let i = 1; i < existingData.length; i++) {
      const val = existingData[i][existingKeyIndex];
      if (val) existingSet.add(String(val).trim().toLowerCase());
    }

    // Get index of deduplication ("Email") key in merged header
    const keyIndex = finalHeaders.findIndex(h =>
      (config.options.caseSensitiveHeaders ? h : h.toLowerCase()) === (config.options.caseSensitiveHeaders ? config.deduplication.key : config.deduplication.key.toLowerCase())
    );

    // Only append rows with emails not already present
    toAppend = cleanedData.filter(row => {
      const val = row[keyIndex];
      if (!val) return true; // allow blank email (unlikely)
      return !existingSet.has(String(val).trim().toLowerCase());
    });

    if (toAppend.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, toAppend.length, finalHeaders.length).setValues(toAppend);
    }
  }

  targetSheet.setFrozenRows(1);

  // Toast & alert summary
  ss.toast("✅ Merge completed", "Done ✅", 5);
  ui.alert(
    "✅ Merge Completed",
        `📄 Total New Rows Imported: ${toAppend.length}

        🧹 Duplicates Removed: ${removedCount}
        ✅ Final Rows Written (NEW): ${cleanedData.length}
        📌 Total Columns Merged: ${finalHeaders.length}`,
    ui.ButtonSet.OK
  );
}






/**
 * 📄 Merge External Spreadsheets: Union Headers + Optional Deduplication, Case Control 🧾🌐
 *
 * 🧠 Use-case:
 * Merge remote sheets from different spreadsheets — auto aligns headers (case-insensitively if set),
 * (optionally) deduplicates based on most complete row per key.
 * 
 * ✅ Keeps most complete row when deduplicating based on non-empty field count.
 *
 * ⚙️ Configuration
 * - sourceFiles: [{ fileId: string, sheetName: string }]
 * - outputSheet: string (target sheet where merged data is written)
 * - deduplication: {
 *     allowDuplicates: boolean,
 *     key: string (required if allowDuplicates is false)
 *   }
 * - options: {
 *     caseSensitiveHeaders: boolean
 *   }
 */
function merge_External_SpreadSheet_Tabs_Header_Union_Static() {
  const config = {
    sourceFiles: [
      {
        fileId: '1UksnP4CUXIMm8NkxiLxi7Yc2tqmXORMbLNcJS71K-xw',
        sheetName: 'Lead_CleanedData'
      },
      {
        fileId: '1zCpceVZRq_X-L0K7czdjHd3p6ME4kkTl2MzD4ev8Icw',
        sheetName: 'Lead_CleanedData'
      }
    ],
    outputSheet: 'Leads_MasterData_Static',
    deduplication: {
      allowDuplicates: false, // ❌ Change to false to deduplicate
      key: 'Email'           // Used only when deduplication is enabled
    },
    options: {
      caseSensitiveHeaders: false // 🔠 Treat 'Region' and 'region' as same?
    }
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const allHeadersMap = new Map(); // Key: header (adjusted case) → original header text
  const sheetData = [];

  ss.toast("🔄 Merging & processing external sources...", "Processing", -1);

  // 🛠 Step 1: Read data and build global header set
  config.sourceFiles.forEach(({ fileId, sheetName }) => {
    const sourceSS = SpreadsheetApp.openById(fileId);
    const sheet = sourceSS.getSheetByName(sheetName);
    if (!sheet) throw new Error(`❌ Sheet "${sheetName}" not found in file: ${fileId}`);

    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const rows = values.slice(1);

    const normalizedHeaders = headers.map(h =>
      config.options.caseSensitiveHeaders ? h.trim() : h.trim().toLowerCase()
    );

    // Map normalized header to a canonical version (use first one seen)
    normalizedHeaders.forEach((norm, i) => {
      if (!allHeadersMap.has(norm)) {
        allHeadersMap.set(norm, headers[i].trim());
      }
    });

    sheetData.push({ normalizedHeaders, originalHeaders: headers, rows });
  });

  const finalHeaders = Array.from(allHeadersMap.values()); // Our merged header structure
  const finalData = [];

  // 🏗 Step 2: Re-align each source row to full column structure
  sheetData.forEach(({ normalizedHeaders, originalHeaders, rows }) => {
    const headerIndex = normalizedHeaders.reduce((map, h, i) => {
      map[h] = i;
      return map;
    }, {});

    rows.forEach(row => {
      const alignedRow = finalHeaders.map(h => {
        const normalizedH = config.options.caseSensitiveHeaders ? h.trim() : h.trim().toLowerCase();
        const i = headerIndex[normalizedH];
        return i !== undefined ? row[i] : '';
      });
      finalData.push(alignedRow);
    });
  });

  // 🧹 Step 3: Optional deduplication
  let cleanedData = finalData;
  let removedCount = 0;
  const { allowDuplicates, key } = config.deduplication;

  if (!allowDuplicates) {
    const dedupKeyNorm = config.options.caseSensitiveHeaders ? key : key.toLowerCase();
    const keyIndex = finalHeaders.findIndex(h => {
      return config.options.caseSensitiveHeaders
        ? h === key
        : h.toLowerCase() === dedupKeyNorm;
    });

    if (keyIndex === -1) throw new Error(`❌ Deduplication key "${key}" not found in headers`);

    const recordMap = new Map();

    for (const row of finalData) {
      const rawKey = String(row[keyIndex]).trim();
      const normKey = config.options.caseSensitiveHeaders ? rawKey : rawKey.toLowerCase();
      if (!normKey) continue;

      const nonEmptyCount = row.filter(cell => cell !== '' && cell !== null).length;
      const existing = recordMap.get(normKey);

      if (!existing || nonEmptyCount > existing.nonEmptyCount) {
        recordMap.set(normKey, { row, nonEmptyCount });
      } else {
        removedCount++;
      }
    }

    cleanedData = Array.from(recordMap.values()).map(x => x.row);
  }

  // 📝 Step 4: Output into structured & cleared target sheet
  let targetSheet = ss.getSheetByName(config.outputSheet);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(config.outputSheet);
  } else {
    // output.clearContents(); // 🔄 Controlled clear Only cell values and formulas (preserves formatting & structure)
    targetSheet.clear(); // 🔄 Everything – values + formatting + notes + merges
  }

  targetSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  if (cleanedData.length > 0) {
    targetSheet.getRange(2, 1, cleanedData.length, finalHeaders.length).setValues(cleanedData);
  }

  targetSheet.setFrozenRows(1); // Freeze header row for better UX

  // ✅ Done: Show summary
  ss.toast("✅ Merge completed", "Done ✅", 5);
  ui.alert(
    "✅ Merge Completed",
    `📄 Total Rows Imported: ${finalData.length}\n🧹 Duplicates Removed: ${removedCount}\n✅ Final Rows Written: ${cleanedData.length}\n📌 Total Columns Merged: ${finalHeaders.length}`,
    ui.ButtonSet.OK
  );
  ui.alert(
    "✅ Merge Complete",
    `📄 Total Rows Imported: ${finalData.length}
    🧹 Duplicates Removed: ${removedCount}
    ✅ Final Rows Written: ${cleanedData.length}
    📌 Total Columns Merged: ${finalHeaders.length}`,
    ui.ButtonSet.OK
  );
}
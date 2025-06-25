import XLSX from 'xlsx';
import { saveAs } from 'file-saver';

// Declare this so our linter knows that tableau is a global object
/* global tableau */


//array_move function from stackoverflow solution:
//https://stackoverflow.com/questions/5306680/move-an-array-element-from-one-array-position-to-another
function array_move(arr, old_index, new_index) {
  if (new_index >= arr.length) {
      var k = new_index - arr.length + 1;
      while (k--) {
          arr.push(undefined);
      }
  }
  arr.splice(new_index, 0, arr.splice(old_index, 1)[0]);
  return arr; // for testing
};

const saveSettings = () => new Promise((resolve, reject) => {
  console.log('[func.js] Saving settings');
  tableau.extensions.settings.set('metaVersion', 2);
  console.log('[func.js] Authoring mode', tableau.extensions.environment.mode);
  if (tableau.extensions.environment.mode === "authoring") {
    tableau.extensions.settings.saveAsync()
    .then(newSavedSettings => {
      //console.log('[func.js] newSavedSettings', newSavedSettings);
      resolve(newSavedSettings);
    }).catch(reject);
  } else {
    resolve();
  }
  
});

const setSettings = (type, value) => new Promise((resolve, reject) => {
  console.log('[func.js] Set settings', type, value);
  let settingKey = '';
  switch(type) {
    case 'sheets':
      settingKey = 'selectedSheets';
      break;
    case 'label':
      settingKey = 'buttonLabel';
      break;
    case 'style':
      settingKey = 'buttonStyle';
      break;
    case 'filename':
      settingKey = 'filename';
      break;
    case 'version':
      settingKey = 'metaVersion';
      break;
    default:
      settingKey = 'unknown';
  }
  tableau.extensions.settings.set(settingKey, JSON.stringify(value));
  resolve();
});

const getSheetColumns = (sheet, existingCols, modified) => new Promise((resolve, reject) => {
  sheet.getSummaryDataAsync({ignoreSelection: true}).then((data) => {
    //console.log('[func.js] Sheet Summary Data', data);
    console.log('[func.js] getSheetColumns existingCols', JSON.stringify(existingCols));
    const columns = data.columns;
    let cols = [];
    const existingIdx = [];
    if (modified) {
      for (var j = 0; j < columns.length; j++) {
        //console.log(columns[j]);
        var col = {};
        col.index = columns[j].index;
        col.name = columns[j].fieldName;
        col.dataType = columns[j].dataType;
        col.changeName = null;
        col.selected = false;
        cols.push(col);
      }
      for (var i = 0; i < existingCols.length; i++) {
        if (existingCols[i] && existingCols[i].hasOwnProperty("name")) {
          existingIdx.push(existingCols[i].name);
        }
      }
      console.log('[func.js] getSheetColumns existingIdx', existingIdx);
      let maxPos = existingIdx.length;
      cols = cols.map((col, idx) => {
        //console.log('[func.js] getSheetColumns Looking for col', col);
        const eIdx = existingIdx.indexOf(col.name);
        const ret = {...col};
        if (eIdx > -1) {
          ret.selected = existingCols[eIdx].selected;
          ret.changeName = existingCols[eIdx].changeName;
          ret.order = eIdx;
        } else {
          ret.order = maxPos;
          maxPos += 1;
        }
        return ret;
      });
    } else {
      for (var k = 0; k < columns.length; k++) {
        var newCol = {};
        newCol.index = columns[k].index;
        newCol.name = columns[k].fieldName;
        newCol.dataType = columns[k].dataType;
        newCol.selected = true;
        newCol.order = k + 1;
        cols.push(newCol);
      }
    }
    cols = cols.sort((a, b) => (a.order > b.order) ? 1 : -1)
    resolve(cols);
  })
  .catch(error => {
    console.log('[func.js] Error with getSummaryDataAsync', sheet, error);
  });
});

const initializeMeta = () => new Promise((resolve, reject) => {
  console.log('[func.js] Initialise Meta');
  var promises = [];
  const worksheets = tableau.extensions.dashboardContent._dashboard.worksheets;
  //console.log('[func.js] Worksheets in dashboard', worksheets);
  var meta = worksheets.map(worksheet => {
    var sheet = worksheet;
    var item = {};
    item.sheetName = sheet.name;
    item.selected = false;
    item.changeName = null;
    item.customCols = false;
    promises.push(getSheetColumns(sheet, null, false));
    return item;
  });

  console.log(`[func.js] Found ${meta.length} sheets`, meta);

  Promise.all(promises).then((sheetArr) => {
    for (var i = 0; i < sheetArr.length; i++) {
      var sheetMeta = meta[i];
      sheetMeta.columns = sheetArr[i];
      meta[i] = sheetMeta;
      console.log(`[func.js] Added ${sheetArr[i].length} columns to ${sheetMeta.sheetName}`, meta);
    }
    //console.log(`[func.js] Meta initialised`, meta);
    resolve(meta);
  });
});

const revalidateMeta = (existing) => new Promise((resolve, reject) => {
  console.log('[func.js] Revalidate Meta');
  var promises = [];
  const worksheets = tableau.extensions.dashboardContent._dashboard.worksheets;
  //console.log('[func.js] Worksheets in dashboard', worksheets);
  var meta = worksheets.map(worksheet => {
    var sheet = worksheet;
    const sheetIdx = existing.findIndex((e) => {
      return e.sheetName === sheet.name;
    });
    if (sheetIdx > -1) {
      console.log(`[func.js] Existing sheet ${sheet.name} columns`, JSON.stringify(existing[sheetIdx].columns));
      promises.push(getSheetColumns(sheet, existing[sheetIdx].columns, true));
      existing[sheetIdx].existed = true;
      return existing[sheetIdx];
    } else {
      var item = {};
      item.sheetName = sheet.name;
      item.selected = false;
      item.changeName = null;
      item.customCols = false;
      item.existed = false;
      promises.push(getSheetColumns(sheet, null, false));
      return item;
    }
  });

  console.log(`[func.js] Found ${meta.length} sheets`, meta);

  Promise.all(promises).then((sheetArr) => {

    for (var i = 0; i < sheetArr.length; i++) {
      var sheetMeta = meta[i];
      sheetMeta.columns = sheetArr[i];
      meta[i] = sheetMeta;
      //console.log(`[func.js] Added ${sheetArr[i].length} columns to ${sheetMeta.sheetName}`, meta);
    }
    meta.forEach((sheet, idx) => {
      if (sheet && sheet.sheetName) {
        const eIdx = existing.findIndex((e) => {
          return e.sheetName === sheet.sheetName;
        });
        meta = array_move(meta, idx, eIdx);
      } else {
        console.log('[func.js] Sheet ordering issue. No sheet defined in idx', idx);
      }
    })
    console.log(`[func.js] Meta revalidated`, JSON.stringify(meta));
    resolve(meta);
  });
});
/*

const exportToExcel = (meta, env, filename) => new Promise((resolve, reject) => {
  let csvFile = "export.csv";
  if (filename && filename.length > 0) {
    csvFile = filename + ".csv";
  }
  buildExcelBlob(meta).then(csvContent => {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    saveAs(blob, csvFile);
    resolve();
  });
});

// Function to build CSV content
const buildExcelBlob = (meta) => new Promise((resolve, reject) => {
  console.log("[func.js] Got Meta", meta);
  const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
  let totalSheets = 0;
  let sheetCount = 0;
  const sheetList = [];
  const columnList = [];
  const tabNames = [];
  const csvRows = [];

  for (let i = 0; i < meta.length; i++) {
    if (meta[i] && meta[i].selected) {
      let tabName = meta[i].changeName || meta[i].sheetName;
      tabName = tabName.replace(/[*?/\\[\]]/gi, '');
      sheetList.push(meta[i].sheetName);
      tabNames.push(tabName);
      columnList.push(meta[i].columns);
      totalSheets = totalSheets + 1;
    }
  }

  sheetList.map((metaSheet, idx) => {
    const sheet = worksheets.find(s => s.name === metaSheet);
    sheet.getSummaryDataAsync({ ignoreSelection: true }).then((data) => {
      const columns = data.columns;
      const columnMeta = columnList[sheetCount];
      const headerOrder = [];

      columnMeta.map((colMeta) => {
        if (colMeta && colMeta.selected) {
          headerOrder.push(colMeta.changeName || colMeta.name);
        }
        return colMeta;
      });

      columns.map((column, idx) => {
        const objCol = columnMeta.find(o => o.name === column.fieldName);
        if (objCol) {
          let col = { ...column, selected: objCol.selected };
          col.outputName = objCol.changeName || objCol.name;
          columns[idx] = col;
          return col;
        } else {
          return null;
        }
      });

      decodeDataset(columns, data.data).then((rows) => {
        // Prepare CSV rows
        const csvHeader = headerOrder.join(",") + "\n";
        const csvData = rows.map(row => {
          return headerOrder.map(header => {
            return JSON.stringify(row[header] || ""); // Handle potential undefined values
          }).join(",");
        }).join("\n");

        csvRows.push(csvHeader + csvData);
        sheetCount = sheetCount + 1;

        if (sheetCount === totalSheets) {
          resolve(csvRows.join("\n\n")); // Separate sheets with a blank line
        }
      });
    });
    return sheet;
  });
});

*/

const exportToExcel = (meta, env, filename) => new Promise((resolve, reject) => {
  let csvFile = "export.csv";
  if (filename && filename.length > 0) {
    csvFile = filename + ".csv";
  }

  buildExcelBlob(meta).then(csvText => {
    const blob = new Blob([csvText], { type: "text/csv;charset=utf-8;" });
    saveAs(blob, csvFile);
    resolve();
  }).catch(reject);
});

const buildExcelBlob = (meta) => new Promise((resolve, reject) => {
  const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
  let csvAll = "";
  let totalSheets = 0;
  let processedSheets = 0;

  for (let i = 0; i < meta.length; i++) {
    if (meta[i] && meta[i].selected) {
      totalSheets++;
      const tabName = (meta[i].changeName || meta[i].sheetName).replace(/[*?/\\[\]]/gi, '');
      const sheet = worksheets.find(s => s.name === meta[i].sheetName);
      const columnMeta = meta[i].columns;

      sheet.getSummaryDataAsync({ ignoreSelection: true }).then(data => {
        const columns = data.columns;
        const rows = data.data;

        // Prepare header row
        let headers = [];
        columnMeta.forEach(col => {
          if (col.selected) {
            headers.push(`"${col.changeName || col.name}"`);
          }
        });

        // Prepare data rows
        const csvRows = rows.map(row => {
          return columnMeta.filter(col => col.selected).map(colMeta => {
            const column = columns.find(c => c.fieldName === colMeta.name);
            const cell = row[columns.indexOf(column)];
            const val = (cell.formattedValue || "").replace(/"/g, '""'); // escape quotes
            return `"${val}"`;
          }).join(",");
        });

        // Combine
        csvAll += `\n\nSheet: ${tabName}\n`;
        csvAll += headers.join(",") + "\n";
        csvAll += csvRows.join("\n");

        processedSheets++;
        if (processedSheets === totalSheets) {
          resolve(csvAll);
        }
      }).catch(reject);
    }
  }

  if (totalSheets === 0) {
    resolve("No sheets selected.");
  }
});


// krisd: Remove recursion to work with larger data sets
// and translate cell data types
const decodeDataset = (columns, dataset) => new Promise((resolve, reject) => {
  let promises = [];
  //for (let i=0; i<dataset.length; i++) {
    promises.push(decodeRow(columns, dataset));
 // }
  Promise.all(promises).then((datasetArr) => {
    //console.log('[func.js] datasetArr', datasetArr);
    resolve(datasetArr);
  });

});

const decodeRow = (columns, row) => new Promise((resolve, reject) => {
  let meta = {};
  for (let j = 0; j < columns.length; j++) {
    if (columns[j].selected) {
      // krisd: let's assign the sheetjs type according to the summary data column type
      let dtype = undefined;
      let dval = undefined;
      // console.log('[func.js] Row', row[j]);
      if (row[j].value === '%null%' && row[j].nativeValue === null && row[j].formattedValue === 'Null') {
        dtype = 'z';
        dval = null;
      } else {
        switch (columns[j]._dataType) {
          case 'int':
          case 'float':
            dtype = 'n';
            dval = Number(row[j].value);  // let nums be raw w/o formatting
            if (isNaN(dval)) dval = row[j].formattedValue;  // protect in case issue
            break;
          case 'date':
          case 'date-time':
            dtype = 's';
            dval = row[j].formattedValue;
            break;
          case 'bool':
            dtype = 'b';
            dval = row[j].value;
            break;
          default:
            dtype = 's';
            dval = row[j].formattedValue;
        }
      }
      let o = {v:dval, t:dtype};
      meta[columns[j].outputName] = o;
    }
  }
  resolve(meta);
});



export {
  initializeMeta,
  revalidateMeta,
  saveSettings,
  setSettings,
  exportToExcel,
}

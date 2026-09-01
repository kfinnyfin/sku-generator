/**
 * Too Much Merch SKU Engine - Google Sheets backend
 *
 * Paste this file into Extensions > Apps Script for the TMM SKU spreadsheet,
 * then deploy (or redeploy) as a Web App.
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return jsonResponse({ status: "error", message: "The catalog is busy. Please try again." });
  }

  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    if (data.action === "fetchAll") return fetchAllCatalogs();
    if (data.action === "saveBrand") return saveBrandCatalog(data);
    if (data.action === "setBrandStatus") return setBrandStatus(data);
    if (data.action === "deleteBrand") return deleteBrandCatalog(data);

    return jsonResponse({ status: "error", message: "Unknown or missing action." });
  } catch (error) {
    return jsonResponse({ status: "error", message: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function safeSheetName(name) {
  var cleaned = String(name || "").replace(/[\[\]\:\*\?\/\\]/g, "-").trim().slice(0, 100);
  if (!cleaned) throw new Error("Brand name cannot be empty.");
  return cleaned;
}

function getIndexSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("_Catalog_Index");
  if (!sheet) {
    sheet = ss.insertSheet("_Catalog_Index");
    sheet.appendRow(["Brand", "Sheet Name", "Brand Code", "Assortment", "Status", "Updated At", "Configuration JSON"]);
    sheet.getRange(1, 1, 1, 7).setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function findIndexRow(sheet, brandName) {
  if (sheet.getLastRow() < 2) return 0;
  var names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var index = 0; index < names.length; index++) {
    if (String(names[index][0]) === String(brandName)) return index + 2;
  }
  return 0;
}

function deleteIndexRecord(brandName) {
  var indexSheet = getIndexSheet();
  var row = findIndexRow(indexSheet, brandName);
  if (row) indexSheet.deleteRow(row);
}

function upsertIndexRecord(catalog, sheetName, status) {
  var indexSheet = getIndexSheet();
  var row = findIndexRow(indexSheet, catalog.name);
  var values = [[
    catalog.name,
    sheetName,
    catalog.code || "",
    catalog.assortment || "Custom",
    status,
    new Date(),
    JSON.stringify(catalog)
  ]];
  if (row) {
    indexSheet.getRange(row, 1, 1, 7).setValues(values);
  } else {
    indexSheet.getRange(indexSheet.getLastRow() + 1, 1, 1, 7).setValues(values);
  }
}

function readCatalogIndex() {
  var indexSheet = getIndexSheet();
  if (indexSheet.getLastRow() < 2) return [];
  return indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 7).getValues().map(function(row) {
    var catalog = {};
    try { catalog = row[6] ? JSON.parse(row[6]) : {}; } catch (error) { catalog = {}; }
    catalog.name = catalog.name || row[0];
    catalog.code = catalog.code || row[2];
    catalog.assortment = catalog.assortment || row[3] || "Custom";
    catalog.active = String(row[4] || "active").toLowerCase() !== "inactive";
    catalog.sheetName = row[1];
    return catalog;
  });
}

function saveBrandCatalog(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var brandName = String(data.brandName || "").trim();
  var previousBrandName = String(data.previousBrandName || "").trim();
  var sheetName = safeSheetName(brandName);
  var status = String(data.catalogStatus || "active").toLowerCase() === "inactive" ? "inactive" : "active";
  var catalog = data.catalog || {};
  var skus = Array.isArray(data.skus) ? data.skus : [];

  if (previousBrandName && previousBrandName !== brandName) {
    var oldSheet = ss.getSheetByName(safeSheetName(previousBrandName));
    if (oldSheet) ss.deleteSheet(oldSheet);
    deleteIndexRecord(previousBrandName);
  }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();

  var headers = ["SKU", "Product Name", "Blank Model", "Size", "Color", "Hits", "Cost (MFA)", "MSRP (Manual)", "Program Tier", "Status"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.setFrozenRows(1);

  if (skus.length) {
    var rows = skus.map(function(item) {
      return [
        item.sku,
        item.product,
        item.model,
        item.size,
        item.color,
        Number(item.hits),
        Number(item.cost),
        item.msrp === "" || item.msrp == null ? "" : Number(item.msrp),
        item.category || "",
        status
      ];
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 7, rows.length, 2).setNumberFormat("$#,##0.00");
  }
  sheet.autoResizeColumns(1, headers.length);

  catalog.name = brandName;
  catalog.active = status === "active";
  catalog.skus = skus;
  upsertIndexRecord(catalog, sheetName, status);
  return jsonResponse({ status: "success", count: skus.length, brandName: brandName, sheetName: sheetName });
}

function setBrandStatus(data) {
  var brandName = String(data.brandName || "").trim();
  var status = String(data.status || "").toLowerCase() === "inactive" ? "inactive" : "active";
  var indexSheet = getIndexSheet();
  var row = findIndexRow(indexSheet, brandName);
  if (!row) throw new Error("Brand catalog not found: " + brandName);

  var configText = indexSheet.getRange(row, 7).getValue();
  var catalog = configText ? JSON.parse(configText) : { name: brandName };
  catalog.active = status === "active";
  indexSheet.getRange(row, 5).setValue(status);
  indexSheet.getRange(row, 6).setValue(new Date());
  indexSheet.getRange(row, 7).setValue(JSON.stringify(catalog));

  var sheetName = indexSheet.getRange(row, 2).getValue() || safeSheetName(brandName);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).setValue(status);
  }
  return jsonResponse({ status: "success", brandName: brandName, catalogStatus: status });
}

function deleteBrandCatalog(data) {
  var brandName = String(data.brandName || "").trim();
  var indexSheet = getIndexSheet();
  var row = findIndexRow(indexSheet, brandName);
  var sheetName = row ? indexSheet.getRange(row, 2).getValue() : safeSheetName(brandName);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  if (row) indexSheet.deleteRow(row);
  return jsonResponse({ status: "success", deleted: brandName });
}

function fetchAllCatalogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var catalogs = readCatalogIndex();
  var catalogBySheet = {};
  catalogs.forEach(function(catalog) { catalogBySheet[catalog.sheetName || safeSheetName(catalog.name)] = catalog; });
  var items = [];

  ss.getSheets().forEach(function(sheet) {
    if (sheet.getName().charAt(0) === "_") return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var catalog = catalogBySheet[sheet.getName()];
    var brandName = catalog ? catalog.name : sheet.getName();
    for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
      var row = values[rowIndex];
      if (!row[0]) continue;
      var status = String(row[9] || (catalog && catalog.active === false ? "inactive" : "active")).toLowerCase();
      items.push({
        brand: brandName,
        brandCode: catalog ? catalog.code : "",
        sku: row[0],
        product: row[1],
        model: row[2],
        size: row[3],
        color: row[4],
        hits: row[5],
        cost: row[6],
        msrp: row[7],
        category: row[8],
        status: status,
        active: status !== "inactive"
      });
    }
  });

  return jsonResponse({ status: "success", items: items, catalogs: catalogs });
}

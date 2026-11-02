function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Yash Marketing - Stock Inward')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

function getDropdownData() {
  var sheetUrl = "https://docs.google.com/spreadsheets/d/1HgFw1etD-9LgPHLOI3fyXjUpOwye_HfP2VYyv7rsKt4/edit";
  var ss = SpreadsheetApp.openByUrl(sheetUrl);
  
  var productSheet = ss.getSheetByName("Products"); 
  var data = productSheet.getDataRange().getDisplayValues();
  var allProducts = [];
  
  for (var i = 1; i < data.length; i++) {
    var brnd = data[i][0] ? data[i][0].toString().trim() : ""; 
    var modl = data[i][1] ? data[i][1].toString().trim() : ""; 
    var cat = data[i][3]  ? data[i][3].toString().trim() : ""; 
    if (brnd || modl || cat) allProducts.push({ b: brnd, m: modl, c: cat });
  }

  var staffSheet = ss.getSheetByName("Staff Master");
  var allStaffs = []; 
  if (staffSheet) {
    var staffData = staffSheet.getDataRange().getDisplayValues();
    for (var j = 1; j < staffData.length; j++) {
      var empId = staffData[j][0] ? staffData[j][0].toString().trim() : ""; 
      var sName = staffData[j][1] ? staffData[j][1].toString().trim() : ""; 
      var sPass = staffData[j][4] ? staffData[j][4].toString().trim() : ""; 
      if (empId && sName) {
        allStaffs.push({ id: empId.toLowerCase(), n: sName, p: sPass });
      }
    }
  }

  var transSheet = ss.getSheetByName("Transporters");
  var allTransporters = [];
  if (transSheet) {
    var transData = transSheet.getDataRange().getDisplayValues();
    for (var k = 1; k < transData.length; k++) {
      var tName = transData[k][0] ? transData[k][0].toString().trim() : "";
      if (tName) allTransporters.push(tName);
    }
  }
  
  return { products: allProducts, staffs: allStaffs, transporters: allTransporters };
}

// 🔥 HISTORY FETCHING MEIN BHI COLUMN MAPPING THEEK KAR DI 🔥
function getInwardHistory() {
  var sheetUrl = "https://docs.google.com/spreadsheets/d/1HgFw1etD-9LgPHLOI3fyXjUpOwye_HfP2VYyv7rsKt4/edit";
  var ss = SpreadsheetApp.openByUrl(sheetUrl);
  var sheet = ss.getSheets()[0]; 
  var data = sheet.getDataRange().getDisplayValues();
  
  var invoices = {};
  var startIdx = data.length > 500 ? data.length - 500 : 1; 
  
  for (var i = data.length - 1; i >= startIdx; i--) {
    var invNo = data[i][6] ? data[i][6].toString().trim() : "";
    if (!invNo) continue;

    var uniqueKey = invNo + "_" + data[i][0];

    if (!invoices[uniqueKey]) {
      invoices[uniqueKey] = {
        invoice: invNo,
        inwardDate: data[i][0],
        transporter: data[i][3],       
        staff: data[i][2],
        commonBrand: data[i][8] ? data[i][8] : "-",                   
        // P (15) Khali Hai
        invoicePhotos: data[i][16] && data[i][16] !== "-" ? data[i][16] : "No Photo", // Q (16)
        boxPhotos: data[i][17] && data[i][17] !== "-" ? data[i][17] : "No Photo",     // R (17)
        pavtiPhotos: data[i][18] && data[i][18] !== "-" ? data[i][18] : "No Photo",   // S (18)
        // T (19) Product Photo niche items array me hai
        missingProd: data[i][20] ? data[i][20] : "-",                 // U (20)
        missingQty: data[i][21] ? data[i][21] : "0",                  // V (21)
        items: []
      };
    }
    
    invoices[uniqueKey].items.push({
      model: data[i][9],
      qty: data[i][11], // Actual Qty
      prodPhoto: data[i][19] ? data[i][19] : "-"                      // T (19): Product Photo
    });
  }

  var historyArray = [];
  for (var key in invoices) {
    historyArray.push(invoices[key]);
  }
  return historyArray;
}

// 🔥 SAVING MEIN AAPKA PERFECT SEQUENCE 🔥
function processFormFast(payloadStr) {
  try {
    var data = JSON.parse(payloadStr);
    
    var sheetUrl = "https://docs.google.com/spreadsheets/d/1HgFw1etD-9LgPHLOI3fyXjUpOwye_HfP2VYyv7rsKt4/edit";
    var sheet = SpreadsheetApp.openByUrl(sheetUrl).getSheets()[0]; 
    var folderId = "1nNpNx0dQDTHarHkZAHWSS8qp-nbcDKFT"; 
    var folder = DriveApp.getFolderById(folderId);
    
    function uploadFiles(base64Array) {
      var urls = [];
      if (!base64Array || !Array.isArray(base64Array)) return "-";
      for (var i = 0; i < base64Array.length; i++) {
        var fileData = base64Array[i];
        if (!fileData || !fileData.data || fileData.data.length < 50) continue;
        try {
          var blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.mimeType || "image/jpeg", fileData.name || ("photo_" + (i + 1) + ".jpg"));
          urls.push(folder.createFile(blob).getUrl());
        } catch (uploadErr) {
          Logger.log("Upload error for file " + i + ": " + uploadErr.message);
        }
      }
      return urls.length > 0 ? urls.join(" \n ") : "-";
    }

    var finalInvoiceUrl = uploadFiles(data.invoicePhotosB64);
    var finalBoxUrl = uploadFiles(data.boxPhotosB64);
    var finalPavtiUrl = uploadFiles(data.pavtiPhotosB64);
    
    var count = data.items.length; 
    var rowsToAppend = [];
    
    for (var i = 0; i < count; i++) {
      var item = data.items[i];
      
      var pPhotoUrl = "-";
      if (item.productPhotoB64 && item.productPhotoB64.data && item.productPhotoB64.data.length > 50) {
        try {
          var pBlob = Utilities.newBlob(Utilities.base64Decode(item.productPhotoB64.data), item.productPhotoB64.mimeType || "image/jpeg", item.productPhotoB64.name || "prod.jpg");
          pPhotoUrl = folder.createFile(pBlob).getUrl();
        } catch (pErr) {
          Logger.log("Product photo upload error: " + pErr.message);
        }
      }

      rowsToAppend.push([
        data.inwardDate || "-",           // A: Inward Date
        data.inwardTime || "-",           // B: Inward Time
        data.staffName || "-",            // C: Staff Name
        data.transporter || "-",          // D: Transporter Name
        data.lrNumber || "-",             // E: LR Number
        data.masterCartons || "0",        // F: Master Cartons
        data.invoiceNo || "-",            // G: Invoice Number
        data.invoiceDate || "-",          // H: Invoice Date
        data.commonBrand || "-",          // I: Brand Name
        item.modelName || "-",            // J: Model Name
        item.category || "-",             // K: Category
        item.piecesPerCarton || "0",      // L: Pieces per Carton (Invoice Qty)
        item.actualQty || "0",            // M: Actual Found Qty
        item.shortageQty || "0",          // N: Shortage Qty
        item.doaQty || "0",               // O: DOA
        "-",                              // P: (KHALI / EMPTY)
        finalInvoiceUrl,                  // Q: Invoice Photo
        finalBoxUrl,                      // R: Box Photo
        finalPavtiUrl,                    // S: Pavti Photo
        pPhotoUrl,                        // T: Product Photo
        data.missingModel || "-",         // U: Missing Product Model
        data.missingQty || "0"            // V: Missing Qty
      ]);
    }
    
    if (rowsToAppend.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }
    
    return "✅ " + count + " Items Saved Super Fast & Successfully!";
  } catch (error) {
    throw new Error(error.toString());
  }
}
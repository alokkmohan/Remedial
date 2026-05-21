// ============================================================
// उपचारात्मक शिक्षण मॉड्यूल — के.आर.पी. सूचना फॉर्म
// Google Apps Script — Code.gs
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('GasForm')
    .setTitle('के.आर.पी. सूचना फॉर्म — उपचारात्मक शिक्षण मॉड्यूल')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --------------- Spreadsheet Setup ---------------

var SPREADSHEET_ID = '1fWx_ePTFglJr25eyFRaTY72splmLHzNEX-mDCx7oYxI';

function getOrCreateSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateDataSheet(ss) {
  var sheet = ss.getSheetByName('KRP डेटा');
  if (!sheet) {
    var defaultSheet = ss.getSheetByName('Sheet1');
    sheet = ss.insertSheet('KRP डेटा');

    var headers = [
      'दिनांक', 'जनपद', 'विद्यालय / संस्थान',
      'कक्षा', 'विषय', 'KRP प्रकार',
      'अध्यापक / अध्यापिका का नाम', 'पदनाम', 'ई-मेल', 'मोबाइल नं.'
    ];
    sheet.appendRow(headers);

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setHorizontalAlignment('center');

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 80);
    sheet.setColumnWidth(7, 250);
    sheet.setColumnWidth(8, 150);
    sheet.setColumnWidth(9, 200);
    sheet.setColumnWidth(10, 130);

    if (defaultSheet && defaultSheet.getName() === 'Sheet1') {
      try { ss.deleteSheet(defaultSheet); } catch (e) {}
    }
  }
  return sheet;
}

// --------------- Form Submission ---------------

function submitForm(formData) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateDataSheet(ss);

    var timestamp = new Date();
    var district = formData.district || '';
    var school   = formData.school   || '';

    var subjects9 = [
      { key: 'hindi',    name: 'हिन्दी' },
      { key: 'english9', name: 'अंग्रेजी' },
      { key: 'math9',    name: 'गणित' },
      { key: 'science',  name: 'विज्ञान' }
    ];

    var subjects11 = [
      { key: 'english11', name: 'अंग्रेजी' },
      { key: 'math11',    name: 'गणित' },
      { key: 'physics',   name: 'भौतिक विज्ञान' },
      { key: 'chemistry', name: 'रसायन विज्ञान' },
      { key: 'biology',   name: 'जीव विज्ञान' }
    ];

    var rowsAdded = 0;

    function saveClass(classData, className, subjects) {
      if (!classData) return;
      subjects.forEach(function (subj) {
        var krps = classData[subj.key] || [];
        krps.forEach(function (krp, idx) {
          if (krp.name) {
            var krpType = idx === 0 ? 'School' : 'DIET';
            sheet.appendRow([
              timestamp, district, school,
              className, subj.name, krpType,
              krp.name || '', krp.designation || '',
              krp.email || '', krp.mobile || ''
            ]);
            rowsAdded++;
          }
        });
      });
    }

    saveClass(formData.class9,  'कक्षा-9',  subjects9);
    saveClass(formData.class11, 'कक्षा-11', subjects11);

    if (rowsAdded === 0) {
      return { success: false, message: 'कृपया कम से कम एक के.आर.पी. की जानकारी भरें।' };
    }

    return {
      success: true,
      message: rowsAdded + ' के.आर.पी. की जानकारी सफलतापूर्वक सहेजी गई!',
      sheetUrl: ss.getUrl()
    };

  } catch (e) {
    return { success: false, message: 'त्रुटि: ' + e.message };
  }
}

// --------------- doPost — called from GitHub Pages form ---------------

function doPost(e) {
  try {
    var formData = JSON.parse(e.postData.contents);
    var result = submitForm(formData);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --------------- Excel Download Menu ---------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KRP Tools')
    .addItem('Excel mein Download karein (.xlsx)', 'downloadAsExcel')
    .addToUi();
}

function downloadAsExcel() {
  var ss = getOrCreateSpreadsheet();
  var id = ss.getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx&id=' + id;
  var html = HtmlService.createHtmlOutput(
    '<script>' +
    'window.open("' + url + '");' +
    'google.script.host.close();' +
    '</script>'
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Downloading...');
}

// --------------- Admin: Get Sheet URL ---------------

function getSheetUrl() {
  try {
    return getOrCreateSpreadsheet().getUrl();
  } catch (e) {
    return '';
  }
}

'use strict';

var HEADER_ROW = 1;
var LABEL_ROW  = 2;
var DATA_ROW   = 3;
var HEADER_COL = 1;

var fixedHeaders = ['date', 'datetime', 'category'];
var measurements = {
  1: 'Weight (kg)',
  4: 'Height (meter)',
  5: 'Fat Free Mass (kg)',
  6: 'Fat Ratio (%)',
  8: 'Fat Mass Weight (kg)',
  9: 'Diastolic Blood Pressure (mmHg)',
  10: 'Systolic Blood Pressure (mmHg)',
  11: 'Heart Pulse (bpm)',
  12: 'Temperature',
  54: 'SP02 (%)',
  71: 'Body Temperature',
  73: 'Skin Temperature',
  76: 'Muscle Mass',
  77: 'Hydration',
  88: 'Bone Mass',
  91: 'Pulse Wave Velocity',
};

function WITHINGSLABEL(numstr) {
  if (!measurements[numstr]) {
    return numstr.charAt(0).toUpperCase() + numstr.slice(1);
  } else {
    return measurements[numstr];
  }
}

function getSheet() {
  var SHEET_NAME = 'Body Measures';
  var id = SpreadsheetApp.getActiveSpreadsheet().getId();
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    initializeSheet(sheet);
  }

  return sheet;
}

function initializeSheet(sheet) {
  if (!sheet) {
    Logger.log('Invalid argument(s)');
    return;
  }

  // Sets the header row and the label row.
  sheet.getRange(HEADER_ROW, 1, 1, fixedHeaders.length).setValues([fixedHeaders]);
  sheet.getRange(LABEL_ROW, 1, 1, sheet.getLastColumn()).setFormulaR1C1('=WITHINGSLABEL(R[-1]C[0])');

  // Hides the header row and column.
  sheet.hideRows(HEADER_ROW);
  sheet.hideColumns(HEADER_COL);
}

function getCurrentHeaders(sheet) {
  if (!sheet) {
    Logger.log('Invalid argument(s)');
    return;
  }

  var rowNum = sheet.getLastRow();
  if (rowNum === 0) {
    initializeSheet(sheet);
  }

  var colNum = sheet.getLastColumn();
  return sheet.getRange(HEADER_ROW, 1, 1, colNum).getValues()[0].map(function(elem) {return String(elem);});
}

function updateSheet(sheet, items) {
  if (!sheet || !items) {
    Logger.log('Invalid argument(s)');
    return;
  }

  // Picks up the all types in the items.
  var typeSet = items.reduce(function(obj, item) {
    for (var key in item) {
      if (/^\d+$/.test(key)) {
        obj[key] = true;
      }
    }
    return obj;
  }, {});

  // Makes an array of types.
  var headers = getCurrentHeaders(sheet);
  var types = Object.keys(typeSet);
  types.sort(function(a, b) {return Number(a) - Number(b);});
  types.forEach(function(type) {
    if (headers.indexOf(type) === -1) {
      headers.push(type);
    }
  });

  // Updates the header row of the sheet.
  sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  initializeSheet(sheet);

  // Gets the keys of the stored data for checking whether the each item has already been stored or not.
  var rowNum = sheet.getLastRow();
  var primaryKeys = sheet.getRange(1, HEADER_COL, rowNum, 1).getValues().map(function(aoa) {return aoa[0];});
  primaryKeys = primaryKeys.slice(DATA_ROW - 1);

  // Appends the new items to the sheet.
  var cells = [];
  items.forEach(function(item) {
    if (primaryKeys.indexOf(item.date) !== -1) {
      return;
    }
    cells.push(headers.map(function(header) {return (item[header] !== undefined) ? item[header] : '';}));
  });
  if (cells.length > 0) {
    sheet.getRange(rowNum + 1, 1, cells.length, headers.length).setValues(cells);
  }
}

function convertWithingsData(data) {
  var items = [];
  if (!data || (data.status && data.status !== 0) || !data.body || !data.body.measuregrps) {
    return items;
  }

  var timezone = data.body.timezone || 'GMT';
  data.body.measuregrps.forEach(function(measuregrp) {
    if (!measuregrp.date || !measuregrp.measures) {
      return;
    }

    var item = measuregrp.measures.reduce(function(obj, measure) {
      if (!measure.type || !measure.value || !measure.unit) {
        return obj;
      }
      if (measure.unit < 0) {
        obj[measure.type] = measure.value / Math.pow(10, -measure.unit);
      } else {
        obj[measure.type] = measure.value * Math.pow(10, measure.unit);
      }
      return obj;
    }, {});

    item.date = measuregrp.date;
    item.datetime = Utilities.formatDate(new Date(measuregrp.date * 1000), timezone, 'yyyy/MM/dd HH:mm:ss');
    item.category = measuregrp.category;

    items.push(item);
  });

  items.sort(function(a, b) {return a.date - b.date;});

  return items;
}

var withings = new WithingsWebService(
  PropertiesService.getScriptProperties().getProperty('withingsConsumerKey'),
  PropertiesService.getScriptProperties().getProperty('withingsConsumerSecret'),
  authCallback);

function authCallback(request) {
  return withings.authCallback(request);
}

function triggerStoringDataFromWithingsApi() {
  // Creates the OAuth1 service for Withings API.
  var service = withings.getService();
  if (!service.hasAccess()) {
    Logger.log('Needs to authorize.');
    withings.logAuthorizationUrl();
    withings.logCallbackUrl();
    return;
  }

  // Accesses to the Withings API to get body measure data.
  var userid = PropertiesService.getScriptProperties().getProperty('userid');
  var url = 'https://wbsapi.withings.net/measure?action=getmeas&userid=' + userid;
  var response = service.fetch(url);
  if (response.getResponseCode() >= 400) {
    Logger.log('Something wrong with HTTP response:');
    Logger.log(response);
    return;
  }

  // Stores the obtained body measure data into the sheet.
  var result = JSON.parse(response.getContentText());
  if (result.status !== 0) {
    Logger.log('Something wrong with Withings API: %s', response.getContentText());
    return;
  }
  var items = convertWithingsData(result);
  updateSheet(getSheet(), items);
}

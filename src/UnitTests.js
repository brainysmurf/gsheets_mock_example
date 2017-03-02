
//Unit tests
function runSheetsUtilsTests_() {
  //Prepare sheet for tests
  var ss = SpreadsheetApp.create("SheetsUtils Test");
  var sheet = ss.getSheets()[0].clear();
  SpreadsheetApp.flush();

  //Variables
  var sheetId = sheet.getSheetId();
  var sheetName = sheet.getName();

  //Global fixtures
  var expectedHeaders = ['Record ID 1', 'Record ID 2', 'test1','Test 2','Test2a','Test 3'];

  var objArray = [
    {
      'Record ID 1': 12345,
      'Record ID 2': 123,
      'test1': 0,
      'Test 2': "Hello",
      'Test2a': '',
      'Test 3': 'World'
    },
    {
      'Record ID 1': 12346,
      'Record ID 2': 123,
      'test1': 1,
      'Test 2': '',
      'Test2a': "Hello",
      'Test 3': 'World'
    },
    {
      'Record ID 1': 12347,
      'Record ID 2': 125,
      'test1': 2,
      'Test 2': "Hello",
      'Test2a': '',
      'Test 3': 'World'
    }
  ];

  var dataArray = [
    expectedHeaders,
    [12345, 123, 0, 'Hello', '', 'World'],
    [12346, 123, 1, '', 'Hello', 'World'],
    [12347, 125, 2, 'Hello','','World']
  ];

  (function() {
    var thisSheet = SheetsUtils.getSheetById(ss, sheetId.toString());
    var thisSheetName = thisSheet.getName();
    if (thisSheetName !== sheetName) {
      throw "getSheetById() doesn't work with sheet Id in string format";
    }
  })();


  (function() {
    var thisDataArray = _.clone(dataArray);
    var headers = thisDataArray.shift();
    var objects = SheetsUtils.convert2DArrayToObjects(thisDataArray, headers);
    if (!_.isEqual(objects, objArray)) {
      throw "convert2DArrayToObjects() doesn't work";
    }
  })();


  (function() {
    sheet.clear();
    var headers = SheetsUtils.getUpsertHeaders(sheet, {expectedHeaders: expectedHeaders, freezeHeaders: true});
    if (
        _.difference(headers, expectedHeaders).length !== 0 ||
        _.difference(expectedHeaders, headers).length !== 0
    ) {
      throw "getUpsertHeaders() doesn't work when writing to blank sheet";
    }
  })();


  (function() {
    sheet.clear();
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    var newHeaders = ['Record ID 1', 'Test 5', 'Record ID 2', 'test1','Test 2','Test2a','Test 3','Test 4'];
    var expected = _.clone(newHeaders);
    var returnedHeaders = SheetsUtils.getUpsertHeaders(sheet, {expectedHeaders: newHeaders, freezeHeaders: true});
    var headersViaGAS = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (
        _.difference(returnedHeaders, expected).length !== 0 ||
        _.difference(expected, returnedHeaders).length !== 0
    ) {
      throw "getUpsertHeaders() doesn't work when updating a sheet with missing headers";
    }
    if (
        _.difference(headersViaGAS, expected).length !== 0 ||
        _.difference(expected, headersViaGAS).length !== 0
    ) {
      throw "getUpsertHeaders() doesn't work when updating a sheet with missing headers";
    }
  })();


  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
    var data = SheetsUtils.getRowsData(sheet);
    if (
        !(_.isEqual(objArray, data))
    ) {
      throw "getRowsData() doesn't work";
    }
  })();


  (function() {
    var thisObjArray = _.cloneDeep(objArray);
    sheet.clear();
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    SheetsUtils.setRowsData(sheet, thisObjArray);
    var theseData = sheet.getDataRange().getValues();
    if (
        !(_.isEqual(dataArray, theseData))
    ) {
      throw "setRowsData() doesn't work";
    }
  })();


  (function() {
    var thisObjArray = _.cloneDeep(objArray);
    sheet.clear();
    sheet.getRange(3, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    var headersRange = sheet.getRange(3, 1, 1, expectedHeaders.length);
    var params = {
      headersRange: headersRange
    };
    SheetsUtils.setRowsData(sheet, thisObjArray, params);
    var theseData = sheet.getRange(3, 1, sheet.getLastRow()-2, sheet.getLastColumn()).getValues();
    if (
        !(_.isEqual(dataArray, theseData))
    ) {
      throw "setRowsData() doesn't work when headersRange is specified as other than the first row";
    }
  })();


  (function() {
    var thisSheet = SheetsUtils.getSheetById(ss, sheetId);
    var thisSheetName = thisSheet.getName();
    if (thisSheetName !== sheetName) {
      throw "getSheetById() doesn't work";
    }
  })();


  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var thisObjArray = _.cloneDeep(objArray);

    var rowToAppend = {
      'Record ID 1': 12348,
      'Record ID 2': 126,
      'test1': 3,
      'Test 2': "Goodbye",
      'Test2a': 'Cruel',
      'Test 3': 'World'
    };

    thisObjArray.push(rowToAppend);

    var expectedValues = thisObjArray;

    SheetsUtils.appendRowsData(sheet, [rowToAppend]);

    var data = SheetsUtils.getRowsData(sheet);
    if (
        !(_.isEqual(expectedValues, data))
    ) {
      throw "appendRowData() doesn't work";
    }
  })();


  (function() {

    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var sourceObject = {
      'Record ID 2': 123,
      'test1': 99999,
      'Test 2': "Goodnight",
      'Test2a': 'Moon',
      'Test 3': 'Chair' //this key should not update
    };
    var matchHeaders = "Record ID 2";

    var expectedReturn = {
      recordsUpdated: 2,
      errors: 0
    };
    var expectedUpdateResult = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 123,
        'test1': 99999,
        'Test 2': "Goodnight",
        'Test2a': 'Moon',
        'Test 3': 'Chair'
      },
      {
        'Record ID 1': 12346,
        'Record ID 2': 123,
        'test1': 99999,
        'Test 2': 'Goodnight',
        'Test2a': "Moon",
        'Test 3': 'Chair'
      },
      {
        'Record ID 1': 12347,
        'Record ID 2': 125,
        'test1': 2,
        'Test 2': "Hello",
        'Test2a': '',
        'Test 3': 'World'
      }
    ];

    var params = {
      requireUnique: false
    };

    var result = SheetsUtils.updateRowsInPlace(sheet, sourceObject, matchHeaders, params);
    if (!_.isEqual(result, expectedReturn)) {
      throw "updateRowsInPlace() doesn't work - updating multiple records, single match key";
    }
    var data = SheetsUtils.getRowsData(sheet);
    if (!_.isEqual(data, expectedUpdateResult)) {
      throw "updateRowsInPlace() doesn't work - updating multiple records, single match key";
    }

    params = {
      requireUnique: true
    };

    var message;
    var expectedErrmessage = "Multiple records in destination match based on your match headers.";
    try {
      SheetsUtils.updateRowsInPlace(sheet, sourceObject, matchHeaders, params);
    } catch (err) {
      message = err.message;
    }
    if (message !== expectedErrmessage) {
      throw "updateRowsInPlace() doesn't throw error on duplicates when require unique is specified.";
    }

  })();

  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var sourceArray = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 4444,
        'test1': 77777,
        'Test 2': 'Goodnight',
        'Test2a': "Moon",
        'Test 3': 'Boy'

      },
      {
        'Record ID 1': 12346,
        'Record ID 2': 5555,
        'test1': 88888,
        'Test 2': 'Hello',
        'Test2a': "Jupiter",
        'Test 3': ''

      }
    ];

    var expectedUpdateResult = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 4444,
        'test1': 77777,
        'Test 2': 'Goodnight',
        'Test2a': "Moon",
        'Test 3': 'Boy'
      },
      {
        'Record ID 1': 12346,
        'Record ID 2': 5555,
        'test1': 88888,
        'Test 2': 'Hello',
        'Test2a': "Jupiter",
        'Test 3': ''
      },
      {
        'Record ID 1': 12347,
        'Record ID 2': 125,
        'test1': 2,
        'Test 2': "Hello",
        'Test2a': '',
        'Test 3': 'World'
      }
    ];

    SheetsUtils.updateRowsData(sheet, sourceArray, ['Record ID 1']);
    var data = SheetsUtils.getRowsData(sheet);
    if (!_.isEqual(data, expectedUpdateResult)) {
      throw "updateRowsData() doesn't work - updating multiple records, single match key";
    }

  })();



  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var sourceArray = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 4444,
        'test1': 77777,
        'Test 2': 'Goodnight',
        'Test2a': "Moon",
        'Test 3': 'Boy'

      },
      {
        'Record ID 1': 12345, //duplicate key value in source
        'Record ID 2': 5555,
        'test1': 88888,
        'Test 2': 'Hello',
        'Test2a': "Jupiter",
        'Test 3': ''

      }
    ];

    var message;
    try {
      SheetsUtils.updateRowsData(sheet, sourceArray, ['Record ID 1']);
    } catch(err) {
      message = err.message;
    }
    var expectedMessage = "updateRowsData - Duplicate record(s) in source data with key 12345";

    if (!_.isEqual(message, expectedMessage)) {
      throw "updateRowsData() doesn't throw error when duplicates exist in source data";
    }

  })();


  (function() {
    var thisDataArray = _.cloneDeep(dataArray);
    thisDataArray[2][0] = 12345; //set duplicate key value in destination
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(thisDataArray);

    var sourceArray = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 4444,
        'test1': 77777,
        'Test 2': 'Goodnight',
        'Test2a': "Moon",
        'Test 3': 'Boy'

      },
      {
        'Record ID 1': 12346,
        'Record ID 2': 5555,
        'test1': 88888,
        'Test 2': 'Hello',
        'Test2a': "Jupiter",
        'Test 3': ''

      }
    ];

    var message;
    try {
      SheetsUtils.updateRowsData(sheet, sourceArray, ['Record ID 1']);
    } catch(err) {
      message = err.message;
    }
    var expectedMessage = "updateRowsData - Duplicate record(s) in destination with key 12345";

    if (!_.isEqual(message, expectedMessage)) {
      throw "updateRowsData() doesn't throw error when duplicates exist in destination data";
    }

  })();




  (function() {
    var thisDataArray = _.cloneDeep(dataArray);
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(thisDataArray);
    var mappedObjArray = [
      {
        'recordId1': 12345,
        'recordId2': 123,
        'test1': 0,
        'test2': "Hello",
        'test2a': '',
        'test3': 'World'
      },
      {
        'recordId1': 12346,
        'recordId2': 123,
        'test1': 1,
        'test2': '',
        'test2a': "Hello",
        'test3': 'World'
      },
      {
        'recordId1': 12347,
        'recordId2': 125,
        'test1': 2,
        'test2': "Hello",
        'test2a': '',
        'test3': 'World'
      }
    ];

    var mappedUpdateHeaders = {
      'test1': 'test1',
      'Test 2': 'test2',
      'Test2a': 'test2a',
      'Test 3': 'test3'
    };

    SheetsUtils.setMappedRowsData(sheet, mappedObjArray, mappedUpdateHeaders);
    var thisObjArray = _.clone(objArray);
    SheetsUtils.setRowsData(sheet, thisObjArray);
    var theseData = sheet.getDataRange().getValues();
    if (
        !(_.isEqual(thisDataArray, theseData))
    ) {
      throw "setMappedRowsData() doesn't work";
    }
  })();




  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
    var dataToAppend = [
      {
        'recordId1': 12348,
        'recordId2': 126,
        'test1': 3,
        'test2': "Greetings",
        'test2a': 'fellow',
        'test3': 'traveler'
      },
      {
        'recordId1': 12349,
        'recordId2': 127,
        'test1': 4,
        'test2': "",
        'test2a': 'Go',
        'test3': 'onward'
      }
    ];

    var mappedUpdateHeaders = {
      'Record ID 1':'recordId1',
      'Record ID 2':'recordId2',
      'test1':'test1',
      'Test 2':'test2',
      'Test2a':'test2a',
      'Test 3':'test3'
    };

    var thisDataArray = _.cloneDeep(dataArray);
    thisDataArray.push([12348, 126, 3, "Greetings","fellow","traveler"]);
    thisDataArray.push([12349, 127, 4, "","Go","onward"]);

    SheetsUtils.appendMappedRowsData(sheet, dataToAppend, mappedUpdateHeaders);
    var theseData = sheet.getDataRange().getValues();

    if (
        !(_.isEqual(thisDataArray, theseData))
    ) {
      throw "appendMappedRowsData() doesn't work";
    }

  })();




  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var sourceObject = {
      'recordId': 12345,
      'recordId2': 123,
      'test1': 99999,
      'test2': "Goodnight",
      'test2a': 'Moon',
      'test3': 'Chair' //this key should not update
    };
    var mappedMatchHeaders = {
      'Record ID 2': 'recordId2'
    };
    var mappedUpdateHeaders = {
      'test1':'test1',
      'Test 2': 'test2',
      'Test2a': 'test2a'
    };
    var expectedReturn = {
      recordsUpdated: 2,
      errors: 0
    };
    var expectedUpdateResult = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 123,
        'test1': 99999,
        'Test 2': "Goodnight",
        'Test2a': 'Moon',
        'Test 3': 'World'
      },
      {
        'Record ID 1': 12346,
        'Record ID 2': 123,
        'test1': 99999,
        'Test 2': 'Goodnight',
        'Test2a': "Moon",
        'Test 3': 'World'
      },
      {
        'Record ID 1': 12347,
        'Record ID 2': 125,
        'test1': 2,
        'Test 2': "Hello",
        'Test2a': '',
        'Test 3': 'World'
      }
    ];

    var params = {
      requireUnique: false
    };

    var result = SheetsUtils.updateMappedRowsInPlace(sheet, sourceObject, mappedMatchHeaders, mappedUpdateHeaders, params);
    if (!_.isEqual(result, expectedReturn)) {
      throw "updateMappedRowsInPlace() doesn't work - updating multiple records, single match key";
    }
    var data = SheetsUtils.getRowsData(sheet);
    if (!_.isEqual(data, expectedUpdateResult)) {
      throw "updateMappedRowsInPlace() doesn't work - updating multiple records, single match key";
    }

    params = {
      requireUnique: true
    };
    var message;
    var expectedErrmessage = "Multiple records in destination match based on your mapped join keys.";
    try {
      result = SheetsUtils.updateMappedRowsInPlace(sheet, sourceObject, mappedMatchHeaders, mappedUpdateHeaders, params);
    } catch (err) {
      message = err.message;
    }
    if (message !== expectedErrmessage) {
      throw "updateMappedRowsInPlace() doesn't throw error on duplicates when require unique is specified.";
    }

  })();



  (function() {

    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var sourceObject = {
      'recordId': 12345,
      'recordId2': 123,
      'test1': 99999,
      'test2': "Goodnight",
      'test2a': 'Moon',
      'test3': 'Chair'
    };
    var mappedMatchHeaders = {
      'Record ID 1': 'recordId',
      'Record ID 2': 'recordId2'
    };
    var mappedUpdateHeaders = {
      'test1':'test1',
      'Test 2': 'test2',
      'Test2a': 'test2a',
      'Test 3': 'test3'
    };
    var expectedReturn = {
      recordsUpdated: 1,
      errors: 0
    };
    var expectedUpdateResult = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 123,
        'test1': 99999,
        'Test 2': "Goodnight",
        'Test2a': 'Moon',
        'Test 3': 'Chair'
      },
      {
        'Record ID 1': 12346,
        'Record ID 2': 123,
        'test1': 1,
        'Test 2': '',
        'Test2a': "Hello",
        'Test 3': 'World'
      },
      {
        'Record ID 1': 12347,
        'Record ID 2': 125,
        'test1': 2,
        'Test 2': "Hello",
        'Test2a': '',
        'Test 3': 'World'
      }
    ];

    var params = {
      requireUnique: false
    };
    var result = SheetsUtils.updateMappedRowsInPlace(sheet, sourceObject, mappedMatchHeaders, mappedUpdateHeaders, params);
    if (!_.isEqual(result, expectedReturn)) {
      throw "updateMappedRowsInPlace() doesn't work - updating single record, multiple match keys";
    }

    var data = SheetsUtils.getRowsData(sheet);
    if (!_.isEqual(data, expectedUpdateResult)) {
      throw "updateMappedRowsInPlace() doesn't work - updating single record, multiple match keys";
    }


  })();



  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var sourceArray = [
      {
        'recordId': 12345,
        'recordId2': 4444,
        'test1': 77777,
        'test2': 'Goodnight',
        'test2a': "Moon",
        'test3': 'Boy'

      },
      {
        'recordId': 12346,
        'recordId2': 5555,
        'test1': 88888,
        'test2': 'Hello',
        'test2a': "Jupiter",
        'test3': ''

      }
    ];

    var mappedMatchHeaders = {
      'Record ID 1': 'recordId',
    };

    var mappedUpdateHeaders = {
      'test1':'test1', //note that Record ID 2 is excluded from update mask
      'Test 2': 'test2',
      'Test2a': 'test2a',
      'Test 3': 'test3'
    };

    var expectedUpdateResult = [
      {
        'Record ID 1': 12345,
        'Record ID 2': 123, //not updated - excluded from update mask
        'test1': 77777,
        'Test 2': 'Goodnight',
        'Test2a': "Moon",
        'Test 3': 'Boy'
      },
      {
        'Record ID 1': 12346,
        'Record ID 2': 123,  //not updated - excluded from update mask
        'test1': 88888,
        'Test 2': 'Hello',
        'Test2a': "Jupiter",
        'Test 3': ''
      },
      {
        'Record ID 1': 12347,
        'Record ID 2': 125,
        'test1': 2,
        'Test 2': "Hello",
        'Test2a': '',
        'Test 3': 'World'
      }
    ];

    SheetsUtils.updateMappedRowsData(sheet, sourceArray, mappedMatchHeaders, mappedUpdateHeaders);
    var data = SheetsUtils.getRowsData(sheet);
    if (!_.isEqual(data, expectedUpdateResult)) {
      throw "updateMappedRowsData() doesn't work - single match key";
    }

  })();

  (function() {
    sheet.clear();
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);

    var sourceArray = [
      {
        'recordId': 12345,
        'recordId2': 4444,
        'test1': 77777,
        'test2': 'Goodnight',
        'test2a': "Moon",
        'test3': 'Boy'

      },
      {
        'recordId': 12345,
        'recordId2': 5555,
        'test1': 88888,
        'test2': 'Hello',
        'test2a': "Jupiter",
        'test3': ''

      }
    ];

    var mappedMatchHeaders = {
      'Record ID 1': 'recordId',
    };

    var mappedUpdateHeaders = {
      'test1':'test1', //note that Record ID 2 is excluded from update mask
      'Test 2': 'test2',
      'Test2a': 'test2a',
      'Test 3': 'test3'
    };

    var expectedMessage = "updateMappedRowsData - Duplicate records in source data with key 12345";

    var message;
    try {
      SheetsUtils.updateMappedRowsData(sheet, sourceArray, mappedMatchHeaders, mappedUpdateHeaders);
    } catch(err) {
      message = err.message;
    }

    var data = SheetsUtils.getRowsData(sheet);
    if (!_.isEqual(message, expectedMessage)) {
      throw "updateMappedRowsData() doesn't throw expected error when there are duplicates in source data";
    }

  })();



  (function() {
    var thisDataArray = _.cloneDeep(dataArray);
    sheet.clear();
    thisDataArray[2][0] = 12345;
    sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(thisDataArray);

    var sourceArray = [
      {
        'recordId': 12345,
        'recordId2': 4444,
        'test1': 77777,
        'test2': 'Goodnight',
        'test2a': "Moon",
        'test3': 'Boy'

      },
      {
        'recordId': 12346,
        'recordId2': 5555,
        'test1': 88888,
        'test2': 'Hello',
        'test2a': "Jupiter",
        'test3': ''

      }
    ];

    var mappedMatchHeaders = {
      'Record ID 1': 'recordId'
    };

    var mappedUpdateHeaders = {
      'test1':'test1', //note that Record ID 2 is excluded from update mask
      'Test 2': 'test2',
      'Test2a': 'test2a',
      'Test 3': 'test3'
    };

    var expectedMessage = "updateMappedRowsData - Duplicate records in destination with key 12345";
    var message;

    try {
      SheetsUtils.updateMappedRowsData(sheet, sourceArray, mappedMatchHeaders, mappedUpdateHeaders);
    } catch(err) {
      message = err.message;
    }

    var data = SheetsUtils.getRowsData(sheet);
    if (!_.isEqual(message, expectedMessage)) {
      throw "updateMappedRowsData() doesn't throw expected error when there are duplicates in destination data";
    }

  })();

}
/*
   GSheetsUtils

   Copyright (c) 2016 Amplified Labs, a division of Amplified IT
   https://www.amplifiedit.com/
   
   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

//Private methods
// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
    return typeof(cellData) === "string" && cellData === "";
}


//Limited mimic of lodash assign -- copies non-inherited key value pairs from source into destination object.
function assign_(destObj, sourceObj) {
    for (var key in sourceObj) {
        if (sourceObj.hasOwnProperty(key)) {
            destObj[key] = sourceObj[key];
        }
    }
    return destObj;
}


//Public methods


/** getSheetById returns a Sheet object when given a Spreadsheet object and sheet Id
 * @example <caption>
 *   Useful when you need to find a sheet by its immutable Sheet Id.
 *   No method available for this exists in native SpreadsheetApp.
 * </caption>
 *   function getStoredSheet() {
 *     var ss = SpreadsheetApp.getActiveSpreadsheet();
 *     var sheetId = PropertiesService.getDocumentProperties().getProperty('sheetId', sheetId);
 *     var sheet = GSheetsUtils.getSheetById(ss, sheetId);
 *     return sheet;
 *   }
 *
 *  @param {Spreadsheet} spreadsheet - the Spreadsheet object where you expect the
 *  sheet with the given ID to live
 *  @param {string} sheetId - the id of the Sheet you want to get. Can be of type
 *  number or string, as long as it can be cast as numeric
 *  @returns {Sheet}
 */
function getSheetById(spreadsheet, sheetId) {
    if (!isNaN(sheetId)) {
        sheetId = parseInt(sheetId);
        var sheets = spreadsheet.getSheets();
        for (var i=0; i<sheets.length; i++) {
            var thisSheetId = sheets[i].getSheetId();
            if (thisSheetId === sheetId) {
                return sheets[i];
            }
        }
    } else {
        throw new Error('Bad value - getSheetById requires numeric sheetId');
    }
}


/** For every row of data in 2-d array, generates an object that contains the data.
 *  @example <caption>
 *    Take a 2-D array data the way it comes from .getValues() and change it into array of objects with specified keys.
 *    Generally not used in most Amplified Labs tools.  Instead we tend to use getRowsData() method.
 *  </caption>
 *    function convertArrayOfQuizScores() {
 *      var headers = ['Quiz 1','Quiz 2','Quiz 3'];
 *      var data = [[93, 87, 75],[91, 83, 71],[61, 63, 70]];
 *      var objects = GSheetsUtils.convertQuizScores(data, headers);
 *      debugger;
 *    }
 *
 *    returns [
 *      {"Quiz 1": 93, "Quiz 2": 91, "Quiz 3": 61},
 *      {"Quiz 1": 87, "Quiz 2": 83, "Quiz 3": 63},
 *      {"Quiz 1": 75, "Quiz 2": 71, "Quiz 3": 70}
 *    ]
 *
 *   @param {string[][]} data - JavaScript 2d array representing spreadsheet data
 *   @param {string[]} keys - Array of strings that define the property names for
 *   the objects to create
 *   @returns {Object[]} - Array of keyed data objects
 */
function convert2DArrayToObjects (data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty_(cellData)) {
                object[keys[j]] = '';
                continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    return objects;
}


/** getUpsertHeaders checks the designated sheet for the existence of the expected headers,
 * sets / upserts any missing headers in the same row as the headers range (default value of row 1)
 * optionally freezes header row(s)
 * returns array of amended headers
 *  @tutorial tutorial1
 *  @example
 *  <caption>
 *    A convenience method to fetch existing headers and also set / repair headers.
 *  </caption>
 *
 *  function getStatusColumnNumber() {
 *     var ss = SpreadsheetApp.getActiveSpreadsheet();
 *     var sheet = ss.getSheetByName('Students');
 *     var headers = GSheetsUtils.getUpsertHeaders(sheet);
 *     //headers are returned as 1-D array
 *     var colNum = headers.indexOf('Status') + 1;
 *     return colNum;
 *  }
 *
 *  function getStatusColumnNumberWithSetFix() {
 *     var ss = SpreadsheetApp.getActiveSpreadsheet();
 *     var sheet = ss.getSheetByName('Students');
 *     var expected = ['First Name','Last Name','Status'];
 *     var headers = GSheetsUtils.getUpsertHeaders(sheet, {expectedHeaders: expected, freezeHeaders: true});
 *     //headers are set / fixed in sheet and returned as 1-D array.
 *     //existing headers not in expected headers array are not disturbed.
 *     var colNum = headers.indexOf('Status') + 1;
 *     return colNum;
 *  }
 *
 *
 *  @param {Sheet} sheet - the Sheet Object where the headers will be read and/or written
 *  @param {Object} [optParams] - optional config parameters
 *  @param {string[]} [optParams.expectedHeaders] - 1-d array of headers you want to insert, or
 *  that you expect to exist in the sheet
 *  @oaram {rowNum} [optParams.columnHeadersRowIndex] - Row number in which headers are to be placed
 *  @param {Boolean} [optParams.freezeHeaders] - Boolean whether you want to freeze header
 *  row(s).  Defaults to false.
 *  @returns {string[]} - 1-D array of headers.
 */
function getUpsertHeaders(sheet, optParams) {
    var params = optParams || {};
    var lastCol = sheet.getLastColumn();
    var headerRow = params.columnHeadersRowIndex ? params.columnHeadersRowIndex : 1;
    var headersRange = (lastCol > 0) ? sheet.getRange(headerRow, 1, 1, lastCol) : null;
    var expectedHeaders = params.expectedHeaders || null;
    var headers = headersRange ? headersRange.getValues()[0] : [];
    if (!expectedHeaders) { //if no expected headers are provided, return existing headers
        return headers;
    }
    if (lastCol > 0) {  //upsert if headers already exist
        for (var i=0; i<expectedHeaders.length; i++) {
            var thisIndex = headers.indexOf(expectedHeaders[i]);
            if (thisIndex === -1) {
                sheet.getRange(headerRow, lastCol+1).setValue(expectedHeaders[i]);
                headers.splice(lastCol-1, 0,  expectedHeaders[i]);
                lastCol++;
            }
        }
    } else { //more efficient approach if sheet is blank
        sheet.getRange(headerRow, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
        headers = expectedHeaders;
    }
    if (params.freezeHeaders) {
        sheet.setFrozenRows(headerRow);
    }
    SpreadsheetApp.flush();
    return headers;
}

/**
 * getRowsData iterates row by row in the input range and returns an array of objects.
 * Assumes that headers in the destination sheet are also the keys of the source data
 * Each object contains all the data for a given row, indexed by its column name.
 * @tutorial tutorial1
 * @example
 * function basicUsage() {
 *    var ss = SpreadsheetApp.getActiveSpreadsheet();
 *    var sheet = ss.getSheetByName('Classes');
 *    var classes = GSheetsUtils.getRowsDate(sheet);
 *    debugger;
 *    //show array of all sheet data with sheet headers as keys
 * }
 *
 * function moreAdvancedUsage() {
 *    var ss = SpreadsheetApp.getActiveSpreadsheet();
 *    var sheet = ss.getSheetByName('Students');
 *    var headerRowNum = 3; //Assume headers aren't in first row
 *    var numColsToInclude = 8; //Assume we want to exclude data from columns after column 8
 *    var dataStartRow = 10; //Assume data doesn't actually start in row 4
 *    var lastRow = sheet.getLastRow();
 *    if (lastRow - 9 > 0) { //check that data actually exists in/below dataStartRow
 *    var dataRange = sheet.getRange(10, 1, lastRow - 9, numColsToInclude);
 *    var headersRange = sheet.getRange(headerRowNum, 1, 1, numColsToInclude);
 *    var data = GSheetsUtils.getRowsData(sheet, {dataRange: dataRange, headersRange: headersRange, columnHeadersRowIndex: headerRowNum});
 *    } else {
 *      data = [];
 *    }
 *    debugger;
 *    //shows array of all sheet data (as objects) starting in row 10, for the first 8 columns in the sheet
 *    //where the column headers are in row 3
 *
 * }
 * @param {Sheet} sheet - the sheet object that contains the data to be processed
 * @param {Object} [optParams] - optional config parameters
 * @param {Range} [optParams.dataRange] - the exact range of cells where the data is stored.
 * This argument is optional and it defaults to all the cells except those in the first row
 * or all the cells below columnHeadersRowIndex (if defined).
 * @param {Range} [optParams.headersRange] - the range that the headers are in - use to
 * limit which columns to read from
 * @param {rowNum} [optParams.columnHeadersRowIndex] - specifies the row number where the
 * column names are stored. This argument is optional and it defaults to the row immediately
 * above range;
 * @returns {Object[]} - array of objects keyed to sheet headers
 */
function getRowsData(sheet, optParams) {
    var params = optParams || {};
    var headersIndex = params.columnHeadersRowIndex || 1;
    if (sheet.getLastRow() < headersIndex + 1) {
        return [];
    }

    var dataRange = params.dataRange ||
        sheet.getRange(headersIndex+1, 1, sheet.getLastRow() - headersIndex, sheet.getLastColumn());
    var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
    var headersRange = params.headersRange || sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
    var headers = headersRange.getValues()[0];
    return this.convert2DArrayToObjects(dataRange.getValues(), headers);
}

/** setRowsData fills in one row of data per object defined in the objects Array.
 * Assumes that headers in the destination sheet are also the keys of the source data
 * For every Column, it checks if data objects define a value for it.
 * @tutorial tutorial2
 * @example
 *
 *    //assume sheet contains headers that match the keys of the objects below...
 *
 *    var dataToWrite = [
 *      {"Quiz 1": 93, "Quiz 2": 91, "Quiz 3": 61},
 *      {"Quiz 1": 87, "Quiz 2": 83, "Quiz 3": 63},
 *      {"Quiz 1": 75, "Quiz 2": 71, "Quiz 3": 70}
 *    ];
 *
 *    GSheetsUtils.setRowsData(sheet, dataToWrite);
 *    //Voila!
 *    //Explore optParams to get more advanced applications
 *    //(e.g. headers not in row 1, data writes starting in arbitrary row, etc.)
 *
 * @param {Sheet} sheet - the Sheet Object where the data will be written
 * @param {Object[]} objects - an Array of Objects, each of which contains data for a row
 * @param {Object} [optParams] - optional config parameters
 * @param {Range} [optParams.headersRange=First Row] - a Range of cells where the column headers are defined. This defaults to the entire first row in sheet.
 * @param {rowNum} [optParams.firstDataRowIndex=Row immediately below headersRange] - index of the first row where data should be written. This defaults to the row immediately below the headers.
 */
function setRowsData(sheet, objects, optParams) {
    var params = optParams || {};
    var headersRange = params.headersRange || sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var firstDataRowIndex = params.firstDataRowIndex || headersRange.getRowIndex() + 1;
    var headers = headersRange.getValues()[0];
    var data = [];
    for (var i = 0; i < objects.length; ++i) {
        var values = [];
        for (var j = 0; j < headers.length; ++j) {
            var header = headers[j];

            // If the header is non-empty and the object value is 0...
            if ((header.length > 0)&&(objects[i][header] === 0)&&(!(isNaN(parseInt(objects[i][header]))))) {
                values.push(0);
            }
            // If the header is empty or the object value is empty...
            else if ( (header.length === 0) || (objects[i][header] === '') || (!objects[i][header])) {
                values.push('');
            }
            else {
                values.push(objects[i][header]);
            }
        }
        data.push(values);
    }

    var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(),
      objects.length, headers.length);
    destinationRange.setValues(data);
}

/** appendRowsData adds an array of JSON objects to the end of a sheet with headers that match the JSON keys
 *  @example
 *    //assume sheet contains headers that match the keys of the objects below...
 *
 *    var dataToWrite = [
 *      {"Quiz 1": 93, "Quiz 2": 91, "Quiz 3": 61},
 *      {"Quiz 1": 87, "Quiz 2": 83, "Quiz 3": 63},
 *      {"Quiz 1": 75, "Quiz 2": 71, "Quiz 3": 70}
 *    ];
 *
 *    GSheetsUtils.appendRowsData(sheet, dataToWrite);
 *    //Voila!
 *    //Explore optParams to get more advanced applications
 *    //(e.g. headers not in row 1, data writes starting in arbitrary row, etc.)
 *  @param {Sheet} sheet - the Sheet Object where the data will be written
 *  @param {Object[]} objects - an array of Objects representing a spreadsheet row
 *  @param {Object} [optParams] - optional config parameters
 *  @param {Range} [optParams.headersRange=First row] - a Range of cells where the column headers are defined.
 *  This defaults to the entire first row in sheet.
 */
function appendRowsData(sheet, objects, optParams) {
    var params = optParams || {};
    var nextRow = sheet.getLastRow() + 1;
    params.firstDataRowIndex = nextRow;
    this.setRowsData(sheet, objects, params);
}



/** updateSingleRowInPlace updates key-value pairs in one row of a sheet by using a row number
 * and an optional match key requirement, specified in a params object via a boolean flag an array of column keys.
 * Performance note: Generally efficient if row number approach and matching doesn't fallback.
 * Gets expensive if a row-shift has occurred in the sheet.
 * Assumes that headers in the destination sheet are also the keys of the source data
 * If the match condition fails on the specified, iterates through all destination data and writes each updates to the sheet individually on all matches.
 * If no match requirement is supplied, updates sheet purely based on row number.
 * @example
 *    //assume sheet contains headers that match the keys of the objects below...
 *     var updatedRecord = {"Student ID":"23456", "Quiz 1": 91, "Quiz 2": 98, "Quiz 3": 61};
 *     var rowNum = 5; //row we expect to find the student record in
 *
 *     GSheetsUtils.updateSingleRowInPlace(sheet, updatedRecord, rowNum, {requireMatch: true, matchHeaders: ['Student ID']});
 *     //will go to row 5, look for a match on student ID, and write the updated values if it matches.
 *     //falls back to updateRowsInPlace method if no match is discovered in the row.
 *
 * @param {Sheet} sheet - the GAS Sheet Object where the data will be written.
 * @param {Object} sourceObject - a single Object containing the key-value pairs to be used
 *   for updating matching records
 * @param {Number} rowNum - the row number of the record you are updating. Required.
 * @param {Object} [optParams] - optional config parameters
 * @param {Object} [optParams.requireMatch=false] - T/F whether to double check that the record matches based on match criteria
 * @param {Object} [optParams.matchHeaders] - a string or array of string values representing the
 *   shared headers to match on
 * @param {Range} [optParams.headersRange=First row] - used to specify which range in the
 *   destination sheet holds the headers.  Expects a GAS Range object.
 */
function updateSingleRowInPlace(sheet, sourceObject, rowNum, params) {
    if (!rowNum) {
        throw "rowNum parameter missing";
    }
    var lastCol = sheet.getLastColumn();
    params = params || {};
    var headersRange = params.headersRange || (lastCol > 0) ? sheet.getRange(1, 1, 1, lastCol) : null;
    var headersIndex = params.headersRange ? params.headersRange.getRow() : 1;
    var requireMatch = params.requireMatch || false;
    var dataRange = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
    var existingDataRow = this.getRowsData(sheet, {dataRange: dataRange, headersIndex: headersIndex})[0];
    var status = {
        recordsUpdated: 0,
        errors: 0
    };
    if (requireMatch) {
        var matchHeaders = params.matchHeaders;
        if (!matchHeaders || !matchHeaders.length) {
            throw "Missing matchHeaders param.  Required if requiring record match."
        }
        var mustMatch = '';
        var thisMatch = '';
        for (var i in matchHeaders) {
            if (!existingDataRow[matchHeaders[i]]) {
                throw "Existing data row is missing a required match parameter."
            } else {
                thisMatch += existingDataRow[matchHeaders[i]];
            }
            if (!sourceObject[matchHeaders[i]]) {
                throw "Source data object is missing a required match parameter."
            } else {
                mustMatch += sourceObject[matchHeaders[i]];
            }
        }
        if (thisMatch === mustMatch) {
            try {
                this.setRowsData(sheet, [sourceObject], { headersRange: headersRange, firstDataRowIndex: rowNum });
                status.recordsUpdated++;
            } catch(err) {
                status.errors++;
            }
            return status;
        } else {
            return this.updateRowsInPlace(sheet, [sourceObject], matchHeaders, {headersRange: headersRange});
        }
    } else if (rowNum) {
        try {
            this.setRowsData(sheet, [sourceObject], { headersRange: headersRange, firstDataRowIndex: rowNum });
            status.recordsUpdated++;
        } catch(err) {
            status.errors++;
        }
        return status;
    }
}


/** updateRowsInPlace updates key-value pairs in one or more rows of a sheet by looking for a
 * match based on a shared header (or headers)
 * and makes updates to matching records.  Performance note: This is an expensive method
 * Assumes that headers in the destination sheet are also the keys of the source data
 * Iterates through all destination data and writes each updates to the sheet individually.
 *  @example
 *
 *   //assume sheet contains headers that match the keys of the objects below...
 *     var updatedRecord = {"Student ID":"23456", "Quiz 1": 91, "Quiz 2": 98, "Quiz 3": 61};
 *
 *     //row number(s) of matching records are not known in advance.
 *     //assume students are supposed to be unique in the sheet
 *
 *     var status = GSheetsUtils.updateRowsInPlace(sheet, updatedRecord, ['Student ID'], {requireUnique: true});
 *     //will loop throush Sheet looking for a match on student ID, and write the updated values if it matches a row.
 *     //Throws an error if more than one match is found based on Student ID
 *
 *   @param {Sheet} sheet - the GAS Sheet Object where the data will be written.
 *   @param {Object} sourceObject - a single Object containing the key-value pairs to be used
 *   for updating matching records
 *   @param {string[]} matchHeaders - a string or array of string values representing the
 *   shared headers to match on
 *   @param {Object} [optParams] - optional config parameters
 *   @param {Range} [optParams.headersRange=First row] - used to specify which range in the
 *   destination sheet holds the headers.  Expects a GAS Range object.
 *   @param {Boolean} [optParams.requireUnique=false] - T/F whether or not to throw an error
 *   if query of destination data results in multiple matches
 *   @returns {updateStatus} updateStatus
 */
function updateRowsInPlace(sheet, sourceObject, matchHeaders, optParams) {
    var params = optParams || {};
    var matchRows = [];
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    matchHeaders = (typeof matchHeaders === 'string') ? [matchHeaders] : matchHeaders;
    var headersRange = params.headersRange || (lastCol > 0) ? sheet.getRange(1, 1, 1, lastCol) : null;
    var headersIndex = params.headersRange ? params.headersRange.getRow() : 1;
    var destData = [];
    if (lastRow - headersIndex > 1) {
        var dataRange = sheet.getRange(headersIndex + 1, 1, lastRow - headersIndex, lastCol);
        destData = lastRow > 1 ? this.getRowsData(sheet, dataRange, {headersIndex: headersIndex}) : [];
    }
    if (!matchHeaders) {
        throw new Error("updateRowsInPlace - missing matchHeaders argument");
    }
    var sourceRecordJoinVal = '';
    for (var j=0; j<matchHeaders.length; j++) {
        if (sourceObject[matchHeaders[j]]) {
            sourceRecordJoinVal += sourceObject[matchHeaders[j]];
        } else {
            throw new Error("updateRowsInPlace - sourceObject missing match key");
        }
    }
    for (var i=0; i<destData.length; i++) {
        var destRecordJoinVal = '';
        for (j=0; j<matchHeaders.length; j++) {
            if (destData[i][matchHeaders[j]]) {
                destRecordJoinVal += destData[i][matchHeaders[j]];
            }
        }
        if (sourceRecordJoinVal === destRecordJoinVal) {
            matchRows.push({rowNum: i+headersIndex+1, arrayIndex: i});
        }
    }
    if ((params.requireUnique)&&(matchRows.length > 0)) {
        throw new Error("Multiple records in destination match based on your match headers.");
    }
    var status = {
        recordsUpdated: 0,
        errors: 0
    };
    for (i=0; i<matchRows.length; i++) {
        assign_(destData[matchRows[i].arrayIndex], sourceObject);
        try {
            this.setRowsData(sheet, [destData[matchRows[i].arrayIndex]], { headersRange: headersRange, firstDataRowIndex: matchRows[i].rowNum });
            status.recordsUpdated++;
        } catch(err) {
            status.errors++;
        }
    }
    return status;
}


/** updateRowsData updates a sheet containing existing records with objects defined in the objects Array.
 * Assumes that headers in the destination sheet are also the keys of the source data
 * Requires that source and destination records are unique based on the provided match key or key combination.
 * Optionally inserts records from source if not found in destination.
 * Optionally removes records in destination if not found in source.
 * @example
 *
 *   //assume sheet contains headers that match the keys of the objects below...
 *     var updatedRecords = [
 *     {"Student ID":"23456", "Quiz 1": 91, "Quiz 2": 98, "Quiz 3": 61},
 *     {"Student ID":"23457", "Quiz 1": 80, "Quiz 2": 58, "Quiz 3": 90}, //this student doesn't currently exist in sheet
 *     {"Student ID":"23458", "Quiz 1": 71, "Quiz 2": 90, "Quiz 3": 65}
 *     ]
 *
 *     var status = GSheetsUtils.updateRowsData(sheet, updatedRecords, ['Student ID'], {upsertNewRecords: true, removeAndArchiveNonMatchingRecords: true});
 *     //will update each matching record in the sheet (match based on Student ID) by clearing and rewriting the contents of the entire sheet
 *     //Based on parameter values...
 *     //Student 23457 will be upserted (e.g. added) to the Sheet.
 *     //Student 23459 will be removed from the sheet and archived.
 *
 * @param {Sheet} sheet - the Sheet Object where the data will be written
 * @param {Object[]} objects - an Array of Objects, each of which contains data for a row
 * @param {string[]} matchFields - a String or an Array of column header values that represent the unique key for
 * updating existing records
 * @param {Object} [optParams] - optional config parameters
 * @param {Boolean} [optParams.upsertNewRecords=false] - a Boolean, whether or not to insert records from source
 * if not found in destination
 * @param {Boolean} [optParams.removeAndArchiveNonMatchingRecords=false] - a Boolean, whether or not to remove
 * data in non-matching rows and archive it in a separate "tabName_archive" tab.
 * @param {Range} [optParams.headersRange=First row] - a Range of cells where the column headers are defined.
 * This defaults to the entire first row in sheet.
 * @returns {updateStatus} updateStatus
 */
function updateRowsData(sheet, objects, matchFields, optParams) {
    var params = optParams || {};
    var lastRow = sheet.getLastRow();
    matchFields = (typeof matchFields === 'string') ? [matchFields] : matchFields;
    var headersIndex = params.headersRange ? params.headersRange.getRow() : 1;
    var existingData = [];
    var dataRange;
    if ((lastRow - headersIndex) > 0) {
        dataRange = sheet.getRange(headersIndex+1, 1, lastRow - headersIndex, sheet.getLastColumn());
        existingData = lastRow>1 ? this.getRowsData(sheet, dataRange, { headersIndex: headersIndex }) : [];
    }

    if (!matchFields) {
        throw new Error("updateRowsData - missing matchFields string or array");
    }

    var status = {
        recordsUpdated: 0,
        recordsInserted: 0,
        recordsArchived: 0,
        error: 0
    };

    //build a hash of source objects keyed to matchFields keys(s)
    var updateHash = {};
    var i, j;
    var thisKey;
    for (i=0; i<objects.length; i++) {
        thisKey = '';
        for (j=0; j<matchFields.length; j++) {
            if (objects[i][matchFields[j]] !== undefined) {
                thisKey += objects[i][matchFields[j]];
            } else {
                throw new Error('Object ' + JSON.stringify(objects[i]) + 'in update object array missing value for match field ' + matchFields[j]);
            }
        }
        if (updateHash[thisKey]) {
            throw new Error('updateRowsData - Duplicate record(s) in source data with key ' + thisKey);
        } else {
            updateHash[thisKey] = objects[i];
        }
    }

    //build a hash of destination objects keyed to matchFields key(s)
    var destHash = {};
    for (i=0; i<existingData.length; i++) {
        thisKey = '';
        for (j=0; j<matchFields.length; j++) {
            if (existingData[i][matchFields[j]] !== undefined) {
                thisKey += existingData[i][matchFields[j]];
            } else {
                throw new Error('Object ' + JSON.stringify(objects[i]) + 'in destination spreadsheet object array missing value for match field ' + matchFields[j]);
            }
        }
        //throw error if duplicates in destination
        if (destHash[thisKey]) {
            throw new Error('updateRowsData - Duplicate record(s) in destination with key ' + thisKey);
        } else {
            destHash[thisKey] = existingData[i];
        }
    }


    var destObjects = [];
    var unmatchedDestObjects = [];

    //handle updates and/or appends to existing records
    for (var key in destHash) {
        if (updateHash[key]) {
            assign_(destHash[key], updateHash[key]);
            destObjects.push(destHash[key]);
            status.recordsUpdated++;
        } else if (!params.removeAndArchiveNonMatchingRecords) { //and preserve unmatched records
            destObjects.push(destHash[key]);
        } else {
            unmatchedDestObjects.push(destHash[key]); //or not
        }
    }

    //optionally append unmatched records from source into destination
    if (params.upsertNewRecords) {
        for (key in updateHash) {
            if (!destHash[key]) {
                destObjects.push(updateHash[key]);
                status.recordsInserted++;
            }
        }
    }

    //handle archiving of unmatched records, do this first, in case of error, as least destructive option
    if (params.removeAndArchiveNonMatchingRecords && unmatchedDestObjects.length) {
        var sheetName = sheet.getName();
        var ss = sheet.getParent();
        var archiveSheet = ss.getSheetByName(sheetName + "_archive");

        if (!archiveSheet) {
            ss.insertSheet(sheetName + "_archive");
            archiveSheet = ss.getSheetByName(sheetName + "_archive");
        }
        var destHeaders = sheet.getRange(headersIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
        this.getUpsertHeaders(archiveSheet,{'expectedHeaders':destHeaders});
        
        status.recordsArchived = unmatchedDestObjects.length;
        this.appendRowsData(archiveSheet, unmatchedDestObjects);
    }

    if (dataRange) {
        dataRange.clear();
    }

    this.setRowsData(sheet, destObjects, { headersRange: params.headersRange, firstDataRowIndex: 2 });
    return status;
}


/** setMappedRowsData fills in one row of data in a sheet per object defined in an objects Array,
 * with header mappings provided between the source data keys and destination sheet headers
 * For every Column, it checks if data objects define a value for it.
 * @example
 * var headerMappings = {
 *    "First Name": "firstName",
 *    "Last Name": "lastName",
 *    "Status": "updateStatus"
 * }
 *
 * var dataToWrite = [
 *   {firstName:'Joe', lastName: 'Blow', status:'Single'},
 *   {firstName':'Jane', lastName:'Doe', status:'Deceased'},
 *   {firstName:'Captain',lastName:'Cook',status:'Lost at sea'}
 * ];
 *
 * //Sheet contains headers 'First Name','Last Name','Status'
 *
 * GSheetsUtils.setMappedRowsData(sheet, dataToWrite, headerMappings);
 * //writes data to sheet based on header mappings...
 *
 * @param {Sheet} sheet - the Sheet Object where the data will be written
 * @param {Object[]} objects - an Array of Objects, each of which contains data for a row
 * @param {Object} headerMappings - a plain Javascript object whose keys are destination headers, and values are source headers.
 * @param {Object} [optParams]
 * @param {Range} [optParams.headersRange] - a Range of cells where the column headers are defined. This defaults to the entire first row in sheet.
 * @param {rowNum} [optParams.firstDataRowIndex] - index of the first row where data should be written. This defaults to the row immediately below the headers.
 */
function setMappedRowsData(sheet, objects, headerMappings, optParams) {
    var params = optParams || {};
    var headersRange = params.headersRange || sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var firstDataRowIndex = params.firstDataRowIndex || headersRange.getRowIndex() + 1;
    var destHeaders = headersRange.getValues()[0];
    var data = [];
    for (var i = 0; i < objects.length; ++i) {
        var values = [];
        for (var j = 0; j < destHeaders.length; ++j) {
            var header = destHeaders[j];

            // If the header is non-empty and the object value is 0...
            if ((header.length > 0)&&(objects[i][headerMappings[header]] === 0)&&(!(isNaN(parseInt(objects[i][headerMappings[header]]))))) {
                values.push(0);
            }
            // If the header is empty or the object value is empty...
            else if ((header.length === 0) || (objects[i][headerMappings[header]] === '') || (!objects[i][headerMappings[header]])) {
                values.push('');
            }
            else {
                values.push(objects[i][headerMappings[header]]);
            }
        }
        data.push(values);
    }

    var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(),
        objects.length, destHeaders.length);
    destinationRange.setValues(data);
}


/** appendRowsData adds an array of JSON objects to the end of a sheet with headers that match the JSON keys
 * @example
 * var headerMappings = {
 *    "First Name": "firstName",
 *    "Last Name": "lastName",
 *    "Status": "updateStatus"
 * }
 *
 * var dataToWrite = [
 *   {firstName:'Joe', lastName: 'Blow', status:'Single'},
 *   {firstName':'Jane', lastName:'Doe', status:'Deceased'},
 *   {firstName:'Captain',lastName:'Cook',status:'Lost at sea'}
 * ];
 *
 * //Sheet contains headers 'First Name','Last Name','Status'
 *
 * GSheetsUtils.appendMappedRowsData(sheet, dataToWrite, headerMappings);
 * //appends data to last data row of sheet based on header mappings...
 *
 *   @param {Sheet} sheet - the Sheet Object where the data will be written
 *   @param {Object[]} objects - an array of Objects representing spreadsheet rows, where keys may be different than those in the destination spreadsheet
 *   @param {Object} headerMappings - a plain Javascript object whose keys are destination headers, and values are source headers.
 *   @param {Object} [optParams] - optional config parameters
 *   @param {Range} [optParams.headersRange] - a Range of cells where the column headers are defined. This defaults to the entire first row in sheet.
 */
function appendMappedRowsData(sheet, objects, headerMappings, optParams) {
    var params = optParams || {};
    var nextRow = sheet.getLastRow() + 1;
    params.firstDataRowIndex = nextRow;
    this.setMappedRowsData(sheet, objects, headerMappings, params);
}


/** updateMappedRecordsInPlace updates key-value pairs in one or more records by looking for a match based on mapped join keys
 * Identical to updateRecordsInPlace but using header mappings
 * and makes updates to based on header mappings.  Performance note: This is an expensive method.
 * Iterates through all destination data and writes each updates to the sheet individually.
 * Arguments:
 *   @param {Sheet} sheet - the GAS Sheet Object where the data will be written.
 *   @param {Object} sourceObject - a single Object containing the key-value pairs to be used for updating matching records
 *   @param {Object} mappedMatchHeaders - to be used for determining matching records, a plain object that maps one or more keys from the destination sheet onto source data keys
 *   @param {Object} mappedUpdateHeaders - to be used for determining the update mask to be applied, and to specify mappings, a plain object that maps keys from destination sheet onto source keys
 *   @param {Object} [optParams] - optional config parameters
 *   @param {Range} [optParams.headersRange] - used to specify which range in the destionation sheet holds the headers.  Expects a GAS Range object.
 *   @param {Boolean} [optParams.requireUnique] - T/F whether or not to throw an error if query of destination data results in multiple matches
 *   @returns {updateStatus} updateStatus
 */
function updateMappedRowsInPlace(sheet, sourceObject, mappedMatchHeaders, mappedUpdateHeaders, optParams) {
    var params = optParams || {};
    var matchRows = [];
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headersRange = params.headersRange || (lastCol > 0) ? sheet.getRange(1, 1, 1, lastCol) : null;
    var headersIndex = params.headersRange ? params.headersRange.getRow() : 1;
    var destData = [];
    if (lastRow - headersIndex > 1) {
        var dataRange = sheet.getRange(headersIndex + 1, 1, lastRow - headersIndex, lastCol);
        destData = lastRow > 1 ? this.getRowsData(sheet, dataRange, {headersIndex: headersIndex}) : [];
    }
    if (!mappedMatchHeaders) {
        throw new Error("updateMatchingRecordsInPlace - missing mappedJoinKeys array");
    }
    var sourceRecordJoinVal = '';
    if (typeof mappedMatchHeaders !== 'object') {
        throw new Error("invalid mappedJoinKeys input");
    }
    for (var destHeader in mappedMatchHeaders) {
        if (mappedMatchHeaders.hasOwnProperty(destHeader)) {
            sourceRecordJoinVal += sourceObject[mappedMatchHeaders[destHeader]];
        }
    }
    for (var i=0; i<destData.length; i++) {
        var destRecordJoinVal = '';
        for (destHeader in mappedMatchHeaders) {
            if (mappedMatchHeaders.hasOwnProperty(destHeader)) {
                destRecordJoinVal += destData[i][destHeader];
            }
        }
        if (sourceRecordJoinVal === destRecordJoinVal) {
            matchRows.push({rowNum: i+headersIndex+1, arrayIndex: i});
        }
    }
    if ((params.requireUnique)&&(matchRows.length > 0)) {
        throw new Error("Multiple records in destination match based on your mapped join keys.");
    }
    var status = {
        recordsUpdated: 0,
        errors: 0
    };
    var invertedMappings = {};
    for (destHeader in mappedUpdateHeaders) {
        if (mappedUpdateHeaders.hasOwnProperty(destHeader)) {
            invertedMappings[mappedUpdateHeaders[destHeader]] = destHeader;
        }
    }
    for (i=0; i<matchRows.length; i++) {
        for (var key in sourceObject) {
            if (sourceObject.hasOwnProperty(key)) {
                destData[matchRows[i].arrayIndex][invertedMappings[key]] = sourceObject[key];
            }
        }
        try {
            this.setRowsData(sheet, [destData[matchRows[i].arrayIndex]], { headersRange: headersRange, firstDataRowIndex: matchRows[i].rowNum });
            status.recordsUpdated++;
        } catch(err) {
            status.errors++;
        }
    }
    return status;
}


/** updateMappedRowsData updates a sheet containing existing records with objects defined in the objects Array.
 * Identical to updateMappedRowsData but using header mappings
 * Requires that both source and destination data be unique based on mapped match headers
 * Optionally inserts records from source if not found in destination.
 * Optionally removes records in destination if not found in source.
 * Requires that destination data have a unique key or key combination.
 *  @param {Sheet} sheet - the Sheet object where the data will be written
 *  @param {Object[]} objects - an Array of Objects, each of which contains data for a row
 *  @param {Object} mappedMatchHeaders - to be used for determining matching records, a plain object that maps one or more keys from the destination sheet onto source data keys
 *  @param {Object} mappedUpdateHeaders - to be used for determining the update mask to be applied, and to specify mappings, a plain object that maps keys from destination sheet onto source keys
 *  @param {Object} optParams - optional config parameters:
 *  @param {Boolean} optParams.upsertNewRecords: a Boolean, whether or not to insert records from source if not found in destination
 *  @param {Boolean} optParams.removeAndArchiveNonMatchingRecords: a Boolean, whether or not to remove data in non-matching rows and archive it in a separate "tabName_archive" tab.
 *  @param {Range} optParams.headersRange: a Range of cells where the column headers are defined. This
 *         defaults to the entire first row in sheet.
 *  @returns {updateStatus} updateStatus
 */
function updateMappedRowsData (sheet, objects, mappedMatchHeaders, mappedUpdateHeaders, optParams) {
    var params = optParams || {};
    var lastRow = sheet.getLastRow();
    var headersIndex = params.headersRange ? params.headersRange.getRow() : 1;
    var existingData = [];
    var dataRange;
    if ((lastRow - headersIndex) > 0) {
        dataRange = sheet.getRange(headersIndex+1, 1, lastRow - headersIndex, sheet.getLastColumn());
        existingData = lastRow>1 ? this.getRowsData(sheet, dataRange, {headersIndex: headersIndex} ) : [];
    }

    if (!mappedMatchHeaders) {
        throw new Error("updateMappedRowsData - missing mappedMatchHeaders object");
    }

    if (!mappedUpdateHeaders) {
        throw new Error("updateMappedRowsData - missing mappedUpdateHeaders object");
    }

    var status = {
        recordsUpdated: 0,
        recordsInserted: 0,
        recordsArchived: 0,
        error: 0
    };

    //build a hash of source objects keyed to mappedMatchHeaders keys(s)
    var updateHash = {};
    var thisKey;
    var destHeader;
    for (var i=0; i<objects.length; i++) {
        thisKey = '';
        for (destHeader in mappedMatchHeaders) {
            if (mappedMatchHeaders.hasOwnProperty(destHeader)) {
                if (objects[i][mappedMatchHeaders[destHeader]] !== undefined) {
                    thisKey += objects[i][mappedMatchHeaders[destHeader]];
                } else {
                    throw new Error('Object ' + JSON.stringify(objects[i]) + 'in update object array missing value for match field ' + mappedMatchHeaders[destHeader]);
                }
            }
        } if (updateHash[thisKey]) {
            throw new Error("updateMappedRowsData - Duplicate records in source data with key " + thisKey);
        } else {
            updateHash[thisKey] = objects[i];
        }

    }

    //build a hash of destination objects keyed to matchFields key(s)
    var destHash = {};
    for (i=0; i<existingData.length; i++) {
        thisKey = '';
        for (destHeader in mappedMatchHeaders) {
            if (mappedMatchHeaders.hasOwnProperty(destHeader)) {
                if (existingData[i][destHeader] !== undefined) {
                    thisKey += existingData[i][destHeader];
                } else {
                    throw new Error('Object ' + JSON.stringify(objects[i]) + 'in destination spreadsheet object array missing value for match field ' + destHeader);
                }
            }
        }
        //throw error if duplicates in destination
        if (destHash[thisKey]) {
            throw new Error('updateMappedRowsData - Duplicate records in destination with key ' + thisKey);
        } else {
            destHash[thisKey] = existingData[i];
        }
    }



    var destObjects = [];
    var unmatchedDestObjects = [];

    //handle updates and/or appends to existing records
    for (var key in destHash) {
        if (destHash.hasOwnProperty(key)) {
            if (updateHash[key]) {
                for (destHeader in mappedUpdateHeaders) {
                    if (mappedUpdateHeaders.hasOwnProperty(destHeader)) {
                        destHash[key][destHeader] = updateHash[key][mappedUpdateHeaders[destHeader]];
                    }
                }
                destObjects.push(destHash[key]);
                status.recordsUpdated++;
            } else if (!params.removeAndArchiveNonMatchingRecords) { //and preserve unmatched records
                destObjects.push(destHash[key]);
            } else {
                unmatchedDestObjects.push(destHash[key]); //or not
            }
        }
    }

    //optionally append unmatched records from source into destination
    if (params.upsertNewRecords) {
        for (key in updateHash) {
            if (updateHash.hasOwnProperty(key)) {
                if (!destHash[key]) {
                    var thisMappedObj = {};
                    for (destHeader in mappedUpdateHeaders) {
                        if (mappedUpdateHeaders.hasOwnProperty(destHeader)) {
                            thisMappedObj[destHeader] = updateHash[key][mappedUpdateHeaders[destHeader]];
                        }
                    }
                    status.recordsInserted++;
                    destObjects.push(thisMappedObj);
                }
            }

        }
    }

    //handle archiving of unmatched records, do this first, in case of error, as least destructive option
    if (params.removeAndArchiveNonMatchingRecords && unmatchedDestObjects.length) {
        var sheetName = sheet.getName();
        var ss = sheet.getParent();
        var archiveSheet = ss.getSheetByName(sheetName + "_archive");
        if (!archiveSheet) {
            ss.insertSheet(sheetName + "_archive");
            archiveSheet = ss.getSheetByName(sheetName + "_archive");
        }
        var destHeaders = sheet.getRange(headersIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
        this.getUpsertHeaders(archiveSheet,{'expectedHeaders':destHeaders});
        status.recordsArchived = unmatchedDestObjects.length;
        this.appendRowsData(archiveSheet, unmatchedDestObjects);
    }

    if (dataRange) {
        dataRange.clear();
    }

    this.setRowsData(sheet, destObjects, {headersRange: params.headersRange, firstDataRowIndex: 2});
}


/**
 *  @typedef {Object} Spreadsheet
 */

/**
 *  @typedef {Object} Sheet
 */

/**
 *  @typedef {Object} Range
 */

/**
 *  @typedef {number} rowNum
 */


/**
 * @typedef updateStatus
 * @type {Object}
 * @property {number} [recordsUpdated] - number of records updated
 * @property {number} [recordsInserted] - number of records inserted
 * @property {number} [recordsArchived] - number of records archived
 * @property {number} [error] - number of errors
 */

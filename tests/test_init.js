const localGass = require('gas-local');  // wrapper to calling GAS code from our local machine
const assert = require('assert');  // sane assert statements for Google Apps yay!
const sinon = require('sinon');

function Range(data) {  // closure
	return {
		getValues: function() { return data; },
		getLastRow: function() { return data.length; },
		getLastColumn: function() { return data[0].length; },
		getColumn: function() { return 1; },  // assume it's a very simple mocked sheet
	}
}


// This kind of mocking is not used for this example,
// left for reference
let defMock = localGass.globalMockDefault;
let stubbedSpreadsheetApp =sinon.stub()
let customMock = { 
    SpreadsheetApp: {},
     __proto__: defMock 
  };

//pass it to require
// TODO: Figure out how this actually works.
let glib = localGass.require('../src', customMock);

(function testGetRowsData() {
	// Test all code paths and assumptions present in getRowsData
	let mockedSheet, mockedParams, actual, expected;

	// Test that empty array is returned when there is no room for data in sheet
	// (i.e. it only contains enough rows for headers or no rows at all)
	expected = [];

	var testSheetData = {
		dataRange: Range([['Data1', 'Data2'], ['Data3', 'Data4']]),
		headersRange: Range([['Column1', 'Column2']]),		
	};

	mockedSheet = {
		getLastRow: sinon.stub().returns(1),
	};
	actual = glib.getRowsData(mockedSheet);
	assert.deepEqual(actual, expected);

	mockedSheet = {
		getLastRow: sinon.stub().returns(3),
		getLastColumn: sinon.stub().returns(2),
		getLastRow: sinon.stub().returns(2),
	};

	mockedParams = {
		columnHeadersRowIndex: 2,
	};
	actual = glib.getRowsData(mockedSheet, mockedParams);
	assert.deepEqual(actual, expected)
	// End expected = []

	// Test raw
	expected = [
		{ Column1: 'Data1', Column2: 'Data2' },
	    { Column1: 'Data3', Column2: 'Data4' },
	];

	mockedParams = {
		dataRange: Range([['Data1', 'Data2'], ['Data3', 'Data4']]),
		headersRange: Range([['Column1', 'Column2']]),
	};

	mockedSheet = {
		getLastRow: sinon.stub().returns(3),
		getLastColumn: sinon.stub().returns(2),
		getRange: sinon.stub(),
	};
	mockedSheet.getRange.withArgs(2, 1, 2, 2).returns(mockedParams.dataRange);  // data
	mockedSheet.getRange.withArgs(1, 1, 1, 2).returns(mockedParams.headersRange);  // headers

	actual = glib.getRowsData(mockedSheet);
	assert.deepEqual(actual, expected);
	actual = glib.getRowsData(mockedSheet, mockedParams);
	assert.deepEqual(actual, expected);
})();





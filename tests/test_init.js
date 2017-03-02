const gas = require('gas-local');  // wrapper to calling GAS code from our local machine
const assert = require('assert');  // sane assert statements for Google Apps yay!

function Range(data) {  // closure
	return {
		getValues: function() { return data; },
		getLastRow: function() { return data.length; },
		getLastColumn: function() { return data[0].length; },
		getColumn: function() { return 1; },  // assume it's a very simple mocked sheet
	}
}

const params = {
	dataRange: Range([['Data1', 'Data2'], ['Data3', 'Data4']]),
	headersRange: Range([['Column1', 'Column2']]),
};
const expected = [
	{ Column1: 'Data1', Column2: 'Data2' },
    { Column1: 'Data3', Column2: 'Data4' },
];

// This kind of mocking is not used for this example,
// left for reference
let defMock = gas.globalMockDefault;
let customMock = { 
    SpreadsheetApp: {
    	mocked: function() {
    		return {
		    	getLastRow: function () { return 3; },
		    	getLastColumn: function() { return 2; },
		    	getRange: function() { return mockedRange; },
		    }
	    }
    },
     __proto__: defMock 
  };

//pass it to require
// TODO: Figure out how this actually works.
let glib = gas.require('../src', customMock);

let mockedSheet = {
	getLastRow: function() { return 3; },  // get us past short circuit code in getRowsData
	getLastColumn: function() { return 2; },
	getRange: function() { return params.mockedData; },
}

//call some function from your app script library working with MailApp 
let actual = glib.getRowsData(mockedSheet, params);

assert.deepEqual(actual, expected);

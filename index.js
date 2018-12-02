///
/* @Author: Alexander P. Cerutti
 * @Description: This script is made to work on Google Apps Script Editor.
 * It uses JS ES5
 */
///

var activeSS = SpreadsheetApp.getActive();

var LocalizedStrings = {
	monthsFull: ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"],
	monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
	columnsName: ["Nome", "Costo", "Data", "Note"],
	totalDescriptor: "Totale mese",
	totalWithPastMonthsDescriptor: "Tot. con mesi precedenti",
};

var Colors = {
	// these are the color for totalWithPastMonthsDescriptor
	positivePastMonthAmount: "#6aa84f", // which will result as negative number
	positivePastMonthAmountText: "#FFF",
	negativePastMonthAmount: "#cc0000", // which will result as positive number
	negativePastMonthAmountText: "#FFF",
	// Total separator rows
	totalSeparator: "#666",
	totalSeparatorText: "#FFF",
};

/*  Events */

/**
 * Google Sheet event onOpen
 * is responsible to check if the current month sheet is available
 * or, otherwise, creates it.
 */

function onOpen() {
	var actualMonth = getMonthName();

	/** DEVELOPMENT ONLY **/
	//activeSS.deleteSheet(activeSS.getSheetByName(actualMonth));
	/** DEVELOPMENT ONLY **/

	// Checking if there is a Sheet with name corresponding to the
	// one associated with this month name

	if (activeSS.getSheetByName(actualMonth) === null) {
		activeSS.insertSheet(actualMonth, activeSS.getNumSheets());
		var activeSheet = activeSS.getSheetByName(actualMonth);

		// Setting cells
		activeSheet.setRowHeight(1, 35);
		activeSheet.setColumnWidths(1, 3, 185);
		activeSheet.setColumnWidth(4, 200);

		// ADColumns and Rows is the header row A-D
		// headerRow is the header row A-Z

		var ADcolumns = activeSheet.getRange("A:D");
		var ADrows = activeSheet.getRange("A1:D1");
		var headerRow = activeSheet.getRange("A1:Z1");

		// Setting header row colors
		headerRow.setBackground(Colors.totalSeparator);
		headerRow.setFontColor(Colors.totalSeparatorText);

		ADrows.setValues([LocalizedStrings.columnsName]);
		ADcolumns.setVerticalAlignment("middle");
		ADcolumns.setHorizontalAlignment("center");

		activeSS.setFrozenRows(1);
		activeSS.setActiveSheet(activeSheet);

		// Creating first row
		addRow();

		// Creating the footer
		var footerRow1 = activeSheet.getRange("A3:Z3");
		footerRow1.setBackground("#666");
		activeSheet.setRowHeight(3, 7);

		// Initially, B2 is the first small row.
		// By adding rows, the second 'B2' will increase itself and change
		var monthResultRow = activeSheet.getRange("A4:D4");
		monthResultRow.setValues([[LocalizedStrings.totalDescriptor, "=SUM($B$2:$B3)", "", ""]]);
		var chainResultRow = activeSheet.getRange("A5:D5");

		var footerRowRange;
		var footerRowNumber;
		if (activeSS.getSheets().length == 1) {
			footerRowNumber = 5;
			footerRowRange = "A" + footerRowNumber + ":Z" + footerRowNumber;
		} else {
			footerRowNumber = 6
			footerRowRange = "A" + footerRowNumber + ":Z" + footerRowNumber;

			chainResultRow.setValues([[LocalizedStrings.totalWithPastMonthsDescriptor, "=" + getPreviousMonthTotalRange() + "+B4", "", ""]]);

			/*
			 * inverted conditional rules
			 * since we are talking about costs and we are assuming to
			 * all the expenses will be positive
			 * while all the entrances will be negative. By doing this
			 * We will see in total rows, a positive excess of money from the past
			 * month with a green color and a minus sign trailing.
			 * Vice-versa for money deficit.
			 */

			var positiveTotalAmount = SpreadsheetApp.newConditionalFormatRule()
				.whenNumberLessThan(0)
				.setBackground(Colors.positivePastMonthAmount)
				.setFontColor(Colors.positivePastMonthAmountText)
				.setRanges([chainResultRow])
				.build();

			var negativeTotalAmount = SpreadsheetApp.newConditionalFormatRule()
				.whenNumberGreaterThan(0)
				.setBackground(Colors.negativePastMonthAmount)
				.setFontColor(Colors.negativePastMonthAmountText)
				.setRanges([chainResultRow])
				.build();

			var rules = activeSheet.getConditionalFormatRules();
			rules.push(positiveTotalAmount, negativeTotalAmount);
			activeSheet.setConditionalFormatRules(rules);
		}

		var footerRow2 = activeSheet.getRange(footerRowRange);
		footerRow2.setBackground(Colors.totalSeparator);
		activeSheet.setRowHeight(footerRowNumber, 7);

		activeSheet.getRange("B2:B6").setNumberFormat("€ ##0.00##")
	}
}

/**
 * Google Event onEdit
 * This is responsible for creating a new row before total when
 * The edited cell is the last one of the expenses rows
 * @param {*} event
 */

function onEdit(event) {
	if (event.range.getRow() == getLastRowRelative()) {
		var month = getMonthName();
		addRow(month);
	}
}


/* Support functions */

/**
 * Gets the range to fetch the previous month total in runtime
 */

function getPreviousMonthTotalRange() {
	var previousName = getMonthName(-1);

	var activeSheet = activeSS.getSheetByName(previousName);
	var maxRows = activeSheet.getLastRow();
	var result = "'" + previousName + "'!" + activeSheet.getRange(maxRows, 2).getA1Notation();
	return result;
}

function getMonthName(override) {
	var month = LocalizedStrings.monthsFull[(new Date()).getMonth() + (override || 0)];
	var year = (new Date()).getFullYear();

	return month + " " + year;
}

/**
 * Finds the last row relative to the end (footer total rows)
 */

function getLastRowRelative() {
	// last row is 2 rows before total

	var relative = activeSS.getActiveSheet().getLastRow();

	if (activeSS.getActiveSheet().getLastRow() == 1) {
		// this is needed when I'm creating the first row
		return relative;
	}

	return relative - (activeSS.getActiveSheet().getIndex() == 1 ? 2 : 3);
}

function parsedDate(keepAmericanFormat) {
	keepAmericanFormat = keepAmericanFormat || false;
	var months = LocalizedStrings.monthsShort;
	// returning a Date string in American format (mm/dd/yyyy)
	var d = (new Date()).toDateString().split(/\s+/);
	d.shift();
	d[0] = months.indexOf(d[0]) + 1;
	d[1] = Number(d[1]);
	// swapping first two positions (day and month)
	if (!keepAmericanFormat) {
		d[1] = d[0] + d[1];
		d[0] = d[1] - d[0];
		d[1] = d[1] - d[0];
	}
	return d.join("/");
}

function addRow() {
	var activeSheet = activeSS.getActiveSheet();

	var LR = getLastRowRelative(); // Last (relative) Row
	activeSheet.insertRowAfter(LR);

	var newRow = activeSheet.getRange(LR == 1 ? LR + 1 : LR, 1, 1, activeSheet.getMaxColumns());
	newRow.setBackground("#FFF");
	newRow.setVerticalAlignment("middle");

	activeSheet.getRange(LR == 1 ? LR + 1 : LR, 1, 1, activeSheet.getMaxColumns()).setFontColor("#000");

	// New row references
	var NR = {
		date: activeSheet.getRange(LR == 1 ? LR + 1 : LR, 3),
		notes: activeSheet.getRange(LR == 1 ? LR + 1 : LR, 4)
	};

	NR.date.setValue(parsedDate(false));
	NR.notes.setValue("-"); // Notes DEFAULT TO "-"
}

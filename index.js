///
/* @Author: Alexander P. Cerutti
 * @Description: This script is made to work on Google Apps Script Editor.
 * It uses JS ES5
 */
///

var activeSS = SpreadsheetApp.getActive();

//////////////////////////
// ***** SETTINGS ***** //
//////////////////////////

var settings = {
	// if true, a new sheet will be created every new month
	trackMonthly: true,
	// if true, every sheet following the first, will have
	// the row which will count the current month total
	// with the past month one.
	// Real value will be calculated if this and trackMonthly
	// will be true
	pastMonthTotal: true,
	excludeSheets: [], // strings,
	// set false to true to keep american format
	columnsDefaultValue: [null, null, parsedDate.bind(this, false), "-"],
	// Expense Name, Cost Amount, Date, Notes
	columnsName: ["Nome", "Costo", "Data", "Note"], // valid only for new sheets
	columnsWidth: [185, 185, 185, 200], // valid only for new sheets
	rowsHeight: 35, // valid only for new sheets
};

var LocalizedStrings = {
	monthsFull: ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"],
	monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
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

//////////////////////////
// *** END SETTINGS *** //
//////////////////////////

///////////////////////////////////////////

/////////////////////////////////
// *** DO NOT TOUCH BELOW^1 *** //
/////////////////////////////////

// ^1: if you want to use it as is.

//////////////////////////////////////////

/*  Events */

settings.pastMonthTotal = settings.pastMonthTotal && settings.trackMonthly;

/**
 * Google Sheet event onOpen
 * is responsible to check if the current month sheet is available
 * or, otherwise, creates it.
 */

function onOpen() {
	if (settings.trackMonthly) {
		var actualMonth = getMonthName();

		if (activeSS.getSheetByName(actualMonth) !== null) {
			// This deletes the last month sheet
			// You can pass to getMonthName() an arbitrary parameter,
			// a positive number, to create the corresponding sheet
			// e.g. month == "Jen"; getMonthName(3) == "Apr";
			// This is useful also for development

			/** DEVELOPMENT ONLY - DISABLE FOR PRODUCTION **/
			//activeSS.deleteSheet(activeSS.getSheetByName(actualMonth));
			/** DEVELOPMENT ONLY **/
		}

		// Checking if there is a Sheet with name corresponding to the
		// one associated with this month name

		if (activeSS.getSheetByName(actualMonth) === null) {
			activeSS.insertSheet(actualMonth, activeSS.getNumSheets());
			var activeSheet = activeSS.getSheetByName(actualMonth);

			// Setting cells
			activeSheet.setRowHeight(1, settings.rowsHeight);

			for (var i = 1; i <= settings.columnsName.length; i++) {
				activeSheet.setColumnWidth(i, settings.columnsWidth[i - 1]);
			}

			// ADColumns and Rows is the header row A-D
			// headerRow is the header row A-Z

			var ADcolumns = activeSheet.getRange("A:D");
			var ADrows = activeSheet.getRange("A1:D1");
			var headerRow = activeSheet.getRange("A1:Z1");

			// Setting header row colors
			headerRow.setBackground(Colors.totalSeparator);
			headerRow.setFontColor(Colors.totalSeparatorText);

			ADrows.setValues([settings.columnsName]);
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

			if (activeSS.getSheets().length === 1 || !settings.pastMonthTotal) {
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
}

/**
 * Google Event onEdit
 * This is responsible for creating a new row before total when
 * The edited cell is the last one of the expenses rows
 * @param {*} event
 */

function onEdit(event) {
	if (event.range.getRow() == getLastRelativeRow()) {
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

	// e.g. 'December 2018'!B61+B4
	return "'" + previousName + "'!" + activeSheet.getRange(maxRows, 2).getA1Notation();
}

/**
 * Retrieves the current month or the one based on position (backward, negative)
 * @function getMonthName
 * @param {number} position -
 */

function getMonthName(position) {
	position = position || (new Date()).getMonth();
	var parsedPosition;
	var month;
	var year;

	// into the months limits
	if (position < 0) {
		parsedPosition = (new Date()).getMonth() + position;

		if (parsedPosition < 0) {
			// We need to reset the index to the greatest month if still under 0
			parsedPosition = LocalizedStrings.monthsFull.length - 1;
		}

		// we need to set the year based on this month index + position (neg. num).
		// If < 0, we are referring to a previous year month
		year = (new Date().getMonth() + position) < 0 ? (new Date()).getFullYear() - 1 : (new Date()).getFullYear();
	} else {
		// position might be > 11. We want to sanitize it.
		// if it is greater than 11, we calculate the amount of times
		// it fits in it, so we can navigate through years
		parsedPosition = position % 11;
		var yearsForward = 0;

		while (position >= LocalizedStrings.monthsFull.length) {
			position = position - LocalizedStrings.monthsFull.length;
			yearsForward++;
		}

		year = (new Date()).getFullYear() + yearsForward;
	}

	month = LocalizedStrings.monthsFull[parsedPosition];
	return month + " " + year;
}

/**
 * Finds the last row relative to the end (footer total rows)
 */

function getLastRelativeRow() {
	// last row is 2 before real rows total

	var relative = activeSS.getActiveSheet().getLastRow();

	if (activeSS.getActiveSheet().getLastRow() == 1 || !settings.pastMonthTotal) {
		// this is needed when I'm creating the first row
		return relative;
	}

	// Last row - Sheet index: if this is not the first sheet
	// you will have another row.

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

	if (settings.excludeSheets && settings.excludeSheets.length && settings.excludeSheets.indexOf(activeSheet.getName()) === -1) {
		return;
	}

	var lastRow = getLastRelativeRow();
	activeSheet.insertRowAfter(lastRow);

	var newRow = activeSheet.getRange(lastRow == 1 ? lastRow + 1 : lastRow, 1, 1, activeSheet.getMaxColumns());
	newRow.setBackground("#FFF");
	newRow.setVerticalAlignment("middle");

	activeSheet
		.getRange(lastRow == 1 ? lastRow + 1 : lastRow, 1, 1, activeSheet.getMaxColumns())
		.setFontColor("#000");

	for (var i = 1; i <= settings.columnsName.length; i++) {
		if (settings.columnsDefaultValue[i - 1]) {
			var columnRange = activeSheet.getRange(lastRow === 1 ? lastRow + 1 : lastRow, i);

			if (typeof settings.columnsDefaultValue[i - 1] === "function") {
				columnRange.setValue(settings.columnsDefaultValue[i - 1]());
			} else {
				columnRange.setValue(settings.columnsDefaultValue[i - 1]);
			}
		}
	}
}

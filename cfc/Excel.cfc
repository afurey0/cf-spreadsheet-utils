/**
 * This component contains commonly-used methods for dealing with Excel spreadsheets.
 * It is primarily used to export query result sets to XLSX and CSV formats.
 * Use {@link #streamQueryAsText} (CSV) or {@link #streamQueryAsSpreadsheet} (XLSX) for the best performance.
 * Avoid using {@link #queryToSpreadsheet} (and {@link #queryToText}, to a lesser extent) for large result sets, as it can use a lot of memory.
 * @author Alex Furey
 * @version 1.1.4
 * @since 2021-02-10
 * @example {@code new cfc.Excel().streamQueryAsSpreadsheet(qRecords);}
 */
component {

	/**
	 * Converts a coordinate to an Excel cell location string.
	 * @column Index of the column, starting at 1.
	 * @row Index of the row, starting at 1.
	 * @return A cell location string.
	 */
	public string function cellToString(required numeric column, required numeric row) {
		return columnToLetter(arguments.column) & arguments.row;
	}

	/**
	 * Converts a column index to Excel column letters.
	 * @columnIndex The index of the column, starting at 1.
	 * @return A column letter.
	 */
	public string function columnToLetter(required numeric columnIndex) {
		local.n = arguments.columnIndex;
		local.s = "";
		local.i = 0;
		while (local.n gt 0) {
			local.r = local.n % 26;
			if (local.r eq 0) {
				local.s &= "Z";
				local.n = int(local.n / 26) - 1;
			} else {
				local.s &= chr(65 + local.r - 1);
				local.n = int(local.n / 26);
			}
		}
		return reverse(local.s);
	}

	/**
	 * Takes date object, and returns it as an Excel numeric date value.
	 * @value A date object.
	 * @return An Excel numeric date.
	 */
	public numeric function dateToNumber(required date value) {
		return dateDiff("d", createDate(1899, 12, 30), arguments.value) + (hour(arguments.value) / 24) + minute(arguments.value) / (24 * 60) + second(arguments.value) / (24 * 60 * 60);
	}

	/**
	 * Takes a query, and returns it as a CSV string.
	 * @data A query result set.
	 * @fieldDelimiter String to print between fields on a row.
	 * @lineSeparator String to print between rows.
	 * @quoteValues If true, always put quotes around fields.
	 * @preserveStrings If true, put quotes around character fields.
	 * @safeDelimiters If true, strip out the fieldDelimiter and lineSeparator from field values.
	 * @return A CSV string.
	 */
	public any function queryToText(required query data, string fieldDelimiter = ",", string lineSeparator = chr(10), boolean quoteValues = false, boolean preserveStrings = false, boolean safeDelimiters = false) {
		savecontent variable="local.csv" {
			queryToOutput(argumentCollection = arguments);
		}
	}

	/**
	 * Takes a query, and outputs it as CSV data.
	 * Use {@code savecontent} to capture the output of this function, if desired.
	 * @data A query result set.
	 * @fieldDelimiter String to print between fields on a row.
	 * @lineSeparator String to print between rows.
	 * @quoteValues If true, always put quotes around fields.
	 * @preserveStrings If true, put quotes around character fields.
	 * @safeDelimiters If true, strip out the fieldDelimiter and lineSeparator from field values.
	 */
	public void function queryToOutput(required query data, string fieldDelimiter = ",", string lineSeparator = chr(10), boolean quoteValues = false, boolean preserveStrings = false, boolean safeDelimiters = false) {
		local.info = getQueryInfo(arguments.data);
		for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
			if (local.c gt 1) {
				writeOutput(arguments.fieldDelimiter);
			}
			printField(local.info.columnLabels[local.c], arguments.fieldDelimiter, arguments.lineSeparator, arguments.quoteValues, arguments.safeDelimiters);
		}
		for (local.r = 1; local.r lte arguments.data.recordCount; local.r++) {
			writeOutput(arguments.lineSeparator);
			for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
				if (local.c gt 1) {
					writeOutput(arguments.fieldDelimiter);
				}
				local.value = arguments.data[arguments.data.getColumnName(local.c)][local.r];
				if (local.info.columnTypes[local.c] eq "date") {
					local.value = dateTimeFormat(local.value, "yyyy-mm-dd HH:nn:ss");
				} else if (local.info.columnTypes[local.c] eq "char" and arguments.preserveStrings) {
					local.value = "=""" & replace(local.value, """", """""", "all") & """";
				} else {
					local.value = local.value;
				}
				printField(local.value, arguments.fieldDelimiter, arguments.lineSeparator, arguments.quoteValues, arguments.safeDelimiters);
			}
		}
	}

	/**
	 * Takes a query, and returns a spreadsheet object containing a sheet containing the query data, including formatting based on the query column types.
	 * Optionally, pass an existing spreadsheet object to the spreadsheet parameter to have the query be added to the specified sheet in that spreadsheet.
	 * If no spreadsheet is passed in, a new spreadsheet will be created.
	 * The specified sheet name will become the active sheet in the spreadsheet object.
	 * The first row of the spreadsheet is frozen and includes filters.
	 * The first row will contain the column labels of the query, which are the same as the column names by default.
	 * To change the column labels for a query use: {@code variables.myQuery.getMetaData().setColumnLabel(["Label 1", "Label 2", "Label 3"])}.
	 * In character columns, any field value that starts with an equals sign will be treated as a formula.
	 * No formatting is applied to formula cells, unless it starts with {@code =HYPERLINK} in which case the text is formatted to be blue and underlined.
	 * @data A query result set.
	 * @sheetName The name of the sheet.
	 * @spreadsheet An existing spreadsheet to append the new sheet to. Omit to create a new spreadsheet instead.
	 * @return A spreadsheet object.
	 */
	public any function queryToSpreadsheet(required query data, required string sheetName, any spreadsheet) {
		local.info = getQueryInfo(arguments.data);
		if (structKeyExists(arguments, "spreadsheet") and isSpreadsheetObject(arguments.spreadsheet)) {
			local.sheetInfo = spreadsheetInfo(arguments.spreadsheet);
			if (listFind(local.sheetInfo.sheetNames, arguments.sheetName) lte 0) {
				spreadsheetCreateSheet(arguments.spreadsheet, arguments.sheetName);
			}
			spreadsheetSetActiveSheet(arguments.spreadsheet, arguments.sheetName);
		} else {
			arguments.spreadsheet = spreadsheetNew(arguments.sheetName, true);
		}
		for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
			spreadsheetSetCellValue(arguments.spreadsheet, local.info.columnLabels[local.c], 1, local.c);
		}
		spreadsheetFormatRow(arguments.spreadsheet, {
			alignment: "left",
			bold: true,
			dataformat: "@"
		}, 1);
		spreadsheetAddFreezePane(arguments.spreadsheet, 0, 1);
		spreadsheetAddAutoFilter(arguments.spreadsheet, rangeToString(1, 1, local.info.columnsCount, 1));
		for (local.r = 1; local.r lte arguments.data.recordCount; local.r++) {
			for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
				local.value = arguments.data[arguments.data.getColumnName(local.c)][local.r];
				if (local.info.columnTypes[local.c] eq "date") {
					local.valid = isValid("date", local.value);
					spreadsheetSetCellValue(arguments.spreadsheet, local.valid ? dateToNumber(local.value) : local.value, local.r + 1, local.c);
					if (local.valid) {
						spreadsheetFormatCell(arguments.spreadsheet, {
							dataformat: "m/d/y"
						}, local.r + 1, local.c);
					}
				} else if (local.info.columnTypes[local.c] eq "numeric") {
					spreadsheetSetCellValue(arguments.spreadsheet, local.value, local.r + 1, local.c);
					if (isNumeric(local.value)) {
						spreadsheetFormatCell(arguments.spreadsheet, {
							dataformat: "General"
						}, local.r + 1, local.c);
					}
				} else if (left(local.value, 1) eq "=") {
					spreadsheetSetCellFormula(arguments.spreadsheet, right(local.value, len(local.value) - 1), local.r + 1, local.c);
					if (mid(local.value, 2, 9) eq "HYPERLINK") {
						spreadsheetFormatCell(arguments.spreadsheet, {
							color: "blue",
							underline: true
						}, local.r + 1, local.c);
					}
				} else {
					spreadsheetSetCellValue(arguments.spreadsheet, local.value, local.r + 1, local.c);
					spreadsheetFormatCell(arguments.spreadsheet, {
						dataformat: "@"
					}, local.r + 1, local.c);
				}
			}
		}
		for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
			spreadsheetSetColumnWidth(arguments.spreadsheet, local.c, min(max(max(local.info.columnWidths[local.c], len(arguments.data.getColumnName(local.c))) + 2, 10), 30));
		}
		return arguments.spreadsheet;
	}

	/**
	 * Generates an Excel range string from the given coordinates.
	 * @startColumn The leftmost column in the range.
	 * @startRow The topmost row in the range.
	 * @endColumn The rightmost column in the range.
	 * @endRow The bottommost row in the range.
	 * @return An Excel range string.
	 */
	public string function rangeToString(required numeric startColumn, required numeric startRow, numeric endColumn, numeric endRow) {
		if (not structKeyExists(arguments, "endColumn")) {
			arguments.endColumn = arguments.startColumn;
		}
		if (not structKeyExists(arguments, "endRow")) {
			arguments.endRow = arguments.startRow;
		}
		return columnToLetter(arguments.startColumn) & arguments.startRow & ":" & columnToLetter(arguments.endColumn) & arguments.endRow;
	}

	/**
	 * Takes a query, and efficiently streams to the browser an XLSX file containing a sheet containing the query data, including formatting based on the query column types.
	 * The first row of the spreadsheet is frozen and includes filters.
	 * The first row will contain the column labels of the query, which are the same as the column names by default.
	 * To change the column labels for a query use: {@code variables.myQuery.getMetaData().setColumnLabel(["Label 1", "Label 2", "Label 3"])}.
	 * In character columns, any field value that starts with an equals sign will be treated as a formula.
	 * No formatting is applied to formula cells, unless it starts with {@code =HYPERLINK} in which case the text is formatted to be blue and underlined.
	 * This function changes the Content Disposition of the reponse and sets the page content using {@code cfcontent}.
	 * Typically, {@code abort} should be used immediately after calling this function or the function call should be at the end of the script.
	 * To generate a workbook with multiple sheets (i.e. multiple query result sets each displayed on a separate sheet), pass arrays to the data and sheetName parameters.
	 * @data A query result set. To generate a spreadsheet with multiple tabs, pass in an array of query result sets.
	 * @fileName The name to suggest for the file.
	 * @sheetName The name of the sheet. To generate a spreadsheet with multiple tabs, pass in an array of names.
	 * @template Path to an XLSX file to use as a template.
	 */
	public void function streamQueryAsSpreadsheet(required any data, string fileName = "export.xlsx", any sheetName = "Export", any template) {
		if (not isArray(arguments.data)) {
			arguments.data = [arguments.data];
		}
		if (not isArray(arguments.sheetName)) {
			arguments.sheetName = [arguments.sheetName];
		}
		local.numberOfSheets = max(arrayLen(arguments.data), arrayLen(arguments.sheetName));
		if (structKeyExists(arguments, "template")) {
			local.spreadsheet = new cfc.BigSpreadsheet(arguments.template);
		} else {
			local.spreadsheet = new cfc.BigSpreadsheet();
		}
		try {
			local.spreadsheet.createStyle("header", {
				alignment: "left",
				bold: true,
				dataformat: "@"
			});
			local.spreadsheet.createStyle("date", {
				dataformat: "m/d/y"
			});
			local.spreadsheet.createStyle("numeric", {
				dataformat: "General"
			});
			local.spreadsheet.createStyle("hyperlink", {
				color: "blue",
				underline: true
			});
			local.spreadsheet.createStyle("text", {
				dataformat: "@"
			});
			for (local.sheetIndex = 1; local.sheetIndex lte local.numberOfSheets; local.sheetIndex++) {
				local.sheetData = arguments.data[local.sheetIndex]?:queryNew("");
				local.spreadsheet.createSheet(arguments.sheetName[local.sheetIndex]?:"Export");
				local.spreadsheet.createRow();
				local.info = getQueryInfo(local.sheetData);
				for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
					local.spreadsheet.setCellValue(local.info.columnLabels[local.c]);
				}
				local.spreadsheet.formatRow("header");
				local.spreadsheet.addFreezePane(0, 1);
				local.spreadsheet.addAutoFilter(1, 1, 1, local.info.columnsCount);
				for (local.r = 1; local.r lte local.sheetData.recordCount; local.r++) {
					local.spreadsheet.createRow();
					for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
						local.value = local.sheetData[local.sheetData.getColumnName(local.c)][local.r];
						if (local.info.columnTypes[local.c] eq "date") {
							local.valid = isValid("date", local.value);
							local.spreadsheet.setCellValue(local.valid ? dateToNumber(local.value) : local.value);
							if (local.valid) {
								local.spreadsheet.formatCell("date");
							}
						} else if (local.info.columnTypes[local.c] eq "numeric") {
							local.spreadsheet.setCellValue(local.value);
							if (isNumeric(local.value)) {
								local.spreadsheet.formatCell("numeric");
							}
						} else if (left(local.value, 1) eq "=") {
							local.spreadsheet.setCellFormula(right(local.value, len(local.value) - 1));
							if (mid(local.value, 2, 9) eq "HYPERLINK") {
								local.spreadsheet.formatCell("hyperlink");
							}
						} else {
							local.spreadsheet.setCellValue(local.value);
							local.spreadsheet.formatCell("text");
						}
					}
				}
				for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
					local.spreadsheet.setColumnWidth(local.c, min(max(max(local.info.columnWidths[local.c], len(local.sheetData.getColumnName(local.c))) + 2, 10), 30));
				}
			}
			local.spreadsheet.setActiveSheetNumber(1);
			local.spreadsheet.stream(arguments.fileName);
		} finally {
			local.spreadsheet.dispose();
		}
	}

	/**
	 * Streams the given query to the browser as CSV data.
	 * This function changes the Content Disposition of the reponse; do not print any more content to the standard output.
	 * Typically, {@code abort} should be used immediately after calling this function or the function call should be at the end of the script.
	 * @data A query result set.
	 * @fileName The name to suggest for the file.
	 * @fieldDelimiter String to print between fields on a row.
	 * @lineSeparator String to print between rows.
	 * @quoteValues If true, always put quotes around fields.
	 * @preserveStrings If true, put quotes around character fields.
	 * @safeDelimiters If true, strip out the fieldDelimiter and lineSeparator from field values.
	 * @flushInterval The output buffer will flush whenever this many bytes are available. Adjusting this may improve memory usage.
	 */
	public void function streamQueryAsText(required query data, string fileName = "export.csv", string fieldDelimiter = ",", string lineSeparator = chr(10), boolean quoteValues = false, boolean preserveStrings = false, boolean safeDelimiters = false, numeric flushInterval = 10240) {
		cfcontent(reset=true);
		cfheader(name="Content-Disposition", value="attachment;filename=" & arguments.fileName);
		cfflush(interval=arguments.flushInterval);
		try {
			queryToOutput(arguments.data, ",", chr(10), true, true);
		} catch (any e) {
			cfheader(name="Content-Disposition", value="inline");
			rethrow;
		}
	}

	/**
	 * Streams a given spreadsheet object to the browser.
	 * @spreadsheet A spreadsheet object.
	 * @fileName The name to suggest for the file.
	 */
	public void function streamSpreadsheet(required any spreadsheet, required string fileName) {
		cfheader(name="Content-Disposition", value="attachment;filename=" & arguments.fileName);
		cfcontent(type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", variable=toBinary(arguments.spreadsheet));
	}

	/**
	 * Retrieves useful information about a query result set.
	 * @data A query result set.
	 * @return Query info data structure with keys: metaData (query meta data), columnTypes (array of strings; each is one of: "date", "numeric", "char"), columnWidths (array of integers).
	 */
	private struct function getQueryInfo(required query data) {
		local.info = {
			metaData: arguments.data.getMetaData(),
			columnTypes: [],
			columnWidths: []
		};
		local.info.columnLabels = local.info.metaData.getColumnLabels();
		local.info.columnsCount = arrayLen(arguments.data.getColumnNames());
		for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
			local.type = local.info.metaData.getColumnTypeName(local.c);
			if (listFindNoCase("DATE,DATETIME,TIMESTAMP", local.type) gt 0) {
				local.info.columnTypes[local.c] = "date";
			} else if (listFindNoCase("INT,INTEGER,TINYINT,SMALLINT,MEDIUMINT,BIGINT,DECIMAL,NUMERIC,FLOAT,DOUBLE,BIT", local.type) gt 0) {
				local.info.columnTypes[local.c] = "numeric";
			} else {
				local.info.columnTypes[local.c] = "char";
			}
			local.info.columnWidths[local.c] = 0;
		}
		for (local.r = 1; local.r lte arguments.data.recordCount; local.r++) {
			for (local.c = 1; local.c lte local.info.columnsCount; local.c++) {
				local.value = len(arguments.data[arguments.data.getColumnName(local.c)][local.r]);
				if (local.value gt local.info.columnWidths[local.c]) {
					local.info.columnWidths[local.c] = local.value;
				}
			}
		}
		return local.info;
	}

	/**
	 * Prints a field to the standard output.
	 * @value The value of the field.
	 * @fieldDelimiter String to print between fields on a row.
	 * @lineSeparator String to print between rows.
	 * @quoteValues If true, always put quotes around fields.
	 * @safeDelimiters If true, strip out the fieldDelimiter and lineSeparator from field values.
	 */
	private void function printField(required any value, required string fieldDelimiter, required string lineSeparator, required boolean quoteValues, required boolean safeDelimiters) {
		if (arguments.quoteValues) {
			writeOutput("""");
			writeOutput(replace(arguments.value, """", """""", "all"));
			writeOutput("""");
		} else {
			if (arguments.safeDelimiters) {
				arguments.value = replace(replace(arguments.value, arguments.fieldDelimiter, "", "all"), arguments.lineSeparator, "", "all");
			}
			writeOutput(arguments.value);
		}
	}

}

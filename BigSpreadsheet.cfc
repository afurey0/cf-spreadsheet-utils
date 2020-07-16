/**
 * Provides a way to manipulate spreadsheets that is similar to ColdFusion's default spreadsheet functions, but it uses a more memory efficient implementation.
 * Use this library instead of the default spreadsheet functions to generate spreadsheets that are too large for the default functions to handle.
 * Note that the BigSpreadsheet component required that you build your spreadsheets one row at a time from top to bottom.
 * Leverages the POI library's SXSSFWorkbook, but all Java classes are abstracted by this component.
 *
 * Example:
 *   variables.employees = queryExecute(...);
 *   variables.bss = new cfc.BigSpreadsheet(); // initialize the BigSpreadsheet
 *   try {
 *
 *     variables.bss.createStyle("header", { // register a style named "header" which makes the text bold
 *       bold: true
 *     });
 *     variables.spreadsheet.createStyle("date", { // register a style named "date" which formats the cell as a date
 *       dataformat: "d-mmm-yy"
 *     });
 *
 *     variables.bss.createSheet("My Sheet"); // add a sheet to the workbook
 *     variables.bss.formatColumn("date", 3); // apply the style named "date" to column 3
 *
 *     variables.bss.createRow(); // start a new row
 *     variables.bss.setCellValue("First Name"); // set the value of the first cell in the row
 *     variables.bss.setCellValue("Last Name"); // set the value of the next cell in the row
 *     variables.bss.setCellValue("Hire Date"); // etc
 *     variables.bss.formatRow("header"); // apply the style named "header" to the current row
 *
 *     for (variables.employee in variables.employees) { // loop over the dataset, starting a new row for each record
 *       variables.bss.createRow();
 *       variables.bss.setCellValue(variables.employee.firstname);
 *       variables.bss.setCellValue(variables.employee.lastname);
 *       variables.bss.setCellValue(variables.employee.hiredate);
 *     }
 *
 *     variables.write(expandPath("/myfolder/myfile.xlsx")); // write the workbook to disk
 *   } finally {
 *     variables.bss.dispose(); // always dispose of the BigSpreadsheet using a try-finally block
 *   }
 *
 * @author Alex Furey
 * @version 1.0.0
 * @since 2020-07-15
 * @see https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/streaming/SXSSFWorkbook.html
 */
component {

	/**
	 * Instantiates a BigSpreadsheet.
	 * Always call the dispose method after instantiating.
	 * @template Path to a template Excel file.
	 */
	public BigSpreadsheet function init(string template) {
		if (structKeyExists(arguments, "path")) {
			variables.workbook = createObject("java", "org.apache.poi.xssf.streaming.SXSSFWorkbook").init(createObject("java", "org.apache.poi.xssf.streaming.XSSFWorkbook").init(arguments.template));
		} else {
			variables.workbook = createObject("java", "org.apache.poi.xssf.streaming.SXSSFWorkbook").init();
		}
		variables.styles = {};
		return this;
	}

	/**
	 * Adds an autofilter to the active sheet.
	 */
	public void function addAutoFilter(required numeric startRow, required numeric endRow, required numeric startColumn, required numeric endColumn) {
		variables.activeSheet.setAutoFilter(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(arguments.startRow - 1, arguments.endRow - 1, arguments.startColumn - 1, arguments.endColumn - 1));
	}

	/**
	 * Adds a freeze pane to the active sheet.
	 * Removes any existing freeze pane.
	 */
	public void function addFreezePane(required numeric columns, required numeric rows) {
		variables.activeSheet.createFreezePane(arguments.columns, arguments.rows);
	}

	/**
	 * Registers a new cell style with the given name.
	 * TODO: Implement all the options provided by ColdFusion's built-in spreadsheetFormatCell function.
	 * @name Name of the style.
	 * @options See https://helpx.adobe.com/coldfusion/cfml-reference/coldfusion-functions/functions-s/spreadsheetformatcell.html
	 */
	public void function createStyle(required string name, required struct options) {
		local.style = variables.workbook.createCellStyle();
		if (structKeyExists(arguments.options, "alignment")) {
			local.style.setAlignment(nameToAlignment(arguments.options.alignment));
		}
		if (structKeyExists(arguments.options, "dataformat")) {
			local.style.setDataFormat(nameToFormat(arguments.options.dataformat));
		}
		if (structKeyExists(arguments.options, "bottomborder")) {
			local.style.setBorderBottom(nameToBorder(arguments.options.bottomborder));
		}
		if (structKeyExists(arguments.options, "leftborder")) {
			local.style.setBorderLeft(nameToBorder(arguments.options.leftborder));
		}
		if (structKeyExists(arguments.options, "rightborder")) {
			local.style.setBorderRight(nameToBorder(arguments.options.rightborder));
		}
		if (structKeyExists(arguments.options, "topborder")) {
			local.style.setBorderTop(nameToBorder(arguments.options.topborder));
		}
		if (structKeyExists(arguments.options, "bottombordercolor")) {
			local.style.setBottomBorderColor(nameToColor(arguments.options.bottombordercolor));
		}
		if (structKeyExists(arguments.options, "leftbordercolor")) {
			local.style.setLeftBorderColor(nameToColor(arguments.options.leftbordercolor));
		}
		if (structKeyExists(arguments.options, "rightbordercolor")) {
			local.style.setRightBorderColor(nameToColor(arguments.options.rightbordercolor));
		}
		if (structKeyExists(arguments.options, "topbordercolor")) {
			local.style.setTopBorderColor(nameToColor(arguments.options.topbordercolor));
		}
		if (structKeyExists(arguments.options, "fgcolor")) {
			local.fgcolor = nameToColor(arguments.options.fgcolor);
			local.style.setFillForegroundColor(local.fgcolor);
			// local.style.setFillBackgroundColor(local.fgcolor);
			local.style.setFillPattern(createObject("java", "org.apache.poi.ss.usermodel.FillPatternType").valueOf("SOLID_FOREGROUND"));
		}
		if (structKeyExists(arguments.options, "textwrap")) {
			local.style.setWrapText(arguments.options.textwrap);
		}
		local.font = variables.workbook.createFont();
		if (structKeyExists(arguments.options, "italic")) {
			local.font.setItalic(arguments.options.italic);
		}
		if (structKeyExists(arguments.options, "bold")) {
			local.font.setBold(arguments.options.bold);
		}
		if (structKeyExists(arguments.options, "color")) {
			local.font.setColor(nameToColor(arguments.options.color));
		}
		if (structKeyExists(arguments.options, "fontsize")) {
			local.font.setFontHeightInPoints(arguments.options.fontsize);
		}
		local.style.setFont(local.font);
		variables.styles[arguments.name] = local.style;
	}

	/**
	 * Creates a new sheet, and makes it the active sheet.
	 * If a sheet with the given name already exists, a new sheet is not created (the named sheet is still made active).
	 */
	public void function createSheet(required string name) {
		arguments.name = nameToSheetName(arguments.name);
		if (isNull(variables.workbook.getSheet(arguments.name))) {
			variables.activeSheet = variables.workbook.createSheet(arguments.name);
		}
		variables.workbook.setActiveSheet(variables.workbook.getSheetIndex(arguments.name));
	}

	/**
	 * Createa a new row in the active sheet, and makes it the active row.
	 */
	public void function createRow() {
		variables.activeRow = variables.activeSheet.createRow(variables.activeSheet.getPhysicalNumberOfRows());
	}

	/**
	 * Applies a registered style to an existing cell in the active row.
	 * If column is omitted, the last column in the row is used.
	 */
	public void function formatCell(required string style, numeric column) {
		if (not structKeyExists(arguments, "column")) {
			local.length = variables.activeRow.getLastCellNum();
			if (local.length gt 0) {
				arguments.column = local.length;
			} else {
				arguments.column = 1;
			}
		}
		getCellOnActiveRow(arguments.column - 1).setCellStyle(variables.styles[arguments.style]);
	}

	/**
	 * Applies a registered style to existing cells in the active row.
	 * If column is omitted, the last column in the row is used.
	 */
	public void function formatCells(required string style, required numeric startColumn, required numeric endColumn) {
		for (local.column = arguments.startColumn - 1; local.column lt arguments.endColumn; local.column++) {
			getCellOnActiveRow(local.column).setCellStyle(variables.styles[arguments.style]);
		}
	}

	/**
	 * Applies a registered style to a column in the active sheet.
	 * Call this function before populating the sheet.
	 */
	public void function formatColumn(required string style, required numeric column) {
		variables.activeSheet.setDefaultColumnStyle(arguments.column - 1, variables.styles[arguments.style]);
	}

	/**
	 * Applies a registered style to the active row.
	 */
	public void function formatRow(required string style) {
		variables.activeRow.setRowStyle(variables.styles[arguments.style]);
		local.iterator = variables.activeRow.cellIterator();
		while (local.iterator.hasNext()) {
			local.iterator.next().setCellStyle(variables.styles[arguments.style]);
		}
	}

	/**
	 * Merges cells in the active sheet.
	 * May cause high memory usage if used excessively; avoid if possible.
	 */
	public void function mergeCells(required numeric startRow, required numeric endRow, required numeric startColumn, required numeric endColumn) {
		variables.activeSheet.addMergedRegion(createObject("java", "org.apache.poi.ss.util.CellRangeAddress").init(arguments.startRow - 1, arguments.endRow - 1, arguments.startColumn - 1, arguments.endColumn - 1));
	}

	/**
	 * Sets the active sheet.
	 */
	public void function setActiveSheet(required string name) {
		arguments.name = nameToSheetName(arguments.name);
		variables.activeSheet = variables.workbook.getSheet(arguments.name);
		variables.workbook.setActiveSheet(variables.workbook.getSheetIndex(arguments.name));
	}

	/**
	 * Sets the active sheet by number (starts at 1).
	 */
	public void function setActiveSheetNumber(required numeric number) {
		variables.workbook.setActiveSheet(arguments.number - 1);
	}

	/**
	 * Sets the formula of a cell on the active row.
	 * If column is omitted, the column after the last column in the row is used.
	 */
	public void function setCellFormula(required any formula, numeric column) {
		if (not structKeyExists(arguments, "column")) {
			local.length = variables.activeRow.getLastCellNum();
			if (local.length gt 0) {
				arguments.column = local.length + 1;
			} else {
				arguments.column = 1;
			}
		}
		local.cell = getCellOnActiveRow(arguments.column - 1);
		local.cell.setCellFormula(arguments.formula);
		if (isNull(local.cell.getCellStyle)) {
			local.style = variables.activeSheet.getColumnStyle(arguments.column - 1);
			if (not isNull(local.style)) {
				local.cell.setCellStyle(local.style);
			}
		}
	}

	/**
	 * Sets the value of a cell on the active row.
	 * If column is omitted, the column after the last column in the row is used.
	 */
	public void function setCellValue(required any value, numeric column) {
		if (not structKeyExists(arguments, "column")) {
			local.length = variables.activeRow.getLastCellNum();
			if (local.length gt 0) {
				arguments.column = local.length + 1;
			} else {
				arguments.column = 1;
			}
		}
		local.cell = getCellOnActiveRow(arguments.column - 1);
		try {
			local.cell.setCellValue(arguments.value);
		} catch (any e) {
			local.cell.setCellValue(javaCast("double", arguments.value));
		}
		if (isNull(local.cell.getCellStyle)) {
			local.style = variables.activeSheet.getColumnStyle(arguments.column - 1);
			if (not isNull(local.style)) {
				local.cell.setCellStyle(local.style);
			}
		}
	}

	/**
	 * Sets the width of a column in the active sheet.
	 * Width is specified in characters, not pixels nor points.
	 */
	public void function setColumnWidth(required numeric column, required numeric width) {
		variables.activeSheet.setColumnWidth(arguments.column - 1, arguments.width * 256);
	}

	/**
	 * Sets the height of the active row in points.
	 */
	public void function setRowHeight(required numeric height) {
		variables.activeRow.setHeight(arguments.height * 20);
	}

	/**
	 * Streams the given query to the browser as an XLSX file.
	 * This function changes the Content Disposition of the reponse and sets the page content using the {@code cfcontent} tag.
	 * Typically, {@code abort} should be used after calling this function.
	 */
	public void function stream(required string fileName) {
		local.path = getTempFile(getTempDirectory(), "sxssf");
		write(local.path);
		cfheader(name="Content-Disposition", value="attachment;filename=" & arguments.fileName);
		cfcontent(type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", file=local.path, deletefile=true);
	}

	/**
	 * Writes the workbook to disk or to an output stream.
	 * If a stream is used, it will be closed by this function.
	 * @output An absolute file path or an output stream. If omitted, the default output stream is used.
	 */
	public void function write(any output) {
		if (structKeyExists(arguments, "output")) {
			if (isSimpleValue(arguments.output)) {
				local.outputStream = createObject("java", "java.io.FileOutputStream").init(arguments.output);
			} else {
				local.outputStream = arguments.output;
			}
		} else {
			local.outputStream = getPageContext().getResponse().getOutputStream();
		}
		try {
			variables.workbook.write(local.outputStream);
		} finally {
			local.outputStream.close();
		}
	}

	/**
	 * Cleans up the spreadsheet and any temporary files that were generated.
	 * Always call this method after initializing a BigSpreadsheet.
	 */
	public void function dispose() {
		variables.workbook.close();
		variables.workbook.dispose();
	}

	private any function getCellOnActiveRow(required numeric column) {
		local.cell = variables.activeRow.getCell(arguments.column);
		if (isNull(local.cell)) {
			local.cell = variables.activeRow.createCell(arguments.column);
		}
		return local.cell;
	}

	private any function nameToAlignment(required string name) {
		if (not structKeyExists(variables, "horizontalAlignments")) {
			variables.horizontalAlignments = createObject("java", "org.apache.poi.ss.usermodel.HorizontalAlignment");
		}
		return variables.horizontalAlignments.valueOf(uCase(arguments.name));
	}

	private any function nameToBorder(required string name) {
		if (not structKeyExists(variables, "borderStyles")) {
			variables.borderStyles = createObject("java", "org.apache.poi.ss.usermodel.BorderStyle");
		}
		return variables.borderStyles.valueOf(uCase(arguments.name));
	}

	private numeric function nameToColor(required string name) {
		return createObject("java", "org.apache.poi.hssf.util.HSSFColor$" & uCase(arguments.name)).init().getIndex();
	}

	private numeric function nameToFormat(required string name) {
		return variables.workbook.createDataFormat().getFormat(arguments.name);
	}

	private any function nameToSheetName(required string name) {
		if (not structKeyExists(variables, "workbookUtil")) {
			variables.workbookUtil = createObject("java", "org.apache.poi.ss.util.WorkbookUtil");
		}
		return variables.workbookUtil.createSafeSheetName(arguments.name);
	}

}

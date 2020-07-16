# cf-spreadsheet-utils
ColdFusion components for working with Excel spreadsheets

## BigSpreadsheet

A CFC wrapper for [SXSSFWorkbook](https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/streaming/SXSSFWorkbook.html) that provides an interface that is somewhat similar to ColdFusion's built-in spreadsheet functions. This component is only for writing spreadsheets. By leveraging SXSSFWorkbook, it can write spreadsheets in a very memory-efficient manner. This is excellent for generating very large spreadsheets. Note that you must build your spreadsheet from top to bottom.

```coldfusion
variables.employees = queryExecute(...);
variables.bss = new cfc.BigSpreadsheet(); // initialize the BigSpreadsheet
try {
  variables.bss.createStyle("header", { // register a style named "header" which makes the text bold
    bold: true
  });
  variables.spreadsheet.createStyle("date", { // register a style named "date" which formats the cell as a date
    dataformat: "d-mmm-yy"
  });
  variables.bss.createSheet("My Sheet"); // add a sheet to the workbook
  variables.bss.formatColumn("date", 3); // apply the style named "date" to column 3
  variables.bss.createRow(); // start a new row
  variables.bss.setCellValue("First Name"); // set the value of the first cell in the row
  variables.bss.setCellValue("Last Name"); // set the value of the next cell in the row
  variables.bss.setCellValue("Hire Date"); // etc
  variables.bss.formatRow("header"); // apply the style named "header" to the current row
  for (variables.employee in variables.employees) { // loop over the dataset, starting a new row for each record
    variables.bss.createRow();
    variables.bss.setCellValue(variables.employee.firstname);
    variables.bss.setCellValue(variables.employee.lastname);
    variables.bss.setCellValue(variables.employee.hiredate);
  }
  variables.write(expandPath("/myfolder/myfile.xlsx")); // write the workbook to disk
} finally {
  variables.bss.dispose(); // always dispose of the BigSpreadsheet using a try-finally block
}
```

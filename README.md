# cf-spreadsheet-utils
ColdFusion components for working with Excel spreadsheets

## Excel CFC

This component contains functions for working with Excel spreadsheets. It includes convenience functions for generating CSV and XLSX files from query result sets as well as streaming them to the browser. Use the `streamQueryAsSpreadsheet` method or the `streamQueryAsText` method for the best performance for spreadsheet downloads. 

### Example

In this example, we use the convenience function, `streamQueryAsSpreadsheet`, to generate a spreadsheet from a query and stream it to the browser on the fly.

```cfml
<cfscript>
  variables.employees = queryExecute("SELECT id, name, hire_date FROM employees"); // retrieve the query result set
  variables.excel = new cfc.Excel(); // instantiate an Excel CFC
  variables.excel.streamQueryAsSpreadsheet(variables.employees, "export.xlsx", "My Sheet"); // call the convenience function; pass in the spreadsheet, desired file name, and sheet name
  abort; // recommend calling abort here to make sure no other output gets streamed to the browser accidentally
</cfscript>
```

## BigSpreadsheet CFC

This is a CFC wrapper for POI's [SXSSFWorkbook](https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/streaming/SXSSFWorkbook.html) that provides an interface that is somewhat similar to ColdFusion's built-in spreadsheet functions. This component is only for *writing* spreadsheets. By leveraging SXSSFWorkbook, it can write spreadsheets with far better memory efficiency than you'd get with ColdFusion's normal spreadsheet functions. It works by building the file in a temporary location without keeping the entire file in memory as ColdFusion normally would. This is excellent for generating very large spreadsheets.

If your data comes from a query, you can use this memory efficient spreadsheet generator easily by using the convenience function `streamQueryAsSpreadsheet` found in the Excel CFC. This function uses BigSpreadsheet, but it handles all the details of reading the query's columns and rows.

### Example

In this example, we execute a query to get a list of employees. Then we generate a spreadsheet and write it to disk.

```cfml
<cfscript>
  variables.employees = queryExecute("SELECT id, name, hire_date FROM employees");
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
    variables.bss.setCellValue("ID"); // set the value of the first cell in the row
    variables.bss.setCellValue("Name"); // set the value of the next cell in the row
    variables.bss.setCellValue("Hire Date"); // etc
    variables.bss.formatRow("header"); // apply the style named "header" to the current row
    for (variables.employee in variables.employees) { // loop over the dataset, starting a new row for each record
      variables.bss.createRow();
      variables.bss.setCellValue(variables.employee.id);
      variables.bss.setCellValue(variables.employee.name);
      variables.bss.setCellValue(variables.employee.hire_date);
    }
    variables.write(expandPath("/myfolder/myfile.xlsx")); // write the workbook to disk
  } finally {
    variables.bss.dispose(); // always dispose of the BigSpreadsheet using a try-finally block
  }
  abort;
</cfscript>
```

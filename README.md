# Description
I needed something out of the box to parse an Excel file via NPOI library without relying on extra 3rd party packages, something that expects me to provide the sheet name (or names as this works with multiple sheets) and column indexes and have the result be returned as a DataTable (or DataSet if multiple sheets where parsed).

## Sample usage:
> [!NOTE]  
> While SheetName is mandatory, the ColumnIndexes isn't, so if it's missing then all columns from said sheet will be parsed.

To parse a single sheet:
```
try
{
    var sheetDetails = new List<ParserItems.SheetDetails>
    {
        new ParserItems.SheetDetails
        {
            SheetName = "Sheet1",
            ColumnIndexes = new List<int> { 0, 1, 2, 5, 9 }
        }
    };

    var parserItems = new ParserItems("MyExcelFile.xlsx", sheetDetails);
    DataTable dataTable = ParseSheet<DataTable>(parserItems);
}
catch (Exception ex)
{
    MessageBox.Show($"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
}
```

To parse multiple sheets:
```
try
{
    var sheetDetails = new List<ParserItems.SheetDetails>
    {
        new ParserItems.SheetDetails
        {
            SheetName = "Sheet1",
        },
        new ParserItems.SheetDetails
        {
            SheetName = "Sheet2",
            ColumnIndexes = new List<int> { 0, 2, 4 }
        }
    };

    var parserItems = new ParserItems("MyExcelFile.xlsx", sheetDetails);
    DataSet dataSet = ParseSheet<DataSet>(parserItems);
    var sheet1 = dataSet.Tables["Sheet1"];
    var sheet2 = dataSet.Tables["Sheet2"];

}
catch (Exception ex)
{
    MessageBox.Show($"{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
}
```
> [!NOTE]  
> A DataSet is returned instead of DataTable, and each parsed sheet can be accessed like dataSet.Tables["Sheet1"] and so on.

## To do:
- [ ] add some async magic so app won't hang while data is parsed?
- [ ] switch to IEnumerable\<T\> instead of DataTable for multiple reasons (strong typing, lazy loading, LINQ, performance to name a few)

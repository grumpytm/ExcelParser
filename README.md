# Description
I needed something out of the box to parse an Excel file via [MiniExcel](https://github.com/mini-software/MiniExcel) library without relying on too many extra 3rd party packages.
> [!WARNING]  
> Currently the MiniExcel library doesn't support .xls files, only .cvs, .xlsx and .xlsm.
## Sample usage:
First create a record:
```
public record Employees()
{
    [Column("A")]
    public int Id { get; set; }

    [Column("B")]
    public string Name { get; set; }

    [Column("D")]
    public string Department { get; set; }
}
```
then inside the app use the parser:
```
try
{
    var helper = new ExcelParser();

    var filePath = "EmployeesData.xlsx";
    var sheetName = "Employees"

    using var stream = File.OpenRead(filePath);
    IEnumerable<Employees> content = helper.ParseSheet<Employees>(stream, sheetName);

    // use a datagrid to display data for example
    BindingSource bindingSource = new BindingSource { DataSource = content.ToList() };
    dataGridView1.DataSource = bindingSource;
}
catch (IOException e) when ((e.HResult & 0x0000FFFF) == 32)
{
    Debug.Print("Can't continue as file is open by another process.");
}
```

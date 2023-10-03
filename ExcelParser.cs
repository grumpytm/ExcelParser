using System.Data;

/* Third party */
using NPOI.SS.UserModel;

public class ExcelParser
{
    private static string[] allowedExtensions = { ".xls", ".xlsx", ".xlsm" };

    public class ParserItems
    {
        public string FileName { get; set; }
        public List<SheetDetails> Sheet { get; set; }

        public class SheetDetails
        {
            public required string SheetName { get; set; }
            public List<int> ColumnIndexes { get; set; }
        }

        public ParserItems(string fileName, List<SheetDetails> sheets)
        {
            FileName = fileName;
            Sheet = sheets;

            var fileExtension = Path.GetExtension(fileName);
            if (!allowedExtensions.Any(s => s == fileExtension))
                throw new FileFormatException($"Error, '{fileName}' is not an accepted file format, only {string.Join(", ", allowedExtensions.ToList())} are allowed.");
        }
    }

    public static T ParseSheet<T>(ParserItems parserItems) where T : class
    {
        int sheetCount = parserItems.Sheet.Count;
        int sheetErrorCount = parserItems.Sheet.Count(sd => string.IsNullOrEmpty(sd.SheetName));

        if (sheetErrorCount >= 1)
            throw new ArgumentException("Error in configuration, parser can't have empty sheet names.");

        if (sheetCount < 1)
            throw new FileFormatException("Error in configuration, parser can't have null sheet names.");

        try
        {
            using (var fileStream = new FileStream(parserItems.FileName, FileMode.Open, FileAccess.Read))
            {
                DataTable dataTable = new DataTable();
                DataSet dataSet = new DataSet();

                IWorkbook workbook = WorkbookFactory.Create(fileStream);

                foreach (var detail in parserItems.Sheet)
                {
                    var sheetName = detail.SheetName;
                    var columnIndexes = detail.ColumnIndexes;

                    ISheet sheet = workbook.GetSheet(sheetName);
                    dataTable = new DataTable(sheetName);
                    PopulateDataTable(sheet, sheetName, columnIndexes, dataTable);

                    if (sheetCount > 1)
                    {
                        dataSet.Tables.Add(dataTable);
                        dataTable.Dispose();
                    }
                }

                if (sheetCount > 1)
                    return dataSet as T;
                else
                    return dataTable as T;
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    private static void PopulateDataTable(ISheet sheet, string sheetName, List<int> columnIndexes, DataTable dataTable)
    {
        if (sheet == null)
            throw new ArgumentException($"Error in configuration, '{sheetName}' sheet does not exist in the workbook.");

        var headerRow = sheet.GetRow(0);
        if (headerRow == null)
            throw new ArgumentException($"Sheet '{sheetName}' has no header row.");

        if (columnIndexes == null)
        {
            foreach (ICell headerCell in sheet.GetRow(0).Cells)
            {
                string columnName = headerCell.StringCellValue;
                dataTable.Columns.Add(columnName, typeof(string));
            }

            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = dataTable.NewRow();

                foreach (DataColumn dataColumn in dataTable.Columns)
                {
                    int columnIndex = dataColumn.Ordinal;
                    ICell cell = row.GetCell(columnIndex);
                    var cellValue = GetCellValue(cell);
                    dataRow[columnIndex] = cellValue;
                }

                dataTable.Rows.Add(dataRow);
            }
        }
        else
        {
            List<ICell> headerCells = sheet.GetRow(0).Cells;
            var columnIntersection = columnIndexes.Where(item => headerCells.Select(i => i.ColumnIndex).Contains(item)).ToList();

            foreach (int columnIndex in columnIntersection)
            {
                var columnName = headerRow.GetCell(columnIndex).StringCellValue;
                dataTable.Columns.Add(columnName);
            }

            for (var i = 1; i <= sheet.LastRowNum; i++)
            {
                DataRow dataRow = dataTable.NewRow();
                IRow row = sheet.GetRow(i);

                foreach (int columnIndex in columnIntersection)
                {
                    ICell cell = row.GetCell(columnIndex);
                    var cellValue = GetCellValue(cell);
                    int position = columnIntersection.IndexOf(columnIndex);
                    dataRow[position] = cellValue;
                }

                dataTable.Rows.Add(dataRow);
            }
        }
    }

    private static object? GetCellValue(ICell cell)
    {
        if (cell == null)
            return null;

        switch (cell.CellType)
        {
            case CellType.Unknown:
            case CellType.Blank:
                return cell.ToString();
            case CellType.Numeric:
                if (DateUtil.IsCellDateFormatted(cell))
                    return cell.DateCellValue;
                else
                    return cell.NumericCellValue;
            case CellType.String:
                return cell.StringCellValue;
            case CellType.Boolean:
                return cell.BooleanCellValue;
            case CellType.Error:
                return cell.ErrorCellValue;
            case CellType.Formula:
                cell.SetCellType(cell.CachedFormulaResultType);
                return GetCellValue(cell);
            default:
                return null;
        }
    }
}

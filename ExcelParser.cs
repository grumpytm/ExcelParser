using MiniExcelLibs;
using System.Reflection;

public class ExcelParser
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class Column : Attribute
    {
        public string Name { get; set; }

        public Column(string value) => Name = value;
    }

    public IEnumerable<T> ParseSheet<T>(Stream stream, string sheetName) where T : new()
    {
        return stream
            .Query(sheetName: sheetName)
            .Cast<IDictionary<string, object>>()
            .Skip(1)
            .Select(sheet =>
            {
                var record = new T();

                foreach (var prop in typeof(T).GetProperties())
                {
                    var columnName = prop.GetCustomAttribute<Column>()?.Name;
                    if (columnName != null && sheet.TryGetValue(columnName, out var value) && prop.CanWrite)
                    {
                        prop.SetValue(record, Convert.ChangeType(value ?? string.Empty, prop.PropertyType));
                    }
                }

                return record;
            });
    }
}

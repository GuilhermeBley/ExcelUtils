using System.Data;
using System.Text.Json;
using IronXL;

namespace ExcelUtils;

public class ExcelRead
{
    public const string DefaultTable = "Sheet1$";

    public static async Task<IEnumerable<T>> ConvertExcelAsync<T>(string fullPath, string tableName = DefaultTable)
    {
        var dataTable = await ConvertExcelToDataTableAsync(fullPath, tableName);

        string serializeddt = SerializeDataTable(dataTable);

        if (string.IsNullOrEmpty(serializeddt))
            return Enumerable.Empty<T>();

        return JsonSerializer.Deserialize<IEnumerable<T>>(serializeddt)?.ToList() ??
            throw new ArgumentNullException("Deserialize failed.");
    }

    public static async Task<IEnumerable<T>> ConvertExcelAsync<T>(string path, string filename, string tableName = DefaultTable)
        => await ConvertExcelAsync<T>(SplitedPathToFull(path, filename), tableName: tableName);
    
    public static async Task<DataTable> ConvertExcelToDataTableAsync(string fullPath, string tableName = DefaultTable)
    {
        var workBook = WorkBook.Load(fullPath);

        WorkSheet workSheet;

        if (tableName == DefaultTable)
            workSheet = workBook.DefaultWorkSheet;
        else
            workSheet = workBook.GetWorkSheet(tableName);

        if (workSheet is null)
            return new DataTable();

        await Task.CompletedTask;
        return workSheet.ToDataTable(true);
    }

    public static async Task<DataTable> ConvertExcelToDataTableAsync(string path, string fileName, string tableName = DefaultTable)
        => await ConvertExcelToDataTableAsync(SplitedPathToFull(path, fileName), tableName);
    
    public static IEnumerable<T> ConvertExcel<T>(string fullPath, string tableName = DefaultTable)
        => ConvertExcelAsync<T>(fullPath, tableName).GetAwaiter().GetResult();

    public static IEnumerable<T> ConvertExcel<T>(string path, string filename, string tableName = DefaultTable)
        => ConvertExcelAsync<T>(path, filename, tableName).GetAwaiter().GetResult();


    public static DataTable ConvertExcelToDataTable(string fullPath, string tableName = DefaultTable)
        => ConvertExcelToDataTableAsync(fullPath, tableName).GetAwaiter().GetResult();

    public static DataTable ConvertExcelToDataTable(string path, string fileName, string tableName = DefaultTable)
        => ConvertExcelToDataTableAsync(path, fileName, tableName).GetAwaiter().GetResult();

    private static string SplitedPathToFull(string path, string fileName)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Null or white space 'path'.");
            
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("Null or white space 'fileName'.");
            
        string fullpath = string.Empty;

        if (path.EndsWith('\\'))
            fullpath = $"{path}\\{fileName}";
        else
            fullpath = $"{path}{fileName}";

        return fullpath;
    }

    [Obsolete("OleDb is not used.")]
    private static string ConvertToConnectionStringOleDb(string fullPath)
    {
        if (string.IsNullOrWhiteSpace(fullPath))
            throw new ArgumentException("Null or white space 'fullpath'.");

        string connectionStringOleDb = string.Empty;
        if (fullPath.EndsWith(".xls"))
        {
            connectionStringOleDb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
        }
        else if (fullPath.EndsWith(".xlsx"))
        {
            connectionStringOleDb = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullPath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
        }
        else
            throw new FormatException("File must be ends with '.xls' or '.xlsx'.");

        return connectionStringOleDb;
    }

    public static string SerializeDataTable(DataTable dataTable)
    {
        if (dataTable is null ||
            !dataTable.AsEnumerable().Any())
        {
            return string.Empty;
        }

        var data = dataTable.Rows.OfType<DataRow>()
                    .Select(row => dataTable.Columns.OfType<DataColumn>()
                        .ToDictionary(col => col.ColumnName, c => row[c]));

        return System.Text.Json.JsonSerializer.Serialize(data);
    }
}

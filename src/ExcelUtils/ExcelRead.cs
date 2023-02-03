using System.Data;
using System.Data.OleDb;
using System.Text.Json;

namespace ExcelUtils;

public class ExcelRead
{
    public const string DefaultTable = "Sheet1$";

    public static async Task<IEnumerable<T>> ConvertExcel<T>(string fullPath, string tableName = DefaultTable)
    {
        var dataTable = await ConvertExcelToDataTable(fullPath, tableName);

        string serializeddt = JsonSerializer.Serialize(dataTable);

        return JsonSerializer.Deserialize<IEnumerable<T>>(serializeddt) ??
            throw new ArgumentNullException("Deserialize failed.");
    }

    public static async Task<IEnumerable<T>> ConvertExcel<T>(string path, string filename, string tableName = DefaultTable)
        => await ConvertExcel<T>(SplitedPathToFull(path, filename), tableName: tableName);
    

    public static async Task<DataTable> ConvertExcelToDataTable(string fullPath, string tableName = DefaultTable)
    {
        using OleDbConnection connection = new OleDbConnection(ConvertToConnectionStringOleDb(fullPath));
        await connection.OpenAsync();

        using var objDA = new System.Data.OleDb.OleDbDataAdapter($"select * from [{tableName}]", connection);
        var excelDataTable = new DataTable();
        objDA.Fill(excelDataTable);

        return excelDataTable;
    }

    public static async Task<DataTable> ConvertExcelToDataTable(string path, string fileName, string tableName = DefaultTable)
        => await ConvertExcelToDataTable(SplitedPathToFull(path, fileName), tableName);

    private static string SplitedPathToFull(string path, string fileName)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Null or white space 'path'.");
            
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("Null or white space 'fileName'.");
            
        string fullpath = string.Empty;

        if (path.EndsWith('/'))
            fullpath = $"{path}/{fileName}";
        else
            fullpath = $"{path}{fileName}";

        return fullpath;
    }

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
}

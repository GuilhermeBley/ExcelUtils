using ExcelUtils.Tests.Model;
using System.Data;

namespace ExcelUtils.Tests;

[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public class ExcelReadOleDbTest
{
    public static string CurrentProject { get; } = Directory.GetParent(Environment.CurrentDirectory)?.Parent?.Parent?.FullName 
        ?? throw new ArgumentNullException("Root not found.");

    public static string Root => CurrentProject + "\\root\\";

    [Fact]
    public async Task ConvertExcelToDataTable_GetTableFromWorkSheet_Success()
    {
        var fullPath = GetRootFile("BookOneTable.xlsx");
        await Task.Delay(1000);
        await WaitUntilUnlocked(fullPath, new CancellationTokenSource(1000).Token);
        var dt = await ExcelReadOleDb.ConvertExcelToDataTableAsync(fullPath);

        Assert.NotEmpty(dt.AsEnumerable());
    }

    [Fact]
    public async Task ConvertExcel_GetTableFromWorkSheet_Success()
    {
        var fullPath = GetRootFile("BookOneTable.xlsx");
        await Task.Delay(1000);
        await WaitUntilUnlocked(fullPath, new CancellationTokenSource(1000).Token);
        var modelEnumerable = await ExcelReadOleDb.ConvertExcelAsync<BookOneTableModel>(fullPath);

        Assert.NotEmpty(modelEnumerable);
        
        Assert.All(modelEnumerable, (model) =>
        {
            Assert.NotNull(model.Column1);
            Assert.NotNull(model.Column2);
            Assert.NotNull(model.Column3);
            Assert.NotNull(model.Column4);
        });
    }

    public static string GetRootFile(string fileName)
    {
        return Root + fileName;
    }

    private async Task WaitUntilUnlocked(string fullPath, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        while (IsFileLocked(fullPath))
        {
            await Task.Delay(100);
            cancellationToken.ThrowIfCancellationRequested();
        }
    }

    private bool IsFileLocked(string fullPath)
    {
        try
        {
            using(FileStream stream = File.Open(fullPath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                stream.Close();
            }
        }
        catch (IOException)
        {
            return true;
        }

        return false;
    }
}
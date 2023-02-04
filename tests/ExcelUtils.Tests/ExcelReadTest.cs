using ExcelUtils.Tests.Model;
using System.Data;

namespace ExcelUtils.Tests;

public class ExcelReadTest
{
    public static string CurrentProject { get; } = Directory.GetParent(Environment.CurrentDirectory)?.Parent?.Parent?.FullName 
        ?? throw new ArgumentNullException("Root not found.");

    public static string Root => CurrentProject + "\\root\\";

    [Fact]
    public void ConvertExcelToDataTable_GetTableFromWorkSheet_Success()
    {
        var dt = ExcelRead.ConvertExcelToDataTable(GetRootFile("BookOneTable.xlsx"));

        Assert.NotEmpty(dt.AsEnumerable());
    }

    [Fact]
    public void ConvertExcel_GetTableFromWorkSheet_Success()
    {
        var modelEnumerable = ExcelRead.ConvertExcel<BookOneTableModel>(GetRootFile("BookOneTable.xlsx"));

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
}
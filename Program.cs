using System.Reflection;
using OfficeOpenXml;

namespace ExcelToObject;
class Program
{
    static void Main(string[] args)
    {
        string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"ExcelFile.xlsx");
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var package = new ExcelPackage(new FileInfo(path));
        ExcelWorksheet sheet = package.Workbook.Worksheets["WorkBook"];

        var table = sheet.Tables.First();
        var x = ExcelExtension.ConvertTableToObjects<dynamic>(table);

        Console.WriteLine("Hello, World!");
    }
}

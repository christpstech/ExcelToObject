using System;
using System.IO;
using System.Text.Json;
using ExcelToObject.Models;
using ExcelToObject.Utils;

namespace ExcelToObject
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string sheetName = "WorkBook";
                string path = Path.Combine(Directory.GetCurrentDirectory(), @"asserts\ExcelFile.xlsx");

                var table = Extensions.ImportToDataTable(path, sheetName);
                if (table != null)
                {
                    var list = Extensions.DataTableToList<Person>(table);
                    if (list != null)
                    {
                        string jsonString = JsonSerializer.Serialize(list);
                        Console.WriteLine(jsonString);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
    }
}

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelToObject.Utils
{
    public static class Extensions
    {
        public static DataTable ImportToDataTable(string path, string sheetName)
        {
            DataTable dt = new DataTable();
            FileInfo fileInfo = new FileInfo(path);

            if (!fileInfo.Exists)
            {
                throw new Exception($"File {path} Does not exist.");
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage xlPackage = new ExcelPackage(fileInfo))
            {
                //Get the worksheets in the workbook 
                var worksheets = xlPackage.Workbook.Worksheets;
                if (worksheets.Count == 0)
                    throw new Exception($"Worksheets Does not empty.");

                ExcelWorksheet worksheet = worksheets[sheetName];

                if (worksheet is null)
                    throw new Exception($"Excel Sheet {sheetName} Does not exist.");

                //Obtain the worksheet size 
                ExcelCellAddress startCell = worksheet.Dimension.Start;
                ExcelCellAddress endCell = worksheet.Dimension.End;

                for (int row = startCell.Row; row <= endCell.Row; row++)
                {
                    DataRow dr = dt.NewRow(); //Create a row
                    int i = 0;
                    for (int col = startCell.Column; col <= endCell.Column; col++)
                    {
                        //Create the data column 
                        if (row == 1)
                            dt.Columns.Add(worksheet.Cells[row, col].Value.ToString());
                        else
                            dr[i++] = worksheet.Cells[row, col].Value.ToString();
                    }
                    if (row > 1)
                        dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        public static List<T> DataTableToList<T>(this DataTable table) where T : new()
        {
            List<T> list = new List<T>();
            var typeProperties = typeof(T).GetProperties().Select(propertyInfo => new
            {
                PropertyInfo = propertyInfo,
                Type = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType
            }).ToList();

            foreach (var row in table.Rows.Cast<DataRow>())
            {
                T obj = new T();
                foreach (var typeProperty in typeProperties)
                {
                    object value = row[typeProperty.PropertyInfo.Name];
                    object? safeValue = value == null || DBNull.Value.Equals(value)
                        ? null
                        : Convert.ChangeType(value, typeProperty.Type);

                    typeProperty.PropertyInfo.SetValue(obj, safeValue, null);
                }
                list.Add(obj);
            }
            return list;
        }
    }
}
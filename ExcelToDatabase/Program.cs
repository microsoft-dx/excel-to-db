using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelToDatabase
{
    class Program
    {
        static readonly XLDataType[] stringTypes =
            { XLDataType.Text, XLDataType.DateTime };
        static readonly Dictionary<string, string> columnNames =
            new Dictionary<string, string>
            {
                {"Something", "SomethingElse" }
            };

        static readonly string SampleFile =
            System.IO.Path.Combine("Files", "Financial Sample.xlsx");

        static void Main(string[] args)
        {
            var result = ReadExcelFile(SampleFile);

            if (result.columns.Count > 0)
            {
                string columnsString = string.Join(",", result.columns);

                foreach (List<string> row in result.rows)
                {
                    string rowString = string.Join(",", row);

                    string finalString =
                        $"INSERT INTO SomeTable({columnsString})" +
                        $"VALUES ({rowString});";
                }
            }

            Console.ReadKey();
        }

        // This will read first sheet
        // We will use ClosedXML
        internal static (List<string> columns, List<List<string>> rows)
            ReadExcelFile(string pathToExcel)
        {
            List<string> cols = new List<string>();
            List<List<string>> rows = new List<List<string>>();

            using (XLWorkbook excelFile =
                XLWorkbook.OpenFromTemplate(pathToExcel))
            {
                IXLWorksheet firstSheet = excelFile.Worksheets.FirstOrDefault();

                if (firstSheet != null)
                {
                    // The header is the first row, supposedly
                    foreach (IXLCell header in
                        firstSheet.Rows().FirstOrDefault().Cells())
                    {
                        string excelColName = header.Value.ToString();
                        cols.Add(columnNames.ContainsKey(excelColName)
                            ? columnNames[excelColName]
                            : excelColName);
                    }

                    foreach (IXLRow currentRow in firstSheet.Rows().Skip(1))
                    {
                        List<string> values = new List<string>();

                        foreach (IXLCell currentCell in currentRow.Cells())
                        {
                            // Make it SQL Proof from reading
                            values.Add(stringTypes.Contains(currentCell.DataType)
                                ? $"'{currentCell.Value}'"
                                : currentCell.Value.ToString());
                        }

                        rows.Add(values);
                    }
                }
            }

            return (cols, rows);
        }
    }
}

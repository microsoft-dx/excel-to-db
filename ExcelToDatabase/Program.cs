using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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

        const string serverPath = "";
        const string dbName = "";
        const string username = "";
        const string password = "";
        const string ourProcedure = "FromExcelToTable";

        static readonly string connectionString = $"Server={serverPath}" +
            $"Initial Catalog={dbName};" +
            "Persist Security Info=False;" +
            $"User ID={username};" +
            $"Password={password};" +
            "MultipleActiveResultSets=False;" +
            "Encrypt=True;" +
            "TrustServerCertificate=False;" +
            "Connection Timeout=30;";

        static void Main(string[] args)
        {
            var (columns, rows) = ReadExcelFile(SampleFile);

            if (columns.Count > 0)
            {
                // Make it nice
                string columnsString = string.Join(",",
                    columns.Select(x => $"[{x.Trim()}]"));

                foreach (List<string> row in rows)
                {
                    string rowString = string.Join(",", row);

                    string finalString =
                        $"INSERT INTO SomeTable({columnsString})" +
                        $"VALUES ({rowString});";

                    CallStoredProcedure(
                        cString: connectionString,
                        spName: ourProcedure,
                        parameters : new Dictionary<string, string>{
                        {"@tableName", "SomeTable"},
                        {"@columns", columnsString},
                        {"@row", rowString}
                    });
                }
            }

            Console.ReadKey();
        }

        internal static void CallStoredProcedure(
            string cString,
            string spName,
            Dictionary<string, string> parameters)
        {
            using var conn = new SqlConnection(cString);
            using var command = new SqlCommand(spName, conn)
            {
                CommandType = CommandType.StoredProcedure
            };
            foreach (KeyValuePair<string, string> parameter in parameters)
            {
                command.Parameters.AddWithValue(parameter.Key, parameter.Value);
            }

            try
            {
                conn.Open();
                command.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                // See if anything happes here, should not
                Console.WriteLine(e.Message);
            }
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

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInjection
{
    class ExcelReader
    {
        public static void ProcessExcel()
        {
            try
            {
                int maxColumns = 0;
                List<string> columnNames = new List<string>();
                List<string> filenames = new List<string>();
                Console.WriteLine("Paste root path of files");
                string path = Console.ReadLine();
                string[] files = Directory.GetFiles(path);
                Console.WriteLine("Files founded: ");
                foreach (string file in files)
                {
                    filenames.Add(path + @"\" + Path.GetFileName(file));
                    Console.WriteLine(path + @"\" + Path.GetFileName(file));
                }
                Console.WriteLine("Press Enter to continue");
                Console.ReadKey();

                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                var missing = System.Reflection.Missing.Value;

                DataTable dataTable = new DataTable("RfcReports");


                //Generate Headers
                foreach (string file in filenames)
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file, false, true, missing, missing, missing, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                    Array myValues = (Array)xlRange.Cells.Value2;

                    int vertical = myValues.GetLength(0);
                    int horizontal = myValues.GetLength(1);
                    bool isGreater = false;

                    if (horizontal >= maxColumns)
                    {
                        maxColumns = horizontal;
                        isGreater = true;
                    }

                    if (isGreater)
                    {
                        //Console.WriteLine("Headers count::" + horizontal);
                        //columnNames.Clear();
                        dataTable.Columns.Clear();
                        for (int i = 1; i <= horizontal; i++)
                        {
                            //dt.Columns.Add(new DataColumn(myValues.GetValue(1, i).ToString()));
                            if (!columnNames.Contains(myValues.GetValue(1, i).ToString()))
                            {
                                columnNames.Add(myValues.GetValue(1, i).ToString());
                            }
                        }
                        for (int col = 1; col <= columnNames.Count; col++)
                        {
                            DataColumn column = new DataColumn();
                            column.AllowDBNull = true;
                            column.DataType = Type.GetType("System.String");
                            column.DefaultValue = string.Empty;
                            column.MaxLength = 200;
                            column.ColumnName = columnNames[col - 1];
                            dataTable.Columns.Add(column);
                        }
                    }
                    xlWorkBook.Close(true, missing, missing);
                    xlApp.Quit();
                }

                //Generate Rows

                foreach (string file in filenames)
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file, false, true, missing, missing, missing, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                    Array myValues = (Array)xlRange.Cells.Value2;
                    int vertical = myValues.GetLength(0);
                    int horizontal = myValues.GetLength(1);

                    for (int a = 2; a <= vertical; a++)
                    {
                        object[] poop = new object[dataTable.Columns.Count];
                        for (int p = 0; p < poop.Length; p++)
                        {
                            poop[p] = null;
                        }
                        for (int b = 1; b <= horizontal; b++)
                        {
                            var currentCol = columnNames.Find(name => name == myValues.GetValue(1, b).ToString());
                            if (!string.IsNullOrEmpty(currentCol))
                            {
                                var index = columnNames.IndexOf(currentCol);
                                poop[index] = myValues.GetValue(a, b);
                            }
                        }
                        DataRow row = dataTable.NewRow();
                        row.ItemArray = poop;
                        dataTable.Rows.Add(row);
                    }

                    xlWorkBook.Close(true, missing, missing);
                    xlApp.Quit();
                }

                Console.WriteLine("Total registers readed: " + dataTable.Rows.Count);
                string queryTable = CreateTABLE(dataTable.TableName, dataTable);
                Connection(queryTable, dataTable);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
            }
        }

        public static string CreateTABLE(string tableName, DataTable table)
        {
            Console.WriteLine("Creating table");
            string sqlsc;
            sqlsc = "CREATE TABLE " + tableName + "(";
            for (int i = 0; i < table.Columns.Count; i++)
            {
                sqlsc += "\n [" + table.Columns[i].ColumnName + "] ";
                string columnType = table.Columns[i].DataType.ToString();
                switch (columnType)
                {
                    case "System.Int32":
                        sqlsc += " int ";
                        break;
                    case "System.Int64":
                        sqlsc += " bigint ";
                        break;
                    case "System.Int16":
                        sqlsc += " smallint";
                        break;
                    case "System.Byte":
                        sqlsc += " tinyint";
                        break;
                    case "System.Decimal":
                        sqlsc += " decimal ";
                        break;
                    case "System.DateTime":
                        sqlsc += " datetime ";
                        break;
                    case "System.String":
                    default:
                        sqlsc += string.Format(" nvarchar({0}) ", table.Columns[i].MaxLength == -1 ? "max" : table.Columns[i].MaxLength.ToString());
                        break;
                }
                if (table.Columns[i].AutoIncrement)
                    sqlsc += " IDENTITY(" + table.Columns[i].AutoIncrementSeed.ToString() + "," + table.Columns[i].AutoIncrementStep.ToString() + ") ";
                if (!table.Columns[i].AllowDBNull)
                    sqlsc += " NOT NULL ";
                sqlsc += ",";
            }
            return sqlsc.Substring(0, sqlsc.Length - 1) + "\n)";
        }

        public static void Connection(string query, DataTable table)
        {
            SqlConnection connection = new SqlConnection(string.Format("Data Source={0}; database={1}; User ID={2}; Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Pass"]));
            connection.Open();
            string queryDrop = string.Format("drop table if exists {0}", table.TableName);
            using (var commandDrop = new SqlCommand(queryDrop, connection))
            {
                int result = commandDrop.ExecuteNonQuery();
                if (result > 0)
                {
                    Console.WriteLine("Table existed, dropped");
                }
            }
            using (var command = new SqlCommand(query, connection))
            {
                int result = command.ExecuteNonQuery();
                if (result > 0)
                {
                    Console.WriteLine("Completed query");
                }
            }
            Console.WriteLine("Table created");
            SqlBulkCopy bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
            bulkCopy.DestinationTableName = table.TableName;
            bulkCopy.WriteToServer(table);

            connection.Close();
            Console.WriteLine("Data Injected");
        }
    }
}

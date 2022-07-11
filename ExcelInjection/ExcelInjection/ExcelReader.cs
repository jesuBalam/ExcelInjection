using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
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
                bool columnsFilled = false;
                bool isUsingExistingTable = false;
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

                

                Console.WriteLine("Enter table name target:");
                string nameTable = Console.ReadLine();

                DataTable dataTable = new DataTable(nameTable.Trim());
                columnNames.Clear();
                Console.WriteLine("Do you need to ceate table? y/n");
                var response = Console.ReadKey();
                if(response.Key == ConsoleKey.N)
                {
                    isUsingExistingTable = true;
                    dataTable.Columns.Clear();
                    SqlConnection connection = new SqlConnection(string.Format("Data Source={0}; database={1}; User ID={2}; Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Pass"]));
                    connection.Open();
                    string query = string.Format("USE {0} SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{1}' AND TABLE_SCHEMA = 'dbo'", ConfigurationManager.AppSettings["Database"], nameTable);
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader oReader = command.ExecuteReader())
                        {
                            while (oReader.Read())
                            {
                                //Console.WriteLine(oReader["COLUMN_NAME"].ToString() + "-" + oReader["COLUMN_NAME"].GetType().ToString());
                                DataColumn column = new DataColumn();
                                //column.AllowDBNull = true;
                                //column.DataType = oReader["COLUMN_NAME"].GetType();
                                if (oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Periodo".ToUpper().Replace(" ", "") 
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Regimen emisor".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Tipo de cambio".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Forma pago".ToUpper().Replace(" ", ""))
                                {
                                    column.DataType = typeof(Int32);
                                }
                                else if (oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Version".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "SubTotal".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Descuento".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "IVA Trasladado 0%".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "IVA Trasladado 16%".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "IVA Retenido".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "ISR Retenido".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "IEPS Trasladado".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Local retenido".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Local trasladado".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Total".ToUpper().Replace(" ", "")
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "IVA Trasladado".ToUpper().Replace(" ",""))
                                {
                                    column.DataType = typeof(Decimal);

                                }
                                else if (oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Fecha emision".ToUpper().Replace(" ", "") 
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Fecha certificacion".ToUpper().Replace(" ", "") 
                                    || oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "") == "Fecha proceso cancelacion".ToUpper().Replace(" ", ""))
                                {

                                    column.DataType = typeof(DateTime);
                                }
                                else
                                {
                                    column.DataType = typeof(String);
                                }
                                column.ColumnName = oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", "");
                                dataTable.Columns.Add(column);
                                columnNames.Add(oReader["COLUMN_NAME"].ToString().ToUpper().Replace(" ", ""));
                            }
                        }
                        
                    }
                    columnsFilled = true;
                    connection.Close();
                }                

                Console.WriteLine("Please wait. Reading and analyzing data. DON'T PRESS ANY KEY");

                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                var missing = System.Reflection.Missing.Value;
                //Generate Headers
                if(!columnsFilled)
                {
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
                                //column.DataType = Type.GetType("System.String");
                                column.ColumnName = columnNames[col - 1];
                                dataTable.Columns.Add(column);
                            }
                        }
                        xlWorkBook.Close(true, missing, missing);
                        xlApp.Quit();
                    }
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
                            var currentCol = columnNames.Find(name => name.ToUpper().Replace(" ", "") == myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", ""));
                            
                            if (!string.IsNullOrEmpty(currentCol))
                            {
                                var index = columnNames.IndexOf(currentCol);
                                if (myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Periodo".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Regimen emisor".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Tipo de cambio".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Forma pago".ToUpper().Replace(" ", ""))
                                {
                                    poop[index] = Convert.ToInt32(myValues.GetValue(a, b));
                                }
                                else if (myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Version".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "SubTotal".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Descuento".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "IVA Trasladado 0%".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "IVA Trasladado 16%".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "IVA Retenido".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "ISR Retenido".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "IEPS Trasladado".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Local retenido".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Local trasladado".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Total".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "IVA Trasladado".ToUpper().Replace(" ", ""))
                                {
                                    if (myValues.GetValue(a, b) != null && !string.IsNullOrEmpty(myValues.GetValue(a, b).ToString()))
                                    {
                                        if (decimal.TryParse(myValues.GetValue(a, b).ToString(), out decimal number))
                                        {
                                            poop[index] = number;
                                        }
                                        else
                                        {
                                            poop[index] = null;
                                        }
                                    }
                                    else
                                    {
                                        poop[index] = null;
                                    }

                                }
                                else if(myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Fecha emision".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Fecha certificacion".ToUpper().Replace(" ", "")
                                    || myValues.GetValue(1, b).ToString().ToUpper().Replace(" ", "") == "Fecha proceso cancelacion".ToUpper().Replace(" ", ""))
                                {
                                    if(myValues.GetValue(a, b) != null)
                                    {
                                        if(!string.IsNullOrEmpty(myValues.GetValue(a, b).ToString()))
                                        {
                                            //Console.WriteLine(myValues.GetValue(a, b).ToString());
                                            DateTime date = DateTime.ParseExact(myValues.GetValue(a, b).ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                                            string formattedDate = date.ToString("yyyy-MM-dd HH:mm:ss");
                                            poop[index] = formattedDate;
                                        }
                                        else
                                        {
                                            poop[index] = null;
                                        }
                                        
                                    }
                                    else
                                    {
                                        poop[index] = null;
                                    }
                                    
                                }
                                else
                                {
                                    if(myValues.GetValue(a,b) != null)
                                    {
                                        poop[index] = myValues.GetValue(a, b).ToString();
                                    }                                    
                                }
                            }
                        }
                        DataRow row = dataTable.NewRow();
                        row.ItemArray = poop;
                        dataTable.Rows.Add(row);
                    }

                    xlWorkBook.Close(true, missing, missing);
                    xlApp.Quit();
                }
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    Console.WriteLine(dataTable.Rows[0][c]);
                }
                Console.WriteLine("Total registers readed: " + dataTable.Rows.Count);
                Console.ReadKey();
                string queryTable = CreateTABLE(dataTable.TableName, dataTable);
                Connection(queryTable, dataTable, isUsingExistingTable);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
            }
        }

        public static string CreateTABLE(string tableName, DataTable table)
        {            
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

        public static void Connection(string query, DataTable table, bool existTable)
        {

            SqlConnection connection = new SqlConnection(string.Format("Data Source={0}; database={1}; User ID={2}; Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Pass"]));
            connection.Open();
            
            if(!existTable)
            {
                Console.WriteLine("Creating table");               
                using (var command = new SqlCommand(query, connection))
                {
                    int result = command.ExecuteNonQuery();
                    if (result > 0)
                    {
                        Console.WriteLine("Completed query");
                    }
                }
                Console.WriteLine("Table created");
            }
           
            SqlBulkCopy bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
            bulkCopy.DestinationTableName = table.TableName;
            bulkCopy.WriteToServer(table);

            connection.Close();
            Console.WriteLine("Data Injected");
        }
    }
}

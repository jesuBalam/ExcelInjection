using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
                List<NewColumn> newColumnsDict = new List<NewColumn>();
                List<string> columnNames = new List<string>();
                List<string> newColumnNames = new List<string>();
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

                

                Console.WriteLine("Enter table name target: (Default - SatData)");
                string nameTable = Console.ReadLine();
                if (string.IsNullOrEmpty(nameTable))
                {
                    nameTable = "SatData";
                }

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
                    string query = string.Format("USE {0} SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{1}' AND TABLE_SCHEMA = 'dbo'", ConfigurationManager.AppSettings["Database"], nameTable);
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader oReader = command.ExecuteReader())
                        {
                            while (oReader.Read())
                            {
                                //Console.WriteLine(oReader["COLUMN_NAME"].ToString() + "-" + oReader["COLUMN_NAME"].GetType().ToString());
                                DataColumn column = new DataColumn();                                
                                
                                if (Utils.GetType(oReader["COLUMN_NAME"].ToString()) == TypeData.Int || oReader["DATA_TYPE"].ToString() == "int")
                                {
                                    column.DataType = typeof(Int32);
                                }
                                else if (Utils.GetType(oReader["COLUMN_NAME"].ToString()) == TypeData.Decimal || oReader["DATA_TYPE"].ToString() == "decimal")
                                {
                                    column.DataType = typeof(Decimal);

                                }
                                else if (Utils.GetType(oReader["COLUMN_NAME"].ToString()) == TypeData.DateTime || oReader["DATA_TYPE"].ToString() == "datetime")
                                {

                                    column.DataType = typeof(DateTime);
                                }
                                else
                                {
                                    column.DataType = typeof(String);
                                }
                                column.ColumnName = oReader["COLUMN_NAME"].ToString();
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

                //Search new headers
                foreach (string file in filenames)
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file, false, true, missing, missing, missing, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                    Array myValues = (Array)xlRange.Cells.Value2;

                    int vertical = myValues.GetLength(0);
                    int horizontal = myValues.GetLength(1);


                    for (int i = 1; i <= horizontal; i++)
                    {
                        //dt.Columns.Add(new DataColumn(myValues.GetValue(1, i).ToString()));
                        string cleanColumn = Regex.Replace(myValues.GetValue(1, i).ToString(), @"[^0-9a-zA-Z]+", "");
                        if (!columnNames.Contains(cleanColumn.ToUpper().Replace(" ", "")))
                        {
                            newColumnNames.Add(cleanColumn.Replace(" ", ""));
                            columnNames.Add(cleanColumn.ToUpper().Replace(" ", ""));
                        }
                    }
                   
                    xlWorkBook.Close(true, missing, missing);
                    xlApp.Quit();
                }

                //Ask types
                if (newColumnNames.Count > 0)
                {
                    Console.WriteLine("New Columns founded in files:");
                    for(int col = 0; col< newColumnNames.Count; col++)
                    {
                        Console.WriteLine(newColumnNames[col]);
                    }
                    Console.WriteLine("Write type of value of each column: (v: varchar, i: int, t: datetime, d: decimal)");
                    for (int col = 0; col < newColumnNames.Count; col++)
                    {
                        Console.WriteLine(newColumnNames[col]);
                        var typeReaded = Console.ReadKey();
                        DataColumn column = new DataColumn();
                        //column.AllowDBNull = true;
                        column.DataType = typeReaded.Key == ConsoleKey.D ? typeof(Decimal) : typeReaded.Key == ConsoleKey.I ? typeof(Int32) : typeReaded.Key == ConsoleKey.T ? typeof(DateTime) : typeof(String);
                        column.ColumnName = newColumnNames[col];
                        dataTable.Columns.Add(column);
                        switch(typeReaded.Key)
                        {
                            case ConsoleKey.D:
                                Utils.columnsDecimal.Add(newColumnNames[col].ToString().ToUpper().Replace(" ", ""));
                                newColumnsDict.Add(new NewColumn { name = newColumnNames[col].Replace(" ", ""), type = "decimal(10,2)" });
                                break;
                            case ConsoleKey.I:
                                Utils.columnsInt.Add(newColumnNames[col].ToString().ToUpper().Replace(" ", ""));
                                newColumnsDict.Add(new NewColumn { name = newColumnNames[col].Replace(" ", ""), type = "int" });
                                break;
                            case ConsoleKey.T:
                                Utils.columnsDate.Add(newColumnNames[col].ToString().ToUpper().Replace(" ", ""));
                                newColumnsDict.Add(new NewColumn { name = newColumnNames[col].Replace(" ", ""), type = "datetime" });
                                break;
                            case ConsoleKey.V:
                                Utils.columnsString.Add(newColumnNames[col].ToString().ToUpper().Replace(" ", ""));
                                newColumnsDict.Add(new NewColumn { name = newColumnNames[col].Replace(" ", ""), type = "varchar(max)" });
                                break;
                            default:
                                Utils.columnsString.Add(newColumnNames[col].ToString().ToUpper().Replace(" ", ""));
                                newColumnsDict.Add(new NewColumn { name = newColumnNames[col].Replace(" ", ""), type = "varchar(max)" });
                                break;
                        }
                    }
                    Console.WriteLine("Building new table, dont press any key");
                }


                //Generate Headers
                if (!columnsFilled)
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
                Console.WriteLine("Generating rows. Wait.");

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
                            string cleanColumn = Regex.Replace(myValues.GetValue(1, b).ToString(), @"[^0-9a-zA-Z]+", "");
                            var currentCol = columnNames.Find(name => name.ToUpper().Replace(" ", "") == cleanColumn.ToUpper().Replace(" ", ""));
                            
                            if (!string.IsNullOrEmpty(currentCol))
                            {
                                var index = columnNames.IndexOf(currentCol);
                                if (Utils.GetType(cleanColumn) == TypeData.Int)
                                {
                                    poop[index] = Convert.ToInt32(myValues.GetValue(a, b));
                                }
                                else if (Utils.GetType(cleanColumn) == TypeData.Decimal)
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
                                else if(Utils.GetType(cleanColumn) == TypeData.DateTime)
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
                Console.WriteLine("Total registers readed: " + dataTable.Rows.Count);
                Console.ReadKey();
                if(newColumnNames.Count>0)
                {
                    UpdateTable(dataTable.TableName, newColumnsDict);
                }                
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

        public static void UpdateTable(string tableName, List<NewColumn> columns)
        {
            Console.WriteLine("Updating sql table..");
            string queryAdd = string.Format("ALTER TABLE {0} ADD ", tableName);
            string columnsToAdd = "";
            for(int col = 0; col<columns.Count; col++)
            {
                if(col==columns.Count-1)
                {
                    columnsToAdd += string.Format("{0} {1} ", columns[col].name, columns[col].type);
                }
                else
                {
                    columnsToAdd += string.Format("{0} {1}, ", columns[col].name, columns[col].type);
                }
            }
            string finalQuery = queryAdd + columnsToAdd;
            SqlConnection connection = new SqlConnection(string.Format("Data Source={0}; database={1}; User ID={2}; Password={3}", ConfigurationManager.AppSettings["ServerDatabase"], ConfigurationManager.AppSettings["Database"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Pass"]));
            connection.Open();
            using (var command = new SqlCommand(finalQuery, connection))
            {
                int result = command.ExecuteNonQuery();
                if (result > 0)
                {
                    Console.WriteLine("SQL Table update with new columns");
                }
            }
            connection.Close();
            Console.WriteLine("SQL Table update with new columns. Press any key");
            Console.ReadKey();
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

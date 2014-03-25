//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;

namespace Ellis.WinApp.Testing.Framework
{
    public class ExcelReader
    {
       public static IEnumerable<DataRow> ImportSpreadsheet(string tablename)
        {
            const string extendedProperties = "Excel 12.0;HDR=YES;IMEX=1";
            //const string path =
            //    "C:\\Users\\kkodandarama\\Documents\\Visual Studio 2013\\Projects\\EllisTestAutomation\\EllisTestAutomation\\Data\\EllisTestData.xls";

            var path = System.IO.Path.GetFullPath(".\\"+ExcelConstants.ExcelName);
            
            var connectionString = string.Format(
                CultureInfo.CurrentCulture,
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"{1}\"",
                path,
                extendedProperties);
            IEnumerable<DataRow> dataRows = null;

            using (var connection = new OleDbConnection(connectionString))
            {
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT * FROM " + tablename;
                    connection.Open();

                    using (var adapter = new OleDbDataAdapter(command))
                    using (var columnDataSet = new DataSet())
                    using (var dataSet = new DataSet())
                    {
                        columnDataSet.Locale = CultureInfo.CurrentCulture;
                        adapter.Fill(columnDataSet);

                        if (columnDataSet.Tables.Count == 1)
                        {
                            //columnDataSet.Clear();
                            var worksheet = columnDataSet.Tables[0];

                            // Now that we have a valid worksheet read in, with column names, we can create a
                            // new DataSet with a table that has preset columns that are all of type string.
                            // This fixes a problem where the OLEDB provider is trying to guess the data types
                            // of the cells and strange data appears, such as scientific notation on some cells.
                            dataSet.Tables.Add("WorksheetData");
                            var tempTable = dataSet.Tables[0];

                            foreach (DataColumn column in worksheet.Columns)
                            {
                                tempTable.Columns.Add(column.ColumnName, typeof (string));
                            }

                            adapter.Fill(dataSet, "WorksheetData");
                            dataRows =
                                (from DataRow dr in dataSet.Tables[0].Rows
                                    where dr["Type"].ToString() == "Valid"
                                    select dr);

                            //if (dataSet.Tables.Count == 1)
                            //{
                            //    worksheet = dataSet.Tables[0];
                            //    foreach (var row in worksheet.Rows)
                            //    {
                            //        // TODO: Consume some data.
                            //    }
                            //}
                        }
                    }
                    connection.Close();
                }
            }
            return dataRows;
        }
    }

    public class ExcelConstants
    {
        //public const string Location =
        //    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kkodandarama\\Documents\\Visual Studio 2013\\Projects\\EllisTestAutomation\\EllisTestAutomation\\Data\\EllisTestData.xls;Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";

        //public const string Location = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\kkodandarama\\Documents\\Visual Studio 2013\\Projects\\EllisAutomation\\TestResults\\Kkodandarama_SQAD-A111215 2013-11-15 11_13_04\\Out\\EllisTestData.xls;Extended Properties=Excel 8.0";

        public const string ExcelName = "TestData.xls";
    }

    //public class DataPath
    //{
    //    public static string GetConnectionString(string name)
    //    {
    //        //var str = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
    //        var str = ConfigurationManager.ConnectionStrings[name].ConnectionString;
    //        return str;
    //    }

    //    public static string GetDirectoryPath()
    //    {
    //        var path =
    //            System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);

    //        return path;
    //    }
    //}

    //public static IEnumerable<DataRow> ImportFromExcel(string tableName)
    //{
    //    //var str = DataPath.GetConnectionString("EllisTestData");
    //    //var path = DataPath.GetDirectoryPath();
    //    //path = path.Substring(6);
    //    //str = str.Replace("|DataDirectory|", path);

    //    var con =
    //        new OleDbConnection(@ExcelConstants.Location);
    //    var cmd = "select * from "+tableName;
    //    var da = new OleDbDataAdapter(cmd, con);
    //    var dt = new DataSet();
    //    da.Fill(dt);

    //    var dataRows =
    //        (from DataRow dr in dt.Tables[0].Rows where dr["Type"].ToString() == "Valid" select dr);
    //    return dataRows;
    //}
}
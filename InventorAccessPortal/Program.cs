using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using ADOX;
using System.Data.OleDb;
using System.Data.SqlTypes;
namespace Convert
{
    class Program
    {
        static void Main(string[] args)
        {
            //The connection strings needed: One for SQL and one for Access
            String accessConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Databases\\test.accdb;";
            String sqlConnectionString = "Data Source=DESKTOP-6P2HDJO\\SQLEXPRESS;Initial Catalog=InventorAccessPortal_SQL;Integrated Security=True";

            //Make adapters for each table we want to export
            SqlDataAdapter adapter1 = new SqlDataAdapter("select * from LoginData", sqlConnectionString);
            //SqlDataAdapter adapter2 = new SqlDataAdapter("select * from Table2", sqlConnectionString);

            //Fills the data set with data from the SQL database
            DataSet dataSet = new DataSet();
            adapter1.Fill(dataSet, "LoginData");
            //adapter2.Fill(dataSet, "Table2");

            //Create an empty Access file that we will fill with data from the data set
            ADOX.Catalog catalog = new ADOX.Catalog();
            catalog.Create(accessConnectionString);

            //Create an Access connection and a command that we'll use
            OleDbConnection accessConnection = new OleDbConnection(accessConnectionString);
            OleDbCommand command = new OleDbCommand();
            command.Connection = accessConnection;
            command.CommandType = CommandType.Text;
            accessConnection.Open();
            //This loop creates the structure of the database
            foreach (DataTable table in dataSet.Tables)
            {
                String columnsCommandText = "(";

                foreach (DataColumn column in table.Columns)
                {
                    String columnName = column.ColumnName;
                    String dataTypeName = column.DataType.Name;
                    String sqlDataTypeName = getAccDataTypeName(dataTypeName);
                    columnsCommandText += "[" + columnName + "] " + sqlDataTypeName + ",";
                }
                columnsCommandText = columnsCommandText.Remove(columnsCommandText.Length - 1);
                columnsCommandText += ")";

                command.CommandText = "CREATE TABLE " + table.TableName + columnsCommandText;

                command.ExecuteNonQuery();
            }

            //This loop fills the database with all information
            foreach (DataTable table in dataSet.Tables)
            {
                foreach (DataRow row in table.Rows)
                {
                    String commandText = "INSERT INTO " + table.TableName + " VALUES (";
                    foreach (var item in row.ItemArray)
                    {
                        commandText += "'" + item.ToString() + "',";
                    }
                    commandText = commandText.Remove(commandText.Length - 1);
                    commandText += ")";

                    command.CommandText = commandText;
                    command.ExecuteNonQuery();
                }
            }

            accessConnection.Close();
        }

        static string getAccDataTypeName(string dataTypeName)
        {
            dataTypeName = dataTypeName + "          ";
            if (String.Equals("char", dataTypeName.Substring(0, 4)) ||
                String.Equals("varchar", dataTypeName.Substring(0, 7)) ||
                String.Equals("text", dataTypeName.Substring(0, 4)) ||
                String.Equals("nchar", dataTypeName.Substring(0, 5)) ||
                String.Equals("nvarchar", dataTypeName.Substring(0, 8)) ||
                String.Equals("ntext", dataTypeName.Substring(0, 5)) ||
                String.Equals("String", dataTypeName.Substring(0, 6)))
            {
                return "Text";
            }
            if (String.Equals("bit", dataTypeName.Substring(0, 3)) ||
                String.Equals("tinyint", dataTypeName.Substring(0, 7))
            )
            {
                return "Byte";
            }
            if (String.Equals("smallint", dataTypeName.Substring(0, 8))
                )
            {
                return "Integer";
            }
            if (String.Equals("int", dataTypeName.Substring(0, 3)) ||
                String.Equals("bigint", dataTypeName.Substring(0, 6)) ||
                String.Equals("Int", dataTypeName.Substring(0, 3))
                )
            {
                return "Long";
            }
            if (String.Equals("float", dataTypeName.Substring(0, 5)))
            {
                return "Single";
            }
            if (String.Equals("real", dataTypeName.Substring(0, 4)))
            {
                return "Double";
            }
            if (String.Equals("datetime", dataTypeName.Substring(0, 8)) ||
                String.Equals("datetime2", dataTypeName.Substring(0, 9)) ||
                String.Equals("smalldatetime", dataTypeName.Substring(0, 13)))
            {
                return "Date/Time";
            }
            if (String.Equals("binary", dataTypeName.Substring(0, 6)) ||
                String.Equals("varbinary", dataTypeName.Substring(0, 6)) ||
                String.Equals("image", dataTypeName.Substring(0, 5)))
            {
                return "Ole Obvject";
            }
            dataTypeName = dataTypeName.Substring(0,dataTypeName.Length - 10);
            return dataTypeName;
        }
    }
}

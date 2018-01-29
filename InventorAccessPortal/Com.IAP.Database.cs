using System;
using System.Data;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;

namespace Com.IAV.Database
{
    public class ConnDbForAcccess
    {
        /// <summary>
        /// String to connect the Database
        /// </summary>
        private string connectionString;

        /// <summary>
        /// Save the connection to Database
        /// </summary>
        protected OleDbConnection Connection;

        /// <summary>
        /// Defult connection of database
        /// </summary>
        public ConnDbForAcccess()
        {
            string connStr;
            connStr = System.Configuration.ConfigurationManager.AppSettings["connStr"]; //read setting from web.config
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + connStr + "Persist Security Info=False;";
            Connection = new OleDbConnection(connectionString);
        }

        /// <summary>
        /// connection with parameters 
        /// </summary>
        /// <param name="newConnectionString"></param>
        public ConnDbForAcccess(string newConnectionString)
        {
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + /*HttpContext.Current.Request.PhysicalApplicationPath + */ newConnectionString;
            Connection = new OleDbConnection(connectionString);
        }

        /// <summary>
        /// Get connection string
        /// </summary>
        public string ConnectionString
        {
            get
            {
                return connectionString;
            }
        }


        /// <summary>
        /// execute SQL command which don't have a feedback result
        /// </summary>
        /// <param name="strSQL"></param>
        /// <returns>the sgin of whether succeed</returns>
        public bool ExeSQL(string strSQL)
        {
            bool resultState = false;

            Connection.Open();
            OleDbTransaction myTrans = Connection.BeginTransaction();
            OleDbCommand command = new OleDbCommand(strSQL, Connection, myTrans);

            try
            {
                command.ExecuteNonQuery();
                myTrans.Commit();
                resultState = true;
            }
            catch
            {
                myTrans.Rollback();
                resultState = false;
            }
            finally
            {
                Connection.Close();
            }
            return resultState;
        }

        /// <summary>
        /// execute SQL command and send result back to DataReader
        /// </summary>
        /// <param name="strSQL"></param>
        /// <returns>dataReader</returns>
        private OleDbDataReader ReturnDataReader(string strSQL)
        {
            Connection.Open();
            OleDbCommand command = new OleDbCommand(strSQL, Connection);
            OleDbDataReader dataReader = command.ExecuteReader();
            Connection.Close();

            return dataReader;
        }

        /// <summary>
        /// execute SQL command and send result back to DataSet
        /// </summary>
        /// <param name="strSQL"></param>
        /// <returns>DataSet</returns>
        public DataSet ReturnDataSet(string strSQL)
        {
            Connection.Open();
            DataSet dataSet = new DataSet();
            OleDbDataAdapter OleDbDA = new OleDbDataAdapter(strSQL, Connection);
            OleDbDA.Fill(dataSet, "objDataSet");

            Connection.Close();
            return dataSet;
        }

        /// <summary>
        /// Execute a search SQL command and return the number of results
        /// </summary>
        /// <param name="strSQL"></param>
        /// <returns>sqlResultCount</returns>
        public int ReturnSqlResultCount(string strSQL)
        {
            int sqlResultCount = 0;

            try
            {
                Connection.Open();
                OleDbCommand command = new OleDbCommand(strSQL, Connection);
                OleDbDataReader dataReader = command.ExecuteReader();

                while (dataReader.Read())
                {
                    sqlResultCount++;
                }
                dataReader.Close();
            }
            catch
            {
                sqlResultCount = 1;
            }
            finally
            {
                Connection.Close();
            }
            return sqlResultCount;
        }


    }
}
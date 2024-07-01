using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectsInfoSystem
{
    // класс "Работа с SQL"
    class MC_SQLDataProvider
    {
        private string dbconnectionString; // строка подключения к БД

        public MC_SQLDataProvider(string db_connectionString)
        {
            dbconnectionString = db_connectionString;
        }

        // вызов хранимой процедуры
        public static DataTable CallStoredProcedure(SqlConnection connection, DataTable dataset, string storedProcedureWParams,
            List<SqlParameter> listSQLparams)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();

            adapter.SelectCommand = new SqlCommand();
            adapter.SelectCommand.Connection = connection;
            adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            adapter.SelectCommand.CommandText = storedProcedureWParams;

            //adapter.SelectCommand.Parameters.AddWithValue("idmapsrc", idmapsrc);

            foreach (SqlParameter p in listSQLparams)
            {
                //adapter.SelectCommand.Parameters.Add(p);
                adapter.SelectCommand.Parameters.AddWithValue(p.ParameterName, p.Value);
            }            

            adapter.SelectCommand.CommandTimeout = 0; //бесконечное время ожидания
            adapter.SelectCommand.ExecuteNonQuery();

            adapter.Fill(dataset);
                        
            return dataset;
        }

        // проверить!!! а то и возвращает Table и получает Table
        // выборка данных SQL-запросом с созданием подключения
        public static DataTable SelectDataFromSQL(DataTable dataset, string connectionString, string queryString)
        {
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = new SqlCommand(queryString, connection);
                        
            adapter.SelectCommand.CommandTimeout = 0; //бесконечное время ожидания

            adapter.Fill(dataset);
            connection.Close();
            return dataset;
        }

        public DataTable SelectDataFromSQL(string queryString)
        {
            SqlConnection connection = new SqlConnection(dbconnectionString);
            connection.Open();

            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = new SqlCommand(queryString, connection);
            adapter.SelectCommand.CommandTimeout = 0; //бесконечное время ожидания
            DataTable dataset = new DataTable();
            adapter.Fill(dataset);

            connection.Close();

            return dataset;
        }

        // выборка данных SQL-запросом без создания подключения
        public static DataTable SelectDataFromSQL(SqlConnection connection, DataTable dataset, string queryString)
        {            
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = new SqlCommand(queryString, connection);

            adapter.SelectCommand.CommandTimeout = 0; //бесконечное время ожидания

            adapter.Fill(dataset);
            
            return dataset;
        }
        
        public static void InsertSQLQuery(string connectionString, string queryString)
        {
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter();

            adapter.InsertCommand = new SqlCommand(queryString, connection);
            adapter.InsertCommand.ExecuteNonQuery();

            connection.Close();
        }

        public static void DeleteSQLQuery(string connectionString, string queryString)
        {
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter();

            adapter.DeleteCommand = new SqlCommand(queryString, connection);
            adapter.DeleteCommand.ExecuteNonQuery();

            connection.Close();
        }

        public static void UpdateSQLQuery(string connectionString, string queryString)
        {
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter();

            adapter.UpdateCommand = new SqlCommand(queryString, connection);
            adapter.UpdateCommand.ExecuteNonQuery();
            
            connection.Close();
        }

        // функция получения значения заданного свойства заданного л/с из OLAP-куба ИЭСБК за заданный период
        // возвращаемые значения:
        // - текстовое значение свойства при наличии
        // - null при отсутствии значения в заданном периоде (или при более чем 1 строке результата - дубли л/с в разных отделениях)
        public static string GetPropValueFromIESBKOLAP(string codels, string periodyear, string periodmonth, int propnumber, SqlConnection connection)
        {
            string queryStringPropValue =
                String.Concat("SELECT codeIESBK,propvalue FROM [iesbk].[dbo].[tblIESBKlspropvalue] ",
                              "WHERE codeIESBK = '", codels, "' AND periodyear = '", periodyear,
                              "' AND periodmonth = '", periodmonth, "' AND lspropertieid ='", propnumber.ToString(), "'");
            DataTable tablePropValue = new DataTable();
            SelectDataFromSQL(connection, tablePropValue, queryStringPropValue);

            string result = null;
            if (tablePropValue.Rows.Count == 1) result = tablePropValue.Rows[0]["propvalue"].ToString();

            tablePropValue.Dispose();

            return result;
        }
    }
}

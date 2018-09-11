using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace IFaceAttReader
{
    class ConnectionFactory
    {
        private static readonly string connectionString =
            ConfigurationManager.ConnectionStrings["mysql"].ConnectionString;
        /// <summary>
        /// mySQl 数据库
        /// </summary>
        /// <returns></returns>
        public static IDbConnection MySqlConnection()
        {
            string mysqlconnectionString = connectionString;  //ConfigurationManager.ConnectionStrings["mysqlconnectionString"].ToString();
            var connection = new MySqlConnection(mysqlconnectionString);
            if (connection.State == ConnectionState.Closed)
            {
                connection.Open();
            }
            return connection;
        }
    }
}

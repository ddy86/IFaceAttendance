using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;

namespace IFaceAttReader
{
    public static class DapperDBContext
    {
        public static List<T> AsList<T>(this IEnumerable<T> source)
        {
            if (source != null && !(source is List<T>))
                return source.ToList();
            return (List<T>)source;
        }
        //参数我们跟后台封装方法保持一致
        public static int Execute(string sql, object param = null, IDbTransaction transaction = null, int? commandTimeout = null, CommandType? commandType = null, int databaseOption = 1)
        {
            using (var conn = ConnectionFactory.MySqlConnection())
            {
                var restult = conn.Execute(sql, param, transaction, commandTimeout, commandType);
                return restult;
            }
        }

        public static List<IFaceAttendance> Query(string sql, object param = null)
        {
            using (var conn = ConnectionFactory.MySqlConnection())
            {
                var restult = conn.Query<IFaceAttendance>(sql, new { Time = param });
                return restult.ToList();
            }
        }


        public static int Execute(CommandDefinition command, int databaseOption = 1)
        {
            using (var conn = ConnectionFactory.MySqlConnection())
            {
                var restult = conn.Execute(command);
                return restult;
            }
        }
    }
}

using System;
using System.Data;
using System.Data.SqlClient;
using static System.String;

namespace TableNaker
{
    public class RuleHelper
    {
        //2100 SQL max paremeters
        public static int CalcBatchSize(int columnCount)
        {
            return (int)1800 / columnCount > 1000 ? 1000 : (int)1800 / columnCount;
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static int CalcBatchSizePageCount(decimal ColumnCount, decimal CalcBatchSize)
        {
            return (int)Math.Ceiling(ColumnCount / CalcBatchSize);
        }

        public static string CheckConnection(ref SqlConnectionStringBuilder sqlConnection)
        {
            var connectionResult = "";
            try
            {
                using (var connection = new SqlConnection(sqlConnection.ConnectionString))
                {
                    connection.Open();
                    connectionResult = connection.State == ConnectionState.Open
                        ? "Connection successful."
                        : "Failed to connect to database.";
                    connection.Close();
                    connection.Dispose();
                }
            }
            catch (SqlException sqlException)
            {
                connectionResult = sqlException.Message;
            }

            return connectionResult;
        }

        public static bool CheckIfTableExists(ref SqlConnectionStringBuilder sqlConnection, string TableName)
        {
            var connection = new SqlConnection(sqlConnection.ConnectionString);
            var cmdText = @"IF EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES 
                            WHERE TABLE_NAME='" + TableName + "') SELECT 1 ELSE SELECT 0";
            connection.Open();
            var dateCheck = new SqlCommand(cmdText, connection);
            var x = Convert.ToInt32(dateCheck.ExecuteScalar());
            connection.Close();
            connection.Dispose();
            return x == 1;
        }
    }
}
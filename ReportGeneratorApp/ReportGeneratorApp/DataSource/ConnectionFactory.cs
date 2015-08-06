using System.Configuration;
using System.Data.SqlClient;

namespace ReportGeneratorApp.DataSource
{
    public class ConnectionFactory
    {
        public static SqlConnection GetConnection()
        {
            string sqlConnectString = ConfigurationManager.ConnectionStrings["reportDBConnectionString"].ConnectionString;
            return new SqlConnection(sqlConnectString);
        }

        public static SqlConnection GetConnection(string connectionString)
        {
            string sqlConnectString = ConfigurationManager.ConnectionStrings[connectionString].ConnectionString;
            return new SqlConnection(sqlConnectString);
        }
    }
}
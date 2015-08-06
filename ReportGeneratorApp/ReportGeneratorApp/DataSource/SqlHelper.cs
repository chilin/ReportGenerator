using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using ReportGeneratorApp.Report;

namespace ReportGeneratorApp.DataSource
{
    public class SqlHelper
    {
        #region 校验数据是否为数字字符
        public static bool IsNumeric(string str)
        {
            if (str == null || str.Length == 0)
                return false;
            System.Text.ASCIIEncoding ascii = new System.Text.ASCIIEncoding();
            byte[] bytestr = ascii.GetBytes(str);
            foreach (byte c in bytestr)
            {
                if (c < 48 || c > 57)
                {
                    return false;
                }
            }
            return true;
        }
        #endregion

        public static object getValue(string sqlstring)
        {
            using (SqlConnection connection = ConnectionFactory.GetConnection())
            {
                object obj = ReportTemplate.ExecuteScalar(connection, sqlstring);
                if (obj != null && IsNumeric(obj.ToString()))
                {
                    return Convert.ToInt32(obj);
                }
                else
                {
                    return (string)obj;
                }
            }            
        }
        public static object getValue(string sqlstring,string[] param)
        {
            sqlstring = string.Format(sqlstring, param);
            using (SqlConnection connection = ConnectionFactory.GetConnection())
            {
                object obj = ReportTemplate.ExecuteScalar(connection, sqlstring);
                if (obj != null && IsNumeric(obj.ToString()))
                {
                    return Convert.ToInt32(obj);
                }
                else
                {
                    return (string)obj;
                }
            }
        }

        public static SqlParameter[] BuildRowToColParameterFromConfig(Dictionary<string, object> paramList)
        {
            int index =0;
            Dictionary<int, SqlParameter> sqlParamList = new Dictionary<int, SqlParameter>();
            SqlParameter oracleParam = new SqlParameter();
            int paramCount =Convert.ToInt32(paramList["paranmerCount"]);
            SqlParameter[] sqlParams = new SqlParameter[paramCount];
            
            foreach (var paramsElement in (IEnumerable<XElement>)paramList["paramsList"])
            {
                if (paramsElement.Name != "parameters") continue;
                foreach (var element in paramsElement.Elements())
                {
                    if (element.Attribute("name") == null) continue;
                    if (element.Attribute("value") == null) continue;

                    if (element.Name == "parameter" + (index+1))
                    {
                        oracleParam = new SqlParameter();
                        oracleParam.ParameterName = element.Attribute("name").Value.ToString();
                        oracleParam.SqlDbType = SqlDbType.VarChar;
                        oracleParam.Value = element.Attribute("value").Value;                       
                        sqlParams[index++] = oracleParam;
                    }
                }
            }
            return sqlParams;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ReportGeneratorApp.Excel;
using ReportGeneratorApp.DataSource;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using ReportGeneratorApp.Log;

namespace ReportGeneratorApp.Report.Generator.Request
{
    public class FillParameterBuilder
    {
        private FillParameter fillParameter = new FillParameter();
        private DateTime reportDate;
        private DateTime startDate, endDate;
        private DateType dateType;
        Dictionary<string, object> ParamList = new Dictionary<string, object>();

        public FillParameterBuilder(DateTime targetReportDate, string key)
        {
            reportDate = targetReportDate;
            dateType = DateType.Fixed;
            ParamList = XmlHelper.GetFillParams("RequestReport", key);
            SheetDataAdapter.BuildFillParameterFromConfig(fillParameter, ParamList);
        }

        public FillParameterBuilder(DateTime startDate, DateTime endDate, string key)
        {
            this.startDate = startDate;
            this.endDate = endDate;
            dateType = DateType.Custom;
            ParamList = XmlHelper.GetFillParams("RequestReport", key);
            SheetDataAdapter.BuildFillParameterFromConfig(fillParameter, ParamList);
        }

        public FillParameter ToFillParameter1()
        {
            fillParameter.QueryBuilderDelegate = this.QueryBuilder;
            fillParameter.RowHandlerDelegate = RowHandler1;
            return fillParameter;
        }

        public FillParameter ToFillParameter2()
        {
            fillParameter.QueryBuilderDelegate = this.QueryBuilder;
            fillParameter.RowHandlerDelegate = RowHandler2;
            return fillParameter;
        }

        private SqlCommand QueryBuilder(SqlConnection connection, string tableName, string[] columnNameArray, string where)
        {
            string sqlSelect = string.Empty;
            if (dateType == DateType.Fixed)
            {
                sqlSelect = string.Format(XmlHelper.GetSql("RequestReport", ParamList["sqlString"].ToString()), DbWhere.GetWhereString(ReportType.Weekly, reportDate));
            }
            else
            {
                sqlSelect = string.Format(XmlHelper.GetSql("RequestReport", ParamList["sqlString"].ToString()), DbWhere.GetWhereString(startDate, endDate));
            }
            SqlCommand command = new SqlCommand(sqlSelect, connection);
            return command;
        }

        private object RowHandler1(object[,] targetArray, SqlDataReader dataReader, int rowIndex, int recordIndex, string[] columnNameArray)
        {
            var columnsCount = columnNameArray.Length;
            int cellColumnIndex = 0;

            try
            {
                for (cellColumnIndex = 0; cellColumnIndex < columnsCount; cellColumnIndex++)
                {
                    var obj = dataReader[cellColumnIndex];
                    if (cellColumnIndex == 10)
                    {
                        string common = HtmlHelper.NoHTML(obj.ToString());
                        if(common.Length > 100)
                        {
                            common = common.Substring(0, 100);
                        }
                        targetArray[rowIndex, cellColumnIndex] = common;
                    }
                    else
                    {
                        targetArray[rowIndex, cellColumnIndex] = obj;
                    }
                }
                //formula
                targetArray[rowIndex, 2] = string.Format("=WEEKNUM(TODAY())-WEEKNUM({0})", SheetDataAdapter.CalExcelColumnName(1) + (fillParameter.RowOffset + recordIndex + 1));
                targetArray[rowIndex, 3] = string.Format("=TODAY()-{0}", SheetDataAdapter.CalExcelColumnName(1) + (fillParameter.RowOffset + recordIndex + 1));
                targetArray[rowIndex, 4] = string.Format(@"=IF({0}>25,""25+"", IF({0}>15,""15+"", IF({0}>10, ""10+"",IF({0}>5,""5+"",IF({0}>3,""3+"")))))", SheetDataAdapter.CalExcelColumnName(3) + (fillParameter.RowOffset + recordIndex + 1));
            }
            catch
            {
                Logger.Info(string.Format("RowHandler Info: RowIndex:{0}, RecordIndex:{1}, ColumnName:{2}", rowIndex, recordIndex + 1, columnNameArray[cellColumnIndex]));
                throw;
            }
            return null;
        }

        private object RowHandler2(object[,] targetArray, SqlDataReader dataReader, int rowIndex, int recordIndex, string[] columnNameArray)
        {
            var columnsCount = columnNameArray.Length;
            int cellColumnIndex = 0;

            try
            {
                for (cellColumnIndex = 0; cellColumnIndex < columnsCount; cellColumnIndex++)
                {
                    var obj = dataReader[cellColumnIndex];
                    if (cellColumnIndex == 7)
                    {
                        string common = HtmlHelper.NoHTML(obj.ToString());
                        if (common.Length > 100)
                        {
                            common = common.Substring(0, 100);
                        }
                        targetArray[rowIndex, cellColumnIndex] = common;
                    }
                    else
                    {
                        targetArray[rowIndex, cellColumnIndex] = obj;
                    }
                }
            }
            catch
            {
                Logger.Info(string.Format("RowHandler Info: RowIndex:{0}, RecordIndex:{1}, ColumnName:{2}", rowIndex, recordIndex + 1, columnNameArray[cellColumnIndex]));
                throw;
            }
            return null;
        }
    }
}

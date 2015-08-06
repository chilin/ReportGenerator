using System;
using System.Data;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Web;
using ReportGeneratorApp.DataSource;
using ReportGeneratorApp.Report;
using Microsoft.Office.Interop.Excel;
using ReportGeneratorApp.CustomizeException;
using ReportGeneratorApp.Log;

namespace ReportGeneratorApp.Excel
{
    public class SheetDataAdapter
    {
        public static void AddSheet(Workbook book, int number = 1)
        {
            book.Sheets.Add(Type.Missing, book.Sheets[book.Sheets.Count], number, Type.Missing);
        }

        //0 based index for sheet range calculation
        //A1 B4 
        //00 31
        public static Range CalRange(Worksheet worksheet, int x1, int y1, int x2, int y2)
        {
            if (x1 < 0 || y1 < 0 || x2 < 0 || y2 < 0)
            {
                return null;
            }
            Debug.WriteLine(CalExcelColumnName(y1) + (x1 + 1) + " " + CalExcelColumnName(y2) + (x2 + 1));
            return worksheet.Range[CalExcelColumnName(y1) + (x1 + 1), CalExcelColumnName(y2) + (x2 + 1)];
        }


        public static string CalExcelColumnName(int index)
        {
            const int alphaCount = 26;
            const char columnA = 'A';

            //when the return value is "ZZ" (index = alphaCount * (alphaCount + 1) - 1 = 701) index = 701 is the maximum allowed
            if (index < 0 || index > alphaCount * (alphaCount + 1) - 1)
            {
                throw new ArgumentException(string.Format("index value is less than 0 or index value is lager than {0}",
                                                          alphaCount*(alphaCount + 1) - 1));
            }

            //When the index equal 26, should return to AA, so change the "=" to ">=" includes the case of index = 26
            if (index >= alphaCount)
            {
                return string.Format("{0}{1}", (char) (columnA + (index - alphaCount)/alphaCount), (char) (columnA + index%alphaCount));
            }
            return string.Format("{0}", (char) (columnA + index%alphaCount));
        }

        public static void BuildFillParameterFromConfig(FillParameter fillParameter, Dictionary<string, object> paramList)
        {
            if(paramList.ContainsKey("sheetIndex"))
            {
                fillParameter.SheetIndex = Convert.ToInt32(paramList["sheetIndex"]);
            }
            if(paramList.ContainsKey("rowOffset"))
            {
                fillParameter.RowOffset = Convert.ToInt32(paramList["rowOffset"]);
            }
            if (paramList.ContainsKey("columnOffset"))
            {
                fillParameter.ColumnOffset = Convert.ToInt32(paramList["columnOffset"]);
            }
            if (paramList.ContainsKey("columnCount"))
            {
                fillParameter.ColumnNameArray = new string[Convert.ToInt32(paramList["columnCount"])];
            }
            if (paramList.ContainsKey("tableName"))
            {
                fillParameter.TableName = paramList["tableName"].ToString();
            }
            if (paramList.ContainsKey("batchSize"))
            {
                fillParameter.BatchSize = Convert.ToInt32(paramList["batchSize"]);
            }
            if (paramList.ContainsKey("customizeBatch"))
            {
                fillParameter.CustomizeBatch = Convert.ToBoolean(paramList["customizeBatch"]);
            }
            if (paramList.ContainsKey("isNeedQueryData"))
            {
                fillParameter.IsNeedQueryData = Convert.ToBoolean(paramList["isNeedQueryData"]);
            }
        }

        public static Worksheet FillSheetRange(Workbook book, FillParameter fillParameter)
        {
            if (fillParameter.IsNeedQueryData)
            {
                //parameter validation check
                if (fillParameter.RowHandlerDelegate == null)
                {
                    throw new Exception("RowHandlerDelegate must be initialized.");
                }
                if (fillParameter.QueryBuilderDelegate == null)
                {
                    throw new Exception("QueryBuilderDelegate must be initialized.");
                }
            }

            Worksheet sheet = book.Sheets[fillParameter.SheetIndex];
            sheet.Activate();

            if (fillParameter.IsNeedQueryData)
            {
                using (SqlConnection connection = ConnectionFactory.GetConnection(fillParameter.connectionString)
                    )
                {
                    Logger.Debug("Start Query Data");
                    SqlCommand command = fillParameter.QueryBuilderDelegate(connection, fillParameter.TableName,
                                                                                fillParameter.ColumnNameArray,
                                                                                fillParameter.WhereClause);

                    Logger.Debug(string.Format("Query String : {0}", command.CommandText));

                    connection.Open();
                    SqlDataReader odr = null;

                    // TODO: 暂时不支持存储过程
                    //if (command.CommandType == CommandType.StoredProcedure)
                    //{
                    //    SqlParameter resultList = new SqlParameter("P_LIST", OracleType.Cursor);
                    //    resultList.Direction = System.Data.ParameterDirection.Output;
                    //    command.Parameters.Add(resultList);

                    //    command.ExecuteNonQuery();
                    //    odr = (SqlDataReader) resultList.Value;

                    //}
                    //else
                    //{
                        odr = command.ExecuteReader(CommandBehavior.CloseConnection);
                    //}
                    Logger.Debug("End Query Data");


                    var startColumn = SheetDataAdapter.CalExcelColumnName(fillParameter.ColumnOffset);
                    var endColumn =
                        SheetDataAdapter.CalExcelColumnName(fillParameter.ColumnOffset +
                                                            fillParameter.ColumnNameArray.Count() - 1);

                    var columnsCount = fillParameter.ColumnNameArray.Count();

                    var cachedResult = new object[fillParameter.BatchSize,columnsCount];

                    var batch = 0;
                    var recordIndex = 0;
                    var rowIndex = 0;

                    Logger.Debug("Start Fill Data");
                    if (!fillParameter.CustomizeBatch) //把是否自定义处理判断提到读取数据之前，否则会丢失第一条数据
                    {
                        try
                        {
                            while (odr.Read())
                            {

                                fillParameter.RowHandlerDelegate(cachedResult, odr, recordIndex%fillParameter.BatchSize,
                                                                 recordIndex, fillParameter.ColumnNameArray);
                                recordIndex++;

                                batch = recordIndex/fillParameter.BatchSize;
                                rowIndex = fillParameter.PerRecordPerRow ? recordIndex : (recordIndex/columnsCount) + 1;

                                if (rowIndex%fillParameter.BatchSize == 0)
                                {
                                    var tempRange = sheet.Range[
                                        startColumn +
                                        (fillParameter.RowOffset + 1 + fillParameter.BatchSize*(batch - 1)),
                                        endColumn + (fillParameter.BatchSize*batch + fillParameter.RowOffset)];
                                    tempRange.Value2 = cachedResult;
                                    cachedResult = new object[fillParameter.BatchSize,columnsCount];
                                }
                                //row index 1 based instead of 0
                            }
                        }
                        catch
                        {
                            Logger.Info("RecordIndex:" + recordIndex);
                            throw;
                        }
                    }
                    else
                    {
                        fillParameter.RowHandlerDelegate(cachedResult, odr, recordIndex%fillParameter.BatchSize,
                                                         recordIndex, fillParameter.ColumnNameArray);
                    }

                    if (!fillParameter.CustomizeBatch)
                    {
                        if (rowIndex%fillParameter.BatchSize != 0)
                        {
                            var finalRange = sheet.Range[
                                startColumn + (fillParameter.RowOffset + 1 + fillParameter.BatchSize*batch),
                                endColumn +
                                ((fillParameter.BatchSize*batch + fillParameter.RowOffset) +
                                 rowIndex%fillParameter.BatchSize)];
                            var shrinkedArray = new object[rowIndex%fillParameter.BatchSize,columnsCount];
                            for (int i = 0; i < shrinkedArray.GetLength(0); i++)
                            {
                                for (int j = 0; j < shrinkedArray.GetLength(1); j++)
                                {
                                    shrinkedArray[i, j] = cachedResult[i, j];
                                }
                            }
                            finalRange.Value2 = shrinkedArray;
                        }
                    }
                    else
                    {
                        var finalRange = CalRange(sheet, fillParameter.RowOffset, fillParameter.ColumnOffset,
                                                  fillParameter.RowOffset + fillParameter.BatchSize - 1,
                                                  fillParameter.ColumnOffset + fillParameter.ColumnNameArray.Count() -
                                                  1);
                        finalRange.Value2 = cachedResult;

                    }
                    Logger.Debug("End Fill Data");
                }

            }
            if (fillParameter.UpdateSheetDelegate != null)
            {
                Logger.Debug("Start Update Sheet");
                fillParameter.UpdateSheetDelegate(sheet);
                Logger.Debug("End Update Sheet");
            }
            //if (rowIndex <= 0)
            //{
            //    throw new NoResultException("No data retrieved from [" + fillParameter.WhereClause+"]");
            //}
            
            return sheet;
        }
    }
}
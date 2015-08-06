using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class SimpleChart : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("FillParameter"))
            {
                throw new ArgumentException("FillParameter");
            }
            if (!paramList.ContainsKey("chartID"))
            {
                throw new ArgumentException("chartID");
            }
            if (!paramList.ContainsKey("plotBy"))
            {
                throw new ArgumentException("plotBy");
            }

            FillParameter parameter = (FillParameter)paramList["FillParameter"];

            int sheetIndex;
            if (paramList.ContainsKey("sheetIndex"))
            {
                sheetIndex = Convert.ToInt32(paramList["sheetIndex"]);
            }
            else
            {
                sheetIndex = parameter.SheetIndex;
            }

            Worksheet sheet = book.Sheets[parameter.SheetIndex];
            Chart chart = null;
            if (book.Sheets[sheetIndex] is Worksheet)
            {
                Worksheet chartSheet = book.Sheets[sheetIndex];
                chart = ((ChartObject)chartSheet.ChartObjects(paramList["chartID"])).Chart;
            }
            else
            {
                chart = book.Sheets[sheetIndex];
            }
            
            int leftOffset = 0;
            int rightOffset = 0;
            int topOffset = 0;
            int bottomOffset = 0;
            if (paramList.ContainsKey("leftOffset"))
            {
                leftOffset = Convert.ToInt32(paramList["leftOffset"]);
            }
            if (paramList.ContainsKey("rightOffset"))
            {
                rightOffset = Convert.ToInt32(paramList["rightOffset"]);
            }
            if (paramList.ContainsKey("topOffset"))
            {
                topOffset = Convert.ToInt32(paramList["topOffset"]);
            }
            if (paramList.ContainsKey("bottomOffset"))
            {
                bottomOffset = Convert.ToInt32(paramList["bottomOffset"]);
            }

            int lastRow = parameter.BatchSize > 100
                                      ? sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).
                                            Row - 1
                                      : parameter.RowOffset + parameter.BatchSize - 1;
            Range targetRange = SheetDataAdapter.CalRange(sheet, parameter.RowOffset + topOffset,
                                                                    parameter.ColumnOffset + leftOffset,
                                                                    lastRow + bottomOffset,
                                                                    parameter.ColumnOffset +
                                                                    parameter.ColumnNameArray.Length - 1 + rightOffset);
            chart.SetSourceData(targetRange, (string)paramList["plotBy"] == "row" ? XlRowCol.xlRows : XlRowCol.xlColumns);
            return null;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class HideRange : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("FillParameter"))
            {
                throw new ArgumentException("FillParameter");
            }

            if (!paramList.ContainsKey("rangeType"))
            {
                throw new ArgumentException("rangeType");
            }

            if (!paramList.ContainsKey("position"))
            {
                throw new ArgumentException("position");
            }

            int leftOffset = 0;
            int rightOffset = 0;
            int topOffset = 0;
            int bottomOffset = 0;
            int size = 0;
            int headerSize = 1;
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
            if (paramList.ContainsKey("size"))
            {
                size = Convert.ToInt32(paramList["size"]);
            }
            if (paramList.ContainsKey("headerSize"))
            {
                headerSize = Convert.ToInt32(paramList["headerSize"]);
            }

            FillParameter parameter = (FillParameter)paramList["FillParameter"];
            Worksheet sheet = book.Sheets[parameter.SheetIndex];

            Range targetRange = null;
            int lastRow = parameter.BatchSize > 100
                                      ? sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).
                                            Row - 1
                                      : parameter.RowOffset + parameter.BatchSize - 1;
            if (lastRow < parameter.RowOffset)
            {
                lastRow = parameter.RowOffset;
            }

            switch ((string)paramList["rangeType"])
            {
                case "row":
                    switch ((string)paramList["position"])
                    {
                        case "top":
                            targetRange = SheetDataAdapter.CalRange(sheet, headerSize,
                                                                    parameter.ColumnOffset + leftOffset,
                                                                    parameter.RowOffset + topOffset  - 1,
                                                                    parameter.ColumnOffset +
                                                                    parameter.ColumnNameArray.Length - 1 + rightOffset);
                            break;
                        case "bottom":
                            targetRange = SheetDataAdapter.CalRange(sheet, headerSize,
                                                                    parameter.ColumnOffset + leftOffset,
                                                                    lastRow + bottomOffset - size,
                                                                    parameter.ColumnOffset +
                                                                    parameter.ColumnNameArray.Length - 1 + rightOffset);
                            break;
                    }
                    break;
                case "column":
                    switch ((string)paramList["position"])
                    {
                        case "left":
                            targetRange = SheetDataAdapter.CalRange(sheet, parameter.RowOffset + topOffset,
                                                                    headerSize,
                                                                    lastRow + bottomOffset,
                                                                    parameter.ColumnOffset + leftOffset - 1);
                            break;
                        case "right":
                            targetRange = SheetDataAdapter.CalRange(sheet, parameter.RowOffset + topOffset,
                                                                    headerSize,
                                                                    lastRow + bottomOffset,
                                                                    parameter.ColumnOffset +
                                                                    parameter.ColumnNameArray.Length - 1 + rightOffset -
                                                                    size);
                            break;
                    }
                    break;
            }

            if (targetRange == null) return null;

            switch ((string)paramList["rangeType"])
            {
                case "row":
                    targetRange.Rows.EntireRow.Hidden = true;
                    break;
                case "column":
                    targetRange.Columns.EntireColumn.Hidden = true;
                    break;
            }

            return null;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class DrawLine : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if(!paramList.ContainsKey("FillParameter"))
            {
                throw new ArgumentException("FillParameter");
            }
            FillParameter parameter = (FillParameter) paramList["FillParameter"];
            Worksheet sheet = book.Sheets[parameter.SheetIndex];
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
            if (parameter.RowOffset + topOffset > lastRow + bottomOffset)
            {
                return null;
            }
            Range targetRange = SheetDataAdapter.CalRange(sheet, parameter.RowOffset + topOffset,
                                                          parameter.ColumnOffset + leftOffset,
                                                          lastRow + bottomOffset,
                                                          parameter.ColumnOffset + parameter.ColumnNameArray.Length - 1 +
                                                          rightOffset);
            targetRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            if (paramList.ContainsKey("color"))
            {
                string[] rgb = paramList["color"].ToString().Split(new char[] {','});
                targetRange.Borders.Color = Color.FromArgb(Convert.ToInt32(rgb[0]), Convert.ToInt32(rgb[1]), Convert.ToInt32(rgb[2]));
            }
            return null;
        }
    }
}
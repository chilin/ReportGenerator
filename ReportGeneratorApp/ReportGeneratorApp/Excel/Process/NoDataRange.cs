using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class NoDataRange : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("sheetIndex"))
            {
                throw new ArgumentException("sheetIndex");
            }

            if (!paramList.ContainsKey("rangeType"))
            {
                throw new ArgumentException("rangeType");
            }

            if (!paramList.ContainsKey("startRow"))
            {
                throw new ArgumentException("startRow");
            }

            if (!paramList.ContainsKey("startColumn"))
            {
                throw new ArgumentException("startColumn");
            }

            Worksheet sheet = book.Sheets[(int)paramList["sheetIndex"]];

            Range targetRange = null;
            string length = "max";
            int size = 0;

            if (paramList.ContainsKey("max"))
            {
                length = paramList["max"].ToString();
            }

            if (paramList.ContainsKey("size"))
            {
                size = Convert.ToInt32(paramList["size"]) - 1;
            }

            int startRow = (int) paramList["startRow"];
            int startColumn = (int) paramList["startColumn"];

            switch ((string)paramList["rangeType"])
            {
                case "row":
                    targetRange = SheetDataAdapter.CalRange(
                        sheet,
                        startRow,
                        startColumn,
                        startRow + size,
                        length == "max"
                            ? sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column - 1
                            : startColumn + Convert.ToInt32(length));
                    break;
                case "column":
                    targetRange = SheetDataAdapter.CalRange(
                        sheet,
                        startRow,
                        startColumn,
                        length == "max"
                            ? sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row - 1
                            : startRow + Convert.ToInt32(length),
                        startColumn + size);
                    break;
            }

            if (targetRange == null) return null;

            //format
            if (paramList.ContainsKey("format"))
            {
                ProcessHelper.FormatRange(targetRange, (string)paramList["format"]);
            }

            return null;
        }
    }
}
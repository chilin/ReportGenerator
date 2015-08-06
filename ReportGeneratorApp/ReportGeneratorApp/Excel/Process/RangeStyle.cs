using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class RangeStyle : IProcessable
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

            FillParameter parameter = (FillParameter)paramList["FillParameter"];
            Worksheet sheet = book.Sheets[parameter.SheetIndex];

            Range targetRange = (Range)ProcessHelper.GetRange(parameter, sheet, paramList);
            if (targetRange == null) return null;

            //format
            if (paramList.ContainsKey("format"))
            {
                ProcessHelper.FormatRange(targetRange, (string) paramList["format"]);
            }

            //merge
            if (paramList.ContainsKey("merge") && Convert.ToBoolean(paramList["merge"]))
            {
                targetRange.Merge();
            }
            return null;
        }
    }
}
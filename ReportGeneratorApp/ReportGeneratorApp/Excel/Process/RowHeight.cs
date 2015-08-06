using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class RowHeight : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("sheetIndex"))
            {
                throw new ArgumentException("sheetIndex");
            }

            double height = 13.5;

            if (paramList.ContainsKey("height"))
            {
                height = Convert.ToDouble(paramList["height"]);
            }

            Worksheet sheet = book.Sheets[Convert.ToInt32(paramList["sheetIndex"])];
            sheet.Activate();
            sheet.Rows.RowHeight = height;

            return null;
        }
    }
}
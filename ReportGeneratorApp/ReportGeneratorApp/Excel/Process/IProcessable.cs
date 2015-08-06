using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public interface IProcessable
    {
        object Process(Workbook book, Dictionary<string, object> paramList);
    }
}
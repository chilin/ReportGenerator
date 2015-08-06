using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class ColumnWidth : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            int[] exceptSheet = null;
            Hashtable exceptColumn = null;
            double width = 8.43;
            if (paramList.ContainsKey("exceptSheet"))
            {
                exceptSheet =
                    Array.ConvertAll(
                        paramList["exceptSheet"].ToString().Split(new char[] {','},
                                                                  StringSplitOptions.RemoveEmptyEntries),
                        element => Convert.ToInt32(element));
            }
            if (paramList.ContainsKey("exceptColumn"))
            {
                string[] exceptColumns = paramList["exceptColumn"].ToString().Split(new char[] {'|'},
                                                                          StringSplitOptions.RemoveEmptyEntries);
                exceptColumn = new Hashtable();
                foreach (string s in exceptColumns)
                {
                    exceptColumn.Add(Convert.ToInt32(s.Substring(0, s.IndexOf(':'))), s.Substring(s.IndexOf(':') + 1));
                }
            }
            if (paramList.ContainsKey("width"))
            {
                width = Convert.ToDouble(paramList["width"]);
            }
            for (int i = 1; i <= book.Sheets.Count; i++)
            {
                if (book.Sheets[i] is Worksheet)
                {
                    Worksheet sheet = book.Sheets[i];
                    if (sheet.Visible != XlSheetVisibility.xlSheetVisible) continue;
                    sheet.Activate();
                    sheet.get_Range("A1").Select();
                    if (exceptSheet == null || !exceptSheet.Contains(sheet.Index))
                    {
                        sheet.Columns.EntireColumn.AutoFit();
                    }
                    if (exceptColumn != null && exceptColumn.ContainsKey(sheet.Index))
                    {
                        string[] columns = exceptColumn[sheet.Index].ToString().Split(new char[] { ',' },
                                                                                      StringSplitOptions.RemoveEmptyEntries);
                        foreach (string column in columns)
                        {
                            Range columnRange = sheet.Columns[column];
                            columnRange.ColumnWidth = width;
                        }
                    }
                }
            }
            return null;
        }
    }
}
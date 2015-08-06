using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class CellValue : IProcessable
    {

        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("sheetIndex"))
            {
                throw new ArgumentException("sheetIndex");
            }

            if (!paramList.ContainsKey("row"))
            {
                throw new ArgumentException("row");
            }

            if (!paramList.ContainsKey("column"))
            {
                throw new ArgumentException("column");
            }

            if (!paramList.ContainsKey("value"))
            {
                throw new ArgumentException("value");
            }

            Worksheet sheet = book.Sheets[Convert.ToInt32(paramList["sheetIndex"])];
            int row = Convert.ToInt32(paramList["row"]);
            int column = Convert.ToInt32(paramList["column"]);
            Range cell = SheetDataAdapter.CalRange(sheet, row, column, row, column);
            System.Globalization.CultureInfo en_us = System.Globalization.CultureInfo.GetCultureInfo("en-US");
            string value = string.Format(en_us, (string)paramList["value"], (DateTime)paramList["StartDate"], (DateTime)paramList["EndDate"]).Replace("\\r", "\r");
            cell.Value = value;
            
            MatchCollection mc = Regex.Matches(value, @"\[\#(?<1>.+?)\#\]");

            //先对ChartTitle赋值
            foreach (Match m in mc)
            {
                GroupCollection gc = m.Groups;
                string source = gc["1"].Value.Trim();
                string target = source.Substring(0, source.IndexOf('|'));
                cell.Value = cell.Value.Replace(m.Value, target);
            }

            foreach (Match m in mc)
            {
                GroupCollection gc = m.Groups;
                string source = gc["1"].Value.Trim();
                string target = source.Substring(0, source.IndexOf('|'));
                int start = value.IndexOf("[#");
                int end = m.Index + target.Length - 1;
                value = value.Replace(m.Value, target);
                string[] param = source.Substring(source.IndexOf('|') + 1).Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in param)
                {
                    switch (p)
                    {
                        case "B":
                            cell.Characters[start, end].Font.Bold = true;
                            break;
                        case "U":
                            cell.Characters[start, end].Font.Underline = true;
                            break;
                        case "I":
                            cell.Characters[start, end].Font.Italic = true;
                            break;
                        case "b":
                            cell.Characters[start, end].Font.Bold = false;
                            break;
                        case "u":
                            cell.Characters[start, end].Font.Underline = false;
                            break;
                        case "i":
                            cell.Characters[start, end].Font.Italic = false;
                            break;
                        default:
                            string[] rgb = p.Split(new char[] { ',' });
                            cell.Characters[start, end].Font.Color = Color.FromArgb(Convert.ToInt32(rgb[0]), Convert.ToInt32(rgb[1]), Convert.ToInt32(rgb[2]));
                            break;
                    }
                }
            }
            return null;
        }
    }
}
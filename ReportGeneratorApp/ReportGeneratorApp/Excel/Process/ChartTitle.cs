using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class ChartTitle : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("sheetIndex"))
            {
                throw new ArgumentException("sheetIndex");
            }

            if (!paramList.ContainsKey("chartID"))
            {
                throw new ArgumentException("chartID");
            }

            if (!paramList.ContainsKey("title"))
            {
                throw new ArgumentException("title");
            }

            Worksheet sheet = book.Sheets[Convert.ToInt32(paramList["sheetIndex"])];
            ChartObject chart = sheet.ChartObjects(paramList["chartID"]);
            System.Globalization.CultureInfo en_us = System.Globalization.CultureInfo.GetCultureInfo("en-US");
            string title = string.Format(en_us, (string)paramList["title"], (DateTime)paramList["StartDate"], (DateTime)paramList["EndDate"]).Replace("\\r", "\r");
            chart.Chart.ChartTitle.Text = title;
            MatchCollection mc = Regex.Matches(title, @"\[\#(?<1>.+?)\#\]");
            
            //先对ChartTitle赋值
            foreach (Match m in mc)
            {
                GroupCollection gc = m.Groups;
                string source = gc["1"].Value.Trim();
                string target = source.Substring(0, source.IndexOf('|'));
                chart.Chart.ChartTitle.Text = chart.Chart.ChartTitle.Text.Replace(m.Value, target);
            }

            foreach (Match m in mc)
            {
                GroupCollection gc = m.Groups;
                string source = gc["1"].Value.Trim();
                string target = source.Substring(0, source.IndexOf('|'));
                int start = title.IndexOf("[#");
                int end = m.Index + target.Length - 1;
                title = title.Replace(m.Value, target);
                string[] param = source.Substring(source.IndexOf('|') + 1).Split(new [] {'|'}, StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in param)
                {
                    switch (p)
                    {
                        case "B":
                            chart.Chart.ChartTitle.Characters[start, end].Font.Bold = true;
                            break;
                        case "U":
                            chart.Chart.ChartTitle.Characters[start, end].Font.Underline = true;
                            break;
                        case "I":
                            chart.Chart.ChartTitle.Characters[start, end].Font.Italic = true;
                            break;
                        case "b":
                            chart.Chart.ChartTitle.Characters[start, end].Font.Bold = false;
                            break;
                        case "u":
                            chart.Chart.ChartTitle.Characters[start, end].Font.Underline = false;
                            break;
                        case "i":
                            chart.Chart.ChartTitle.Characters[start, end].Font.Italic = false;
                            break;
                        default:
                            string[] rgb = p.Split(new char[] {','});
                            chart.Chart.ChartTitle.Characters[start, end].Font.Color = Color.FromArgb(Convert.ToInt32(rgb[0]), Convert.ToInt32(rgb[1]), Convert.ToInt32(rgb[2]));
                            break;
                    }
                }
            }
            return null;
        }
    }
}
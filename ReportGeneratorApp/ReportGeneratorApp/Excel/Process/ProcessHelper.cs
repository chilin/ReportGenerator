using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class ProcessHelper
    {
        public static object GetRange(FillParameter parameter, Worksheet sheet, XElement element)
        {
            Dictionary<string, object> paramList = new Dictionary<string, object>();
            foreach (var attribute in element.Attributes())
            {
                paramList.Add(attribute.Name.LocalName, attribute.Value);
            }
            return GetRange(parameter, sheet, paramList);
        }

        public static object GetRange(FillParameter parameter, Worksheet sheet, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("rangeType"))
            {
                throw new ArgumentException("rangeType");
            }

            if (!paramList.ContainsKey("position"))
            {
                throw new ArgumentException("position");
            }

            if((string)paramList["rangeType"] == "cell")
            {
                return sheet.Range[paramList["position"], paramList["position"]];
            }

            int leftOffset = 0;
            int rightOffset = 0;
            int topOffset = 0;
            int bottomOffset = 0;
            int size = 0;
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
                size = Convert.ToInt32(paramList["size"]) - 1;
            }

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
                            targetRange = SheetDataAdapter.CalRange(sheet, parameter.RowOffset + topOffset,
                                                                    parameter.ColumnOffset + leftOffset,
                                                                    parameter.RowOffset + topOffset + size,
                                                                    parameter.ColumnOffset +
                                                                    parameter.ColumnNameArray.Length - 1 + rightOffset);
                            break;
                        case "bottom":
                            targetRange = SheetDataAdapter.CalRange(sheet, lastRow + bottomOffset,
                                                                    parameter.ColumnOffset + leftOffset,
                                                                    lastRow + bottomOffset + size,
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
                                                                    parameter.ColumnOffset + leftOffset,
                                                                    lastRow + bottomOffset,
                                                                    parameter.ColumnOffset + leftOffset + size);
                            break;
                        case "right":
                            targetRange = SheetDataAdapter.CalRange(sheet, parameter.RowOffset + topOffset,
                                                                    parameter.ColumnOffset +
                                                                    parameter.ColumnNameArray.Length - 1 + rightOffset,
                                                                    lastRow + bottomOffset,
                                                                    parameter.ColumnOffset +
                                                                    parameter.ColumnNameArray.Length - 1 + rightOffset +
                                                                    size);
                            break;
                    }
                    break;
            }

            return targetRange;
        }

        public static object FormatRange(Range range, string formatString)
        {
            string[] formats = formatString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string format in formats)
            {
                string[] pair = format.Trim().Split(new[] {':'}, StringSplitOptions.RemoveEmptyEntries);
                switch (pair[0])
                {
                    case "bgColor":
                        string[] bgColor = pair[1].Split(new[] { ',' });
                        range.Interior.Color = Color.FromArgb(Convert.ToInt32(bgColor[0]), Convert.ToInt32(bgColor[1]),
                                                              Convert.ToInt32(bgColor[2]));
                        break;
                    case "font":
                        string[] param = pair[1].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var p in param)
                        {
                            switch (p)
                            {
                                case "B":
                                    range.Font.Bold = true;
                                    break;
                                case "U":
                                    range.Font.Underline = true;
                                    break;
                                case "I":
                                    range.Font.Italic = true;
                                    break;
                                case "b":
                                    range.Font.Bold = false;
                                    break;
                                case "u":
                                    range.Font.Underline = false;
                                    break;
                                case "i":
                                    range.Font.Italic = false;
                                    break;
                                default:
                                    string[] rgb = p.Split(new[] {','});
                                    range.Font.Color = Color.FromArgb(Convert.ToInt32(rgb[0]), Convert.ToInt32(rgb[1]),
                                                                      Convert.ToInt32(rgb[2]));
                                    break;
                            }
                        }
                        break;
                    case "hAlign":
                        switch (pair[1])
                        {
                            case "center":
                                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                break;
                            case "left":
                                range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                break;
                            case "right":
                                range.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                break;
                        }
                        break;
                    case "vAlign":
                        switch (pair[1])
                        {
                            case "center":
                                range.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                break;
                            case "top":
                                range.VerticalAlignment = XlVAlign.xlVAlignTop;
                                break;
                            case "right":
                                range.VerticalAlignment = XlVAlign.xlVAlignBottom;
                                break;
                        }
                        break;
                }
            }
            return null;
        }

        public static object FormatCondition(FormatCondition condition, string formatString)
        {
            string[] formats = formatString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string format in formats)
            {
                string[] pair = format.Trim().Split(new[] {':'}, StringSplitOptions.RemoveEmptyEntries);
                switch (pair[0])
                {
                    case "bgColor":
                        string[] bgColor = pair[1].Split(new[] {','});
                        condition.Interior.Color = Color.FromArgb(Convert.ToInt32(bgColor[0]), Convert.ToInt32(bgColor[1]),
                                                              Convert.ToInt32(bgColor[2]));
                        break;
                    case "font":
                        string[] param = pair[1].Split(new[] {'|'}, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var p in param)
                        {
                            switch (p)
                            {
                                case "B":
                                    condition.Font.Bold = true;
                                    break;
                                case "U":
                                    condition.Font.Underline = true;
                                    break;
                                case "I":
                                    condition.Font.Italic = true;
                                    break;
                                case "b":
                                    condition.Font.Bold = false;
                                    break;
                                case "u":
                                    condition.Font.Underline = false;
                                    break;
                                case "i":
                                    condition.Font.Italic = false;
                                    break;
                                default:
                                    string[] rgb = p.Split(new[] {','});
                                    condition.Font.Color = Color.FromArgb(Convert.ToInt32(rgb[0]), Convert.ToInt32(rgb[1]),
                                                                      Convert.ToInt32(rgb[2]));
                                    break;
                            }
                        }
                        break;
                }
            }
            return null;
        }
    }
}
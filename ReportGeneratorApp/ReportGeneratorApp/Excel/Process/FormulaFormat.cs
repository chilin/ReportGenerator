using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class FormulaFormat : IProcessable
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
            if (!paramList.ContainsKey("SubParams"))
            {
                throw new ArgumentException("SubParams");
            }

            FillParameter parameter = (FillParameter)paramList["FillParameter"];
            Worksheet sheet = book.Sheets[parameter.SheetIndex];

            Range targetRange = (Range)ProcessHelper.GetRange(parameter, sheet, paramList);
            if (targetRange == null) return null;

            for (int x = 1; x <= targetRange.Cells.Count; x++ )
            {
                Range targetCell = targetRange.Cells[x];

                targetCell.FormatConditions.Delete();

                foreach (var formulaElement in (IEnumerable<XElement>)paramList["SubParams"])
                {
                    if (formulaElement.Name != "formula") continue;
                    if (formulaElement.Attribute("expression") == null) continue;
                    if (formulaElement.Attribute("format") == null) continue;
                    string expression = formulaElement.Attribute("expression").Value;
                    string format = formulaElement.Attribute("format").Value;
                    int paramCount = formulaElement.Elements().Count();
                    if (paramCount != 0)
                    {
                        Range[] paramRanges = new Range[paramCount];
                        int i = 0;
                        foreach (var param in formulaElement.Elements())
                        {
                            int rowOffset = 0;
                            int columnOffset = 0;
                            if (param.Attribute("rowOffset") != null)
                            {
                                rowOffset = Convert.ToInt32(param.Attribute("rowOffset").Value);
                            }
                            if (param.Attribute("columnOffset") != null)
                            {
                                columnOffset = Convert.ToInt32(param.Attribute("columnOffset").Value);
                            }
                            if (rowOffset == 0 && columnOffset == 0)
                            {
                                paramRanges[i] = targetCell;
                            }
                            else
                            {
                                paramRanges[i] = targetCell.Offset[rowOffset, columnOffset];
                            }
                            i++;
                        }
                        string[] paramStrings = new string[paramCount];
                        for (int j = 0; j < paramCount; j++)
                        {
                            paramStrings[j] = paramRanges[j].AddressLocal;
                        }
                        FormatCondition condition = targetCell.FormatConditions.Add(XlFormatConditionType.xlExpression,
                                                                                    Missing.Value,
                                                                                    string.Format(
                                                                                        expression.Replace("'", "\""),
                                                                                        paramStrings),
                                                                                    Missing.Value, Missing.Value);
                        ProcessHelper.FormatCondition(condition, format);
                        condition.StopIfTrue = true;
                    }
                    else
                    {
                        FormatCondition condition = targetCell.FormatConditions.Add(XlFormatConditionType.xlExpression,
                                                                                    Missing.Value,
                                                                                    expression.Replace("'", "\""),
                                                                                    Missing.Value, Missing.Value);
                        ProcessHelper.FormatCondition(condition, format);
                        condition.StopIfTrue = true;
                    }
                }
            }
            return null;
        }
    }
}
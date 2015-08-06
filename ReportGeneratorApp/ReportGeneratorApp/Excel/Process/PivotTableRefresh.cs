using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class PivotTableRefresh : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("id"))
            {
                throw new ArgumentException("id");
            }
            if (!paramList.ContainsKey("sheetIndex"))
            {
                throw new ArgumentException("sheetIndex");
            }
            if (!paramList.ContainsKey("sourceSheet"))
            {
                throw new ArgumentException("sourceSheet");
            }

            int leftOffset = 0;
            int rightOffset = 0;
            int topOffset = 0;
            int bottomOffset = 0;
            string sourceTable = string.Empty;
            Hashtable except = null;
            Hashtable include = null;
            Hashtable defaultValues = null;
            Hashtable fieldName = new Hashtable();
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
            if (paramList.ContainsKey("sourceTable"))
            {
                sourceTable = paramList["sourceTable"].ToString();
            }
            if (paramList.ContainsKey("except"))
            {
                except = new Hashtable();
                string[] excepts = ((string)paramList["except"]).Split(new string[] { "|||" },
                                                                        StringSplitOptions.RemoveEmptyEntries);
                foreach (string s in excepts)
                {
                    string key = s.Substring(0, s.IndexOf("||"));
                    string value = s.Substring(s.IndexOf("||") + 2);
                    if(!except.ContainsKey(key))
                    {
                        except.Add(key, value);
                    }
                    if(!fieldName.ContainsKey(key))
                    {
                        fieldName.Add(key, key);
                    }
                }
            }
            if (paramList.ContainsKey("include"))
            {
                include = new Hashtable();
                string[] includes = ((string) paramList["include"]).Split(new string[] {"|||"},
                                                                          StringSplitOptions.RemoveEmptyEntries);
                foreach (string s in includes)
                {
                    string key = s.Substring(0, s.IndexOf("||"));
                    string value = s.Substring(s.IndexOf("||") + 2);
                    if (!include.ContainsKey(key))
                    {
                        include.Add(key, value);
                    }
                    if (!fieldName.ContainsKey(key))
                    {
                        fieldName.Add(key, key);
                    }
                }
            }
            if (paramList.ContainsKey("default"))
            {
                defaultValues = new Hashtable();
                string[] defaults = ((string)paramList["default"]).Split(new string[] { "|||" },
                                                                          StringSplitOptions.RemoveEmptyEntries);
                foreach (string s in defaults)
                {
                    string key = s.Substring(0, s.IndexOf("||"));
                    string value = s.Substring(s.IndexOf("||") + 2);
                    if (!defaultValues.ContainsKey(key))
                    {
                        defaultValues.Add(key, value);
                    }
                    if (!fieldName.ContainsKey(key))
                    {
                        fieldName.Add(key, key);
                    }
                }
            }

            Worksheet sheet = book.Sheets[Convert.ToInt32(paramList["sheetIndex"])];
            Worksheet sourceSheet = book.Sheets[Convert.ToInt32(paramList["sourceSheet"])];
            PivotTable pivotTable = sheet.PivotTables(paramList["id"]);
            if(pivotTable == null) return null;

            if (string.IsNullOrWhiteSpace(sourceTable))
            {
                int lastRow = sourceSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row + bottomOffset;
                int lastColumn = sourceSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column + rightOffset;

                if (lastRow <= topOffset + 1) return null;

                pivotTable.SourceData = string.Format(@"{0}!R{1}C{2}:R{3}C{4}",
                                                      sourceSheet.Name,
                                                      topOffset + 1,
                                                      leftOffset + 1,
                                                      lastRow,
                                                      lastColumn);
            }
            else
            {
                pivotTable.SourceData = string.Format(@"{0}!{1}",
                                                      sourceSheet.Name,
                                                      sourceTable);
            }
            pivotTable.PivotCache().Refresh();
            
            if(fieldName.Count == 0) return null;

            foreach (DictionaryEntry entry in fieldName)
            {
                PivotField field = pivotTable.PivotFields(entry.Key);
                field.ClearAllFilters();
                //field.CurrentPage = "(All)";
                foreach (PivotItem item in field.PivotItems())
                {
                    if(item.RecordCount == 0 && !item.IsCalculated)
                    {
                        item.Delete();
                        continue;
                    }
                    if (except != null && except.ContainsKey(entry.Key) && Regex.IsMatch(item.Name, except[entry.Key].ToString()))
                    {
                        item.Visible = false;
                    }
                    if (include != null && include.ContainsKey(entry.Key) && Regex.IsMatch(item.Name, include[entry.Key].ToString()))
                    {
                        item.Visible = true;
                    }
                    if (defaultValues != null && defaultValues.ContainsKey(entry.Key) && Regex.IsMatch(item.Name, defaultValues[entry.Key].ToString()))
                    {
                        field.CurrentPage = item.Name;
                    }
                }
            }
            return null;
        }
    }
}
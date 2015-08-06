using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;

namespace ReportGeneratorApp.Excel.Process
{
    public class ComplexChart : IProcessable
    {
        public object Process(Workbook book, Dictionary<string, object> paramList)
        {
            if (!paramList.ContainsKey("FillParameter"))
            {
                throw new ArgumentException("FillParameter");
            }
            if (!paramList.ContainsKey("sheetIndex"))
            {
                throw new ArgumentException("sheetIndex");
            }
            if (!paramList.ContainsKey("SubParams"))
            {
                throw new ArgumentException("SubParams");
            }

            FillParameter parameter = (FillParameter)paramList["FillParameter"];
            Worksheet dataSheet = book.Sheets[parameter.SheetIndex];
            Chart chart = null;
            int sheetIndex = Convert.ToInt32(paramList["sheetIndex"]);
            if (book.Sheets[sheetIndex] is Worksheet)
            {
                Worksheet chartSheet = book.Sheets[sheetIndex];
                chart = ((ChartObject)chartSheet.ChartObjects(paramList["chartID"])).Chart;
            }
            else
            {
                chart = book.Sheets[sheetIndex];
            }
            SeriesCollection collection = chart.SeriesCollection();
            //foreach (Series series in collection)
            //{
            //    series.Delete();
            //}
            int seriesIndex = 0;
            foreach (var seriesElement in (IEnumerable<XElement>)paramList["SubParams"])
            {
                seriesIndex ++;
                if(seriesElement.Name != "series") continue;
                Range nameRange = null;
                Range xvalueRange = null;
                Range valueRange = null;
                foreach (var element in seriesElement.Elements())
                {
                    if (element.Attribute("rangeType") == null) continue;
                    if (element.Attribute("position") == null) continue;
                    if(element.Name == "name")
                    {
                        nameRange = (Range)ProcessHelper.GetRange(parameter, dataSheet, element);
                    }
                    else if(element.Name == "xvalue")
                    {
                        xvalueRange = (Range)ProcessHelper.GetRange(parameter, dataSheet, element);
                    }
                    else if (element.Name == "value")
                    {
                        valueRange = (Range)ProcessHelper.GetRange(parameter, dataSheet, element);
                    }
                }
                if(xvalueRange == null || valueRange == null) continue;

                Series series = collection.Item(seriesIndex);
                if (nameRange != null) series.Name = string.Format(@"='{0}'!{1}", dataSheet.Name, nameRange.Address);
                series.XValues = xvalueRange;
                series.Values = valueRange;
            }
            return null;
        }
    }
}
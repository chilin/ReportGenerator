using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ReportGeneratorApp.Excel;
using ReportGeneratorApp.DataSource;
using ReportGeneratorApp.Report.Generator.Request;

namespace ReportGeneratorApp.Report.Generator
{
    public enum DateType
    {
        Fixed,
        Custom
    }

    public class RequestReport : ReportTemplate
    {
        DateType dateType;
        public RequestReport(DateTime reportDate)
            : base("RequestReport", reportDate, ReportType.Weekly)
        {
            dateType = DateType.Fixed;
        }

        public RequestReport(DateTime startDate, DateTime endDate)
            : base("RequestReport", startDate, ReportType.Weekly)
        {
            StartDate = startDate;
            EndDate = endDate;
            dateType = DateType.Custom;
        }

        protected override FillParameter[] InitializeFillParameters()
        {
            Dictionary<string, object> ParamList = XmlHelper.GetFillParams("RequestReport", "paramCount");
            int paramCount = Convert.ToInt32(ParamList["value"]);
            FillParameter[] fillParameters = new FillParameter[paramCount];

            for (int paramIndex = 1; paramIndex <= paramCount; paramIndex++)
            {
                if (paramIndex == 1)
                {
                    if (dateType == DateType.Fixed)
                    {
                        fillParameters[paramIndex - 1] = new FillParameterBuilder(
                            this.ReportDate, "param" + paramIndex.ToString()).ToFillParameter1();
                    }
                    else
                    {
                        fillParameters[paramIndex - 1] = new FillParameterBuilder(
                            this.StartDate, this.EndDate, "param" + paramIndex.ToString()).ToFillParameter1();
                    }
                }
                else
                {
                    if (dateType == DateType.Fixed)
                    {
                        fillParameters[paramIndex - 1] = new FillParameterBuilder(
                            this.ReportDate, "param" + paramIndex.ToString()).ToFillParameter2();
                    }
                    else
                    {
                        fillParameters[paramIndex - 1] = new FillParameterBuilder(
                            this.StartDate, this.EndDate, "param" + paramIndex.ToString()).ToFillParameter2();
                    }
                }
            }

            return fillParameters;
        }
    }
}

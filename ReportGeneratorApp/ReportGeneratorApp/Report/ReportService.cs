using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace ReportGeneratorApp.Report
{
    class ReportService
    {
        public static GenerateResult GenerateReport(string reportName, int year, int month, int day)
        {
            return GenerateReport(reportName, new DateTime(year, month, day));
        }

        public static GenerateResult GenerateReport(string reportName, DateTime date)
        {
            var type = Type.GetType("ReportGeneratorApp.Report.Generator." + reportName + "Report", false);
            if (type != null)
            {
                var instance = Activator.CreateInstance(type, date);
                var method = type.GetMethod("Generate", BindingFlags.Instance | BindingFlags.Public);
                var ret = method.Invoke(instance, null) as GenerateResult;
                GC.Collect();
                return ret;
            }

            var result = new GenerateResult
            {
                ErrorMessage = string.Format("Not Found Report {0}", reportName),
                Status = -1
            };
            return result;
        }

        public static GenerateResult GenerateReport(string reportName, DateTime startDate, DateTime endDate)
        {
            var type = Type.GetType("ReportGeneratorApp.Report.Generator." + reportName + "Report", false);
            if (type != null)
            {
                var instance = Activator.CreateInstance(type, startDate, endDate);
                var method = type.GetMethod("Generate", BindingFlags.Instance | BindingFlags.Public);
                var ret = method.Invoke(instance, null) as GenerateResult;
                GC.Collect();
                return ret;
            }

            var result = new GenerateResult
            {
                ErrorMessage = string.Format("Not Found Report {0}", reportName),
                Status = -1
            };
            return result;
        }
    }
}

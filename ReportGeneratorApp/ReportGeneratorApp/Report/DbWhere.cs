using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;

namespace ReportGeneratorApp.Report
{
    public class DbWhere
    {
        public static string GetWhereString(DateTime startDate, DateTime endDate)
        {
            return string.Format(" [System.CreatedDate] >= '{0}' AND [System.CreatedDate] < '{1}'", startDate.ToString("yyyy-MM-dd"), endDate.AddDays(1).ToString("yyyy-MM-dd"));
        }

        public static string GetWhereString(ReportType ReportType, DateTime ReportDate, String tablealias)
        {   
            String alias = "";
            if (tablealias != null)
            {
                alias = tablealias+".";
            }
            StringBuilder strSqlWhere = new StringBuilder();
            switch (ReportType)
            {
                case ReportType.Daily:  // yesterday
                    {
                        // yesterday
                        strSqlWhere = strSqlWhere.Append(" [System.CreatedDate] = '" + ReportDate.AddDays(-1).ToString("yyyy-MM-dd") + "'");
                        break;
                    }
                case ReportType.Weekly:
                    {
                        // today(2013-3-12):2013-3-5 ~ 2013-3-11
                        strSqlWhere = strSqlWhere.Append(" " + alias + "[System.CreatedDate] >= '"
                            + ReportDate.AddDays(-7).ToString("yyyy-MM-dd") + "'");
                        strSqlWhere = strSqlWhere.Append(" AND " + alias + "[System.CreatedDate] < '"
                            + ReportDate.ToString("yyyy-MM-dd") + "'");
                        break;
                    }
                case ReportType.Monthly:
                    {
                        strSqlWhere = strSqlWhere.Append(" " + alias + "[System.CreatedDate] >= '"
                            + ReportDate.AddMonths(-1).ToString("yyyy-MM") + "-01'");
                        strSqlWhere = strSqlWhere.Append(" AND " + alias + "[System.CreatedDate] < '"
                            + ReportDate.ToString("yyyy-MM") + "-01'");
                        break;
                    }
                default:
                    {
                        // Daily
                        strSqlWhere = strSqlWhere.Append(" [System.CreatedDate] = '" + ReportDate.AddDays(-1).ToString("yyyy-MM-dd") + "'");
                        break;
                    }
            }

            return strSqlWhere.ToString();
        }

        public static string GetWhereString(ReportType ReportType, DateTime ReportDate)
        {
            return GetWhereString(ReportType, ReportDate, null);
        }

        public static DateTime GetStartDate(ReportType ReportType, DateTime ReportDate)
        {
            // DateTime is a struct, without null value.
            DateTime dtime = DateTime.Now;
            if (ReportDate != DateTime.MinValue)
            {
                // DateTime.MinValue is "00001-1-1"
                dtime = ReportDate;
            }

            // ReportType is a enum.
            switch (ReportType)
            {
                case ReportType.Daily:
                    {
                        dtime = dtime.AddDays(-1);
                        break;
                    }
                case ReportType.Weekly:
                    {
                        dtime = ReportDate.AddDays(-7);
                        break;
                    }
                case ReportType.Monthly:
                    {
                        dtime = dtime.AddDays(1 - dtime.Day).AddMonths(-1); // the first day of last month
                        break;
                    }
                default:
                    {
                        // Daily
                        dtime = dtime.AddDays(-1);
                        break;
                    }
            }
            return dtime;

        }

        public static DateTime GetEndDate(ReportType ReportType, DateTime ReportDate)
        {
            // DateTime is a struct, without null value.
            DateTime dtime = DateTime.Now;
            if (ReportDate != DateTime.MinValue)
            {
                // DateTime.MinValue is "00001-1-1"
                dtime = ReportDate;
            }

            // ReportType is a enum.
            switch (ReportType)
            {
                case ReportType.Daily:
                    {
                        dtime = dtime.AddDays(-1); // get yesterday.
                        break;
                    }
                case ReportType.Weekly:
                    {
                        dtime = ReportDate.AddDays(-1);
                        break;
                    }
                case ReportType.Monthly:
                    {
                        dtime = new DateTime(ReportDate.Year, ReportDate.Month, 1).AddDays(-1);
                        break;
                    }
                default:
                    {
                        dtime = dtime.AddDays(-1); // get yesterday.
                        break;
                    }
            }

            return dtime;
        }

        public static DateTime GetMonday(DateTime objDt)
        {
            int week = Convert.ToInt32(objDt.DayOfWeek.ToString("d"));
            //When Sunday set week to 7
            return objDt.AddDays(1 - (week == 0 ? 7 : week));
        }

        public static DateTime GetSunday(DateTime objDt)
        {
            int week = Convert.ToInt32(objDt.DayOfWeek.ToString("d"));
            return objDt.AddDays(0 - week);
        }

        public static int GetQuarter(int month)
        {
            switch (month)
            {
                case 11:
                case 12:
                case 1:
                    return 1;
                case 2:
                case 3:
                case 4:
                    return 2;
                case 5:
                case 6:
                case 7:
                    return 3;
                case 8:
                case 9:
                case 10:
                    return 4;
                default:
                    return 0;
            }
        }

        public static bool IsFirstWorkDayPerMonth(DateTime dateTime)
        {
            if (dateTime.Day == 1)
            {
                return (!dateTime.DayOfWeek.Equals(DayOfWeek.Sunday) && !dateTime.DayOfWeek.Equals(DayOfWeek.Saturday));
            }
            else if(dateTime.Day == 2 || dateTime.Day == 3)
            {
                return dateTime.DayOfWeek.Equals(DayOfWeek.Monday);
            }
            return false;
        }

        public static bool IsFirstMondayPerMonth(DateTime dateTime)
        {
            return dateTime.Equals(GetFirstWeekMonday(dateTime));
        }

        public static DateTime GetFirstWeekMonday(DateTime dateTime)
        {
            DateTime date = new DateTime(dateTime.Year, dateTime.Month, 1);
            date = date.AddDays((7 - (int)date.DayOfWeek + 1) % 7);
            return date;
        } 

        public static int GetQuarterLastMonth(int quarter)
        {
            switch (quarter)
            {
                case 1:
                    return 1;
                case 2:
                    return 4;
                case 3:
                    return 7;
                case 4:
                    return 10;
                default:
                    return 0;
            }
        }

        public static string GetFY(int year, int month)
        {
            if (month >= 11)
            {
                return (year + 1).ToString().Substring(2);
            }
            else
            {
                return year.ToString().Substring(2);
            }
        }

        public static bool IsMonday(DateTime datetime)
        {
            string weekName = datetime.DayOfWeek.ToString();

            if ("Monday".Equals(weekName))
            {
                return true;
            }

            return false;
        }
        public static bool IsTuesday(DateTime datetime)
        {
            string weekName = datetime.DayOfWeek.ToString();

            if ("Tuesday".Equals(weekName))
            {
                return true;
            }

            return false;
        }
        public static string GetOrdinalDay(DateTime datetime)
        {
            string ordinalDay = datetime.Day.ToString();

            switch (datetime.Day)
            { 
                case 1:
                case 21:
                case 31:
                    ordinalDay += "st";
                    break;
                case 2:
                case 22:
                    ordinalDay += "nd";
                    break;
                case 3:
                case 23:
                    ordinalDay += "rd";
                    break;
                default:
                    ordinalDay += "th";
                    break;
            }

            return ordinalDay;
        }

        public static string GetMonthDay(DateTime datetime)
        {
            // Format: May 1st, May 30th
            string date = string.Format("{0} {1}",
                datetime.ToString("MMMM", System.Globalization.CultureInfo.GetCultureInfo("en-US")), 
                GetOrdinalDay(datetime));

            return date;
        }
    }
}
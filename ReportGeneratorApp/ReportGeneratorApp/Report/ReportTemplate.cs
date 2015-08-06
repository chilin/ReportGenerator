using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ReportGeneratorApp.Excel;
using ReportGeneratorApp.CustomizeException;
using ReportGeneratorApp.Report;
using ReportGeneratorApp.DataSource;
using System.Data;
using ReportGeneratorApp.Log;


namespace ReportGeneratorApp.Report
{
    public delegate object RowHandlerDelegate(object[,] targetArray, SqlDataReader dataReader, int rowIndex, int recordIndex,  string[] columnNameArray);
    public delegate SqlCommand QueryBuilderDelegate(SqlConnection connection, string tableName, string[] columnNameArray, string where);
    public delegate object UpdateSheetDelegate(Worksheet sheet);

    /// <summary>
    /// 
    /// </summary>
    public enum ReportType { Daily, Weekly, Monthly}

    public abstract class ReportTemplate
    {
        public event EventHandler<EventArgs> BeginFillBookEvent;
        public event EventHandler<EventArgs> EndFillBookEvent;
        public string TemplateFolder { get; set; }
        public string TemplateName { get; set; }
        public string ReportName { get; set; }
        public string CustomReportName { get; set; }
        public string ReportNameSuffix { get; set; }
        public string OutputFolder { get; set; }
        public string TimeFolder { get; set; }
        public string ShotName { get; set; }

        public string SourceFile { get; set; }
        public string TargetFile { get; set; }
        public DateTime ReportDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime StartDate { get; set; }
        public ReportType ReportType { get; set; }

        public static string DefaultOutputFolder;
        public static string DefaultTemplateFolder;
        public static string DefaultNameSuffix;

        public int flagSheetIndex { get; set; }

        static ReportTemplate()
        {
            //DefaultOutputFolder = ConfigurationManager.AppSettings["Default_Output_Folder"];
            //DefaultTemplateFolder = ConfigurationManager.AppSettings["Default_Template_Folder"];
            //DefaultNameSuffix = ConfigurationManager.AppSettings["Default_Suffix"];
            DefaultOutputFolder = string.Format(ConfigurationManager.AppSettings["Default_Output_Folder"], 
                AppDomain.CurrentDomain.BaseDirectory);
            DefaultTemplateFolder = string.Format(ConfigurationManager.AppSettings["Default_Template_Folder"],
                AppDomain.CurrentDomain.BaseDirectory);
            DefaultNameSuffix = ConfigurationManager.AppSettings["Default_Suffix"];
        }

        public ReportTemplate() { }

        protected ReportTemplate(string shortName, DateTime date, ReportType type)
        {
            this.ReportDate = date;
            this.ShotName = shortName;
            this.ReportType = type;

            this.StartDate = DbWhere.GetStartDate(type, date);
            this.EndDate = DbWhere.GetEndDate(type, date);

            Logger.Debug(string.Format("============================ Start Generate {0}, ReportDate: {1}============================", shortName, ReportDate.ToString("yyyy-MM-dd HH:mm:ss")));
        }
        
        private void InitializeParameters()
        {
            ReportName = ConfigurationManager.AppSettings[ShotName + "_Name"];
            if (string.IsNullOrEmpty(ReportName))
            {
                throw (new Exception(ShotName + "_Name node needed under <appSettings> in App.config file"));
            }
            var regex = new Regex(@"\{[yYMm\-_dD]*\}");
            System.Globalization.CultureInfo en_us = System.Globalization.CultureInfo.GetCultureInfo("en-US"); 

            MatchCollection colls = regex.Matches(ReportName);

            if (colls != null && colls.Count > 0)
            {
                int matchcounter = 0;
                foreach (Match match in colls){
                    matchcounter = matchcounter + 1;

                    var dateFormatString = match.Value;
                    var dateString = "";
                    if (matchcounter == 1)
                    {
                        try
                        {
                            dateString = this.StartDate.ToString(dateFormatString.Substring(1, dateFormatString.Length - 2), en_us);
                        }
                        catch (Exception e)
                        {
                            try
                            {
                                dateString = this.ReportDate.ToString(dateFormatString.Substring(1, dateFormatString.Length - 2), en_us);
                            }
                            catch (Exception ee)
                            {
                                dateString = DateTime.Now.ToString(dateFormatString.Substring(1, dateFormatString.Length - 2), en_us);
                            }
                        }

                        ReportName = regex.Replace(ReportName, dateString, 1);

                    }
                    else if (matchcounter == 2)
                    {
                        try
                        {
                            dateString = this.EndDate.ToString(dateFormatString.Substring(1, dateFormatString.Length - 2), en_us);
                        }
                        catch (Exception e)
                        {
                            dateString = DateTime.Now.ToString(dateFormatString.Substring(1, dateFormatString.Length - 2), en_us);
                        }

                        ReportName = regex.Replace(ReportName, dateString, 1);
                    }
                }
            }

            if (!string.IsNullOrEmpty(CustomReportName)) ReportName = CustomReportName;

            TemplateFolder = ConfigurationManager.AppSettings[ShotName + "_Template_Folder"];
            ReportNameSuffix = ConfigurationManager.AppSettings[ShotName + "_Suffix"];
            OutputFolder = ConfigurationManager.AppSettings[ShotName + "_Output_Folder"];
            TemplateName = ConfigurationManager.AppSettings[ShotName + "_Template_Name"];

            TemplateFolder = string.IsNullOrEmpty(TemplateFolder) ? DefaultTemplateFolder : TemplateFolder;
            ReportNameSuffix = string.IsNullOrEmpty(ReportNameSuffix) ? DefaultNameSuffix : ReportNameSuffix;
            OutputFolder = string.IsNullOrEmpty(OutputFolder) ? DefaultOutputFolder : OutputFolder;
            TimeFolder = string.Format(@"{0}\{1}\{2}\{3}\", 
                DateTime.Now.ToString("yyyy"), 
                DateTime.Now.ToString("MM"),
                DateTime.Now.ToString("dd"),
                DateTime.Now.ToString("HH-mm"));
            SourceFile = TemplateFolder + (string.IsNullOrEmpty(TemplateName) ? ShotName : TemplateName) + ReportNameSuffix;
            TargetFile = OutputFolder + TimeFolder + ReportName + ReportNameSuffix;
            if(!System.IO.Directory.Exists(OutputFolder + TimeFolder))
            {
                System.IO.Directory.CreateDirectory(OutputFolder + TimeFolder);
            }
        }

        protected abstract FillParameter[] InitializeFillParameters();
        protected object SimpleRowHandlerHelper(object[,] targetArray, SqlDataReader dataReader, int rowIndex, int recordIndex, string[] columnNameArray)
        {
            var columnsCount = columnNameArray.Length;
            int cellColumnIndex = 0;

            try
            {
                for (cellColumnIndex = 0; cellColumnIndex < columnsCount; cellColumnIndex++)
                {
                    var obj = dataReader[cellColumnIndex];
                    targetArray[rowIndex, cellColumnIndex] = obj;
                }
            }
            catch
            {
                Logger.Info(string.Format("SimpleRowHandlerHelper Info: RowIndex:{0}, RecordIndex:{1}, ColumnName:{2}", rowIndex, recordIndex + 1, columnNameArray[cellColumnIndex]));
                throw;
            }
            return null;
        }
        protected SqlCommand SimpleQueryBuilderHelper(SqlConnection connection, string tableName, string[] columnNameArray, string whereClause)
        {
            string sqlSelect = string.Format(@"SELECT {0} FROM {1} " + whereClause, string.Join(",", columnNameArray), tableName);
            SqlCommand command = new SqlCommand(sqlSelect, connection);
            return command;
        }
        
        virtual protected void EndFillSheetHandler(Worksheet worksheet)
        {
            
        }
        protected virtual GenerateResult ResultGenerate()
        {
            var result = new GenerateResult();
            result.Status = 0;
            result.FileName = TargetFile.Substring(TargetFile.LastIndexOf(@"\"));
            result.FilePath = TargetFile.Substring(0, TargetFile.LastIndexOf(@"\"));
            return result;
        }

        protected String auditReport()
        {
            using (SqlConnection connection = ConnectionFactory.GetConnection())
            {
                connection.Open();

                SqlCommand command = connection.CreateCommand();
                //SqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);
                //command.Transaction = transaction;

                String guid = Guid.NewGuid().ToString();

                try
                {
                    command.CommandText = "insert into t_report_audit (uuid, start_date, end_date, report_name, report_type, generate_time, requestor, report_shot_name) values (:uuid, :startDate, :endDate, :reportName, :reportType, :generateTime, :requestor, :reportShotName)";

                    SqlParameter parauuid = new SqlParameter("uuid", SqlDbType.VarChar);
                    parauuid.Value = guid;
                    parauuid.Direction = ParameterDirection.Input;
                    command.Parameters.Add(parauuid);

                    SqlParameter parastartDate = new SqlParameter("startDate", SqlDbType.DateTime);
                    parastartDate.Value = this.StartDate;
                    parastartDate.Direction = ParameterDirection.Input;
                    command.Parameters.Add(parastartDate);

                    SqlParameter paraendDate = new SqlParameter("endDate", SqlDbType.DateTime);
                    paraendDate.Value = this.EndDate;
                    paraendDate.Direction = ParameterDirection.Input;
                    command.Parameters.Add(paraendDate);

                    SqlParameter parareportName = new SqlParameter("reportName", SqlDbType.VarChar);
                    parareportName.Value = this.ReportName;
                    parareportName.Direction = ParameterDirection.Input;
                    command.Parameters.Add(parareportName);

                    SqlParameter parareportType = new SqlParameter("reportType", SqlDbType.VarChar);
                    parareportType.Value = this.ReportType;
                    parareportType.Direction = ParameterDirection.Input;
                    command.Parameters.Add(parareportType);

                    SqlParameter paragenerateTime = new SqlParameter("generateTime", SqlDbType.DateTime);
                    paragenerateTime.Value = DateTime.Now;
                    paragenerateTime.Direction = ParameterDirection.Input;
                    command.Parameters.Add(paragenerateTime);

                    SqlParameter pararequestor = new SqlParameter("requestor", SqlDbType.VarChar);
                    pararequestor.Value = getRequestor();
                    pararequestor.Direction = ParameterDirection.Input;
                    command.Parameters.Add(pararequestor);

                    SqlParameter parareportShotName = new SqlParameter("reportShotName", SqlDbType.VarChar);
                    parareportShotName.Value = this.ShotName;
                    parareportShotName.Direction = ParameterDirection.Input;
                    command.Parameters.Add(parareportShotName);

                    command.ExecuteNonQuery();
                    //transaction.Commit();
                }
                catch (Exception e)
                {
                    //transaction.Rollback();
                    throw e;
                }

                return guid;
            }
        }

        protected virtual string getRequestor()
        {
            return this.ShotName;
        }

        protected virtual void BeforeGenerateReport()
        {

        }

        protected virtual void EndOpenTemplateFile(Workbook book)
        {

        }

        protected virtual void AfterGenerateReport()
        {
           
        }

        protected virtual void specialActionAfterGenerateReport(Worksheet sheet)
        {

        }

        protected virtual void EndSaveReport()
        {
        }

        protected void RunProcedure(SqlConnection connection, string procedureName, SqlParameter[] parameters)
        {
            SqlTransaction transaction = null;
            try
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                command.CommandText = procedureName;
                command.Transaction = transaction;
                command.CommandType = CommandType.StoredProcedure;

                if (parameters != null)
                {
                    foreach (SqlParameter p in parameters)
                    {
                        //check for derived output value with no value assigned
                        if ((p.Direction == ParameterDirection.InputOutput) && (p.Value == null))
                        {
                            p.Value = DBNull.Value;
                        }
                        command.Parameters.Add(p);
                    }
                }
                command.ExecuteNonQuery();
                transaction.Commit();
            }
            catch
            {
                if (transaction != null)
                {
                    transaction.Rollback();
                }
                connection.Close();
                throw;
            }
            finally
            {
                connection.Close();
            }
        }

        protected void RunSql(SqlConnection connection, string[] sqls)
        {
            SqlTransaction transaction = null;
            try
            {
                connection.Open();

                SqlCommand command = connection.CreateCommand();
                transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

                command.Transaction = transaction;

                Logger.Debug("Start Run SQL");
                foreach (string sql in sqls)
                {
                    Logger.Debug(sql);
                    command.CommandText = sql;
                    command.ExecuteNonQuery();
                }
                transaction.Commit();
                Logger.Debug("End Run SQL");
            }
            catch
            {
                if (transaction != null)
                {
                    transaction.Rollback();
                }
                connection.Close();
                throw;
            }
            finally
            {
                connection.Close();
            }
        }

        public static object ExecuteScalar(SqlConnection connection, string sql)
        {
            object retval;
            try
            {
                connection.Open();

                SqlCommand command = connection.CreateCommand();

                Logger.Debug("Start ExecuteScalar");
                Logger.Debug(sql);
                command.CommandText = sql;
                retval = command.ExecuteScalar();
                Logger.Debug("End ExecuteScalar");

                return retval;
            }
            catch
            {
                connection.Close();
                throw;
            }
            finally
            {
                connection.Close();
            }
        }

        public GenerateResult Generate()
        {
            Workbook book = null;
            Application excel = null;
            try
            {
                InitializeParameters();
                BeforeGenerateReport();

                excel = new Application();
                File.Copy(SourceFile, TargetFile, true);
                object missing = Missing.Value;
                book = excel.Workbooks.Open(TargetFile, missing, missing, missing,
                                                     missing, missing, missing, missing,
                                                     missing, missing);
                EndOpenTemplateFile(book);
                var fillParameters = InitializeFillParameters();


                if (flagSheetIndex > 0)
                {
                    Worksheet sheet1 = null;
                    if (flagSheetIndex == 99999)
                    {
                        sheet1 = book.Sheets["hiddendata"];
                    }
                    else
                    {
                        sheet1 = book.Sheets[flagSheetIndex];
                    }
                    try
                    {
                        sheet1.Activate();
                        specialActionAfterGenerateReport(sheet1);
                    }
                    catch
                    {
                        throw;
                    }
                    finally
                    {
                        if (sheet1 != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
                            sheet1 = null;
                        }
                    }
                    
                }

                foreach (var fillParameter in fillParameters)
                {
                    //BeginFillBookEvent(book, null);
                    Worksheet sheet = null;
                    try
                    {
                        Logger.Debug(string.Format("======== Start Fill Sheet {0} ==========", fillParameter.SheetIndex));
                        sheet = SheetDataAdapter.FillSheetRange(book, fillParameter);
                        Logger.Debug(string.Format("======== End Fill Sheet {0} ==========", fillParameter.SheetIndex));

                    }
                    catch (NoResultException ne)
                    {
                        Logger.Info(ne.Message);
                        Logger.Info("[" + fillParameter.TableName + "] [" + fillParameter.WhereClause + "][sheetindex=" + fillParameter.SheetIndex + "]");
                        Logger.Error(ne);
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                            sheet = null;
                        }
                        
                    }
                    

                    //EndFillBookEvent(book, null);
                }

                //process
                Dictionary<string, object> processList = XmlHelper.GetProcess(ShotName);
                if (processList != null)
                {
                    foreach (var process in processList)
                    {
                        var paramList = process.Value as Dictionary<string, object>;
                        var type = Type.GetType("ReportGeneratorApp.Excel.Process." + paramList["type"], false);
                        if (type != null)
                        {
                            if(paramList.ContainsKey("paramIndex"))
                            {
                                paramList.Add("FillParameter", fillParameters[Convert.ToInt32(paramList["paramIndex"])]);
                            }
                            paramList.Add("StartDate", StartDate);
                            paramList.Add("EndDate", EndDate);
                            var instance = Activator.CreateInstance(type, null);
                            var method = type.GetMethod("Process", BindingFlags.Instance | BindingFlags.Public);
                            method.Invoke(instance, new[] {book, process.Value});
                        }
                    }
                }

                AfterGenerateReport();
                if (this.ReportNameSuffix == ".xlsm")
                {
                    ExcelMacroHelper.RunMacro(excel, new object[] { "auto_open_by2010" });
                }

                book.Sheets[1].Activate();
            }
            catch (Exception e)
            {
                Logger.Error(e);
                return new GenerateResult { Status=-1, ErrorMessage=e.Message };
            }
            finally
            {
                if (book != null)
                {
                    
                    book.Save();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                    book = null;

                }
                if (excel != null)
                {
                    excel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                    excel = null;
                }
                GC.Collect();
            }

            try
            {
                EndSaveReport();
            }
            catch (Exception e)
            {
                Logger.Error(e);
                return new GenerateResult { Status = -1, ErrorMessage = e.Message };
            }

            Logger.Debug(string.Format("============================ End Generate {0} ============================", this.ShotName));

            return ResultGenerate();
            
        }
    }
}
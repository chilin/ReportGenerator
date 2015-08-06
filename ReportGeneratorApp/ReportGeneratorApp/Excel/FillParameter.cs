using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ReportGeneratorApp.Report;

namespace ReportGeneratorApp.Excel
{
    public class FillParameter
    {
        public FillParameter()
        {
            SheetIndex = 1;
            BatchSize = 10000;
            PerRecordPerRow = true;
            CustomizeBatch = false;
            IsNeedQueryData = true;
            connectionString = "reportDBConnectionString";
        }
        /// <summary>
        /// Excel Sheet Index start with 1 instead of 0
        /// </summary>
        public int SheetIndex { get; set; }
        public int RowOffset { get; set; }
        public int ColumnOffset { get; set; }
        public string[] ColumnNameArray { get; set; }
        public string TableName { get; set; }
        public int BatchSize { get; set; }
        public bool PerRecordPerRow { get; set; }
        public bool CustomizeBatch { get; set; }
        public bool IsNeedQueryData { get; set; }
        public QueryBuilderDelegate QueryBuilderDelegate { get; set; }
        public RowHandlerDelegate RowHandlerDelegate { get; set; }
        public UpdateSheetDelegate UpdateSheetDelegate { get; set; }
        public string WhereClause { get; set; }

        public String connectionString { get; set; }
    }
}
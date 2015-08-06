using System.Runtime.Serialization;

namespace ReportGeneratorApp.Report
{
    public class GenerateResult
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        /**
         * Failed:-1
         * Success:0
         * 
         * 
         */
        public int Status { get; set; }
        public string ErrorMessage { get; set; }
        public long ReportLogId { get; set; }
    }
}
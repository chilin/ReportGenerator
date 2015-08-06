using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ReportGeneratorApp.Report;

namespace ReportGeneratorApp
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btn_generate_Click(object sender, EventArgs e)
        {
            this.label1.Text = "Generating, Please wait...";
            this.btn_generate.Enabled = false;

            GenerateResult ret = null;
            if (this.radioFixed.Checked)
            {
                DateTime lastSunday = DbWhere.GetSunday(DateTime.Now);
                ret = ReportService.GenerateReport("Request", lastSunday);
            }
            else if (this.radioCustom.Checked)
            {
                ret = ReportService.GenerateReport("Request", this.dt_StartDate.Value, this.dt_EndDate.Value);
            }

            
            if (ret.Status > -1)
            {
                this.label1.Text = "Successful to generate report.";
                System.Diagnostics.Process.Start(ret.FilePath);
            }
            else
            {
                this.label1.Text = "Failed to generate report.";
                MessageBox.Show(ret.ErrorMessage);
            }

            this.btn_generate.Enabled = true;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            showReportDate();
        }

        private void showReportDate()
        {
            DateTime lastSunday = DbWhere.GetSunday(DateTime.Now);
            this.lbl_tip.Text = String.Format("ReportTime: {0} ～ {1}",
                lastSunday.AddDays(-7).ToString("yyyy-MM-dd"),
                lastSunday.AddDays(-1).ToString("yyyy-MM-dd"));
            this.dt_StartDate.Value = lastSunday.AddDays(-7);
            this.dt_EndDate.Value = lastSunday.AddDays(-1);
        }
    }
}

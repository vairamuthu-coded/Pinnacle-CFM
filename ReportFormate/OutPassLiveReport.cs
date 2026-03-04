using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace Pinnacle.ReportFormate
{
    public partial class OutPassLiveReport : Form,ToolStripAccess
    {
        private static OutPassLiveReport _instance;
        public static OutPassLiveReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new OutPassLiveReport();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

        Models.UserRights sm = new Models.UserRights(); Models.Master mas = new Models.Master();
        Report.OutPassLiveReport rd = new Report.OutPassLiveReport();
        public OutPassLiveReport()
        {
            InitializeComponent();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }
        public void ReadOnlys()
        {

        }

        private void OutPassLiveReport_Load(object sender, EventArgs e)
        {
          

            this.txtsearch.Focus(); comboformate.SelectedIndex = 1;
            frmdate.Value = DateTime.Now.AddDays(0); todate.Value = DateTime.Now.AddDays(0);
            COMBCODELOAD();
            
        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

        public void DownLoads()
        {
            if (comboformate.Text != "")
            {

                DialogResult result = MessageBox.Show("Do you want to '" + comboformate.Text + "' Formate ??", "" + comboformate.Text + "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    switch (comboformate.Text)
                    {
                        case "Word":
                            rd.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + combocompcode.Text + "'OutPassLiveReprt.doc");
                            break;

                        case "Excel":
                            rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + combocompcode.Text + "'OutPassLiveReprt.xls");
                            break;

                        case "PDF":
                            rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + combocompcode.Text + "'OutPassLiveReprt.pdf");
                            break;

                        case "CSV":
                            rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "d:\\'" + combocompcode.Text + "'OutPassLiveReprt.csv");
                            break;
                    }

                }
                else
                {

                }
            }
            else
            {
                MessageBox.Show("Pls Select Combo Box Value");
            }

        }

        private void News_Click(object sender, EventArgs e)
        {

        }
        private void COMBCODELOAD()
        {
            try
            {
                string sel3 = "SELECT DISTINCT  A.HOSTELNAME as COMPCODE FROM ASPTBLHOSTELGATEPASS A ";
                DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "GTCOMPMAST");
                DataTable dt3 = ds3.Tables["GTCOMPMAST"];
                if (dt3.Rows.Count > 0)
                {
                    combocompcode.DisplayMember = "COMPCODE";
                    combocompcode.ValueMember = "COMPCODE";
                    combocompcode.DataSource = dt3;


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
        }
        private void butView_Click(object sender, EventArgs e)
        {
            try
            {
                if (combocompcode.Text != "")
                {
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,D.MIDCARD ,C.FNAME , D.DEPT AS DISPNAME,substr(A.SYSTEMDATE,1,10) AS CONTACTNO,'' as HOSTELNAME,'' as BLOCKFLOOR , '' as ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,A.OUTTIME, A.REMARKS   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE      JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO       JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT     JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID where A.HOSTELNAME='"+combocompcode.Text+"'  AND  a.intime is  not null  and A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dt2);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh(); txtsearch.Text = "";
                //}
                //else
                //{


                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME,substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME,F.BLOCKFLOOR , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,A.OUTTIME, A.REMARKS   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID   WHERE A.HOSTELNAME='" + combocompcode.Text + "'  AND a.intime is  not null  and A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh(); txtsearch.Text = "";
                }
            }
            catch (Exception EX)
            { MessageBox.Show(EX.Message); }
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtsearch.Text.Length>3)
                {

                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPCODE,D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME,substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME,F.BLOCKFLOOR , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,A.OUTTIME, A.REMARKS   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE  A.INTIME IS  not NULL and D.MIDCARD LIKE'%" + txtsearch.Text + "%'  and A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh();txtsearch.Focus();
                }
                else
                {
                    butView_Click(sender,e);
                }

            }
            catch (Exception EX)
            { MessageBox.Show(EX.Message); }
        }

        public void Prints()
        {
            try
            {
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME,substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME,F.BLOCKFLOOR , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,A.OUTTIME, A.REMARKS   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID   WHERE a.intime is  not null  and A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";

                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                    reportdocument.Load(Application.StartupPath + "\\Report\\OutPassLiveReport.rpt");
                    reportdocument.SetDataSource(dt2);
                    reportdocument.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                    reportdocument.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Pdfs()
        {
            //DialogResult result = MessageBox.Show("Do you want to PDF Formate ??", "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //if (result.Equals(DialogResult.OK))
            //{
            //    rd.ExportToDisk(ExportFormatType.PortableDocFormat, "c:\\'" + combocompcode.Text + "'CustomerReport1.pdf");
            //}
            //else
            //{

            //}
        }

     

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void comboformate_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPCODE,D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME,H.CONTACTNO,F.HOSTELNAME,F.BLOCKFLOOR , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.OUTTIME,A.INTIME, A.REMARKS   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT AND E.COMPCODE=B.GTCOMPMASTID   JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE a.outtime is  null   ORDER BY 1";

                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh(); txtsearch.Focus();
             
                

            }
            catch (Exception EX)
            { MessageBox.Show(EX.Message); }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        public void News()
        {
           
        }

        public void Saves()
        {
           
        }

     

        public void Searchs()
        {
           
        }

        public void Deletes()
        {
           
        }

        public void Imports()
        {
           
        }

     

        public void ChangePasswords()
        {
           
        }

       

        public void ChangeSkins()
        {
           
        }

        public void Logins()
        {
           
        }

        public void GlobalSearchs()
        {
           
        }

        public void TreeButtons()
        {
           
        }

        public void Exit()
        {
            this.Hide();
            crystalReportViewer1.ReportSource = null;
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
        }

        public void GridLoad()
        {
           
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

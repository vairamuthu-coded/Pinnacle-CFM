using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace Pinnacle.ReportFormate
{
    public partial class SampleCollectionReport : Form,ToolStripAccess
    {
        private static SampleCollectionReport _instance;
        public static SampleCollectionReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new SampleCollectionReport();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public SampleCollectionReport()
        {
            InitializeComponent();
        }
        Report.SampleCollectionReport rd = new Report.SampleCollectionReport();
        Report.Sample.SampleCorrection sc = new Report.Sample.SampleCorrection ();
        Report.SampleCollectionReport3 rchart = new Report.SampleCollectionReport3();
        string path = "";
        string folderLocation = "d:\\SampleCollections-Download\\";
        public void StyleLoad()
        {
            string sel = "SELECT distinct asptblbuysamid,STYLENAME    FROM  asptblbuysam    WHERE ACTIVE='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbuysam");
            DataTable dt = ds.Tables["asptblbuysam"];
            combostyle.DataSource = dt;
            combostyle.DisplayMember = "STYLENAME";
            combostyle.ValueMember = "asptblbuysamid";

        }
        private void stockreport()
        {
            string sel0 = "SELECT   A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE   JOIN ASPTBLBUYSAMinw Q ON Q.AGFSAMPLE=A.AGFSAMPLE  where   Q.OUTWARD='T'   ORDER by a.AGFSAMPLE ASC ";
            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
            DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
            rd.SetDataSource(dt0);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd;
            crystalReportViewer1.Refresh();
            lblcount.Refresh();
            lblcount.Text = "CompCode : " + Class.Users.HCompcode + ",  Total Stock Qty " + dt0.Rows.Count;

        }
        
        private void outreport()
        {
            string sel0 = "SELECT A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE JOIN ASPTBLBUYSAMINW  R ON R.AGFSAMPLE=A.AGFSAMPLE  WHERE R.OUTWARD='F'    ORDER by a.AGFSAMPLE ASC";
            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
            DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
            rd.SetDataSource(dt0);
            crystalReportViewer2.ReportSource = null;
            crystalReportViewer2.ReportSource = rd;
            crystalReportViewer2.Refresh(); lblcount.Refresh();
            lblcount.Text = "CompCode : " + Class.Users.HCompcode + ",  Total  Outward  Qty " + dt0.Rows.Count;

        }

        private void reportchart()
        {
            string sel0 = "SELECT A.MFYEAR,  B.COMPCODE,b.compname,A.PCS FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE = B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID = A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID = A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID = A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID = A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID = A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID = A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID = A.ORDERPACKTYPE where a.active='T' ";

            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
            DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
            rchart.SetDataSource(dt0);
            crystalReportViewer3.ReportSource = null;
            crystalReportViewer3.ReportSource = rchart;
            crystalReportViewer3.Refresh();
        }

        private void Pendingreport()
        {
           

            string sel0 = "SELECT A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE WHERE a.inward is null    ORDER by a.AGFSAMPLE ASC";
            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
            DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
            if (dt0.Rows.Count > 0)
            {
                rd.SetDataSource(dt0);
                crystalReportViewer5.ReportSource = null;
                crystalReportViewer5.ReportSource = rd;
                crystalReportViewer5.Refresh(); lblcount.Refresh();
                lblcount.Text = "CompCode : " + Class.Users.HCompcode + ",  Total  Outward  Qty " + dt0.Rows.Count;
            }
            else
            {
                MessageBox.Show("No Data Found.");
            }
        }
        private void correctionreport()
        {
            string sel0 = "";
           
            if (txtrefcode.Text == "")
            {
                sel0 = "SELECT A.AGFSAMPLE, B.COMPNAME,A.GARMENTIMAGE, A.REMARKS, A.FABRICCOMPLIANT , A.RISK1, A.RISK2, A.RISK3, A.RISK4, A.RISK5, A.RISK6  FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE A.REMARKS IS NOT NULL  OR A.FABRICCOMPLIANT IS NOT NULL  OR A.RISK1 IS NOT NULL  OR A.RISK2 IS NOT NULL  OR A.RISK3 IS NOT NULL  OR A.RISK4 IS NOT NULL  OR A.RISK5 IS NOT NULL   OR A.RISK6 IS NOT NULL  ORDER BY 1";
             
            }
            if (txtrefcode.Text.Length >= 5)
            {
                sel0 = "SELECT A.AGFSAMPLE, B.COMPNAME,A.GARMENTIMAGE, A.REMARKS, A.FABRICCOMPLIANT , A.RISK1, A.RISK2, A.RISK3, A.RISK4, A.RISK5, A.RISK6  FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE A.AGFSAMPLE='" + txtrefcode.Text + "'   ORDER BY 1";

            }
            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
            DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
            if (dt0.Rows.Count > 0)
            {
                sc.SetDataSource(dt0);
                crystalReportViewer4.ReportSource = null;
                crystalReportViewer4.ReportSource = sc;
                crystalReportViewer4.Refresh(); lblcount.Refresh();
                lblcount.Text = "CompCode : " + Class.Users.HCompcode + ",  Total  Outward  Qty " + dt0.Rows.Count;
                txtrefcode.Text = ""; txtrefcode.Select();
            }
            else
            {
                MessageBox.Show("No Data Found.");
            }
        }
        private void combostyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combostyle.Text != "ALL" && Class.Users.HCompcode == "ALL" )
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE   where A.ACTIVE='T' AND  a.stylename='" + combostyle.Text + "'  ORDER by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"]; rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combostyle.Text = ""; return;
            }
            if (combostyle.Text != "ALL" && Class.Users.HCompcode != "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE   where A.ACTIVE='T' AND  a.stylename='" + combostyle.Text + "' and b.compcode='" + Class.Users.HCompcode + "' ORDER by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"]; rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combostyle.Text = ""; return;
            }
        }

        private void SampleCollectionReport_Load(object sender, EventArgs e)
        {

            compcode(); FabricLoad();  DepartmentLoad(); SeasonLoad(); OrderPackLoad(); BrandLoad();StyleLoad();
         
            News();
            stockreport();
        }
        public void compcode()
        {
            string sel = " SELECT 0 AS GTCOMPMASTID, 'ALL' AS COMPCODE FROM DUAL UNION ALL SELECT A.GTCOMPMASTID,A.COMPCODE    FROM  GTCOMPMAST A    WHERE A.ACTIVE='T'  ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCOMPMAST");
            DataTable dt = ds.Tables["GTCOMPMAST"];
            combocompcode.DataSource = dt;
            combocompcode.DisplayMember = "COMPCODE";
            combocompcode.ValueMember = "GTCOMPMASTID";
          
        }


      
        public void BrandLoad()
        {


            string sel = "SELECT distinct  A.ASPTBLBRANDMASID,A.BRAND    FROM  ASPTBLBRANDMAS A  JOIN asptblbuysam B ON A.ASPTBLBRANDMASID=B.brand  join gtcompmast c on C.GTCOMPMASTID=B.COMPCODE      WHERE A.ACTIVE='T' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLBRANDMAS");
            DataTable dt = ds.Tables["ASPTBLBRANDMAS"];
            if (dt.Rows.Count > 0)
            {
                combobrand.DataSource = dt;
                combobrand.DisplayMember = "BRAND";
                combobrand.ValueMember = "ASPTBLBRANDMASID";
              
            }
        }
        public void OrderPackLoad()
        {
            string sel = " SELECT 0 AS ASPTBLORDPACKMASid, 'ALL' AS ORDERPACKTYPE FROM DUAL UNION ALL SELECT A.ASPTBLORDPACKMASid,A.ORDERPACKTYPE    FROM  ASPTBLORDPACKMAS A   WHERE A.ACTIVE='T' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLORDPACKMAS");
            DataTable dt = ds.Tables["ASPTBLORDPACKMAS"];
            combopacktype1.DataSource = dt;
            combopacktype1.DisplayMember = "ORDERPACKTYPE";
            combopacktype1.ValueMember = "ASPTBLORDPACKMASID";
                 combopacktype1.DataSource = dt;
            combopacktype1.DisplayMember = "ORDERPACKTYPE";
            combopacktype1.ValueMember = "ASPTBLORDPACKMASID";
        }
        public void SeasonLoad()
        {
            string sel = " SELECT DISTINCT  A.ASPTBLSEASONMASID, A.SEASON      FROM  ASPTBLSEASONMAS  A  JOIN asptblbuysam B ON A.ASPTBLSEASONMASID=B.SEASON    join gtcompmast c on C.GTCOMPMASTID=B.COMPCODE         WHERE A.ACTIVE='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLSEASONMAS");
            DataTable dt = ds.Tables["ASPTBLSEASONMAS"];
            comboseason.DataSource = dt;
            comboseason.DisplayMember = "SEASON";
            comboseason.ValueMember = "ASPTBLSEASONMASID";
            comboseason.DataSource = dt;
            comboseason.DisplayMember = "SEASON";
            comboseason.ValueMember = "ASPTBLSEASONMASID";
      
        }
       

        public void DepartmentLoad()
        {
            string sel = "SELECT 0 AS ASPTBLSAMDEPTMASID, 'ALL' AS DEPARTMENT FROM DUAL UNION ALL  SELECT distinct ASPTBLSAMDEPTMASID,DEPARTMENT    FROM  ASPTBLSAMDEPTMAS    WHERE ACTIVE='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLSAMDEPTMAS");
            DataTable dt = ds.Tables["ASPTBLSAMDEPTMAS"];
            combodept.DataSource = dt;
            combodept.DisplayMember = "DEPARTMENT";
            combodept.ValueMember = "ASPTBLSAMDEPTMASID";
        
        }

        public void FabricLoad()
        {
            string sel = "SELECT distinct b.category    FROM  ASPTBLSAMDEPTMAS  a    join asptblsamcatmas b on a.ASPTBLSAMDEPTMASID=b.asptblsamcatmasid   WHERE a.ACTIVE='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLSAMDEPTMAS");
            DataTable dt = ds.Tables["ASPTBLSAMDEPTMAS"];
            combocategory.DataSource = dt;
            combocategory.DisplayMember = "category";
            combocategory.ValueMember = "category";

          
        }
        void empty()
        {
            combostyle.SelectedIndex = -1;         
            combocompcode.SelectedIndex = -1;lblcount.Text = "";
            combocategory.SelectedIndex = -1; combobrand.SelectedIndex = -1; combopacktype1.SelectedIndex = -1; comboseason.SelectedIndex = -1; combodept.SelectedIndex = -1; txtrefcode.Text = "";
            lblcount.Text = "";
        }
        public void News()
        {
        
            empty();
        }

        public void Saves()
        {

        }

        public void Prints()
        {

        }

        public void Searchs()
        {

        }

        public void Deletes()
        {

        }

        public void ReadOnlys()
        {
        }

        public void Imports()
        {

        }
   
        public void Pdfs()
        {
            DialogResult result = MessageBox.Show("Do you want to  PDF Formate ??", "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result.Equals(DialogResult.OK))
            {
                if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
                {
                    rd.ExportToDisk(ExportFormatType.PortableDocFormat, folderLocation + "-" + path.ToString() + " " + Class.Users.HCompcode + " SampleStock.pdf");

                    /////  rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + path.ToString() + " " + Class.Users.HCompcode + "'SampleCollection.pdf");

                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
                {
                    rd.ExportToDisk(ExportFormatType.PortableDocFormat, folderLocation + "-" + path.ToString() + " " + Class.Users.HCompcode + " SampleOutWard.pdf");

                    /////  rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + path.ToString() + " " + Class.Users.HCompcode + "'SampleCollection.pdf");

                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
                {

                    //sc.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + Class.Users.HCompcode + "'SampleComments.doc");

                    // sc.ExportToDisk(ExportFormatType.ExcelWorkbook, folderLocation + "-" + path.ToString() + " " + Class.Users.HCompcode + "'SampleComments.xls");
                    rchart.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + Class.Users.HCompcode + "'Samplechart.pdf");



                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])//your specific tabname
                {
                    sc.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + Class.Users.HCompcode + "'SampleComments.pdf");

                }
            }
            else
            {

            }
        }

        public void ChangePasswords()
        {

        }

        public void DownLoads()
        {


            DialogResult result = MessageBox.Show("Do you want to  EXCEL Formate ??", "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result.Equals(DialogResult.OK))
            {
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
                {
                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }

                    rd.ExportToDisk(ExportFormatType.ExcelWorkbook, folderLocation +  Class.Users.HCompcode + "'StockReport.xls");
                    return;
                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
                {
                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }

                    rd.ExportToDisk(ExportFormatType.ExcelWorkbook, folderLocation +  " " + Class.Users.HCompcode + "'OutwardReport.xls");
                    return;
                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
                {                



                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])//your specific tabname
                {
                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }
                    sc.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + Class.Users.HCompcode + "'SampleComments.doc");

                    // sc.ExportToDisk(ExportFormatType.ExcelWorkbook, folderLocation + "-" + path.ToString() + " " + Class.Users.HCompcode + "'SampleComments.xls");
                    sc.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + Class.Users.HCompcode + "'SampleComments.pdf");

                    return;

                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])//your specific tabname
                {

                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }

                    rd.ExportToDisk(ExportFormatType.ExcelWorkbook, folderLocation + Class.Users.HCompcode + "'PendingInwardReport.xls");
                }
                
            }
            else
            {

            }
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
            GlobalVariables.MdiPanel.Show();
            this.Hide(); GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
        }

        public void GridLoad()
        {
            string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE WHERE A.ACTIVE='T'   order by a.ASPTBLBUYSAMID desc ";

            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
            DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"]; rd.SetDataSource(dt0);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd;
            crystalReportViewer1.Refresh();
            lblcount.Text = "Total Row's : " + dt0.Rows.Count;
        }


        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            outreport();
            tabControl1.SelectTab("tabPage2");
        }

        private void comboseason_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboseason.Text == "ALL")
            {
                GridLoad();
            }
            if (comboseason.Text != "ALL" && Class.Users.HCompcode != "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where A.ACTIVE='T' AND  e.season='" + comboseason.Text + "' and b.compcode='" + Class.Users.HCompcode + "'  order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Season : " + comboseason.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "Season " + comboseason.Text + "CompCode  ";
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh();
                lblcount.Text = "Season : " + comboseason.Text + ",  Total  Row's: " + dt0.Rows.Count; comboseason.Select();
            }
            if (comboseason.Text != "ALL" && Class.Users.HCompcode == "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where A.ACTIVE='T' AND  e.season='" + comboseason.Text + "'   order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Season : " + comboseason.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "Season " + comboseason.Text + "CompCode  ";
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); comboseason.Select();

            }
        }

        private void combodept_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combodept.Text == "ALL")
            {
                GridLoad();
            }
            if (combodept.Text != "ALL" && Class.Users.HCompcode != "ALL")
            {
                string sel0 = "SELECT  A.ASPTBLBUYSAMID, A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where A.ACTIVE='T' AND  f.department='" + combodept.Text + "' and b.compcode='" + Class.Users.HCompcode + "' order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"]; rd.SetDataSource(dt0);
                lblcount.Text = "Department : " + combodept.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "Department " + combodept.Text + "CompCode  ";
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combodept.Select();

            }
            if (combodept.Text != "ALL" && Class.Users.HCompcode == "ALL")
            {
                string sel0 = "SELECT  A.ASPTBLBUYSAMID, A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where A.ACTIVE='T' AND  f.department='" + combodept.Text + "'  order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Department : " + combodept.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "Department " + combopacktype1.Text + "CompCode  ";
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combodept.Select();

            }
        }

        private void combopacktype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopacktype1.Text == "ALL")
            {
                GridLoad();
            }
            if (combopacktype1.Text != "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where A.ACTIVE='T' AND  P.ORDERPACKTYPE='" + combopacktype1.Text + "' and b.compcode='" + Class.Users.HCompcode + "'  order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Order PackType : " + combopacktype1.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "OrderPackType " + combopacktype1.Text + "CompCode  "; 
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combopacktype1.Select();

            }
            if (combopacktype1.Text != "ALL" && Class.Users.HCompcode == "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where A.ACTIVE='T' AND  P.ORDERPACKTYPE='" + combopacktype1.Text + "'  order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh();
                lblcount.Text = "Order PackType : " + combopacktype1.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "OrderPackType " + combopacktype1.Text + "CompCode  "; combopacktype1.Select();
            }
        }

        private void txtrefcode_TextChanged(object sender, EventArgs e)
        {
            if (txtrefcode.Text.Length >= 5)
            {
                string sel0 = "SELECT  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where  A.AGFSAMPLE='" + txtrefcode.Text + "'   order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"]; rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh();

                crystalReportViewer2.ReportSource = null;
                crystalReportViewer2.ReportSource = rd;
                crystalReportViewer2.Refresh();
               
                lblcount.Text = "RefCode : " + txtrefcode.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "RefCode " + txtrefcode.Text + "CompCode  "; txtrefcode.Select();

                correctionreport();
            }
        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Class.Users.HCompcode == "ALL")
            {
                GridLoad();
            }
            if (Class.Users.HCompcode != "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where A.ACTIVE='T' AND  B.COMPCODE='" + Class.Users.HCompcode + "' and b.compcode='" + Class.Users.HCompcode + "'  order by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh();
                lblcount.Text = "CompCode : " + Class.Users.HCompcode + ",  Total  Row's: " + dt0.Rows.Count;
              
            }

        }

        private void combobrand_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combobrand.Text == "ALL")
            {
                GridLoad();
            }
            if (combobrand.Text != "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where a.active='T' AND  C.BRAND='" + combobrand.Text + "' and b.compcode='" + Class.Users.HCompcode + "'  order by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Brand Name : " + combobrand.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "Category Name " + combobrand.Text + "CompCode  ";
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combobrand.Select();
            }
            if (combobrand.Text != "ALL" && Class.Users.HCompcode == "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where a.active='T' AND  C.BRAND='" + combobrand.Text + "'  order by a.ASPTBLBUYSAMID desc ";

                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Brand Name : " + combobrand.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "Category Name " + combobrand.Text + "CompCode  ";
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combobrand.Select();

            }
        }

        private void combofabric_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combocategory.Text == "ALL")
            {
                GridLoad();
            }
            if (Class.Users.HCompcode != "ALL" && combocategory.Text != "ALL")
            {
                combodept.Text = "";
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where a.active='T' AND   O.CATEGORY='" + combocategory.Text + "' and b.compcode='" + Class.Users.HCompcode + "'  order by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Category Name : " + combocategory.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";
                path = "Category Name " + combocategory.Text+ "CompCode  ";
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combocategory.Select();

            }

            if (combocategory.Text != "ALL" && Class.Users.HCompcode == "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where  a.active='T' AND   O.CATEGORY='" + combocategory.Text + "'  order by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                lblcount.Text = "Category Name : " + combocategory.Text + ",  Total  Row's: " + dt0.Rows.Count;
                path = "";              
                path = "Category Name " + combocategory.Text + "CompCode  ";
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); combocategory.Select();

            }
        }

        
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
          
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
                stockreport();
               
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            {
                outreport();

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
            {
                reportchart();

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])//your specific tabname
            {
                correctionreport();

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])//your specific tabname
            {
                Pendingreport();

            }
        }

        private void butdatewise_Click(object sender, EventArgs e)
        {
            if (Class.Users.HCompcode == "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where a.active='T' AND  a.date1='" + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "'   order by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                path = "";
                path = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                lblcount.Text = "Date : " + dateTimePicker1.Value.ToString("dd-MM-yyyy") + ",  Total  Row's: " + dt0.Rows.Count;
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh();
              
            }
            if (Class.Users.HCompcode != "ALL")
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where a.active='T' AND  a.date1='" + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "' and b.compcode='" + Class.Users.HCompcode + "'  order by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                path = "";
                path = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                lblcount.Text = "CompCode : " + Class.Users.HCompcode + ",  Total  Row's: " + dt0.Rows.Count;
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh();
               
            }

        }

        private void combogsm_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void txtgsm_TextChanged(object sender, EventArgs e)
        {

            if (txtgsm.Text != "System.Data.DataRowView" && txtgsm.Text.Length>=3)
            {
                string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,b.compname,C.BRAND,A.AGFSAMPLE,E.SEASON || ' /' ||  A.MFYEAR || '' AS SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE where a.active='T' AND  a.gsm='" + txtgsm.Text + "' and b.compcode='" + Class.Users.HCompcode + "'  order by a.ASPTBLBUYSAMID desc ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
                DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];
                path = "";
                path = dateTimePicker1.Value.ToString("dd-MM-yyyy");
                lblcount.Text = "CompCode : " + Class.Users.HCompcode + ",  Total  Row's: " + dt0.Rows.Count;
                rd.SetDataSource(dt0);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh(); txtgsm.Select();
            }
        }

        private void txtgsm_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void stockRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            stockreport(); tabControl1.SelectTab("tabPage1");
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

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
    public partial class MovementReport : Form,ToolStripAccess
    {
      

        private static MovementReport _instance;
        public static MovementReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new MovementReport();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public MovementReport()
        {
            InitializeComponent();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }

  

        Models.UserRights sm = new Models.UserRights(); Models.Master mas = new Models.Master();
        Report.MovementReport rd = new Report.MovementReport();
        Report.MovementReport2 rd2 = new Report.MovementReport2();
       // Report.MovementReport1 rd1 = new Report.MovementReport1();
        DataTable dtgeneral = new DataTable();


        public void ReadOnlys()
        {

        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void COMBCODELOAD()
        {
            try
            {
                DataTable dt3 = new DataTable();
                //if (combohostel.Text == "MENS HOTEL" || combohostel.Text == "WORKING GENTS HOSTEL" || combohostel.Text == "WOMENS HOSTEL" || combohostel.Text == "GENTS STAFF HOSTEL" || combohostel.Text == "BOYS HOSTEL")
               // {
                    string sel3 = " SELECT DISTINCT B.GTCOMPMASTID, B.COMPCODE FROM HOSTELLIVEDATA A JOIN GTCOMPMAST B ON B.COMPCODE= A.COMPCODE  UNION ALL SELECT DISTINCT 1 AS GTCOMPMASTID, 'PASS MISSED' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 2 AS GTCOMPMASTID, 'NATIVE' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 3 AS GTCOMPMASTID, 'SECURITY' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 4 AS GTCOMPMASTID, 'TOTAL COUNT' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 5 AS GTCOMPMASTID, 'REMARKS' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 6 AS GTCOMPMASTID, 'LEAVE' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 7 AS GTCOMPMASTID, 'HOSTEL OUTING' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 8 AS GTCOMPMASTID, 'RESIGNATION' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 9 AS GTCOMPMASTID, 'WITHOUT-PHOTO' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 10 AS GTCOMPMASTID, 'AGF' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 11 AS GTCOMPMASTID, 'FLF' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 12 AS GTCOMPMASTID, 'FLFD' AS COMPCODE  FROM DUAL A ORDER BY 2";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "GTCOMPMAST");
                    dt3 = ds3.Tables["GTCOMPMAST"];
                //}
                //if (combohostel.Text != "MENS HOTEL" || combohostel.Text != "WORKING GENTS HOSTEL" || combohostel.Text != "WOMENS HOSTEL" || combohostel.Text != "GENTS STAFF HOSTEL" || combohostel.Text != "BOYS HOSTEL")
                //{
                //    string sel3 = "SELECT DISTINCT  A.HOSTELNAME as COMPCODE  FROM ASPTBLHOSTELGATEPASS A where A.HOSTELNAME not in 'MENS HOTEL'  and A.HOSTELNAME not in 'GENTS STAFF HOSTEL'      and A.HOSTELNAME not in 'BOYS HOSTEL'  and A.HOSTELNAME not in 'WOMENS HOSTEL' and A.HOSTELNAME not in 'WORKING GENTS HOSTEL'";
                //    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLHOSTELGATEPASS");
                //     dt3 = ds3.Tables["ASPTBLHOSTELGATEPASS"];
                //}
                //if (Class.Users.HUserName == "VAIRAM")
                //{
                //    string sel3 = " SELECT DISTINCT B.GTCOMPMASTID, B.COMPCODE FROM HOSTELLIVEDATA A JOIN GTCOMPMAST B ON B.COMPCODE= A.COMPCODE  UNION ALL SELECT DISTINCT 1 AS GTCOMPMASTID, 'PASS MISSED' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 2 AS GTCOMPMASTID, 'NATIVE' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 3 AS GTCOMPMASTID, 'SECURITY' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 4 AS GTCOMPMASTID, 'TOTAL COUNT' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 5 AS GTCOMPMASTID, 'REMARKS' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 6 AS GTCOMPMASTID, 'LEAVE' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 7 AS GTCOMPMASTID, 'HOSTEL OUTING' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 8 AS GTCOMPMASTID, 'RESIGNATION' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 9 AS GTCOMPMASTID, 'WITHOUT-PHOTO' AS COMPCODE  FROM DUAL A  UNION ALL SELECT DISTINCT 10 AS GTCOMPMASTID, 'AGF' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 11 AS GTCOMPMASTID, 'FLF' AS COMPCODE  FROM DUAL A UNION ALL SELECT DISTINCT 12 AS GTCOMPMASTID, 'FLFD' AS COMPCODE  FROM DUAL A ORDER BY 2";
                //    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "GTCOMPMAST");
                //    dt3 = ds3.Tables["GTCOMPMAST"];
                //}
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
        private void Exit_Click(object sender, EventArgs e)
        {
            crystalReportViewer1.ReportSource = null;
            this.Hide();
        }



        private void MovementReport_Load(object sender, EventArgs e)
        {

            this.txtsearch.Focus(); comboformate.SelectedIndex = 1;
            frmdate.Value = DateTime.Now.AddDays(0); todate.Value = DateTime.Now.AddDays(0);
            hostelload();
            COMBCODELOAD();
            idcardsearch(); comboidcardsearch.Text = ""; comboidcardsearch.SelectedIndex = -1;
            if (Class.Users.HUserName == "VAIRAM" || Class.Users.HostelName == "MENS HOTEL" || Class.Users.HostelName == "WORKING GENTS HOSTEL" || Class.Users.HostelName == "WOMENS HOSTEL" || Class.Users.HostelName == "GENTS STAFF HOSTEL" || Class.Users.HostelName == "BOYS HOSTEL")
            {
                combohostel.Enabled = true;
            }
            else
            {
                combohostel.Enabled = false;
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (combocompcode.Text == "")
            //{
            //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME,substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE   B.COMPCODE='"+combocompcode.Text+"' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";

            //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
            //    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
            //    rd.SetDataSource(dt2);
            //    crystalReportViewer1.ReportSource = null;
            //    crystalReportViewer1.ReportSource = rd;

            //    crystalReportViewer1.Refresh(); txtsearch.Text = "";
            //}
        }

        public void DownLoads()
        {
           

            if (comboformate.Text != "")
            {

                DialogResult result = MessageBox.Show("Do you want to '" + comboformate.Text + "' Formate ??", "" + comboformate.Text + "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    if (combocompcode.Text == "TOTAL COUNT")//your specific tabname
                    {
                        switch (comboformate.Text)
                        {
                            case "Word":
                                rd2.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + combocompcode.Text + "'MonthwiseReport.doc");
                                break;

                            case "Excel":                               
                                rd2.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + combocompcode.Text + "'MonthwiseReport.xls");
                                break;

                            case "PDF":
                                rd2.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + combocompcode.Text + "'MonthwiseReport.pdf");
                                break;

                            case "CSV":
                                rd2.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + combocompcode.Text + "'MonthwiseReport.csv");
                                break;
                        }
                    }
                    else
                    {
                        switch (comboformate.Text)
                        {
                            case "Word":
                                rd.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + combocompcode.Text + "'MonthwiseReport.doc");
                                break;

                            case "Excel":
                                // formatType = ExportFormatType.Excel;
                                rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy")+"  " + combocompcode.Text + "'MonthwiseReport.xls");
                                break;

                            case "PDF":
                                rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + combocompcode.Text + "'MonthwiseReport.pdf");
                                break;

                            case "CSV":
                                rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "d:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + combocompcode.Text + "'MonthwiseReport.csv");
                                break;
                        }
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
        private void hostelload()
        {
            try
            {
                DataTable dt3 = new DataTable();
               // if (combohostel.Text == "MENS HOTEL" || combohostel.Text == "WORKING GENTS HOSTEL" || combohostel.Text == "WOMENS HOSTEL" || combohostel.Text == "GENTS STAFF HOSTEL" || combohostel.Text == "BOYS HOSTEL")
               // {
                    string sel3 = "SELECT  '' AS HOSTELNAME FROM DUAL  UNION ALL SELECT DISTINCT  A.HOSTELNAME FROM ASPTBLHOSTELGATEPASS A where A.HOSTELNAME not in 'AGF'  and A.HOSTELNAME not in 'AGFM'      and A.HOSTELNAME not in 'FLF'  and A.HOSTELNAME not in 'FLFD' ";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLHOSTELGATEPASS");
                    dt3 = ds3.Tables["ASPTBLHOSTELGATEPASS"];
                //}
                //if (combohostel.Text != "MENS HOTEL" || combohostel.Text != "WORKING GENTS HOSTEL" || combohostel.Text != "WOMENS HOSTEL" || combohostel.Text != "GENTS STAFF HOSTEL" || combohostel.Text != "BOYS HOSTEL")
                //{
                //    string sel3 = "SELECT DISTINCT  A.HOSTELNAME FROM ASPTBLHOSTELGATEPASS A where A.HOSTELNAME not in 'MENS HOTEL'  and A.HOSTELNAME not in 'GENTS STAFF HOSTEL'      and A.HOSTELNAME not in 'BOYS HOSTEL'  and A.HOSTELNAME not in 'WOMENS HOSTEL' and A.HOSTELNAME not in 'WORKING GENTS HOSTEL'";
                //    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLHOSTELGATEPASS");
                //    dt3 = ds3.Tables["ASPTBLHOSTELGATEPASS"];
                //}
                if (dt3.Rows.Count > 0)
                {
                    combohostel.DisplayMember = "HOSTELNAME";
                    combohostel.ValueMember = "HOSTELNAME";
                    combohostel.DataSource = dt3;
                    

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
        }
        private void idcardsearch()
        {
            try
            {
                string sel3 = ""; DataTable dt3 = new DataTable();
                 sel3 = "SELECT distinct A.IDCARDNO FROM ASPTBLHOSTELGATEPASS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID= A.COMPCODE   JOIN HREMPLOYMAST C ON C.COMPCODE = B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON C.HREMPLOYMASTID = D.HREMPLOYMASTID    AND D.HOSTEL='YES' JOIN GTDEPTDESGMAST E ON E.GTDEPTDESGMASTID = D.DEPTNAME       JOIN HRECONTACTDETAILS F ON F.HREMPLOYMASTID = C.HREMPLOYMASTID AND F.HREMPLOYMASTID = D.HREMPLOYMASTID        and D.IDACTIVE='YES'  AND B.COMPCODE='"+Class.Users.HCompcode+ "' join ASPTBLMACHINEMAS I on I.COMPCODE=B.GTCOMPMASTID  WHERE A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLHOSTELGATEPASS");
                 dt3 = ds3.Tables["ASPTBLHOSTELGATEPASS"];
                if (dt3.Rows.Count <= 0)
                {
                     sel3 = "SELECT distinct A.IDCARDNO FROM ASPTBLHOSTELGATEPASS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID= A.COMPCODE   JOIN HREMPLOYMAST C ON C.COMPCODE = B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON C.HREMPLOYMASTID = D.HREMPLOYMASTID  and D.IDACTIVE='YES'  AND B.COMPCODE='" + Class.Users.HCompcode + "'  JOIN HRECONTACTDETAILS F ON F.HREMPLOYMASTID = C.HREMPLOYMASTID AND F.HREMPLOYMASTID = D.HREMPLOYMASTID   join ASPTBLMACHINEMAS I on I.COMPCODE=B.GTCOMPMASTID WHERE A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                    DataSet ds4 = Utility.ExecuteSelectQuery(sel3, "ASPTBLHOSTELGATEPASS");
                    dt3 = ds4.Tables["ASPTBLHOSTELGATEPASS"];
                }
                if (dt3.Rows.Count > 0)
                {
                    comboidcardsearch.DisplayMember = "MIDCARD";
                    comboidcardsearch.ValueMember = "MIDCARD";
                    comboidcardsearch.DataSource = dt3;


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
                lblcount.Visible = false; dtgeneral = null;
                //if (combohostel.Text != "")
                //{
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE  A.HOSTELNAME='" + combohostel.Text + "' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dtgeneral);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh();
                //    lblcount.Visible = true; lblcount.Text = "Total PASS MISSED Employee Count:-" + dtgeneral.Rows.Count.ToString();
                //    txtsearch.Text = ""; return;
                //}
                if (combohostel.Text == "MENS HOTEL" || combohostel.Text == "WORKING GENTS HOSTEL" || combohostel.Text == "WOMENS HOSTEL" || combohostel.Text == "GENTS STAFF HOSTEL" || combohostel.Text == "BOYS HOSTEL")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE  B.COMPCODE='" + combocompcode.Text + "'  AND A.HOSTELNAME='" + combohostel.Text + "' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                    lblcount.Visible = true; lblcount.Text = "Total PASS MISSED Employee Count:-" + dtgeneral.Rows.Count.ToString();
                    txtsearch.Text = ""; return;
                }
                //if (combocompcode.Text == "AGF" || combocompcode.Text == "FLF" || combocompcode.Text == "FLFD"  || combohostel.Text == "")
                //{
                //if (combocompcode.Text != "MENS HOTEL" || combocompcode.Text != "WORKING GENTS HOSTEL" || combocompcode.Text != "WOMENS HOSTEL" || combocompcode.Text != "GENTS STAFF HOSTEL" || combocompcode.Text != "BOYS HOSTEL")
                else
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,A.HOSTELNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,A.HOSTELNAME , '' AS ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE   A.HOSTELNAME='" + combocompcode.Text + "' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                    lblcount.Visible = true; lblcount.Text = "Total  Employee Count:-" + dtgeneral.Rows.Count.ToString();
                    txtsearch.Text = ""; return;
                }
                //if (combocompcode.Text == "AGFC" && combohostel.Text != "")
                //{
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE  B.COMPCODE='" + combocompcode.Text + "'  AND A.HOSTELNAME='" + combohostel.Text + "' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dtgeneral);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh();
                //    lblcount.Visible = true; lblcount.Text = "Total PASS MISSED Employee Count:-" + dtgeneral.Rows.Count.ToString();
                //    txtsearch.Text = ""; return;
                //}
                //if (combocompcode.Text == "AGFM" && combohostel.Text != "")
                //{
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE  B.COMPCODE='" + combocompcode.Text + "'  AND A.HOSTELNAME='" + combohostel.Text + "'  AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dtgeneral);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh();
                //    lblcount.Visible = true; lblcount.Text = "Total PASS MISSED Employee Count:-" + dtgeneral.Rows.Count.ToString();
                //    txtsearch.Text = ""; return;
                //}

                //if (combocompcode.Text == "AGFMGII" && combohostel.Text != "")
                //{
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE  B.COMPCODE='" + combocompcode.Text + "'  AND A.HOSTELNAME='" + combohostel.Text + "' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dtgeneral);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh();
                //    lblcount.Visible = true; lblcount.Text = "Total PASS MISSED Employee Count:-" + dtgeneral.Rows.Count.ToString();
                //    txtsearch.Text = ""; return;
                //}
                if (combocompcode.Text == "PASS MISSED" && combohostel.Text != "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE A.PASSMISSED='T' AND A.HOSTELNAME='" + combohostel.Text + "' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                    lblcount.Visible = true; lblcount.Text = "Total PASS MISSED Employee Count:-" + dtgeneral.Rows.Count.ToString();
                    txtsearch.Text = ""; return;
                }

                if (combocompcode.Text == "NATIVE" && combohostel.Text != "")
                {

                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE NOT E.mnname1='SECURITY'  AND A.NATIVE='T'   AND A.HOSTELNAME='" + combohostel.Text + "'  AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    // dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd; lblcount.Visible = true; lblcount.Text = "Total Native Employee Count:-" + dtgeneral.Rows.Count.ToString();
                    crystalReportViewer1.Refresh(); txtsearch.Text = ""; return;
                }
                if (combocompcode.Text == "TOTAL COUNT" && combohostel.Text != "")
                {
                    //  dtgeneral = null;
                    // string sel2 = "SELECT B.COMPCODE,G.REASON AS PLACE,  COUNT(A.ASPTBLHOSTELGATEPASSID)COUNT,1 ORD   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE JOIN HREMPLOYMAST    C ON C.COMPCODE = B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID = C.HREMPLOYMASTID AND D.IDCARD = C.IDCARDNO AND D.DEPTNAME = A.DEPARTMENT AND D.MIDCARD = A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE = B.COMPCODE AND F.IDCARDNO = C.IDCARDNO AND F.IDCARDNO = D.IDCARD AND F.HOSTELNAME = A.HOSTELNAME  AND D.IDCARD = F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID = A.REASON JOIN HRECONTACTDETAILS H ON H.HREMPLOYMASTID = C.HREMPLOYMASTID WHERE NOT E.mnname1 = 'SECURITY' AND A.HOSTELNAME = 'WORKING GENTS HOSTEL' and A.MODIFIED between TO_DATE('11-10-2021', 'dd-MM-yyyy') and TO_DATE('11-10-2021','dd-MM-yyyy') GROUP BY B.COMPCODE,G.REASON UNION ALL SELECT B.COMPCODE,'TOTAL' PLACE,  COUNT(A.ASPTBLHOSTELGATEPASSID)COUNT,2 ORD FROM ASPTBLHOSTELGATEPASS A  JOIN GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN HREMPLOYMAST    C ON C.COMPCODE = B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID = C.HREMPLOYMASTID AND D.IDCARD = C.IDCARDNO AND D.DEPTNAME = A.DEPARTMENT AND D.MIDCARD = A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE = B.COMPCODE AND F.IDCARDNO = C.IDCARDNO AND F.IDCARDNO = D.IDCARD AND F.HOSTELNAME = A.HOSTELNAME  AND D.IDCARD = F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID = A.REASON JOIN HRECONTACTDETAILS H ON H.HREMPLOYMASTID = C.HREMPLOYMASTID  WHERE NOT E.mnname1 = 'SECURITY' AND A.HOSTELNAME = 'WORKING GENTS HOSTEL' and A.MODIFIED between TO_DATE('11-10-2021', 'dd-MM-yyyy') and TO_DATE('11-10-2021','dd-MM-yyyy')     GROUP BY B.COMPCODE     ORDER BY 1,3";
                    string sel2 = "SELECT b.compcode,  G.REASON AS PLACE, A.ASPTBLHOSTELGATEPASSID COUNT   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE     JOIN HREMPLOYMAST    C ON C.COMPCODE = B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID = C.HREMPLOYMASTID AND D.IDCARD = C.IDCARDNO AND D.DEPTNAME = A.DEPARTMENT     AND D.MIDCARD = A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE = B.COMPCODE AND F.IDCARDNO = C.IDCARDNO     AND F.IDCARDNO = D.IDCARD AND F.HOSTELNAME = A.HOSTELNAME  AND D.IDCARD = F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID = A.REASON   JOIN HRECONTACTDETAILS H ON H.HREMPLOYMASTID = C.HREMPLOYMASTID  WHERE NOT E.mnname1='SECURITY' AND A.HOSTELNAME='" + combohostel.Text + "'  AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')           ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    crystalReportViewer1.ReportSource = null;
                    rd2.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = rd2;
                    crystalReportViewer1.Refresh();

                    //lblcount.Text = "Total Employee Count:-" + dtgeneral.Rows[0]["COUNT"].ToString();

                    lblcount.Visible = true;
                    txtsearch.Text = ""; return;
                }

                if (combocompcode.Text == "SECURITY" && combohostel.Text != "")
                {

                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE E.mnname1='SECURITY'    AND A.HOSTELNAME='" + combohostel.Text + "'  AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    // dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                    lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString();
                    txtsearch.Text = ""; return;
                }



                if (combocompcode.Text != "" && combohostel.Text == "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE NOT E.mnname1='SECURITY'  AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    //  dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                }

                if (combocompcode.Text == "REMARKS" && combohostel.Text != "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE NOT E.mnname1='SECURITY'  AND A.HOSTELNAME='" + combohostel.Text + "'  AND  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') AND A.REMARKS IS NOT NULL ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    //  dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                }
                if (combocompcode.Text == "RESIGNATION" || combocompcode.Text == "LEAVE" || combocompcode.Text == "HOSTEL OUTING" && combohostel.Text != "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE NOT E.mnname1='SECURITY' AND G.REASON='"+combocompcode.Text+"' AND A.HOSTELNAME='" + combohostel.Text + "'  AND  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    //  dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                }
                //if (combocompcode.Text == "LEAVE" && combohostel.Text != "")
                //{
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE NOT E.mnname1='SECURITY'  AND G.REASON='" + combocompcode.Text + "' AND A.HOSTELNAME='" + combohostel.Text + "' AND  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    dtgeneral = null;
                //    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dtgeneral);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;

                //}
                //if (combocompcode.Text == "RESIGNATION" && combohostel.Text != "")
                //{
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE NOT E.mnname1='SECURITY'  AND G.REASON='" + combocompcode.Text + "' AND A.HOSTELNAME='" + combohostel.Text + "'  AND  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')  ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    //  dtgeneral = null;
                //    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dtgeneral);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                //}
                if (combocompcode.Text == "WITHOUT-PHOTO" && combohostel.Text != "")
                {
                    //   string sel2 = "SELECT  x.MIDCARD,x.HREMPLOYMASTID,x.FNAME,x.DISPNAME,x.photo,x.REMARKS from ( SELECT B.MIDCARD,A.HREMPLOYMASTID,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as FNAME,E.MNNAME1 as DISPNAME,  H.IMAGEBYTES AS PHOTO,C.COMPCODE as REMARKS FROM    HREMPLOYMAST    A   JOIN HREMPLOYDETAILS B ON A.HREMPLOYMASTID=B.HREMPLOYMASTID AND B.IDACTIVE='YES' AND A.IDCARDNO=B.IDCARD   JOIN   GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = B.DEPTNAME    JOIN HRECONTACTDETAILS F ON E.COMPCODE = A.COMPCODE AND E.COMPCODE = C.GTCOMPMASTID   AND F.HREMPLOYMASTID = A.HREMPLOYMASTID AND F.HREMPLOYMASTID = B.HREMPLOYMASTID    join hostellivedata g on G.COMPCODE=C.COMPCODE and G.IDCARDNO=B.IDCARD and B.HOSTEL='YES'  left outer join  ASPTBLEMP h on  H.COMPCODE=c.GTCOMPMASTID and H.IDCARDNO=B.MIDCARD   and H.EMPID=A.HREMPLOYMASTID   )x where x.photo is null   ORDER BY 6";
                    string sel2 = "SELECT  x.MIDCARD,x.HREMPLOYMASTID,x.FNAME,x.DISPNAME,x.photo,x.REMARKS , '-' as ASPTBLHOSTELGATEPASSID,x.COMPCODE,x.DESIGNATION,x.CONTACTNO,x.HOSTELNAME ,x.ROOMNO,x.REASON,x.PERMISSIONHRS,x.INTIME,x.OUTTIME, x.BLOCKFLOOR from ( SELECT B.MIDCARD,a.HREMPLOYMASTID,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as FNAME,E.MNNAME1 as DISPNAME,  H.IMAGEBYTES AS PHOTO, '-' as ASPTBLHOSTELGATEPASSID,c.compname AS COMPCODE,'ALL' as DESIGNATION,C.COMPCODE as REMARKS,'-' as CONTACTNO,'-' as HOSTELNAME ,'-' as ROOMNO,'-' as REASON,'-' as PERMISSIONHRS,'-' as INTIME,'-' as OUTTIME,  '-' as  BLOCKFLOOR  FROM    HREMPLOYMAST    A   JOIN HREMPLOYDETAILS B ON A.HREMPLOYMASTID=B.HREMPLOYMASTID AND B.IDACTIVE='YES' AND A.IDCARDNO=B.IDCARD   JOIN   GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = B.DEPTNAME    JOIN HRECONTACTDETAILS F ON E.COMPCODE = A.COMPCODE AND E.COMPCODE = C.GTCOMPMASTID   AND F.HREMPLOYMASTID = A.HREMPLOYMASTID AND F.HREMPLOYMASTID = B.HREMPLOYMASTID    join hostellivedata g on G.COMPCODE=C.COMPCODE and G.IDCARDNO=B.IDCARD and B.HOSTEL='YES'  left outer join  ASPTBLEMP h on  H.COMPCODE=c.GTCOMPMASTID and H.IDCARDNO=B.MIDCARD   and H.EMPID=A.HREMPLOYMASTID   )x where x.photo is null   ORDER BY 6";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "HREMPLOYMAST");
                    //  dtgeneral = null;
                    dtgeneral = ds2.Tables["HREMPLOYMAST"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                }

                if (comboidcardsearch.Text != "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE, B.COMPCODE as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT   LEFT JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID   WHERE  D.MIDCARD='" + comboidcardsearch.Text + "'   ORDER BY A.ASPTBLHOSTELGATEPASSID DESC";

                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                  //  dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = "";comboidcardsearch.Text = ""; return;
                }
                if (combocompcode.Text == "" && combohostel.Text != "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE NOT E.mnname1='SECURITY'  AND A.HOSTELNAME='" + combohostel.Text + "' and  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') AND A.REMARKS IS NOT NULL ORDER BY 1";

                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                }
                //if (combocompcode.Text == "" && combohostel.Text == "")
                //{
                //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID   WHERE NOT E.mnname1='SECURITY'  AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                //    //  dtgeneral = null;
                //    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                //    rd.SetDataSource(dtgeneral);
                //    crystalReportViewer1.ReportSource = null;
                //    crystalReportViewer1.ReportSource = rd;
                //    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                //}

                if (combocompcode.Text == "" && combohostel.Text != "" && comboidcardsearch.Text == "" )
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE NOT E.mnname1='SECURITY'  AND A.HOSTELNAME='" + combohostel.Text + "' and  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                     dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh(); lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dtgeneral.Rows.Count.ToString(); txtsearch.Text = ""; return;
                }
                if (combocompcode.Text == "" && combohostel.Text == "" && comboidcardsearch.Text != "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID     WHERE  D.MIDCARD='" + comboidcardsearch.Text + "'    ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                    lblcount.Visible = true; lblcount.Text = "Total PASS MISSED Employee Count:-" + dtgeneral.Rows.Count.ToString();
                    txtsearch.Text = ""; return;
                }
            }
            catch (Exception EX)
            { MessageBox.Show(EX.Message); }
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtsearch.Text != "")
                {

                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE NOT E.mnname1='SECURITY'  AND A.HOSTELNAME='" + combohostel.Text + "' and  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') AND  D.MIDCARD LIKE'%" + txtsearch.Text + "%' OR  C.FNAME LIKE'%" + txtsearch.Text + "%' ORDER BY 1";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh();   txtsearch.Focus();
                }
                if (combocompcode.Text != null && txtsearch.Text != "")
                {
                    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE NOT E.mnname1='SECURITY'  AND A.HOSTELNAME='" + combohostel.Text + "' and  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') AND  D.MIDCARD LIKE'%" + txtsearch.Text + "%'  OR  C.FNAME LIKE'%" + txtsearch.Text + "%' ORDER BY 1";

                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                    dtgeneral = null;
                    dtgeneral = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                    rd.SetDataSource(dtgeneral);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh(); txtsearch.Focus();
                }
                else
                {
                    // butView_Click(sender,e);
                }
            }
            catch (Exception EX)
            { MessageBox.Show(EX.Message); }
        }

        private void comboformate_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

       

        public void Prints()
        {
            try
            {
                
                    if (printDialog1.ShowDialog() == DialogResult.OK)
                    {
                        //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'" + combocompcode.Text + "' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID    WHERE  B.COMPCODE='" + combocompcode.Text + "' AND A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
                        //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                        //    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];

                        if (combocompcode.Text == "TOTAL COUNT" && combohostel.Text != "")
                        {

                            //rd2.SetDataSource(dtgeneral);                       
                            //crystalReportViewer1.ReportSource = null;
                            //crystalReportViewer1.ReportSource = rd2;
                            //crystalReportViewer1.Refresh();
                            //.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                            //rd2.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

                            CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                            reportdocument.Load(Application.StartupPath + "\\Report\\MovementReport2.rpt");
                            reportdocument.SetDataSource(dtgeneral);
                            reportdocument.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                            reportdocument.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

                        }
                        else
                        {



                            CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                            rd.Load(Application.StartupPath + "\\Report\\MovementReport.rpt");
                            rd.SetDataSource(dtgeneral);
                            rd.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                            rd.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

                           
                        }
                    }
                
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }

        private void rToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void outTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            combocompcode.Text = "OutTime Only";
            string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE NOT E.mnname1='SECURITY'  AND A.INTIME IS   NULL    AND A.OUTTIME IS NOT  NULL  AND A.HOSTELNAME='" + combohostel.Text + "' and  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";

            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
            DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
            rd.SetDataSource(dt2);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd;
            lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dt2.Rows.Count.ToString();
            crystalReportViewer1.Refresh(); txtsearch.Text = "";

        }

        private void News_Click(object sender, EventArgs e)
        {

        }

        private void combohostel_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboidcardsearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            combocompcode.Text = "";combocompcode.SelectedIndex = -1;
            //if (comboidcardsearch.Text != "")
            //{
            //    string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID   WHERE   A.IDCARDNO='" + comboidcardsearch.Text + "'  ORDER BY A.ASPTBLHOSTELGATEPASSID DESC";

            //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
            //    DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
            //    rd.SetDataSource(dt2);
            //    crystalReportViewer1.ReportSource = null;
            //    crystalReportViewer1.ReportSource = rd;
            //    crystalReportViewer1.Refresh(); txtsearch.Text = ""; return;
            //}
        }

        private void Prints_Click_1(object sender, EventArgs e)
        {

        }

        private void Pdfs_Click(object sender, EventArgs e)
        {

        }

        private void DownLoads_Click_1(object sender, EventArgs e)
        {

        }

        private void comboformate_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void withoutInTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE NOT E.mnname1='SECURITY'  AND A.INTIME IS   NOT NULL  AND A.OUTTIME IS NOT  NULL AND A.HOSTELNAME='" + combohostel.Text + "' and  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";
            combocompcode.Text = "InTime & OutTime Only";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
            DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
            rd.SetDataSource(dt2);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd;

            crystalReportViewer1.Refresh();
            lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dt2.Rows.Count.ToString();
            txtsearch.Text = "";
        }

        private void withoutInTimeOutTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,'ALL' as DESIGNATION, D.MIDCARD ,C.FNAME ,D.DEPT AS DISPNAME, substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME,  A.REMARKS,A.TIMEDIFF as BLOCKFLOOR   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID  JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO     JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO    AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE NOT E.mnname1='SECURITY'  AND A.INTIME IS  NULL AND A.OUTTIME IS  NULL AND A.HOSTELNAME='" + combohostel.Text + "' and  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') ORDER BY 1";

            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
            DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
            combocompcode.Text = "Without InTime & OutTime";
            rd.SetDataSource(dt2);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd;
            lblcount.Visible = true; lblcount.Text = "Total  Count:-" + dt2.Rows.Count.ToString();
            crystalReportViewer1.Refresh(); txtsearch.Text = "";
        }

        public void News()
        {
            hostelload();
            COMBCODELOAD();
            idcardsearch();
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

        public void Pdfs()
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
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
            this.Hide();
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

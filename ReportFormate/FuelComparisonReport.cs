using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.ReportFormate
{
    public partial class FuelComparisonReport : Form,ToolStripAccess
    {
        public FuelComparisonReport()
        {
            InitializeComponent();

         
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }

        public void ReadOnlys()
        {

        }

        private static FuelComparisonReport _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        byte[] bytes; private int rowIndex = 0;
        public static FuelComparisonReport Instance
        {
            get { if (_instance == null) _instance = new FuelComparisonReport(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }


        Report.FuelComparison rd = new Report.FuelComparison();
   

        private void FuelComparisonReport_Load(object sender, EventArgs e)
        {
            DataTable dt1 = mas.findcomcode(Class.Users.HCompcode);
            if (dt1.Rows.Count > 0)
            {
                combocompcode.DisplayMember = "COMPCODE";
                combocompcode.ValueMember = "GTCOMPMASTID";
                combocompcode.DataSource = dt1;

            }

            this.txtsearch.Focus(); comboformate.SelectedIndex = 1;
        }

        private void News_Click(object sender, EventArgs e)
        {

        }

        private void butView_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    string sel2 = "SELECT    XY.COMPCODE,XY.ITEMNAME,XY.LITRES ,XY.FUELRATE2 ,SUM(XY.LITRES*XY.FUELRATE2) AS TOTAL FROM (SELECT    X.COMPCODE,X.ITEMNAME,SUM(X.LITRES) AS LITRES,SUM(X.FUELRATE2)AS FUELRATE2 FROM  (SELECT A.ASPTBLVEHTOKENID, D.COMPCODE,A.TOKENNO,  A.TOKENDATE,A.VEHICLENO, G.VEHTYPE ,E.FNAME, F.DEPT,  H.ITEMNAME,  CASE B.LITRES  WHEN  'FULL' THEN '0'    ELSE B.LITRES END as LITRES,      (SELECT to_number(AA.FUELRATE2)AS FUELRATE2 FROM ASPTBLFUELRATEMASDET AA   JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID       JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME       WHERE AB.COMPCODE = D.COMPCODE AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE=A.TOKENDATE ) AS FUELRATE2        FROM  ASPTBLVEHTOKEN A     JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID     JOIN  HRVEHMAST C ON A.VEHICLENO=C.HRVEHMASTID         JOIN GTCOMPMAST D ON D.GTCOMPMASTID=A.COMPCODE  JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME      JOIN hremploydetails F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID        join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid=C.VEHICLETYPE   JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID=B.ITEMNAME             WHERE  D.COMPCODE = '" + combocompcode.Text + "'  AND A.TOKENDATE between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy'))X GROUP BY  X.COMPCODE,X.ITEMNAME)XY GROUP BY  XY.COMPCODE,XY.ITEMNAME ,XY.LITRES ,XY.FUELRATE2 ORDER BY 1";
            //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLFUELRATEMASDET");
            //    DataTable dt2 = ds2.Tables["ASPTBLFUELRATEMASDET"];

            //    rd.SetDataSource(dt2);
            //    crystalReportViewer1.ReportSource = null;
            //    crystalReportViewer1.ReportSource = rd;

            //    crystalReportViewer1.Refresh();
            //}
            //catch (Exception EX)
            //{ MessageBox.Show(EX.Message); }
        }

        public void DownLoads()
        {
            //if (comboformate.Text != "")
            //{

            //    DialogResult result = MessageBox.Show("Do you want to '" + comboformate.Text + "' Formate ??", "" + comboformate.Text + "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //    if (result.Equals(DialogResult.OK))
            //    {
            //        // ExportFormatType formatType = ExportFormatType.NoFormat;                    
            //        switch (comboformate.Text)
            //        {
            //            case "Word":
            //                rd.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + combocompcode.Text + "'MonthwiseReport.doc");
            //                break;

            //            case "Excel":
            //                // formatType = ExportFormatType.Excel;
            //                rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + combocompcode.Text + "'MonthwiseReport.xls");
            //                break;

            //            case "PDF":
            //                rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + combocompcode.Text + "'MonthwiseReport.pdf");
            //                break;

            //            case "CSV":
            //                rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "d:\\'" + combocompcode.Text + "'MonthwiseReport.csv");
            //                break;
            //        }

            //    }
            //    else
            //    {

            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Pls Select Combo Box Value");
            //}
        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    string sel2 = "SELECT    XY.COMPCODE,XY.ITEMNAME,XY.LITRES ,XY.FUELRATE2 ,SUM(XY.LITRES*XY.FUELRATE2) AS TOTAL FROM (SELECT    X.COMPCODE,X.ITEMNAME,SUM(X.LITRES) AS LITRES,SUM(X.FUELRATE2)AS FUELRATE2 FROM  (SELECT A.ASPTBLVEHTOKENID, D.COMPCODE,A.TOKENNO,  A.TOKENDATE,A.VEHICLENO, G.VEHTYPE ,E.FNAME, F.DEPT,  H.ITEMNAME,  CASE B.LITRES  WHEN  'FULL' THEN '0'    ELSE B.LITRES END as LITRES,      (SELECT to_number(AA.FUELRATE2)AS FUELRATE2 FROM ASPTBLFUELRATEMASDET AA   JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID       JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME       WHERE AB.COMPCODE = D.COMPCODE AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE=A.TOKENDATE ) AS FUELRATE2        FROM  ASPTBLVEHTOKEN A     JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID     JOIN  HRVEHMAST C ON A.VEHICLENO=C.HRVEHMASTID         JOIN GTCOMPMAST D ON D.GTCOMPMASTID=A.COMPCODE  JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME      JOIN hremploydetails F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID        join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid=C.VEHICLETYPE   JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID=B.ITEMNAME             WHERE  D.COMPCODE = '" + combocompcode.Text + "'  AND A.TOKENDATE between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy'))X GROUP BY  X.COMPCODE,X.ITEMNAME)XY GROUP BY  XY.COMPCODE,XY.ITEMNAME ,XY.LITRES ,XY.FUELRATE2 ORDER BY 1";
            //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLFUELRATEMASDET");
            //    DataTable dt2 = ds2.Tables["ASPTBLFUELRATEMASDET"];

            //    rd.SetDataSource(dt2);
            //    crystalReportViewer1.ReportSource = null;
            //    crystalReportViewer1.ReportSource = rd;

            //    crystalReportViewer1.Refresh();
            //}
            //catch (Exception EX)
            //{ MessageBox.Show(EX.Message); }
        }

        public void Exit()
        {
            this.Hide();
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        public void News()
        {
           
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

        public void GridLoad()
        {
           
        }

        private void comboformate_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

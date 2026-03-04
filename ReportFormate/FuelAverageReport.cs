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
    public partial class FuelAverageReport : Form,ToolStripAccess
    {
        public FuelAverageReport()
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

        private static FuelAverageReport _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        byte[] bytes; private int rowIndex = 0;
        public static FuelAverageReport Instance
        {
            get { if (_instance == null) _instance = new FuelAverageReport(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }


        Report.FuelAverage rd = new Report.FuelAverage();
        

        private void FuelAverageReport_Load(object sender, EventArgs e)
        {

          

          comboformate.SelectedIndex = 1;
            // var ss = System.DateTime.Now.ToShortDateString();

            todate.Value = DateTime.Today.AddDays(0);
            frmdate.Value = DateTime.Today.AddDays(-1);
            //try
            //{
            //    string sel2 = "SELECT Y.COMPCODE, Y.FUELDATE,Y.ITEMNAME,Y.LITRES,Y.FUELRATE2 ,'0' AS TOTAL FROM(SELECT AB.COMPCODE,  AA.FUELDATE,AC.ITEMNAME,'0' AS LITRES, AA.FUELRATE2,'0'AS TOTAL FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE=AB.GTCOMPMASTID JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID=AA.ITEMNAME )Y";
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

        private void News_Click(object sender, EventArgs e)
        {

        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            butView_Click(sender,e);
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    if (txtsearch.Text != "")
            //    {
            //        //A.ASPTBLHOSTELGATEPASSID LIKE'%" + txtsearch.Text + "%' OR D.MIDCARD LIKE'%" + txtsearch.Text + "%' OR C.FNAME LIKE'%" + txtsearch.Text + "%'
            //        string sel2 = "SELECT AA.ASPTBLVEHTOKENID,AA.ASPTBLVEHTOKENID1,AA.COMPCODE,AA.TOKENNO,AA.TOKENDATE,AA.VEHICLENO, AA.VEHTYPE ,AA.EMPNAME,AA.DEPT,AA.ACTIVE,AA.TOKENCANCEL, AA.ITEMNAME,AA.FUELDATE,AA.FUELRATE2,AA.LITRES,SUM(AA.FUELRATE2 * AA.LITRES) AS TOTAL FROM(  SELECT A.ASPTBLVEHTOKENID, A.ASPTBLVEHTOKENID1, D.COMPCODE, A.TOKENNO, A.TOKENDATE, C.VEHICLENO, G.VEHTYPE, CONCAT(E.fname, concat('-', F.MIDCARD))  AS EMPNAME, I.DISPNAME AS DEPT,A.ACTIVE, A.TOKENCANCEL,H.ITEMNAME, (SELECT  max(AA.FUELDATE) FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME WHERE AB.COMPCODE = D.COMPCODE AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT max(AAA.FUELDATE) FROM ASPTBLFUELRATEMASDET AAA JOIN GTCOMPMAST AAB  ON AAA.COMPCODE = AAB.GTCOMPMASTID JOIN GTGENITEMMAST AAC ON AAC.GTGENITEMMASTID = AAA.ITEMNAME WHERE AAB.COMPCODE = AB.COMPCODE AND AAC.ITEMNAME = AC.ITEMNAME AND AAA.FUELDATE = A.TOKENDATE)) AS FUELDATE,(SELECT  avg(AA.FUELRATE2)FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID  JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME  WHERE AB.COMPCODE = D.COMPCODE AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT max(AAA.FUELDATE) FROM ASPTBLFUELRATEMASDET AAA JOIN GTCOMPMAST AAB  ON AAA.COMPCODE = AAB.GTCOMPMASTID  JOIN GTGENITEMMAST AAC ON AAC.GTGENITEMMASTID = AAA.ITEMNAME  WHERE AAB.COMPCODE = AB.COMPCODE AND AAC.ITEMNAME = AC.ITEMNAME AND AAA.FUELDATE = A.TOKENDATE  )) AS FUELRATE2, B.LITRES   FROM ASPTBLVEHTOKEN A JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID = B.ASPTBLVEHTOKENID  JOIN HRVEHMAST C ON A.VEHICLENO = C.HRVEHMASTID JOIN GTCOMPMAST D ON D.GTCOMPMASTID = A.COMPCODE JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID = A.EMPNAME JOIN hremploydetails F ON F.HREMPLOYMASTID = E.HREMPLOYMASTID  join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid = C.VEHICLETYPE JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID = B.ITEMNAME   JOIN GTDEPTDESGMAST I ON I.GTDEPTDESGMASTID=F.DEPTNAME  WHERE D.COMPCODE = '" + combocompcode.Text + "' AND A.TOKENNO LIKE'%" + txtsearch.Text + "%' OR A.TOKENDATE LIKE'%" + txtsearch.Text + "%' OR C.VEHICLENO LIKE'%" + txtsearch.Text + "%')AA GROUP BY AA.ASPTBLVEHTOKENID,AA.ASPTBLVEHTOKENID1,AA.COMPCODE,AA.TOKENNO,AA.TOKENDATE,AA.VEHICLENO, AA.VEHTYPE ,AA.EMPNAME,AA.DEPT,AA.ACTIVE,AA.TOKENCANCEL ,AA.ITEMNAME,AA.FUELDATE,AA.FUELRATE2,AA.LITRES order by 1,12";
            //        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
            //        DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
            //        rd.SetDataSource(dt2);
            //        crystalReportViewer1.ReportSource = null;
            //        crystalReportViewer1.ReportSource = rd;

            //        crystalReportViewer1.Refresh(); txtsearch.Focus();
            //    }
            //    else
            //    {
            //        string sel2 = "SELECT AA.ASPTBLVEHTOKENID,AA.ASPTBLVEHTOKENID1,AA.COMPCODE,AA.TOKENNO,AA.TOKENDATE,AA.VEHICLENO, AA.VEHTYPE ,AA.EMPNAME,AA.DEPT,AA.ACTIVE,AA.TOKENCANCEL, AA.ITEMNAME,AA.FUELDATE,AA.FUELRATE2,AA.LITRES,SUM(AA.FUELRATE2 * AA.LITRES) AS TOTAL FROM(  SELECT A.ASPTBLVEHTOKENID, A.ASPTBLVEHTOKENID1, D.COMPCODE, A.TOKENNO, A.TOKENDATE, C.VEHICLENO, G.VEHTYPE, CONCAT(E.fname, concat('-', F.MIDCARD))  AS EMPNAME, I.DISPNAME AS DEPT,A.ACTIVE, A.TOKENCANCEL,H.ITEMNAME, (SELECT  max(AA.FUELDATE) FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME WHERE AB.COMPCODE = D.COMPCODE AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT max(AAA.FUELDATE) FROM ASPTBLFUELRATEMASDET AAA JOIN GTCOMPMAST AAB  ON AAA.COMPCODE = AAB.GTCOMPMASTID JOIN GTGENITEMMAST AAC ON AAC.GTGENITEMMASTID = AAA.ITEMNAME WHERE AAB.COMPCODE = AB.COMPCODE AND AAC.ITEMNAME = AC.ITEMNAME AND AAA.FUELDATE = A.TOKENDATE)) AS FUELDATE,(SELECT  avg(AA.FUELRATE2)FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID  JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME  WHERE AB.COMPCODE = D.COMPCODE AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT max(AAA.FUELDATE) FROM ASPTBLFUELRATEMASDET AAA JOIN GTCOMPMAST AAB  ON AAA.COMPCODE = AAB.GTCOMPMASTID  JOIN GTGENITEMMAST AAC ON AAC.GTGENITEMMASTID = AAA.ITEMNAME  WHERE AAB.COMPCODE = AB.COMPCODE AND AAC.ITEMNAME = AC.ITEMNAME AND AAA.FUELDATE = A.TOKENDATE  )) AS FUELRATE2, B.LITRES   FROM ASPTBLVEHTOKEN A JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID = B.ASPTBLVEHTOKENID  JOIN HRVEHMAST C ON A.VEHICLENO = C.HRVEHMASTID JOIN GTCOMPMAST D ON D.GTCOMPMASTID = A.COMPCODE JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID = A.EMPNAME JOIN hremploydetails F ON F.HREMPLOYMASTID = E.HREMPLOYMASTID  join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid = C.VEHICLETYPE JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID = B.ITEMNAME   JOIN GTDEPTDESGMAST I ON I.GTDEPTDESGMASTID=F.DEPTNAME  WHERE D.COMPCODE = '" + combocompcode.Text + "')AA GROUP BY AA.ASPTBLVEHTOKENID,AA.ASPTBLVEHTOKENID1,AA.COMPCODE,AA.TOKENNO,AA.TOKENDATE,AA.VEHICLENO, AA.VEHTYPE ,AA.EMPNAME,AA.DEPT,AA.ACTIVE,AA.TOKENCANCEL ,AA.ITEMNAME,AA.FUELDATE,AA.FUELRATE2,AA.LITRES order by 1,12";
            //        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
            //        DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
            //        rd.SetDataSource(dt2);
            //        crystalReportViewer1.ReportSource = null;
            //        crystalReportViewer1.ReportSource = rd;

            //        crystalReportViewer1.Refresh(); txtsearch.Focus();
            //    }

            //}
            //catch (Exception EX)
            //{ MessageBox.Show(EX.Message); }
        }

        private void butView_Click(object sender, EventArgs e)
        {

            try
            {
                if (combocompcode.Text == "TOKEN CANCEL")
                {
                    string sel2 = "SELECT    XY.COMPCODE,XY.ITEMNAME,XY.LITRES ,ROUND(XY.FUELRATE2,3) FUELRATE2 , TOTAL FROM (SELECT    X.COMPCODE,X.ITEMNAME,SUM(X.LITRES) AS LITRES,AVG(X.FUELRATE2) AS FUELRATE2,SUM(X.LITRES*X.FUELRATE2) TOTAL FROM  (SELECT A.ASPTBLVEHTOKENID, D.COMPCODE,A.TOKENNO,  A.TOKENDATE,A.VEHICLENO, G.VEHTYPE ,E.FNAME, F.DEPT,  H.ITEMNAME,  CASE B.LITRES  WHEN  'FULL' THEN '0'    ELSE B.LITRES END as LITRES, ( SELECT AVG(AA.FUELRATE2) AS FUELRATE2 FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME WHERE AB.COMPCODE = D.COMPCODE AND AC.GTGENITEMMASTID = AA.ITEMNAME AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT MAX(ZAA.FUELDATE) FROM ASPTBLFUELRATEMASDET ZAA JOIN GTCOMPMAST ZAB  ON ZAA.COMPCODE = ZAB.GTCOMPMASTID JOIN GTGENITEMMAST ZAC ON ZAC.GTGENITEMMASTID = ZAA.ITEMNAME WHERE ZAB.COMPCODE = AB.COMPCODE AND ZAC.ITEMNAME = AC.ITEMNAME AND ZAA.FUELDATE <= A.TOKENDATE  )  ) AS FUELRATE2        FROM  ASPTBLVEHTOKEN A     JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID  AND A.TOKENCANCEL='T'   JOIN GTCOMPMAST D ON D.GTCOMPMASTID=A.COMPCODE  JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME      JOIN hremploydetails F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID        join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid=A.VEHICLETYPE   JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID=B.ITEMNAME   WHERE  D.COMPCODE = '" + Class.Users.HCompcode + "'  AND A.TOKENDATE between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy'))X GROUP BY  X.COMPCODE,X.ITEMNAME   )XY   WHERE XY.TOTAL > 0 ORDER BY 2";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLFUELRATEMASDET");
                    DataTable dt2 = ds2.Tables["ASPTBLFUELRATEMASDET"];
                    crystalReportViewer1.ReportSource = null;
                    rd.SetDataSource(dt2);                  
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                }
                else
                {
                   string sel2 = "SELECT    XY.COMPCODE,XY.ITEMNAME,XY.LITRES ,ROUND(XY.FUELRATE2,3) FUELRATE2 , TOTAL,XY.FROMDATE,XY.TODATE  FROM (SELECT    X.COMPCODE,X.ITEMNAME,SUM(X.LITRES) AS LITRES,AVG(X.FUELRATE2) AS FUELRATE2,SUM(X.LITRES*X.FUELRATE2) TOTAL,X.FROMDATE,X.TODATE   FROM  (SELECT A.ASPTBLVEHTOKENID, D.COMPCODE,A.TOKENNO,TO_CHAR(TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')) AS FROMDATE, to_char(TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')) AS TODATE,A.VEHICLENO, G.VEHTYPE ,E.FNAME, F.DEPT,     H.ITEMNAME,  CASE B.LITRES  WHEN  'FULL' THEN '0'    ELSE B.LITRES END as LITRES, ( SELECT AVG(AA.FUELRATE2) AS FUELRATE2     FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME     WHERE AB.COMPCODE = D.COMPCODE AND AC.GTGENITEMMASTID = AA.ITEMNAME AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT MAX(ZAA.FUELDATE) FROM ASPTBLFUELRATEMASDET ZAA     JOIN GTCOMPMAST ZAB  ON ZAA.COMPCODE = ZAB.GTCOMPMASTID JOIN GTGENITEMMAST ZAC ON ZAC.GTGENITEMMASTID = ZAA.ITEMNAME     WHERE ZAB.COMPCODE = AB.COMPCODE AND ZAC.ITEMNAME = AC.ITEMNAME AND ZAA.FUELDATE <= A.TOKENDATE  )  ) AS FUELRATE2        FROM  ASPTBLVEHTOKEN A     JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID  AND A.TOKENCANCEL='F'   JOIN GTCOMPMAST D ON D.GTCOMPMASTID=A.COMPCODE  JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME      JOIN hremploydetails F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID        join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid=A.VEHICLETYPE   JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID=B.ITEMNAME    WHERE  D.COMPCODE = '" + Class.Users.HCompcode + "'  AND A.TOKENDATE between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy'))X GROUP BY  X.COMPCODE,X.ITEMNAME,X.FROMDATE,X.TODATE   )XY   WHERE XY.TOTAL > 0 ORDER BY 2";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLFUELRATEMASDET");
                    DataTable dt2 = ds2.Tables["ASPTBLFUELRATEMASDET"];
                    crystalReportViewer1.ReportSource = null;
                    rd.SetDataSource(dt2);
                  
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh();
                }
            }
            catch (Exception EX)
            { MessageBox.Show(EX.Message); }
           
        }

        public void DownLoads()
        {
            if (comboformate.Text != "")
            {

                DialogResult result = MessageBox.Show("Do you want to '" + comboformate.Text + "' Formate ??", "" + comboformate.Text + "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    // ExportFormatType formatType = ExportFormatType.NoFormat;                    
                    switch (comboformate.Text)
                    {
                        case "Word":
                            rd.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + combocompcode.Text + "'AverageReport.doc");
                            break;

                        case "Excel":
                            // formatType = ExportFormatType.Excel;
                            rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + combocompcode.Text + "'AverageReport.xls");
                            break;

                        case "PDF":
                            rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + combocompcode.Text + "'AverageReport.pdf");
                            break;

                        case "CSV":
                            rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "d:\\'" + combocompcode.Text + "'AverageReport.csv");
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

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

     

        private void ChangePasswords_Click(object sender, EventArgs e)
        {

        }

        private void Deletes_Click(object sender, EventArgs e)
        {

        }

     
        public void Prints()
        {
            try
            {
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {

                    try
                    {
                        if (combocompcode.Text == "TOKEN CANEL")
                        {
                            string sel2 = "SELECT    XY.COMPCODE,XY.ITEMNAME,XY.LITRES ,ROUND(XY.FUELRATE2,3) FUELRATE2 , TOTAL FROM (SELECT    X.COMPCODE,X.ITEMNAME,SUM(X.LITRES) AS LITRES,AVG(X.FUELRATE2) AS FUELRATE2,SUM(X.LITRES*X.FUELRATE2) TOTAL FROM  (SELECT A.ASPTBLVEHTOKENID, D.COMPCODE,A.TOKENNO,  A.TOKENDATE,A.VEHICLENO, G.VEHTYPE ,E.FNAME, F.DEPT,  H.ITEMNAME,  CASE B.LITRES  WHEN  'FULL' THEN '0'    ELSE B.LITRES END as LITRES, ( SELECT AVG(AA.FUELRATE2) AS FUELRATE2 FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME WHERE AB.COMPCODE = D.COMPCODE AND AC.GTGENITEMMASTID = AA.ITEMNAME AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT MAX(ZAA.FUELDATE) FROM ASPTBLFUELRATEMASDET ZAA JOIN GTCOMPMAST ZAB  ON ZAA.COMPCODE = ZAB.GTCOMPMASTID JOIN GTGENITEMMAST ZAC ON ZAC.GTGENITEMMASTID = ZAA.ITEMNAME WHERE ZAB.COMPCODE = AB.COMPCODE AND ZAC.ITEMNAME = AC.ITEMNAME AND ZAA.FUELDATE <= A.TOKENDATE  )  ) AS FUELRATE2        FROM  ASPTBLVEHTOKEN A     JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID  AND A.TOKENCANCEL='T'   JOIN GTCOMPMAST D ON D.GTCOMPMASTID=A.COMPCODE  JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME      JOIN hremploydetails F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID        join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid=A.VEHICLETYPE   JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID=B.ITEMNAME   WHERE  D.COMPCODE = '" + Class.Users.HCompcode + "'  AND A.TOKENDATE between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy'))X GROUP BY  X.COMPCODE,X.ITEMNAME   )XY  ORDER BY 2";
                            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLFUELRATEMASDET");
                            DataTable dt2 = ds2.Tables["ASPTBLFUELRATEMASDET"];
                            CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                            reportdocument.Load(Application.StartupPath + "\\Report\\FuelAverage.rpt");
                            reportdocument.SetDataSource(dt2);
                            reportdocument.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                            reportdocument.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

                        }
                        else
                        {
                            string sel2 = "SELECT    XY.COMPCODE,XY.ITEMNAME,XY.LITRES ,ROUND(XY.FUELRATE2,3) FUELRATE2 , TOTAL FROM (SELECT    X.COMPCODE,X.ITEMNAME,SUM(X.LITRES) AS LITRES,AVG(X.FUELRATE2) AS FUELRATE2,SUM(X.LITRES*X.FUELRATE2) TOTAL FROM  (SELECT A.ASPTBLVEHTOKENID, D.COMPCODE,A.TOKENNO,  A.TOKENDATE,A.VEHICLENO, G.VEHTYPE ,E.FNAME, F.DEPT,  H.ITEMNAME,  CASE B.LITRES  WHEN  'FULL' THEN '0'    ELSE B.LITRES END as LITRES, ( SELECT AVG(AA.FUELRATE2) AS FUELRATE2 FROM ASPTBLFUELRATEMASDET AA JOIN GTCOMPMAST AB  ON AA.COMPCODE = AB.GTCOMPMASTID JOIN GTGENITEMMAST AC ON AC.GTGENITEMMASTID = AA.ITEMNAME WHERE AB.COMPCODE = D.COMPCODE AND AC.GTGENITEMMASTID = AA.ITEMNAME AND AC.ITEMNAME = H.ITEMNAME AND AA.FUELDATE = (SELECT MAX(ZAA.FUELDATE) FROM ASPTBLFUELRATEMASDET ZAA JOIN GTCOMPMAST ZAB  ON ZAA.COMPCODE = ZAB.GTCOMPMASTID JOIN GTGENITEMMAST ZAC ON ZAC.GTGENITEMMASTID = ZAA.ITEMNAME WHERE ZAB.COMPCODE = AB.COMPCODE AND ZAC.ITEMNAME = AC.ITEMNAME AND ZAA.FUELDATE <= A.TOKENDATE  )  ) AS FUELRATE2        FROM  ASPTBLVEHTOKEN A     JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID  AND A.TOKENCANCEL='F'   JOIN GTCOMPMAST D ON D.GTCOMPMASTID=A.COMPCODE  JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME      JOIN hremploydetails F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID        join HRVEHTYPEMAST G on G.HRVEHTYPEMASTid=A.VEHICLETYPE   JOIN GTGENITEMMAST H ON H.GTGENITEMMASTID=B.ITEMNAME   WHERE  D.COMPCODE = '" + Class.Users.HCompcode + "'  AND A.TOKENDATE between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "','dd-MM-yyyy'))X GROUP BY  X.COMPCODE,X.ITEMNAME   )XY  ORDER BY 2";
                            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLFUELRATEMASDET");
                            DataTable dt2 = ds2.Tables["ASPTBLFUELRATEMASDET"];
                            CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                            reportdocument.Load(Application.StartupPath + "\\Report\\FuelAverage.rpt");
                            reportdocument.SetDataSource(dt2);
                            reportdocument.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                            reportdocument.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

                        }
                    }
                    catch (Exception EX)
                    { MessageBox.Show(EX.Message); }



                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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

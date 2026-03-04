using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.Shared;
namespace Pinnacle.ReportFormate
{
    public partial class VehicleTableMatch : Form,ToolStripAccess
    {
        public VehicleTableMatch()
        {
            InitializeComponent();
          
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }

      

        private static VehicleTableMatch _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        byte[] bytes; private int rowIndex = 0;
        public static VehicleTableMatch Instance
        {
            get { if (_instance == null) _instance = new VehicleTableMatch(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }

        Report.VehicleTableMatch rd = new Report.VehicleTableMatch();
  

   

        private void butView_Click(object sender, EventArgs e)
        {
            try
            {
                if (comboformate.Text == "AUVIT")
                {
                    string sel2 = "SELECT  TO_CHAR(A.ASPTBLVEHTOKENID) AS ASPTBLVEHTOKENID,B.COMPCODE,A.VEHICLENO AS VEHICLENOS,TO_CHAR(C.HRVEHMASTID) AS VEHICLEMASID,  C.VEHICLENO,A.TOKENDATE FROM ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN HRVEHMAST C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE AND C.HRVEHMASTID=A.VEHICLENO  WHERE  B.COMPCODE='FLF'  ORDER BY 1";

                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLVEHTOKEN");
                    DataTable dt2 = ds2.Tables["ASPTBLVEHTOKEN"];

                    rd.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;

                    crystalReportViewer1.Refresh();
                }
                else
                {
                    string sel2 = "SELECT  TO_CHAR(A.ASPTBLVEHTOKENID)AS ASPTBLVEHTOKENID,B.COMPCODE,A.VEHICLENO AS VEHICLENOS,TO_CHAR(C.ASPTBLVEHMASID) AS VEHICLEMASID,  C.VEHICLENO,A.TOKENDATE FROM ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN ASPTBLVEHMAS C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE AND C.ASPTBLVEHMASID=A.VEHICLENO  WHERE  B.COMPCODE='FLF'  ORDER BY 1";

                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLVEHTOKEN");
                    DataTable dt2 = ds2.Tables["ASPTBLVEHTOKEN"];

                    rd.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
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


                        case "AUVIT":
                            // formatType = ExportFormatType.Excel;
                            rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + combocompcode.Text + "'AuvitTable.xls");
                            break;

                        case "DOTNET":
                            rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + combocompcode.Text + "'DotNetTable.xls");
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

        private void comboformate_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void Prints_Click(object sender, EventArgs e)
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

        public void Exit()
        {
            this.Hide();
            crystalReportViewer1.ReportSource = null;
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
        }

        public void GridLoad()
        {
           
        }

        public void ReadOnlys()
        {
            
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CrystalDecisions.Shared;
namespace Pinnacle.Report.CFM
{
    public partial class RawMaterialReport : Form,ToolStripAccess
    {
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        public RawMaterialReport()
        {
            InitializeComponent();
           
        }
       
        private static RawMaterialReport _instance;
        public static RawMaterialReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new RawMaterialReport();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }

        string ss = System.DateTime.Now.ToShortDateString();
       
        private void RawMaterialReport_Load(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(txtsearch, "Search...");
        }

     
        private void butsumbit_Click(object sender, EventArgs e)
        {
         
            try
            {
                Cursor = Cursors.WaitCursor;
                ReportFormate.CFM.RawMaterialCrystalReport rd = new ReportFormate.CFM.RawMaterialCrystalReport();

                if (txtsearch.Text != "")
                {


                    string sel2 = "select a.asptblrawmaterialid,a.asptblrawmaterialid1,a.finyear,b.compcode,b.compname,a.certificateno, a.vechileno,date_format(a.datetime1,'%d-%m-%Y') as datetime1,c.partyname as  receivedFrom,a.itemname,d.itemname as itemnamevarity,a.grossweight,a.tareweight,a.netweight, a.noofbag,e.godownname,a.thirdpartyweight,a.tripwagonno,a.lotno,a.sampledby,a.certifiedby,a.visualstatus,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid   where b.compcode='" + combocompcode.Text + "'  and a.certificateno like'%" + txtsearch.Text + "%' or a.vechileNo like'%" + txtsearch.Text + "%';";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblrawmaterial");
                    DataTable dt2 = ds2.Tables["asptblrawmaterial"];
                    if (dt2.Rows.Count <= 0)
                    {
                        MessageBox.Show("Invalid Vehicle NO");txtsearch.Focus();
                    }
                    rd.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd; 
                    crystalReportViewer1.Refresh(); txtsearch.Select();
                }
                else
                {
                    string sel2 = "select a.asptblrawmaterialid,a.asptblrawmaterialid1,a.finyear,b.compcode,b.compname,a.certificateno, a.vechileno, date_format(a.datetime1,'%d-%m-%Y') as datetime1,c.partyname as  receivedFrom,a.itemname,d.itemname as itemnamevarity,a.grossweight,a.tareweight,a.netweight, a.noofbag,e.godownname,a.thirdpartyweight,a.tripwagonno,a.lotno,a.sampledby,a.certifiedby,a.visualstatus,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid where b.compcode='" + combocompcode.Text + "' and  a.datetime1 between date_format('" + frmdate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') and date_format('" + todate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d');";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblrawmaterial");
                    DataTable dt2 = ds2.Tables["asptblrawmaterial"];
                    rd.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                }
            }
            catch (Exception ex) { }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            News();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);


        }

        public void News()
        {
            crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();

            frmdate.Value = System.DateTime.Now.AddDays(0);
            todate.Value = System.DateTime.Now.AddDays(0);
            DataTable dt3 = mas.comcode();
            if (dt3.Rows.Count > 0)
            {
                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt3;
            }
            frmdate.Value = DateTime.Now.AddDays(0); todate.Value = DateTime.Now.AddDays(0);
            txtsearch.Select();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
           
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            
        }

        //private void combocustomer_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    string sel2 = "select a.asptblrawmaterialid,a.asptblrawmaterialid1,f.finyear,b.compcode,b.compname,a.certificateno, a.vechileno,date_format(a.datetime1,'%d-%m-%Y') as datetime1,c.partyname as  receivedFrom,a.itemname,d.itemname as itemnamevarity,a.grossweight,a.tareweight,a.netweight, a.noofbag,e.godownname,a.thirdpartyweight,a.tripwagonno,a.lotno,a.sampledby,a.certifiedby,a.visualstatus,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid join gtfinancialyear f on a.finyear=f.gtfinancialyearid where b.compcode='" + combocompcode.Text + "';";
        //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblrawmaterial");
        //    DataTable dt2 = ds2.Tables["asptblrawmaterial"];
        //    rd.SetDataSource(dt2);
        //    crystalReportViewer1.ReportSource = null;
        //    crystalReportViewer1.ReportSource = rd;
        //    crystalReportViewer1.Refresh();
        //}

        private void comboformate_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void DownLoads()
        {
            //if (comboformate.Text != "")
            //{
            //    ReportFormate.CFM.RawMaterialCrystalReport rd = new ReportFormate.CFM.RawMaterialCrystalReport();

            //    DialogResult result = MessageBox.Show("Do you want to '" + comboformate.Text + "' Formate ??", "" + comboformate.Text + "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //    if (result.Equals(DialogResult.OK))
            //    {
            //        // ExportFormatType formatType = ExportFormatType.NoFormat;                    
            //        switch (comboformate.Text)
            //        {
            //            case "Word":
            //                rd.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + combocompcode.Text + "'RawMaterialReport.doc");
            //                break;

            //            case "Excel":
            //                // formatType = ExportFormatType.Excel;
            //                rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + combocompcode.Text + "'RawMaterialReport.xls");
            //                break;

            //            case "PDF":
            //                rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + combocompcode.Text + "'RawMaterialReport.pdf");
            //                break;

            //            case "CSV":
            //                rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "d:\\'" + combocompcode.Text + "'RawMaterialReport.csv");
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

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        public void Prints()
        {
            try
            {
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    string sel2 = "select a.asptblrawmaterialid,a.asptblrawmaterialid1,a.finyear,b.compcode,b.compname,a.certificateno, a.vechileno, date_format(a.datetime1,'%d-%m-%Y') as datetime1,c.partyname as  receivedFrom,a.itemname,d.itemname as itemnamevarity,a.grossweight,a.tareweight,a.netweight, a.noofbag,e.godownname,a.thirdpartyweight,a.tripwagonno,a.lotno,a.sampledby,a.certifiedby,a.visualstatus,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid  where b.compcode='" + combocompcode.Text + "' and  a.datetime1 between date_format('" + frmdate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') and date_format('" + todate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d');";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblrawmaterial");
                    DataTable dt2 = ds2.Tables["asptblrawmaterial"];
                    CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                    reportdocument.Load(Application.StartupPath + "\\ReportFormate\\RawMaterialCrystalReport.rpt");
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

        public void Saves()
        {
            throw new NotImplementedException();
        }

        public void Searchs()
        {
            throw new NotImplementedException();
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void Deletes()
        {
            throw new NotImplementedException();
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        public void Imports()
        {
            throw new NotImplementedException();
        }

        public void Pdfs()
        {
            throw new NotImplementedException();
        }

        public void ChangePasswords()
        {
            throw new NotImplementedException();
        }

        public void ChangeSkins()
        {
            throw new NotImplementedException();
        }

        public void Logins()
        {
            throw new NotImplementedException();
        }

        public void GlobalSearchs()
        {
            throw new NotImplementedException();
        }

        public void TreeButtons()
        {
            throw new NotImplementedException();
        }



        public void GridLoad()
        {
            throw new NotImplementedException();
        }
    }
}

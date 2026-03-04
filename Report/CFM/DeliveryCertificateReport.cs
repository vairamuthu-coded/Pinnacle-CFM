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
    public partial class DeliveryCertificateReport : Form,ToolStripAccess
    {

        private static DeliveryCertificateReport _instance;
        public static DeliveryCertificateReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DeliveryCertificateReport();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        DataTable dtgeneral = new DataTable();
        DataTable reversedDt; DataTable reversedDt1;
        DataTable reversedDt2; int i, cnt = 0; string idcardcount = "";
        public DeliveryCertificateReport()
        {
            InitializeComponent();
        }

        

        string ss = System.DateTime.Now.ToShortDateString();

        private void DeliveryCertificateReport_Load(object sender, EventArgs e)
        {

            toolTip1.SetToolTip(txtsearch,"Search...");

        }
        private void butsumbit_Click(object sender, EventArgs e)
        {

            try
            {

                ReportFormate.CFM.Delivery rd1 = new ReportFormate.CFM.Delivery();
                DataTable dtcom = new DataTable();
                Cursor = Cursors.WaitCursor;
                // ---------- Create schema ----------
                string sel0 = @" SELECT  '' as asptbldeliveryid, '' as ascreatedon, '' as modifiedon,   '' as sendto,'' as certifiedby,'' as createdby, '' as compcode, '' as compname, '' as address,  '' as finyear, '' as certificateno, '' as vechileNo, '' as  deliverydate, '' as firstweight, '' as  secondweight, '' as  netweight, '' as weightdeffer,'' as productname, '' as productweight,'' as productweight1, 0 as productkgs FROM asptbldelivery a JOIN gtcompmast c ON a.compcode=c.gtcompmastid    WHERE a.asptbldeliveryid <= 0";
                DataTable schema = Utility.ExecuteSelectQuery(sel0, "asptbldelivery").Tables[0];
                foreach (DataColumn col in schema.Columns)
                    dtcom.Columns.Add(col.ColumnName);
                DataRow r; string sel = ""; DataTable dt;
                 sel = "select a.compcode,a.compname,a.address,a.companylogo from gtcompmast a where a.active='T'  AND A.COMPCODE='" + combocompcode.Text + "'";
                DataTable schema0 = Utility.ExecuteSelectQuery(sel, "gtcompmast").Tables[0];
                if (txtsearch.Text != "")
                {
                    sel = "";
                    sel = @"SELECT a.asptbldeliveryid, a.createdon,  date_format(a.deliverydate, '%d-%m-%Y') AS modifiedon,  a.sendto,a.certifiedby,  a.createdby, c.compcode, c.compname, c.address,  a.finyear, a.certificateno, a.vechileNo,  a.deliverydate, a.firstweight, a.secondweight,  a.netweight, a.weightdeffer,'' AS productname,'' as productweight,  a.productweight as productweight1,'' AS productkgs  FROM asptbldelivery a        JOIN gtcompmast c ON a.compcode=c.gtcompmastid        WHERE  c.compcode = '" + combocompcode.Text + "'  and a.vechileNo= '" + txtsearch.Text.Trim() + "'     ORDER BY a.asptbldeliveryid";
                    dt = Utility.ExecuteSelectQuery(sel, "asptbldelivery").Tables[0];
                    if (dt.Rows.Count <= 0)
                    {
                        MessageBox.Show("Invalid Vehicle No");  txtsearch.Focus();
                        return;
                    }
                }
                else
                {
                    sel = "";
                    sel = @"SELECT a.asptbldeliveryid, a.createdon,  date_format(a.deliverydate, '%d-%m-%Y') AS modifiedon, a.sendto,a.certifiedby, a.createdby, c.compcode, c.compname, c.address,  a.finyear, a.certificateno, a.vechileNo,  a.deliverydate, a.firstweight, a.secondweight,  a.netweight,  a.weightdeffer,'' AS productname,'' as productweight,  a.productweight as productweight1  ,'' AS productkgs  FROM asptbldelivery a        JOIN gtcompmast c ON a.compcode=c.gtcompmastid        WHERE  c.compcode = '" + combocompcode.Text + "' and a.deliverydate between date_format('" + Convert.ToDateTime(frmdate.Value.ToString().Substring(0, 10)).ToString("yyyy-MM-dd") + "', '%Y-%m-%d')  and date_format('" + Convert.ToDateTime(todate.Value.ToString().Substring(0, 10)).ToString("yyyy-MM-dd") + "', '%Y-%m-%d')     ORDER BY a.asptbldeliveryid";
                    dt = Utility.ExecuteSelectQuery(sel, "asptbldelivery").Tables[0];
                }
                // ---------- Merge products ----------
                foreach (DataRow master in dt.Rows)
                {
                    sel = "";
                     sel = @"SELECT a.asptbldeliveryid, '' as createdon, '' as modifiedon,   '' as sendto,'' as certifiedby, '' as createdby, '' as compcode, '' as compname, '' as address,  '' as finyear, '' as certificateno, '' as vechileNo, '' as deliverydate, '' as firstweight, '' as secondweight, '' as netweight, '' as certifiedby, '' as weightdeffer,  b.productname1 as productname,a.productweight,'' as productweight1,a.productweight, a.productkgs FROM asptbldeliverydet a JOIN asptblproductweightmas b ON b.asptblproductweightmasid=a.productname  join asptbldelivery c on c.asptbldeliveryid=a.asptbldeliveryid WHERE a.asptbldeliveryid='" + master["asptbldeliveryid"] + "'";
                    DataTable det = Utility.ExecuteSelectQuery(sel, "asptbldeliverydet").Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        r = dtcom.NewRow();
                        r.ItemArray = master.ItemArray;
                        dtcom.Rows.Add(r);
                    }
                    if (det.Rows.Count > 0)
                    {
                      
                        foreach (DataRow prod in det.Rows)
                        {                           
                            
                                DataRow rr = dtcom.NewRow();

                                rr.ItemArray = master.ItemArray;
                                rr["productname"] = prod["productname"];
                                rr["productweight"] = prod["productweight"];
                            rr["productweight1"] = prod["productweight1"];
                            rr["weightdeffer"] = prod["weightdeffer"];
                            rr["productkgs"] = prod["productkgs"];
                                dtcom.Rows.Add(rr);
                            
                           
                        }
                    }

                }
                if (dtcom.Rows.Count > 0)
                {
                    crystalReportViewer1.ReportSource = null;
                    rd1.Database.Tables["DataTable5"].SetDataSource(schema0);
                    rd1.Database.Tables["DataTable1"].SetDataSource(dtcom);

                    crystalReportViewer1.ReportSource = rd1;
                    crystalReportViewer1.Refresh();
                    Cursor = Cursors.Default;
                }

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show(ex.Message, "Delivery Report Error");
            }
            finally
            {
                Cursor = Cursors.Default;
            }



        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        public void News()
        {
     

            DataTable dt3 = mas.comcode();
            if (dt3.Rows.Count > 0)
            {
                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt3;
            }
            frmdate.Value = DateTime.Now.AddDays(0); todate.Value = DateTime.Now.AddDays(0);

        }

        //private void combocustomer_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    string sel2 = "select  c.compcode, c.compname ,a.sendto,a.certificateno,a.vechileNo,   date_format(a.deliverydate,'%d-%m-%Y') as deliverydate, a.deliverytime,    a.firstweight,    a.secondweight,    a.netweight, a.productweight   from asptbldelivery a  join gtfinancialyear b on a.finyear=b.gtfinancialyearid  join gtcompmast c on c.gtcompmastid=a.compcode join gtcitymast d on d.gtcitymastid=c.city  where c.compcode='" + combocompcode.Text + "';";
        //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldelivery");
        //    DataTable dt2 = ds2.Tables["asptbldelivery"];
        //    rd.SetDataSource(dt2);
        //    crystalReportViewer1.ReportSource = null;
        //    crystalReportViewer1.ReportSource = rd;
        //    crystalReportViewer1.Refresh(); txtsearch.Select();
        //}

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            //if (txtsearch.Text != "")
            //{
            //    ReportFormate.CFM.DeliveryChallanCrystalReport rd = new ReportFormate.CFM.DeliveryChallanCrystalReport();

            //    string sel2 = "select  c.compcode, c.compname ,a.sendto,a.certificateno,a.vechileNo,   date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,  a.deliverytime,    a.firstweight,    a.secondweight,    a.netweight, a.productweight   from asptbldelivery a  join gtcompmast c on c.gtcompmastid=a.compcode join gtcitymast d on d.gtcitymastid=c.city  where c.compcode='" + combocompcode.Text + "'  and a.certificateno like'%" + txtsearch.Text + "%' or a.vechileNo like'%" + txtsearch.Text + "%';";
            //    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldelivery");
            //    DataTable dt2 = ds2.Tables["asptbldelivery"];
            //    rd.SetDataSource(dt2);
            //    crystalReportViewer1.ReportSource = null;
            //    crystalReportViewer1.ReportSource = rd;
            //    crystalReportViewer1.Refresh(); txtsearch.Select();
            //}
        }

        public void DownLoads()
        {
            //if (comboformate.Text != "")
            //{
            //    ReportFormate.CFM.Delivery rd = new ReportFormate.CFM.Delivery();

            //    DialogResult result = MessageBox.Show("Do you want to '" + comboformate.Text + "' Formate ??", "" + comboformate.Text + "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //    if (result.Equals(DialogResult.OK))
            //    {
            //        // ExportFormatType formatType = ExportFormatType.NoFormat;                    
            //        switch (comboformate.Text)
            //        {
            //            case "Word":
            //                rd.ExportToDisk(ExportFormatType.WordForWindows, "d:\\'" + combocompcode.Text + "'DeliveryCertificateReport.doc");
            //                break;

            //            case "Excel":
            //                // formatType = ExportFormatType.Excel;
            //                rd.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\'" + combocompcode.Text + "'DeliveryCertificateReport.xls");
            //                break;

            //            case "PDF":
            //                rd.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\'" + combocompcode.Text + "'DeliveryCertificateReport.pdf");
            //                break;

            //            case "CSV":
            //                rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "d:\\'" + combocompcode.Text + "'DeliveryCertificateReport.csv");
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

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            //printDocument1.PrinterSettings = printDialog1.PrinterSettings;
            //printDocument1.Print();
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {

        }

        public void Prints()
        {
            try
            {
                ReportFormate.CFM.Delivery rd = new ReportFormate.CFM.Delivery();

                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    
                        string sel2 = "select  c.compcode, c.compname ,a.sendto,a.certificateno,a.vechileNo,   date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,  a.deliverytime,    a.firstweight,    a.secondweight,    a.netweight, a.productweight   from asptbldelivery a   join gtcompmast c on c.gtcompmastid=a.compcode join gtcitymast d on d.gtcitymastid=c.city  where c.compcode='" + combocompcode.Text + "' and  a.deliverydate between date_format('" + frmdate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') and date_format('" + todate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d');";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldelivery");
                        DataTable dt2 = ds2.Tables["asptbldelivery"];
                        CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                        reportdocument.Load(Application.StartupPath + "\\ReportFormate\\Delivery.rpt");
                        reportdocument.SetDataSource(dt2);
                        reportdocument.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                        reportdocument.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

                    
                }
            }
            catch(Exception ex)
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

        private void button1_Click(object sender, EventArgs e)
        {

        }

        public void GridLoad()
        {
            throw new NotImplementedException();
        }

    
    }
}

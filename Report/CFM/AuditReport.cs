using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Report.CFM
{
    public partial class AuditReport : Form,ToolStripAccess
    {
        public AuditReport()
        {
            InitializeComponent();
        }
        string ss = System.DateTime.Now.ToShortDateString();
        private static AuditReport _instance;
        public static AuditReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new AuditReport();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        DataTable dtgeneral = new DataTable();
        DataTable reversedDt; DataTable reversedDt1;
        DataTable reversedDt2; int i, cnt = 0; string idcardcount = "";

        private void button1_Click(object sender, EventArgs e)
        {
           
            if (combocompcode.Text != "" && txtsearch.Text != "")
            {
              
                try
                {
                    Cursor = Cursors.WaitCursor;
                    ReportFormate.CFM.RawMaterialPivotReport rd1 = new ReportFormate.CFM.RawMaterialPivotReport();

                    string sel3 = "select a.asptblrawmaterialid2,a.asptblrawmaterialid,a.asptblrawmaterialid1,a.finyear,b.compcode,b.compname,b.address, a.certificateno, a.vechileno, date_format(a.datetime1,'%d-%m-%Y') as datetime1,c.partyname as  receivedFrom,a.itemname,d.itemname as itemnamevarity,a.grossweight,a.tareweight,a.netweight, a.noofbag,e.godownname,a.thirdpartyweight,a.tripwagonno,a.lotno,a.sampledby,a.certifiedby,a.visualstatus,a.delayreason,a.operation,a.createdon,a.modifiedon,a.remarks,a.createdby as test1 from asptblrawmaterial2 a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid  where b.compcode='" + combocompcode.Text + "' and  a.vechileno='" + txtsearch.Text + "' order by 1;";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblrawmaterial2");
                    DataTable dt3 = ds3.Tables["asptblrawmaterial2"];
                    if (dt3.Rows.Count > 0)
                    {
                        rd1.SetDataSource(dt3);
                        crystalReportViewer1.ReportSource = null;
                        crystalReportViewer1.ReportSource = rd1;
                        crystalReportViewer1.Refresh();
                    }
                    else
                    {
                        crystalReportViewer1.ReportSource = null; Cursor = Cursors.Default;
                    }
                    
                }
                catch (Exception ex) { Cursor = Cursors.Default; }
                finally { Cursor = Cursors.Default; }

                try
                {
                    Cursor = Cursors.WaitCursor;
                    string sel = "select a.asptbldeliveryid, a.operation, a.createdon,a.modifiedon,a.createdby,b.compcode, b.compname,b.address, a.finyear,a.certificateno,a.vechileNo,a.deliverydate,a.firstweight,a.secondweight,a.netweight,a.productweight,a.sendto,a.certifiedby,a.weightdeffer from asptbldelivery2 a join gtcompmast b on a.compcode=b.gtcompmastid  where b.compcode='" + combocompcode.Text + "' and   a.vechileno='" + txtsearch.Text + "' order by a.asptbldeliveryid asc;";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptbldelivery2");
                    DataTable dt = ds2.Tables["asptbldelivery2"];
                    if (dt.Rows.Count > 0)
                    {
                        crystalReportViewer2.ReportSource = null;//ReportFormate\CFM\DeliveryChallanPivotReport _Left.rpt D:\Software\Backup-24-07-2021\CFM\Pinnacle-20\ReportFormate\CFM\DeliveryChallanPivotReport.rpt
                        CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                        reportdocument.Load(Application.StartupPath + "\\ReportFormate\\CFM\\DeliveryChallanPivotReport _Left.rpt");
                        reportdocument.Database.Tables["DataTable1"].SetDataSource(dt);
                        crystalReportViewer2.ReportSource = reportdocument;
                        crystalReportViewer2.Refresh();
                    }
                    else
                    {
                        crystalReportViewer2.ReportSource = null; Cursor = Cursors.Default;
                    }
                }
                catch (Exception ex) { Cursor = Cursors.Default; }
                finally { Cursor = Cursors.Default; }
                
            }
            if (combocompcode.Text != "" && txtsearch.Text == "")
            {
                try
                {
                    Cursor = Cursors.WaitCursor;
                    ReportFormate.CFM.RawMaterialPivotReport rd1 = new ReportFormate.CFM.RawMaterialPivotReport();

                    string sel3 = "select a.asptblrawmaterialid2,a.asptblrawmaterialid,a.asptblrawmaterialid1,a.finyear,c.compcode,c.compname,c.address, a.certificateno, a.vechileno, a.datetime1,d.partyname as  receivedFrom,a.itemname,e.itemname as itemnamevarity,a.grossweight,a.tareweight,a.netweight, a.noofbag,f.godownname,a.thirdpartyweight,a.tripwagonno,a.lotno,a.sampledby,a.certifiedby,a.visualstatus,a.delayreason,a.operation,a.createdon,a.modifiedon,a.remarks,a.createdby as test1 from asptblrawmaterial2 a right outer join asptblrawmaterial2 b on a.asptblrawmaterialid2=b.asptblrawmaterialid2 and a.asptblrawmaterialid=b.asptblrawmaterialid and a.compcode=b.compcode and a.finyear=b.finyear join gtcompmast c on a.compname=c.gtcompmastid  join asptblpartymas d on a.receivedFrom =d.asptblpartymasid join asptblitemmast e on a.itemnamevarity=e.asptblitemmastid join asptblgodwonmas f on a.godownname=f.asptblgodwonmasid   where c.compcode='" + combocompcode.Text + "' and  b.datetime2 between date_format('" + frmdate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') and date_format('" + todate.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by 1;";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblrawmaterial2");
                    DataTable dt3 = ds3.Tables["asptblrawmaterial2"];
                    if (dt3.Rows.Count > 0)
                    {
                        rd1.SetDataSource(dt3);
                        crystalReportViewer1.ReportSource = null;
                        crystalReportViewer1.ReportSource = rd1;
                        crystalReportViewer1.Refresh();
                    }
                    else
                    {
                        crystalReportViewer1.ReportSource = null;
                    }
                }
                catch (Exception ex) { Cursor = Cursors.Default; }
                finally
                {
                    Cursor = Cursors.Default;
                }

                try
                {
                    Cursor = Cursors.WaitCursor;
                    string sel = "select a.asptbldeliveryid2, a.asptbldeliveryid, a.operation ,a.createdon, a.deliverydate as modifiedon,concat(concat(a.sendto , ' -User : ') , a.certifiedby) as sendto, a.createdby,c.compcode, c.compname,c.address, a.finyear,a.certificateno,a.vechileNo,a.deliverydate,a.firstweight,a.secondweight,a.netweight,a.productweight,a.sendto,a.certifiedby,a.productweight-a.netweight as  weightdeffer from asptbldelivery2 a  right outer join asptbldelivery2 b on a.asptbldeliveryid2=b.asptbldeliveryid2 and a.asptbldeliveryid=b.asptbldeliveryid and a.finyear=b.finyear and a.compcode=b.compcode and a.asptbldeliveryid1=b.asptbldeliveryid1 join gtcompmast c on a.compcode=c.gtcompmastid   where c.compcode='" + combocompcode.Text + "' and b.deliverydate1 between date_format('" + Convert.ToDateTime(frmdate.Value.ToString().Substring(0, 10)).ToString("yyyy-MM-dd") + "', '%Y-%m-%d')  and date_format('" + Convert.ToDateTime(todate.Value.ToString().Substring(0, 10)).ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptbldeliveryid2 asc;";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptbldelivery2");
                    DataTable dt = ds2.Tables["asptbldelivery2"];
                    if (dt.Rows.Count > 0)
                    {
                        crystalReportViewer2.ReportSource = null;//ReportFormate\CFM\DeliveryChallanPivotReport _Left.rpt D:\Software\Backup-24-07-2021\CFM\Pinnacle-20\ReportFormate\CFM\DeliveryChallanPivotReport.rpt
                        CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                        reportdocument.Load(Application.StartupPath + "\\ReportFormate\\CFM\\DeliveryChallanPivotReport _Left.rpt");
                        reportdocument.Database.Tables["DataTable1"].SetDataSource(dt);
                        crystalReportViewer2.ReportSource = reportdocument;
                        crystalReportViewer2.Refresh();
                    }
                    else
                    {
                        crystalReportViewer2.ReportSource = null; Cursor = Cursors.Default;
                    }
                }
                catch (Exception ex) { Cursor = Cursors.Default; }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
            Cursor = Cursors.Default;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
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
            Cursor = Cursors.Default;
        }

        public void Saves()
        {
            throw new NotImplementedException();
        }

        public void Prints()
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

        public void DownLoads()
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

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            News();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            
        }

        public void GridLoad()
        {
            throw new NotImplementedException();
        }
    }
}

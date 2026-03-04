using CrystalDecisions.Shared;
using Pinnacle.Master;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static QRCoder.PayloadGenerator;

namespace Pinnacle.ReportFormate.Lyla
{
    public partial class ProductionStatusReport : Form,ToolStripAccess
    {
        private static ProductionStatusReport _instance;
        public static ProductionStatusReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ProductionStatusReport();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public ProductionStatusReport()
        {
            InitializeComponent();
        }

        private void ProductionStatusReport_Load(object sender, EventArgs e)
        {


            
        }
        private void empty()
        {
 combotype.SelectedIndex = 0;comboBox1.SelectedIndex = 0;combotype1.SelectedIndex = 0;
  pictureBox1.Image = null; combopcsmonthly.SelectedIndex = 0;
            //frmdate.Enabled = false;
            //todate.Enabled = false;
            //butdatewise.Enabled = false;
            //combopono.Enabled = false;
            //button1.Enabled = false;
            combotype.SelectedIndex = 0;
            reportwise(); reportwise1();
            this.BackColor = Class.Users.BackColors; crystalReportViewer1.ReportSource = null;
            butheader.BackColor = Class.Users.BackColors;crystalReportViewer2.ReportSource = null;
            this.Font = Class.Users.FontName;//dataGridView1.Rows.Clear(); 
            string sel1 = "SELECT distinct B.GTCOMPMASTID,b.compcode,b.compname,a.finyear FROM  asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + Class.Users.HCompcode + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
            DataTable dt = ds.Tables["asptblpur"];
            //combocompcode.DisplayMember= "compcode";
            //combocompcode.ValueMember= "GTCOMPMASTID";
            //combocompcode.DataSource=dt;
            //combofinyear.DisplayMember = "finyear";
            //combofinyear.ValueMember = "finyear";
            //combofinyear.DataSource = dt;
            string sel2 = "SELECT  0 asptblpurID,  ''pono FROM  dual union all SELECT  0 asptblpurID,  'ALL'pono FROM  dual union all SELECT distinct X.asptblpurID,  X.pono FROM ( SELECT distinct A.asptblpurID,  a.pono FROM  asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + Class.Users.HCompcode+"' ORDER BY A.asptblpurid DESC) X ";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
            DataTable dt2 = ds2.Tables["asptblpur"];
            combopono.DisplayMember = "pono";
            combopono.ValueMember = "pono";
            combopono.DataSource = dt2;Class.Users.UserTime = 0;
            combopopcsrate.DisplayMember = "pono";
            combopopcsrate.ValueMember = "pono";
            combopopcsrate.DataSource = dt2;
            //foreach(DataGridViewColumn co in dataGridView2.Columns)
            //{
            //    co.HeaderText = "";
            //}
        }
        public void GridLoad()
        {

        }
        public void News()
        {
            empty();
            Class.Users.Query = "SELECT distinct a.asptblpurid,  C.Asptblsizmasid, B.Compcode, A.PoNo,D.ColorName, C.SizeName,E.OrderQty,e.ExcessQty,'' Inward,'' Rework, '' Delivery,'' CheckingInward, ''CheckingRework,'' StitchingRework,''CheckingDelivery,''Notes,b.CompName,b.Address,b.companylogo,''BalQty,''inward1,''rework1,''delivery1,''checkin1,''checkcheck1,''checkstitch1,''checkdel1,''balqty1,j.modified as FromDate,''Todate,e.orderno,e.styleref FROM asptblpurdet1 A JOIN gtcompmast B ON A.compcode=B.gtcompmastid JOIN asptblsizmas C ON C.sizename=A.sizename JOIN asptblcolmas D ON D.colorname=A.colorname JOIN asptblpur E ON E.compcode=A.compcode  AND E.pono=A.pono   and e.asptblpurid=a.asptblpurid join asptblbuymas h on h.asptblbuymasid=e.buyer  join asptblstymas i on i.asptblstymasid=e.stylename  join asptblprolot j on j.compcode=b.gtcompmastid and j.pono=a.pono WHERE a.asptblpurdet1id<0;";
            Class.Users.dt = CC.select(Class.Users.Query, "ASPTBLPROLOT");
            Class.Users.dt1 = Class.Users.dt.Clone();
            // Class.Users.TableName = "asptblprolot";
            // Class.Users.TableNameGrid = "asptblprolotdet";
            //GlobalVariables.HideCols = new string[] { "Asptblsizmasid", "Compcode","CompName", "OrderQty", "ExcessQty","Address", "CompanyLogo"};
            // GlobalVariables.WidthCols = new Int32[] {0,0,100, 120, 70, 0, 0, 70, 70, 70, 120, 120, 120, 125, 80,0,0,0,100,100,80,80};
            //CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols);
            //CommonFunctions.AddGridColumn(dataGridView2, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols);

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

        public void Searchs(int EditID)
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
            
        }

        public void ChangePasswords()
        {
            
        }

        public void DownLoads()
        {
            if (comboBox2.Text != "" && tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])
            {
                DialogResult result = MessageBox.Show("Do you want to Formate ??", "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    if (comboBox2.Text == "EXCEL")
                    {
                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage2"])//your specific tabname
                        {

                            rd4.ExportToDisk(ExportFormatType.ExcelRecord, "d:\\Nyla-PoWise.xls");
                        }

                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])//your specific tabname
                        {
                            rd5.ExportToDisk(ExportFormatType.ExcelRecord, "d:\\Nyla-DateWise.xls");


                        }
                    }
                    if (comboBox2.Text == "PDF")
                    {
                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage2"])//your specific tabname
                        {

                            rd4.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\Nyla-PoWiseRate.pdf");
                        }

                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])//your specific tabname
                        {
                            rd5.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\Nyla-DateWiseRate.pdf");


                        }
                    }
                }
            }
            if (comboBox1.Text != "")
            {
                DialogResult result = MessageBox.Show("Do you want to Formate ??", "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    if (comboBox1.Text == "EXCEL")
                    {
                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage2"])//your specific tabname
                        {

                            rd2.ExportToDisk(ExportFormatType.ExcelRecord, "d:\\Nyla-PoWise.xls");
                        }

                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])//your specific tabname
                        {
                            rd3.ExportToDisk(ExportFormatType.ExcelRecord, "d:\\Nyla-DateWise.xls");


                        }
                    }
                    if (comboBox1.Text == "PDF")
                    {
                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage2"])//your specific tabname
                        {

                            rd2.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\Nyla-PoWise.pdf");
                        }

                        if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])//your specific tabname
                        {
                            rd3.ExportToDisk(ExportFormatType.PortableDocFormat, "d:\\Nyla-DateWise.pdf");


                        }
                    }
                }
            }
            else
            {
                
                MessageBox.Show("Pls Select Combo Box Value");
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }
        Models.CommonClass CC = new Models.CommonClass();
        private void button1_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); DataTable dt = new DataTable(); DataTable dt1 = new DataTable();
            if (combopono.Text == "ALL")
            {
                string sel = "select a.podate, a.pono,a.colorname,a.SIZENAME,a.ASPTBLSIZMASID,a.compname,a.address,a.orderqty,a.excessqty,a.orderno,a.styleref, b.inward,ifnull(b.rework,0) rework, b.delivery delivery,ifnull(d.cinward,0)  as checkinginward, ifnull(c.cdefqty,0) as checkingrework, ifnull(c.csdefqty,0) stitchingrework,  ifnull(d.cdelivery,0) as checkingdelivery,  d.cinward-d.cdelivery-ifnull(d.crework,0) balqty, ifnull(d.crework,0) as delivery1,a.compname,a.address from (    select a.asptblpurid,a.pono,c.colorname,d.SIZENAME,d.ASPTBLSIZMASID,b.orderqty,a.podate,a.orderno,a.styleref,e.compname,e.address,a.excessqty from asptblpur a  join asptblpurdet b on a.asptblpurid = b.asptblpurid  join asptblcolmas c on c.asptblcolmasid = b.colorname  join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  join gtcompmast e on e.gtcompmastid = a.compcode  ) a left join ( select aa.asptblpurid, aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework from (  select 0 asptblpurid,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward, case when a.issuetype = 'DELIVERY'   then sum(b.orderqty) else 0 end  delivery, case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename   group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype ) aa group by aa.asptblpurid,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID ) b on a.pono = b.pono and a.colorname = b.colorname and a.sizename = b.sizename left join ( select a.asptblpurid,a.pono,a.colorname,a.sizename,a.ASPTBLSIZMASID, sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty, sum(a.csdefqty) csdefqty from ( select 0 asptblpurid,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID, case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty, case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,  case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.ISSUETYPE,a.remarks ) a group by a.asptblpurid,a.pono,a.colorname,a.sizename, a.ASPTBLSIZMASID) c on a.pono = c.pono and a.colorname = c.colorname and a.sizename = c.sizename left join ( select aa.asptblpurid,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select 0 asptblpurid,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward, case when a.issuetype = 'DELIVERY'  then sum(b.orderqty) else 0 end  delivery, case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a join asptblchkdet b on a.asptblchkid = b.asptblchkid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype ) aa group by aa.asptblpurid,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID ) d on a.pono = d.pono and a.colorname = d.colorname and a.sizename = d.sizename order by 1";

                ds = Utility.ExecuteSelectQuery(sel, "asptblpurdet1");
                dt = ds.Tables["asptblpurdet1"];
            }
            else
            {
                // string sel = "select a.podate, a.pono,a.colorname,a.SIZENAME,a.ASPTBLSIZMASID,a.compname,a.address,a.orderqty,a.excessqty,a.orderno,a.styleref, b.inward,ifnull(b.rework,0) rework, b.delivery delivery,ifnull(d.cinward,0)  as checkinginward, ifnull(c.cdefqty,0) as checkingrework, ifnull(c.csdefqty,0) stitchingrework,  ifnull(d.cdelivery,0) as checkingdelivery,  ifnull(d.cinward,0)-ifnull(d.cdelivery,0)-ifnull(d.crework,0) balqty,ifnull(d.crework,0) as chereworkdelivery delivery1 from (   select a.pono,c.colorname,d.SIZENAME,d.ASPTBLSIZMASID,b.orderqty,a.podate,a.orderno,a.styleref,e.compname,e.address,a.excessqty from asptblpur a  join asptblpurdet b on a.asptblpurid = b.asptblpurid  join asptblcolmas c on c.asptblcolmasid = b.colorname  join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  join gtcompmast e on e.gtcompmastid = a.compcode   where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "' )     ) a   left join (    select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework from (  select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward, case when a.issuetype = 'DELIVERY'   then sum(b.orderqty) else 0 end  delivery, case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "') group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype ) aa group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID ) b on a.pono = b.pono and a.colorname = b.colorname and a.sizename = b.sizename left join ( select a.pono,a.colorname,a.sizename,a.ASPTBLSIZMASID, sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty, sum(a.csdefqty) csdefqty from ( select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID, case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty, case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,  case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "') group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.ISSUETYPE,a.remarks ) a group by a.pono,a.colorname,a.sizename, a.ASPTBLSIZMASID) c on a.pono = c.pono and a.colorname = c.colorname and a.sizename = c.sizename left join ( select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward, case when a.issuetype = 'DELIVERY'  then sum(b.orderqty) else 0 end  delivery, case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a join asptblchkdet b on a.asptblchkid = b.asptblchkid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "' ) group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype ) aa group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID ) d on a.pono = d.pono and a.colorname = d.colorname and a.sizename = d.sizename order by 1,2,3,5";

                string sel = "select a.podate, a.pono,a.colorname,a.SIZENAME,a.ASPTBLSIZMASID,a.compname,a.address,a.orderqty,a.excessqty,a.orderno,a.styleref, b.inward,ifnull(b.rework,0) rework, b.delivery delivery,ifnull(d.cinward,0)  as checkinginward, ifnull(c.cdefqty,0) as checkingrework, ifnull(c.csdefqty,0) stitchingrework,  ifnull(d.cdelivery,0) as checkingdelivery,  ifnull(d.cinward,0)-ifnull(d.cdelivery,0)-ifnull(d.crework,0) balqty,ifnull(d.crework,0) as delivery1 from (   select a.pono,c.colorname,d.SIZENAME,d.ASPTBLSIZMASID,b.orderqty,a.podate,a.orderno,a.styleref,e.compname,e.address,a.excessqty from asptblpur a  join asptblpurdet b on a.asptblpurid = b.asptblpurid  join asptblcolmas c on c.asptblcolmasid = b.colorname  join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  join gtcompmast e on e.gtcompmastid = a.compcode   where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "' )     ) a   left join (    select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework from (  select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward, case when a.issuetype = 'DELIVERY'   then sum(b.orderqty) else 0 end  delivery, case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "') group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype ) aa group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID ) b on a.pono = b.pono and a.colorname = b.colorname and a.sizename = b.sizename left join ( select a.pono,a.colorname,a.sizename,a.ASPTBLSIZMASID, sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty, sum(a.csdefqty) csdefqty from ( select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID, case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty, case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,  case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "') group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.ISSUETYPE,a.remarks ) a group by a.pono,a.colorname,a.sizename, a.ASPTBLSIZMASID) c on a.pono = c.pono and a.colorname = c.colorname and a.sizename = c.sizename left join ( select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward, case when a.issuetype = 'DELIVERY'  then sum(b.orderqty) else 0 end  delivery, case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a join asptblchkdet b on a.asptblchkid = b.asptblchkid join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where (a.pono='" + combopono.Text + "' or 'ALL' ='" + combopono.Text + "' ) group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype ) aa group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID ) d on a.pono = d.pono and a.colorname = d.colorname and a.sizename = d.sizename order by 1,2,3,5";
                ds = Utility.ExecuteSelectQuery(sel, "asptblpurdet1");
                dt = ds.Tables["asptblpurdet1"];
            }
            string sel1 = "select a.companylogo from gtcompmast a where a.compcode='" + Class.Users.HCompcode + "'";
            ds = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
            dt1 = ds.Tables["gtcompmast"];
            rd2.Database.Tables["DataTable2"].SetDataSource(dt);
            rd2.Database.Tables["DataTableImage"].SetDataSource(dt1);
           
            crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd2;
                crystalReportViewer1.Refresh();
          
           
            Cursor.Current = Cursors.Default;
            tabControl2.SelectTab("TabPage2");
           
        }
        Report.Lyla.ProductionStatus rd2 = new Report.Lyla.ProductionStatus();
        Report.Lyla.DateWisePcsRate rd5 = new Report.Lyla.DateWisePcsRate();
        Report.Lyla.PcsRate1 rd4 = new Report.Lyla.PcsRate1();
        Report.Lyla.ProductionStatusDate rd3 = new Report.Lyla.ProductionStatusDate();
        private void BindGrid(DataGridView grid, DataTable dt1)
        {

            int j = 0, i = 0; Int64 inqty=0, bqty=0, total = 0,inqty2=0, inchdel2=0, inqty1 = 0, inrew1 = 0, indel1 = 0, inchin1 = 0, inchch1 = 0, inchst1 = 0, inchdel1 = 0;
            foreach (DataRow row in dt1.Rows)
            {
                
                grid.Rows.Add();
                grid.Rows[i].Cells[0].Value = Convert.ToInt64("0" + row["ASPTBLSIZMASID"].ToString());
                grid.Rows[i].Cells[1].Value = row["Compcode"].ToString();
                grid.Rows[i].Cells[2].Value = row["PoNo"].ToString();
                grid.Rows[i].Cells[3].Value = row["ColorName"].ToString();
                grid.Rows[i].Cells[4].Value = row["SizeName"].ToString();
                grid.Rows[i].Cells[5].Value = row["OrderQty"].ToString();
                grid.Rows[i].Cells[6].Value = row["ExcessQty"].ToString();
                grid.Rows[i].Cells[7].Value = row["Inward"].ToString();
                grid.Rows[i].Cells[8].Value = row["Rework"].ToString();
                grid.Rows[i].Cells[9].Value = row["Delivery"].ToString();
                grid.Rows[i].Cells[10].Value = row["CheckingInward"].ToString();
                grid.Rows[i].Cells[11].Value = row["CheckingRework"].ToString();
                grid.Rows[i].Cells[12].Value = row["StitchingRework"].ToString();
                grid.Rows[i].Cells[13].Value = row["CheckingDelivery"].ToString();
                grid.Rows[i].Cells[14].Value = row["Notes"].ToString();
                inqty2 += Convert.ToInt64("0" + row["Inward"].ToString());
                inqty1 = Convert.ToInt64("0" + row["Inward"].ToString());
                inrew1 += Convert.ToInt64("0" + row["Rework"].ToString());
                indel1 += Convert.ToInt64("0" + row["Delivery"].ToString());
                inchin1 += Convert.ToInt64("0" + row["CheckingInward"].ToString());
                inchch1 += Convert.ToInt64("0" + row["CheckingRework"].ToString());
                inchst1 += Convert.ToInt64("0" + row["StitchingRework"].ToString());
                inchdel1 = Convert.ToInt64("0" + row["CheckingDelivery"].ToString());
                inchdel2 += Convert.ToInt64("0" + row["CheckingDelivery"].ToString());
                total = inqty1 - inchdel1;
               
                grid.Rows[i].Cells[18].Value = total.ToString();
                grid.Rows[i].Cells[19].Value = row["FromDate"].ToString();
                grid.Rows[i].DefaultCellStyle.BackColor = j % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                Class.Users.dt1.Rows[i]["inward1"] = inqty2;
                Class.Users.dt1.Rows[i]["rework1"] = inrew1;
                Class.Users.dt1.Rows[i]["delivery1"] = indel1;
                Class.Users.dt1.Rows[i]["checkin1"] = inchin1;
                Class.Users.dt1.Rows[i]["checkcheck1"] = inchch1;
                Class.Users.dt1.Rows[i]["checkstitch1"] = inchst1;
                Class.Users.dt1.Rows[i]["checkdel1"] = inchdel2;
                Class.Users.dt1.Rows[i]["Balqty"] = total;
                bqty += total;             
                j++; i++;
            }
            Cursor = Cursors.Default;
        }
        byte[] stdbytes; Int64 std;
        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(combopono.Text != "ALL")
            {

            }
            string sel="select a.garmentimage from asptblpur a where a.pono='"+combopono.Text+"'";
            DataTable dt = CC.select(sel, "asptblpur");
            int c = dt.Rows.Count;
            if (c >= 1)
            {
                if (c >= 1 && dt.Rows[0]["garmentimage"].ToString() !="")
                {
                    pictureBox1.Image = null; stdbytes = null;
                    stdbytes = (byte[])dt.Rows[0]["garmentimage"];
                    Image img = Models.Device.ByteArrayToImage(stdbytes);
                    pictureBox1.Image = img;
                }
                else
                {
                    pictureBox1.Image = null; stdbytes = null;
                }
            }
        }

        private void downLoadToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void butdatewise_Click(object sender, EventArgs e)
        {
            // string sel = "select  b.asptblpurid,  date_format(a.modified,'%d-%m-%Y') FromDate,b.orderno,b.styleref,b.compname,b.address,b.orderqty,b.excessqty,a.pono,a.colorname,a.sizename,a.inward,ifnull(a.rework,0) rework,ifnull(a.delivery,0) delivery,ifnull(a.cinward,0) as checkinginward, ifnull(a.cdefqty,0) as checkingrework, a.cdefqty as stitchingrework,a.cdelivery as checkingdelivery,ifnull(a.crework,0) as chereworkdelivery delivery1 from  ( select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.rework) rework,sum(aa.delivery) delivery,  sum(aa.sdefqty) sdefqty,sum(aa.cdefqty) cdefqty,sum(aa.csdefqty) csdefqty,sum(aa.cinward) cinward,sum(aa.cdelivery) cdelivery,sum(aa.crework) crework from (  select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework,0 sdefqty,0 cdefqty,0 csdefqty,  0 cinward, 0 cdelivery,0 crework from (   select a.modified,a.pono,c.colorname,d.sizename, case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid   join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "' group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype ) aa  group by aa.modified,aa.pono,aa.colorname,aa.sizename   union all   select a.modified,a.pono,a.colorname,a.sizename,0 inward,0 delivery,0 rework, sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty,   sum(a.csdefqty) csdefqty,0 cinward,0 cdelivery,0 crework from ( select a.modified,a.pono,c.colorname,d.sizename,  case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty,   case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,    case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a   join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname   join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "'   group by a.modified,a.pono,c.colorname,d.sizename,a.ISSUETYPE,a.remarks ) a group by a.modified,a.pono,a.colorname,a.sizename  union all select aa.modified,aa.pono,aa.colorname,aa.sizename,0 inward,0 delivery,0 rework,0 sdefqty,0 cdefqty,0 csdefqty,  sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select a.modified,a.pono,c.colorname,d.sizename,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a   join asptblchkdet b on a.asptblchkid = b.asptblchkid join asptblcolmas c on c.asptblcolmasid = b.colorname   join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "'   group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype ) aa   group by aa.modified,aa.pono,aa.colorname,aa.sizename ) aa   group by aa.modified,aa.pono,aa.colorname,aa.sizename ) a join (  select  a.asptblpurid,a.pono,a.orderno,a.styleref,e.compname,e.address,a.orderqty,a.excessqty from asptblpur a  join gtcompmast e on e.gtcompmastid = a.compcode   ) b on a.pono = b.pono order by 1";
            string sel = "select  b.asptblpurid,  date_format(a.modified,'%d-%m-%Y') FromDate,b.orderno,b.styleref,b.compname,b.address,b.orderqty,b.excessqty,a.pono,a.colorname,a.sizename,a.inward,ifnull(a.rework,0) rework,ifnull(a.delivery,0) delivery,ifnull(a.cinward,0) as checkinginward, ifnull(a.cdefqty,0) as checkingrework, a.cdefqty as stitchingrework,a.cdelivery as checkingdelivery,ifnull(a.crework,0) as delivery1 from  ( select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.rework) rework,sum(aa.delivery) delivery,  sum(aa.sdefqty) sdefqty,sum(aa.cdefqty) cdefqty,sum(aa.csdefqty) csdefqty,sum(aa.cinward) cinward,sum(aa.cdelivery) cdelivery,sum(aa.crework) crework from (  select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework,0 sdefqty,0 cdefqty,0 csdefqty,  0 cinward, 0 cdelivery,0 crework from (   select a.modified,a.pono,c.colorname,d.sizename, case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid   join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "' group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype ) aa  group by aa.modified,aa.pono,aa.colorname,aa.sizename   union all   select a.modified,a.pono,a.colorname,a.sizename,0 inward,0 delivery,0 rework, sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty,   sum(a.csdefqty) csdefqty,0 cinward,0 cdelivery,0 crework from ( select a.modified,a.pono,c.colorname,d.sizename,  case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty,   case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,    case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a   join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname   join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "'   group by a.modified,a.pono,c.colorname,d.sizename,a.ISSUETYPE,a.remarks ) a group by a.modified,a.pono,a.colorname,a.sizename  union all select aa.modified,aa.pono,aa.colorname,aa.sizename,0 inward,0 delivery,0 rework,0 sdefqty,0 cdefqty,0 csdefqty,  sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select a.modified,a.pono,c.colorname,d.sizename,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a   join asptblchkdet b on a.asptblchkid = b.asptblchkid join asptblcolmas c on c.asptblcolmasid = b.colorname   join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "'   group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype ) aa   group by aa.modified,aa.pono,aa.colorname,aa.sizename ) aa   group by aa.modified,aa.pono,aa.colorname,aa.sizename ) a join (  select  a.asptblpurid,a.pono,a.orderno,a.styleref,e.compname,e.address,a.orderqty,a.excessqty from asptblpur a  join gtcompmast e on e.gtcompmastid = a.compcode   ) b on a.pono = b.pono order by 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpurdet1");
            DataTable dt = ds.Tables["asptblpurdet1"];
            rd3.SetDataSource(dt);
            string sel1 = "select a.companylogo from gtcompmast a where a.compcode='" + Class.Users.HCompcode + "'";
            ds = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
          DataTable  dt1 = ds.Tables["gtcompmast"];
            rd2.Database.Tables["DataTable2"].SetDataSource(dt);
            rd2.Database.Tables["DataTableImage"].SetDataSource(dt1);
            crystalReportViewer2.ReportSource = null;
            crystalReportViewer2.ReportSource = rd3;
            crystalReportViewer2.Refresh();
            Cursor.Current = Cursors.Default;
            tabControl2.SelectTab("TabPage4");

          

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            reportwise();
        }
        void reportwise()
        {
            if (combotype.Text == "PO WISE")
            {
                frmdate.Enabled = false;
                todate.Enabled = false;
                butdatewise.Enabled = false;
                combopono.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                combopono.Enabled = false;
                button1.Enabled = false;
                frmdate.Enabled = true;
                todate.Enabled = true;
                butdatewise.Enabled = true;
            }
        }
        void reportwise1()
        {
            if (combotype1.Text == "PO WISE")
            {
                datepcsrate1.Enabled = false;
                datepcsrate2.Enabled = false;
                butdatewise.Enabled = false;
                combopopcsrate.Enabled = true;
                butpcsmonthly.Enabled = true;
                button2.Enabled = false;
            }
            else
            {
                butpcsmonthly.Enabled = false;
                combopopcsrate.Enabled = false;
                button2.Enabled = true;
                datepcsrate1.Enabled = true;
                datepcsrate2.Enabled = true;
                butdatewise.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           




        }

        private void butpcsmonthly_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); DataTable dt = new DataTable();
            if (combopono.Text == "ALL")
            {
                //  string sel = "select aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) inward,sum(aa.delivery) delivery,  sum(aa.rework) rework,aa.test2,aa.transactions,aa.sno from (   select e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,    case when a.issuetype = 'DELIVERY'   then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework ,a.notes as pronotes,'PRODUCTION ENTRY' transactions,'1' sno  from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid   join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename join gtcompmast e on e.gtcompmastid=a.compcode  where  a.notes='" + combopcsmonthly.Text + "'  group by e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa  group by aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.pronotes,aa.transactions,aa.sno  union all  select aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework,aa.pronotes ,aa.transactions,aa.sno from (   select e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY'  then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework,a.notes as pronotes,'CHECKING ENTRY' transactions,'2'sno from asptblchk a   join asptblchkdet b on a.asptblchkid = b.asptblchkid  join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  join gtcompmast e on e.gtcompmastid=a.compcode where  a.notes='" + combopcsmonthly.Text + "' group by e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa  group by aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.pronotes,aa.transactions,aa.sno  order by 12,3,7";
                string sel = "select a.podate, a.pono,a.colorname,a.SIZENAME,a.ASPTBLSIZMASID,a.compname,a.address,a.orderqty,a.excessqty,a.orderno,a.styleref, b.inward,ifnull(b.rework,0) rework, b.delivery delivery,ifnull(d.cinward,0)  as checkinginward, ifnull(c.cdefqty,0) as checkingrework, ifnull(c.csdefqty,0) stitchingrework,  ifnull(d.cdelivery,0) as checkingdelivery,  ifnull(d.cinward,0)-ifnull(d.cdelivery,0)-ifnull(d.crework,0) balqty,ifnull(d.crework,0) as delivery1,b.test1,c.test2,d.test3 from (    select a.pono,c.colorname,d.SIZENAME,d.ASPTBLSIZMASID,b.orderqty, a.podate,a.orderno,a.styleref,e.compname,e.address,a.excessqty from asptblpur a  join asptblpurdet b on a.asptblpurid = b.asptblpurid   join asptblcolmas c on c.asptblcolmasid = b.colorname  join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  join gtcompmast e on e.gtcompmastid = a.compcode    where (a.pono='" + combopopcsrate.Text + "' or 'ALL' ='" + combopopcsrate.Text + "') ) a   left join (     select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.test1, sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework from (   select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.notes as test1, case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,  case when a.issuetype = 'DELIVERY'   then sum(b.orderqty) else 0 end  delivery,  case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid  join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  where  a.notes='" + combopcsmonthly.Text + "' group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa  group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.test1 ) b on a.pono = b.pono and a.colorname = b.colorname and a.sizename = b.sizename left join (  select a.pono,a.colorname,a.sizename,a.ASPTBLSIZMASID,a.test2 ,sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty, sum(a.csdefqty) csdefqty from (  select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.notes as test2, case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty,  case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,   case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a  join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname  join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where  a.notes='" + combopcsmonthly.Text + "' group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.ISSUETYPE,a.remarks,a.notes ) a  group by a.pono,a.colorname,a.sizename, a.ASPTBLSIZMASID,a.test2) c on a.pono = c.pono and a.colorname = c.colorname and a.sizename = c.sizename left join (  select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.test3, sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.notes as test3, case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,  case when a.issuetype = 'DELIVERY'  then sum(b.orderqty) else 0 end  delivery,  case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a join asptblchkdet b on a.asptblchkid = b.asptblchkid  join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where   a.notes='" + combopcsmonthly.Text + "' group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,a.aa.test3 ) d on a.pono = d.pono and a.colorname = d.colorname and a.sizename = d.sizename order by 1,2,3,5";
                ds = Utility.ExecuteSelectQuery(sel, "asptblpurdet1");
                dt = ds.Tables["asptblpurdet1"];
            }
            else
            {
                // string sel = "select aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) inward,sum(aa.delivery) delivery,  sum(aa.rework) rework,aa.pronotes,aa.transactions,aa.sno from (   select e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,    case when a.issuetype = 'DELIVERY'   then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework ,a.notes as pronotes,'PRODUCTION ENTRY' transactions,'1' sno  from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid   join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename join gtcompmast e on e.gtcompmastid=a.compcode  where (a.pono='" + combopono.Text+"' or 'ALL' ='"+ combopono.Text + "') and a.notes='"+combopcsmonthly.Text+ "'  group by e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa  group by aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.pronotes,aa.transactions,aa.sno  union all  select aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework,aa.pronotes ,aa.transactions,aa.sno from (   select e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY'  then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework,a.notes as pronotes,'CHECKING ENTRY' transactions,'2'sno from asptblchk a   join asptblchkdet b on a.asptblchkid = b.asptblchkid  join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  join gtcompmast e on e.gtcompmastid=a.compcode where (a.pono='" + combopono.Text+"' or 'ALL' ='"+combopono.Text+"' ) and a.notes='"+combopcsmonthly.Text+ "' group by e.compname,e.address,a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa  group by aa.compname,aa.address,aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.pronotes,aa.transactions,aa.sno  order by 12,7";
                string sel = "select a.podate, a.pono,a.colorname,a.SIZENAME,a.ASPTBLSIZMASID,a.compname,a.address,a.orderqty,a.excessqty,a.orderno,a.styleref, b.inward,ifnull(b.rework,0) rework, b.delivery delivery,ifnull(d.cinward,0)  as checkinginward, ifnull(c.cdefqty,0) as checkingrework, ifnull(c.csdefqty,0) stitchingrework,  ifnull(d.cdelivery,0) as checkingdelivery,  ifnull(d.cinward,0)-ifnull(d.cdelivery,0)-ifnull(d.crework,0) balqty,ifnull(d.crework,0) as delivery1,b.test1,c.test2,d.test3 from (    select a.pono,c.colorname,d.SIZENAME,d.ASPTBLSIZMASID,b.orderqty, a.podate,a.orderno,a.styleref,e.compname,e.address,a.excessqty from asptblpur a  join asptblpurdet b on a.asptblpurid = b.asptblpurid   join asptblcolmas c on c.asptblcolmasid = b.colorname  join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  join gtcompmast e on e.gtcompmastid = a.compcode    where (a.pono='" + combopopcsrate.Text + "' or 'ALL' ='" + combopopcsrate.Text + "') ) a   left join (     select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.test1, sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework from (   select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.notes as test1, case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,  case when a.issuetype = 'DELIVERY'   then sum(b.orderqty) else 0 end  delivery,  case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid  join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  where (a.pono='" + combopopcsrate.Text + "' or 'ALL' ='" + combopopcsrate.Text + "')  and a.notes='" + combopcsmonthly.Text + "' group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa  group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.test1 ) b on a.pono = b.pono and a.colorname = b.colorname and a.sizename = b.sizename left join (  select a.pono,a.colorname,a.sizename,a.ASPTBLSIZMASID,a.test2 ,sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty, sum(a.csdefqty) csdefqty from (  select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.notes as test2, case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty,  case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,   case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a  join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname  join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where (a.pono='" + combopopcsrate.Text + "' or 'ALL' ='" + combopopcsrate.Text + "')  and a.notes='" + combopcsmonthly.Text + "' group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.ISSUETYPE,a.remarks,a.notes ) a  group by a.pono,a.colorname,a.sizename, a.ASPTBLSIZMASID,a.test2) c on a.pono = c.pono and a.colorname = c.colorname and a.sizename = c.sizename left join (  select aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,aa.test3, sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.notes as test3, case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,  case when a.issuetype = 'DELIVERY'  then sum(b.orderqty) else 0 end  delivery,  case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a join asptblchkdet b on a.asptblchkid = b.asptblchkid  join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where  (a.pono='" + combopopcsrate.Text + "' or 'ALL' ='" + combopopcsrate.Text + "')  and a.notes='" + combopcsmonthly.Text + "' group by a.pono,c.colorname,d.sizename,d.ASPTBLSIZMASID,a.issuetype,a.notes ) aa group by aa.pono,aa.colorname,aa.sizename,aa.ASPTBLSIZMASID,a.aa.test3 ) d on a.pono = d.pono and a.colorname = d.colorname and a.sizename = d.sizename order by 1,2,3,5";
                ds = Utility.ExecuteSelectQuery(sel, "asptblpurdet1");
                dt = ds.Tables["asptblpurdet1"];
            }

            string sel1 = "select a.companylogo from gtcompmast a where a.compcode='" + Class.Users.HCompcode + "'";
            ds = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
           DataTable dt1 = ds.Tables["gtcompmast"];
            rd4.Database.Tables["DataTable2"].SetDataSource(dt);
            rd4.Database.Tables["DataTableImage"].SetDataSource(dt1);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd4;
            crystalReportViewer1.Refresh();


            Cursor.Current = Cursors.Default;
            tabControl2.SelectTab("TabPage2");
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            // string sel = "select  b.asptblpurid,  date_format(a.modified,'%d-%m-%Y') FromDate,b.orderno,b.styleref,b.compname,b.address,b.orderqty,b.excessqty,a.pono,a.colorname,a.sizename,a.inward,ifnull(a.rework,0) rework,ifnull(a.delivery,0) delivery,ifnull(a.cinward,0) as checkinginward, ifnull(a.cdefqty,0) as checkingrework, a.cdefqty as stitchingrework,a.cdelivery as checkingdelivery,ifnull(a.crework,0) as delivery1 from  ( select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.rework) rework,sum(aa.delivery) delivery,  sum(aa.sdefqty) sdefqty,sum(aa.cdefqty) cdefqty,sum(aa.csdefqty) csdefqty,sum(aa.cinward) cinward,sum(aa.cdelivery) cdelivery,sum(aa.crework) crework from (  select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework,0 sdefqty,0 cdefqty,0 csdefqty,  0 cinward, 0 cdelivery,0 crework from (   select a.modified,a.pono,c.colorname,d.sizename, case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid   join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "' group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype ) aa  group by aa.modified,aa.pono,aa.colorname,aa.sizename   union all   select a.modified,a.pono,a.colorname,a.sizename,0 inward,0 delivery,0 rework, sum(a.sdefqty) sdefqty, sum(a.cdefqty) cdefqty,   sum(a.csdefqty) csdefqty,0 cinward,0 cdelivery,0 crework from ( select a.modified,a.pono,c.colorname,d.sizename,  case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty,   case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,    case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty  from asptblcutpanret a   join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname   join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "'   group by a.modified,a.pono,c.colorname,d.sizename,a.ISSUETYPE,a.remarks ) a group by a.modified,a.pono,a.colorname,a.sizename  union all select aa.modified,aa.pono,aa.colorname,aa.sizename,0 inward,0 delivery,0 rework,0 sdefqty,0 cdefqty,0 csdefqty,  sum(aa.inward) cinward,sum(aa.delivery) cdelivery,sum(aa.rework) crework from ( select a.modified,a.pono,c.colorname,d.sizename,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,   case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,   case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework from asptblchk a   join asptblchkdet b on a.asptblchkid = b.asptblchkid join asptblcolmas c on c.asptblcolmasid = b.colorname   join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "'   group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype ) aa   group by aa.modified,aa.pono,aa.colorname,aa.sizename ) aa   group by aa.modified,aa.pono,aa.colorname,aa.sizename ) a join (  select  a.asptblpurid,a.pono,a.orderno,a.styleref,e.compname,e.address,a.orderqty,a.excessqty from asptblpur a  join gtcompmast e on e.gtcompmastid = a.compcode   ) b on a.pono = b.pono order by 1";
            string sel = "select  b.asptblpurid,  date_format(a.modified,'%d-%m-%Y') FromDate,b.orderno,b.styleref,b.compname,b.address,b.orderqty,b.excessqty,a.pono,a.colorname,a.sizename,a.inward,ifnull(a.rework,0) rework,ifnull(a.delivery,0) delivery,ifnull(a.cinward,0) as checkinginward, ifnull(a.cdefqty,0) as checkingrework, a.cdefqty as stitchingrework,a.cdelivery as checkingdelivery,ifnull(a.crework,0) as delivery1,a.notes from (select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.rework) rework,sum(aa.delivery) delivery,  sum(aa.sdefqty) sdefqty,sum(aa.cdefqty) cdefqty,sum(aa.csdefqty) csdefqty,sum(aa.cinward) cinward,sum(aa.cdelivery) cdelivery,sum(aa.crework) crework ,aa.notes from ( select aa.modified,aa.pono,aa.colorname,aa.sizename,sum(aa.inward) inward,sum(aa.delivery) delivery,sum(aa.rework) rework,0 sdefqty,0 cdefqty,0 csdefqty,  0 cinward, 0 cdelivery,0 crework,aa.notes from (   select a.modified,a.pono,c.colorname,d.sizename,  case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,    case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,     case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework,a.notes from asptblprolot a join asptblprolotdet b on a.asptblprolotid = b.asptblprolotid     join asptblcolmas c on c.asptblcolmasid = b.colorname join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename  where a.modified between '" + datepcsrate1.Value.ToString("yyyy-MM-dd") + "' and  '" + datepcsrate2.Value.ToString("yyyy-MM-dd") + "' and a.notes='" + combopcsmonthly.Text + "'   group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype,a.notes ) aa   group by aa.modified,aa.pono,aa.colorname,aa.sizename,aa.notes     union all  select bb.modified,bb.pono,bb.colorname,bb.sizename,0 inward,0 delivery,0 rework, sum(bb.sdefqty) sdefqty, sum(bb.cdefqty) cdefqty,    sum(bb.csdefqty) csdefqty,0 cinward,0 cdelivery,  0 crework,bb.notes from (   select a.modified,a.pono,c.colorname,d.sizename,    case when a.ISSUETYPE = 'STITCHING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end sdefqty,     case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'CHECKING MISTAKE' then sum(b.pcs) else 0 end cdefqty,      case when a.ISSUETYPE = 'CHECKING MISTAKE' and a.remarks = 'STITCHING MISTAKE' then sum(b.pcs) else 0 end csdefqty ,a.notes from asptblcutpanret a     join asptblcutpanretdet b on a.asptblcutpanretid = b.asptblcutpanretid join asptblcolmas c on c.asptblcolmasid = b.colorname     join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + datepcsrate1.Value.ToString("yyyy-MM-dd") + "' and  '" + datepcsrate2.Value.ToString("yyyy-MM-dd") + "' and a.notes='" + combopcsmonthly.Text + "'     group by a.modified,a.pono,c.colorname,d.sizename,a.ISSUETYPE,a.remarks,a.notes ) bb group by bb.modified,bb.pono,bb.colorname,bb.sizename ,bb.notes  union all  select cc.modified,cc.pono,cc.colorname,cc.sizename,0 inward,0 delivery,0 rework,0 sdefqty,0 cdefqty,0 csdefqty,  sum(cc.inward) cinward,sum(cc.delivery) cdelivery,  sum(cc.rework) crework,cc.notes from ( select a.modified,a.pono,c.colorname,d.sizename,    case when a.issuetype = 'INWARD' then sum(b.orderqty) else 0 end inward,     case when a.issuetype = 'DELIVERY' then sum(b.orderqty) else 0 end  delivery,     case when a.issuetype = 'REWORK' then sum(b.orderqty) else 0 end   rework,a.notes from asptblchk a   join asptblchkdet b on a.asptblchkid = b.asptblchkid   join asptblcolmas c on c.asptblcolmasid = b.colorname   join asptblsizmas d on d.ASPTBLSIZMASID = b.sizename where a.modified between '" + datepcsrate1.Value.ToString("yyyy-MM-dd") + "' and  '" + datepcsrate2.Value.ToString("yyyy-MM-dd") + "' and a.notes='"+combopcsmonthly.Text+"'     group by a.modified,a.pono,c.colorname,d.sizename,a.issuetype,a.notes ) cc   group by cc.modified,cc.pono,cc.colorname,cc.sizename,cc.notes) aa  group by aa.modified,aa.pono,aa.colorname,aa.sizename,aa.notes)a  join (  select  a.asptblpurid,a.pono,a.orderno,a.styleref,e.compname,e.address,a.orderqty,a.excessqty   from asptblpur a   join gtcompmast e on e.gtcompmastid = a.compcode   ) b on a.pono = b.pono";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpurdet1");
            DataTable dt = ds.Tables["asptblpurdet1"];
            rd5.SetDataSource(dt);
            crystalReportViewer2.ReportSource = null;
            crystalReportViewer2.ReportSource = rd5;
            crystalReportViewer2.Refresh();
            Cursor.Current = Cursors.Default;
            tabControl2.SelectTab("TabPage4");
        }

        private void combotype1_SelectedIndexChanged(object sender, EventArgs e)
        {
            reportwise1();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            crystalReportViewer2.ReportSource = null; crystalReportViewer1.ReportSource = null;
        }
    }
}

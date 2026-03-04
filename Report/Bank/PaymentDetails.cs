using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Report.Bank
{
    public partial class PaymentDetails : Form,ToolStripAccess
    {
        public PaymentDetails()
        {
            InitializeComponent();
                    }
        private static PaymentDetails _instance;
        //Models.Master mas = new Models.Master();
        //Models.UserRights sm = new Models.UserRights();
        //Models.Validate va = new Models.Validate();

        public static PaymentDetails Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PaymentDetails();
                GlobalVariables.CurrentForm = _instance; return _instance;

            }
        }
        public void loadproduct()
        {
            string sel = "SELECT  a.asptblpartymasid,a.partyname from asptblpartymas a where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
            DataTable dt = ds.Tables["asptblpartymas"];
            if (dt.Rows.Count > 0)
            {
                comboparty.DisplayMember = "partyname";
                comboparty.ValueMember = "asptblpartymasid";
                comboparty.DataSource = dt;
            }
        }
        private void PaymentDetails_Load(object sender, EventArgs e)
        {
           
        }

        string tabname = "";
           
            void reportPayment(CrystalDecisions.Windows.Forms.CrystalReportViewer cry,string tabname)
            {
              
            cry.ReportSource = null; cry.Refresh(); ;
            if (tabname == "tabPage1")
                {
                    Report.Bank.BankDetails rd = new BankDetails();
                Report.Bank.PartyDetails rd1 = new PartyDetails();
                if (checkall.Checked == false)
                {
                    string sel = "select b.asptblledmasid,  a.accountholdername as compcode,a.partyname,f.bankname,a.accno ,g.branch,g.ifsc, b.ledgermonth,round(b.amount) as amount,b.ledgerDate,b.remarks1,'" + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "' as  fromdate,'" + dateTimePicker2.Value.ToString("dd-MM-yyyy") + "' as todate from asptblpartymas a join asptblledmas b on a.asptblpartymasid=b.partyname join gtcitymast c on c.gtcitymastid=a.city  join gtstatemast d on d.gtstatemastid=a.state join gtcountrymast e on e.gtcountrymastid=a.country  join asptblbanmas f on f.asptblbanmasid=a.bankname   join asptblifscmas g on g.asptblifscmasid=a.ifsc and g.asptblifscmasid=a.branch" +
                        " where a.partyname ='" + comboparty.Text + "' and b.ledgerdate between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' order by b.asptblledmasid desc ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];
                    string sel1 = "select a.compname,a.address, a.companylogo from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd.Database.Tables["DataTable1"].SetDataSource(dt2);
                    rd.Database.Tables["DataTable11"].SetDataSource(dt3);
                    cry.ReportSource = null;
                    cry.ReportSource = rd;
                    cry.Refresh();
                    tabControl1.SelectTab(tabPage1);
                }
                else
                {
                    string sel = "select b.asptblledmasid,  a.accountholdername as compcode,a.partyname,f.bankname,a.accno ,g.branch,g.ifsc, b.ledgermonth,round(b.amount) as amount,b.ledgerDate,b.remarks1,'" + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "' as  fromdate,'" + dateTimePicker2.Value.ToString("dd-MM-yyyy") + "' as todate from asptblpartymas a join asptblledmas b on a.asptblpartymasid=b.partyname join gtcitymast c on c.gtcitymastid=a.city  join gtstatemast d on d.gtstatemastid=a.state join gtcountrymast e on e.gtcountrymastid=a.country  join asptblbanmas f on f.asptblbanmasid=a.bankname   join asptblifscmas g on g.asptblifscmasid=a.ifsc and g.asptblifscmasid=a.branch" +
                       " where  b.ledgerdate between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' order by b.asptblledmasid desc ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];
                    string sel1 = "select a.compname,a.address, a.companylogo from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd.Database.Tables["DataTable1"].SetDataSource(dt2);
                    rd.Database.Tables["DataTable11"].SetDataSource(dt3);
                    cry.ReportSource = null;
                    cry.ReportSource = rd;
                    cry.Refresh();
                    tabControl1.SelectTab(tabPage1);
                }
                cry.Zoom(115);
            }
            if (tabname == "tabPage2")
            {
                Report.Bank.BankDetails rd = new BankDetails();
                Report.Bank.PartyDetails rd1 = new PartyDetails();

                if (checkall.Checked == false)
                {
                    string sel = "select a.asptblpartymasid,a.accountholdername,a.partyname, f.bankname,a.accno ,g.branch,g.ifsc, b.ledgermonth,round(b.amount) as amount,b.ledgerDate,a.gstno,a.email,b.paymenttype,'VENPAY' AS test2,b.remarks1,b.remarks2,b.remarks3,'11' as test1 from asptblpartymas a join asptblledmas b on a.asptblpartymasid=b.partyname join gtcitymast c on c.gtcitymastid=a.city  join gtstatemast d on d.gtstatemastid=a.state join gtcountrymast e on e.gtcountrymastid=a.country  join asptblbanmas f on f.asptblbanmasid=a.bankname   join asptblifscmas g on g.asptblifscmasid=a.ifsc and g.asptblifscmasid=a.branch" +
                    " where a.partyname ='" + comboparty.Text + "' and b.ledgerdate between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' order by a.asptblpartymasid desc";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];

                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1,a.phoneno from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd1.Database.Tables["DataTable12"].SetDataSource(dt2);
                    rd1.Database.Tables["DataTable11"].SetDataSource(dt3);
                    cry.ReportSource = null;
                    cry.ReportSource = rd1;
                    cry.Refresh();
                    tabControl1.SelectTab(tabPage2);
                }
                else
                {
                    string sel = "select a.asptblpartymasid,  a.accountholdername,a.partyname,f.bankname,a.accno ,g.branch,g.ifsc, b.ledgermonth,round(b.amount) as amount,b.ledgerDate,a.gstno,b.paymenttype,'VENPAY' AS test2,b.remarks1,b.remarks2,b.remarks3,'11' as test1 from asptblpartymas a join asptblledmas b on a.asptblpartymasid=b.partyname join gtcitymast c on c.gtcitymastid=a.city  join gtstatemast d on d.gtstatemastid=a.state join gtcountrymast e on e.gtcountrymastid=a.country  join asptblbanmas f on f.asptblbanmasid=a.bankname   join asptblifscmas g on g.asptblifscmasid=a.ifsc and g.asptblifscmasid=a.branch" +
                       " where  b.ledgerdate between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' order by a.asptblpartymasid desc ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];
                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1,a.phoneno from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");//,a.website,a.email,a.accno
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd1.Database.Tables["DataTable12"].SetDataSource(dt2);
                    rd1.Database.Tables["DataTable11"].SetDataSource(dt3);
                    cry.ReportSource = null;
                    cry.ReportSource = rd1;
                    cry.Refresh();
                    tabControl1.SelectTab(tabPage2);
                }
                cry.Zoom(115);
            }
            if (tabname == "tabPage3")
            {
                Report.Bank.AdvanceReport rd2 = new AdvanceReport();
                if (checkall.Checked == false)
                {
                    string sel = "select  a.asptbladvpaymasid,a.dateofpayment,b.accountholdername as partyname,c.department,a.orderno,a.itemdesc,d.resonseperson as personname,a.amount,a.gst,a.invoiceamt  ,a.tds,a.tdsvalue,a.deductionamt,a.advanceterms,round(a.advanceamount) as advanceamt from asptbladvpaymas a join asptblpartymas b on a.partyname=b.asptblpartymasid join asptbldeptmas c on c.asptbldeptmasid=a.department join asptblresmas d on d.asptblresmasid=a.responseperson where b.partyname ='" + comboparty.Text + "' and a.dateofpayment between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' order by a.asptbladvpaymasid desc ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];

                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1 from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd2.Database.Tables["Advance1"].SetDataSource(dt2);
                    rd2.Database.Tables["DataTable11"].SetDataSource(dt3);
                  
                    cry.ReportSource = rd2;
                   
                   
                }
                else
                {
                    string sel = "select  a.asptbladvpaymasid,a.dateofpayment,b.accountholdername as partyname,c.department,a.orderno,a.itemdesc,d.resonseperson as personname,a.amount,a.gst,a.invoiceamt  ,a.tds,a.tdsvalue,a.deductionamt,a.advanceterms,round(a.advanceamount) as advanceamt from asptbladvpaymas a join asptblpartymas b on a.partyname=b.asptblpartymasid join asptbldeptmas c on c.asptbldeptmasid=a.department join asptblresmas d on d.asptblresmasid=a.responseperson where  a.dateofpayment between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' order by a.asptbladvpaymasid desc ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];
                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1 from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd2.Database.Tables["Advance1"].SetDataSource(dt2);
                    rd2.Database.Tables["DataTable11"].SetDataSource(dt3);                   
                    cry.ReportSource = rd2;
                   
                   
                }
                cry.Zoom(115);
                tabControl1.SelectTab(tabPage3);
            }
            if (tabname == "tabPage4")
            {
                Report.Bank.ApprovalStatus rd4 = new ApprovalStatus();
                cry.ReportSource = null; cry.Refresh();
                if (checkall.Checked == false)
                {
                    string sel = "select  a.asptbladvpaymasid,a.dateofpayment,b.accountholdername as partyname,c.department,a.orderno,a.itemdesc,d.resonseperson as personname,a.amount,a.gst,a.invoiceamt  ,a.tds,a.tdsvalue,a.deductionamt,a.advanceterms,round(a.advanceamount) as advanceamt,a.ManagerRemarks,a.MDRemarks,case when a.approval='R' then 'Reject' when a.approval='T' THEN 'Approved'  ELSE 'Pending' end 'approval',case when a.mdapproval='R' then 'Reject' when a.mdapproval='T' THEN 'Approved'  ELSE 'Pending' end 'MDapproval',case when  a.conformed='T' THEN 'Paid' when a.conformed='R' THEN 'Reject' when a.conformed='D' THEN 'Deleted'  ELSE 'UnPaid' end 'conformed' from asptbladvpaymas a join asptblpartymas b on a.partyname=b.asptblpartymasid join asptbldeptmas c on c.asptbldeptmasid=a.department join asptblresmas d on d.asptblresmasid=a.responseperson where b.partyname ='" + comboparty.Text + "' and a.dateofpayment between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'  order by a.asptbladvpaymasid desc ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];
                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1 from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd4.Database.Tables["Advance1"].SetDataSource(dt2);
                    rd4.Database.Tables["DataTable11"].SetDataSource(dt3);
                    cry.ReportSource = rd4;
                    cry.Refresh();
                    tabControl1.SelectTab(tabPage4);
                }
                else
                {
                    string sel = "select  a.asptbladvpaymasid,a.dateofpayment,b.accountholdername as partyname,c.department,a.orderno,a.itemdesc,d.resonseperson as personname,a.amount,a.gst,a.invoiceamt  ,a.tds,a.tdsvalue,a.deductionamt,a.advanceterms,round(a.advanceamount) as advanceamt,a.ManagerRemarks,a.MDRemarks,case when a.approval='R' then 'Reject' when a.approval='T' THEN 'Approved'  ELSE 'Pending' end 'approval',case when a.mdapproval='R' then 'Reject' when a.mdapproval='T' THEN 'Approved'  ELSE 'Pending' end 'MDapproval',case when  a.conformed='T' THEN 'Paid' when a.conformed='R' THEN 'Reject' when a.conformed='D' THEN 'Deleted'  ELSE 'UnPaid' end 'conformed' from asptbladvpaymas a join asptblpartymas b on a.partyname=b.asptblpartymasid join asptbldeptmas c on c.asptbldeptmasid=a.department join asptblresmas d on d.asptblresmasid=a.responseperson where  a.dateofpayment between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "' and a.ManagerRemarks is not null or a.MDRemarks is not null order by a.asptbladvpaymasid desc ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];
                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1 from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd4.Database.Tables["Advance1"].SetDataSource(dt2);
                    rd4.Database.Tables["DataTable11"].SetDataSource(dt3);
                    cry.ReportSource = rd4;
                    cry.Refresh();
                    tabControl1.SelectTab(tabPage4);
                }
                cry.Zoom(115);
            }
            if (tabname == "tabPage5")
            {
                Report.Bank.UTR rd5 = new UTR();
                cry.ReportSource = null; cry.Refresh();
                if (checkall.Checked == false)
                {
                    string sel = "select a.asptblutrid, a.utrdate,b.ledgerDate, d.department,e.accountholdername as partyname,c.orderno,c.itemdesc,round(b.amount) as amount,a.utrno,a.billno,c.ManagerRemarks ,c.MDRemarks  from asptblutr a join asptblledmas b on a.asptblledmasid=b.asptblledmasid join asptbladvpaymas c on c.asptbladvpaymasid=b.asptbladvpaymasid join asptbldeptmas d on d.asptbldeptmasid=c.department join asptblpartymas e on e.asptblpartymasid=a.partyname  where e.partyname ='" + comboparty.Text + "' and a.utrdate between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'   order by a.asptblutrid desc";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];

                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1 from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd5.Database.Tables["Advance1"].SetDataSource(dt2);
                    rd5.Database.Tables["DataTable11"].SetDataSource(dt3);

                    cry.ReportSource = rd5;
                    cry.Refresh();
                    tabControl1.SelectTab(tabPage5);
                }
                else
                {
                    string sel = "select a.asptblutrid, a.utrdate,b.ledgerDate, d.department,e.accountholdername as partyname,c.orderno,c.itemdesc,round(b.amount) as amount,a.utrno,a.billno,c.ManagerRemarks ,c.MDRemarks from asptblutr a join asptblledmas b on a.asptblledmasid=b.asptblledmasid join asptbladvpaymas c on c.asptbladvpaymasid=b.asptbladvpaymasid join asptbldeptmas d on d.asptbldeptmasid=c.department join asptblpartymas e on e.asptblpartymasid=a.partyname where  a.utrdate between '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' and  '" + dateTimePicker2.Value.ToString("yyyy-MM-dd") + "'   order by a.asptblutrid desc";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                    DataTable dt2 = ds2.Tables["asptblpartymas"];
                    string sel1 = "select a.compname,a.address, a.companylogo,a.website as website1,a.accno as accno1 from gtcompmast a where a.active='T'";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt3 = ds3.Tables["asptblpartymas"];
                    rd5.Database.Tables["Advance1"].SetDataSource(dt2);
                    rd5.Database.Tables["DataTable11"].SetDataSource(dt3);
                    cry.ReportSource = rd5;
                    cry.Refresh();

                    tabControl1.SelectTab(tabPage5);
                }
                cry.Zoom(115);
            }
            tabname = "";
        }
        private void butreport_Click(object sender, EventArgs e)
        {
            tabname = ""; tabname = "tabPage1";
            reportPayment(crystalReportViewer1, tabname);
            
        }
        private void butpartydetails_Click(object sender, EventArgs e)
        {
            tabname = ""; tabname = "tabPage2";
            reportPayment(crystalReportViewer2, tabname);
        }
        private void butadvance_Click(object sender, EventArgs e)
        {
            tabname = ""; tabname = "tabPage3";
            reportPayment(crystalReportViewer3, tabname);
        }
        private void butRejection_Click(object sender, EventArgs e)
        {
            tabname = ""; tabname = "tabPage4";
            reportPayment(crystalReportViewer4, tabname);
        }
        

        private void bututr_Click(object sender, EventArgs e)
        {
            tabname = ""; tabname = "tabPage5";
            reportPayment(crystalReportViewer5, tabname);
          
        }
        public void GridLoad()
        {

        }
        public void News()
        {

            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            DateTime dateForButton = DateTime.Parse("01" + "-" + System.DateTime.Now.Month + "-" + System.DateTime.Now.Year);
            int totdays = DateTime.DaysInMonth(System.DateTime.Now.Year, System.DateTime.Now.Month - 1);
            dateTimePicker1.Value = dateForButton.AddDays(-totdays);

            loadproduct();
            reportPayment(crystalReportViewer1, tabname);
          
           
        }
        public void ChangePasswords()
        {
            
        }

        public void ChangeSkins()
        {
            
        }

        public void Deletes()
        {
            
        }

        public void DownLoads()
        {
            
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
          
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        public void GlobalSearchs()
        {
            throw new NotImplementedException();
        }

       
        public void Imports()
        {
            throw new NotImplementedException();
        }

        public void Logins()
        {
            throw new NotImplementedException();
        }

    

        public void Pdfs()
        {
            throw new NotImplementedException();
        }

        public void Prints()
        {
            throw new NotImplementedException();
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
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

        public void TreeButtons()
        {
            throw new NotImplementedException();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            loadproduct();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                tabname = ""; tabname = "tabPage1";
                reportPayment(crystalReportViewer1, tabname);
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                tabname = ""; tabname = "tabPage2";
                reportPayment(crystalReportViewer2, tabname);
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                tabname = ""; tabname = "tabPage3";
                reportPayment(crystalReportViewer3, tabname);
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                tabname = ""; tabname = "tabPage4";
                reportPayment(crystalReportViewer4, tabname);
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {
                tabname = ""; tabname = "tabPage5";
                reportPayment(crystalReportViewer5, tabname);
            }
        }

        
    }
}

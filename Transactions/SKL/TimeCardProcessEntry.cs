using CrystalDecisions.Shared;
using Pinnacle.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions.SKL
{
    public partial class TimeCardProcessEntry : Form, ToolStripAccess
    {
        private static TimeCardProcessEntry _instance;

        Models.Master mas = new Models.Master();

        ListView allip = new ListView();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        public static TimeCardProcessEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new TimeCardProcessEntry();
                GlobalVariables.CurrentForm = _instance;
                return _instance;

            }
        }

        public TimeCardProcessEntry()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToShortTimeString().ToString());
            tabControl1.TabPages.Remove(tabPage2); tabControl1.TabPages.Remove(tabPage4);
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }
        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            int cnt = dt1.Rows.Count;
            if (cnt >= 1)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {
                                       if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { GlobalVariables.TreeButtons.Visible = false; tabControl1.TabPages.Remove(tabPage4); } else { GlobalVariables.TreeButtons.Visible = false; tabControl1.TabPages.Add(tabPage4); }
                       
                    }
                }



            }
            else
            {

                MessageBox.Show("No Screen Rights Defined .Pls Contact Your Administrator...");

                this.Hide();


            }


        }
       
        public void News()
        {
            GridLoad();
            companyload(); companyload1(); companyloadsearch();
            crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();
            empty(); tabControl1.SelectTab(tabPage1);
            txtsearch1.Select(); lbltoalrecord.Text = "";
        }
        public void companyload()
        {
            try
            {
                string sel = "select '-1' as  gtcompmastid,'-------' as compcode from dual union all select distinct a.gtcompmastid,a.compcode from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ptransaction ='COMPANY' and a.active='T'  and a.compcode='" + Class.Users.HCompcode + "' order by 2";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "hrpaydetails");
                DataTable dt = ds.Tables["hrpaydetails"];
                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;

            }
            catch (Exception ex)
            {
                MessageBox.Show("companyload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void companyload1()
        {
            try
            {
                string sel1 = "select '-1' as  gtcompmastid,'-------' as compcode from dual union all select distinct a.gtcompmastid,a.compcode from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ptransaction ='COMPANY' and a.active='T'  and a.compcode='" + Class.Users.HCompcode + "' order by 2";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                DataTable dt1 = ds1.Tables["hrpaydetails"];
                combocompcode1.DisplayMember = "compcode";
                combocompcode1.ValueMember = "gtcompmastid";
                combocompcode1.DataSource = dt1;

                combocompcodemuster.DisplayMember = "compcode";
                combocompcodemuster.ValueMember = "gtcompmastid";
                combocompcodemuster.DataSource = dt1;

            }
            catch (Exception ex)
            {
                MessageBox.Show("companyload1: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void companyloadsearch()
        {
            //try
            //{
            //    string sel = "select distinct a.gtcompmastid,a.compcode from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ptransaction ='COMPANY'  AND a.compcode='" + combocompcode1.Text + "'    and a.payperiod='" + comboPayPeriod1.Text + "' order by 2;";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel, "hrpaydetails");
            //    DataTable dt = ds.Tables["hrpaydetails"];
            //    combounitsearch.DisplayMember = "compcode";
            //    combounitsearch.ValueMember = "gtcompmastid";
            //    combounitsearch.DataSource = dt;
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("companyloadsearch: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

        }
        void empty()
        {
            txtsearch.Text = ""; txtsearch1.Text = "";
            
            lbltotalcount.Text = "Count :"+ 0;
            combocompcode.Text = ""; comboseq.Text = ""; comboseq.SelectedIndex = -1;
            combocompcode.SelectedIndex = -1;
            comboPayPeriod.Text = "";
            comboPayPeriod.SelectedIndex = -1;
            comboPayPeriod1.Text = "";
            comboPayPeriod1.SelectedIndex = -1;
            combocompcode1.Text = "";
            combocompcode1.SelectedIndex = -1;
            dataGridView1.AllowUserToAddRows = true;
            do
            {
                int i = 0;
                for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }

            }
            while (dataGridView1.Rows.Count > 1);
            dataGridView1.AllowUserToAddRows = false; checkall.Checked = false;
            this.Font = Class.Users.FontName;
            this.BackColor = Class.Users.BackColors;
            lbltotalcount.ForeColor= Class.Users.Color1;
            panel1.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors; 
            panel5.BackColor = Class.Users.BackColors;
            label1.ForeColor= Class.Users.Color1;
            label2.ForeColor = Class.Users.Color1;
            label3.ForeColor = Class.Users.Color1;
            label4.ForeColor = Class.Users.Color1;
            label5.ForeColor = Class.Users.Color1;
            label8.ForeColor = Class.Users.Color1;
            label9.ForeColor = Class.Users.Color1;
            label10.ForeColor = Class.Users.Color1;
            label11.ForeColor = Class.Users.Color1;
            label13.ForeColor = Class.Users.Color1;
            checkBox1.ForeColor = Class.Users.Color1;
            checkall.ForeColor = Class.Users.Color1;
            panel8.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            checkword.ForeColor = Class.Users.Color1;
            checkexcel.ForeColor = Class.Users.Color1;
            chkall.ForeColor = Class.Users.Color1;
            checkall.ForeColor = Class.Users.Color1;
            panel6.BackColor = Class.Users.BackColors;
            Class.Users.UserTime = 0;
            butheader.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;

        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combocompcode.SelectedValue) > 0)
                {
                    Cursor = Cursors.WaitCursor;
                    string sel1 = "select '------' as  payperiod from dual union all select distinct a.payperiod from hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + combocompcode.Text + "'";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];
                    comboPayPeriod.DisplayMember = "payperiod";
                    comboPayPeriod.ValueMember = "payperiod";
                    comboPayPeriod.DataSource = dt1;
                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {

            }
        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (System.Data.OleDb.OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    System.Data.OleDb.OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }
        public void Saves()
        {

        }

        public void Prints()
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)//,b.intime,b.intime1, b.outtime,b.outtime1,c.address,c.pincode 
            {
                string sel2 = "select c.tpr as category, c.intime as from1,c.outtime as to1,c.lunchfrom as cityname,c.lunchto as statename,c.breakfrom as country,c.breakto as pincode,substr(a.docdate,3,5) as docdate1, CONCAT(date_format(a.docdate,'%d-%m-%Y  '),substr(dayname(a.docdate),1,3)) as docdate,c.compname,c.address, d.cityname,e.statename,f.countryname as country,c.pincode, b.midcard as idcardno , b.empname, b.doj, b.uanno , b.esino , b.fathername, b.united , b.category,  b.department,b.designation,  b.orjpayabledays , b.nhdays , b.payabledays , b.govtdaysalary, b.otwages, b.basicda,  b.basic, b.da, b.hra, b.others, b.ebasic,  b.eda, b.ebasicda, b.ehra, b.eothers, b.payableothrs, b.otamount, b.incentive,  b.govtgross,   b.pfamount, b.esiamount, b.messamount, b.deduction,  b.netamount,   b.bankaccountno , b.bankname , b.ifsccode, b.payperiod , b.fromdate, b.todate,CASE    WHEN a.intime is null  then 'Absent' ELSE  a.intime END intime, CASE    WHEN a.outtime is null  then 'Absent' ELSE  a.outtime END outtime,a.otintime,a.otouttime,  a.shiftcnt*8 totalhrs,ifnull(a.ot,0) as totalothrs,ifnull(a.shiftcnt,0) as present ,b.advance, a.ot, b.othersexp,b.creditdate,(  select COUNT(A.INTIME) totalweeklyoff from pldatta a where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "' and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' and a.intime='WO') as totalweeklyoff,(  select sum(aa.totaldays-cc.payabledays)  as totalabsent from( select count(z.empid) totaldays from pldatta z  where z.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "' and z.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and z.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "') aa,(   select bb.payabledays from(  SELECT SUM(x.INTIME)  as payabledays from (select y.orjpayabledays as INTIME from hrpaydetails y  where y.idcardno='" + listView2.SelectedItems[0].SubItems[3].Text + "' and y.united='" + listView2.SelectedItems[0].SubItems[6].Text + "' and y.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' union all select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "'  and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "'and a.intime='WO'  union all  select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "' and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' and a.intime='NH'  )x)bb)cc) as totalabsent,b.orjpayabledays as totalpresent  from pldatta a  right join hrpaydetails b on a.payperiod=b.payperiod and a.compcode=b.united and b.idcardno=a.empid Join gtcompmast c on c.compcode=a.compcode and c.gtcompmastid=b.compcode join gtcitymast d on d.gtcitymastid=c.city join gtstatemast e on e.gtstatemastid=c.state join gtcountrymast f on f.gtcountrymastid=c.country where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "'  and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' order by 1;  ";

                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                rd.Database.Tables["DataTable1"].SetDataSource(ds2.Tables[0]);
                rd.Database.Tables["DataTable2"].SetDataSource(ds2.Tables[0]);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd;
                crystalReportViewer1.Refresh();
                rd.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                rd.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);


            }
            else
            {

            }
        }
        private void txtsearch1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; listView2.Items.Clear();
                if (txtsearch1.Text.Length >= 1)
                {

                    foreach (ListViewItem item in listfilter2.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter2.Items[item0].SubItems[4].ToString().Contains(txtsearch1.Text.ToUpper()) || listfilter2.Items[item0].SubItems[5].ToString().Contains(txtsearch1.Text.ToUpper()))
                        {
                            list.Text = listfilter2.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[7].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView2.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal2.Text = "Total Count: " + listView2.Items.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    listView2.Items.Clear();
                    foreach (ListViewItem item in listfilter2.Items)
                    {


                        this.listView2.Items.Add((ListViewItem)item.Clone());

                        if (item0 % 2 == 0)
                        {
                            item.BackColor = Color.WhiteSmoke;
                        }
                        else
                        {
                            item.BackColor = Color.White;
                        }

                        item0++;
                    }
                    lbltotal2.Text = "Total Count: " + listView2.Items.Count;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }
        }


        private void comboPayPeriod1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combocompcode1.SelectedValue) > 0 && comboPayPeriod1.Text != "")
                {
                    lbltotal2.Text = "";
                    listView2.Items.Clear(); listfilter2.Items.Clear();
                    txtsearch1.Text = ""; Cursor = Cursors.WaitCursor; checkword.Checked = false; chkall.Checked = false;
                    string sel1 = "select distinct a.hrpaydetailsid,a.idcardno ,a.midcard, a.empname, b.compcode,a.payperiod from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid  join pldatta c on a.united=c.compcode and  a.idcardno=c.empid  and c.payperiod=a.payperiod where a.payperiod='" + comboPayPeriod1.Text + "' and a.united='" + combocompcode1.Text + "' order by 1;";

                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRPayDetails");
                    DataTable dt = ds.Tables["HRPayDetails"];
                    if (dt.Rows.Count > 0)
                    {
                        int j = 1;

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list1 = new ListViewItem();
                            list1.SubItems.Add(j.ToString());
                            list1.SubItems.Add(myRow["hrpaydetailsid"].ToString());
                            list1.SubItems.Add(myRow["idcardno"].ToString());
                            list1.SubItems.Add(myRow["midcard"].ToString());
                            list1.SubItems.Add(myRow["empname"].ToString());
                            list1.SubItems.Add(myRow["compcode"].ToString());
                            list1.SubItems.Add(myRow["payperiod"].ToString());
                            this.listfilter2.Items.Add((ListViewItem)list1.Clone());
                            if (j % 2 == 0)
                            {
                                list1.BackColor = Color.WhiteSmoke;
                            }
                            else
                            {
                                list1.BackColor = Color.White;
                            }
                            listView2.Items.Add(list1);
                            j++;
                        }

                        lbltotal2.Text = "" + comboPayPeriod1.Text + " Total Record's : " + listView2.Items.Count;
                        crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();
                    }

                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        private void combocompcode1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combocompcode1.SelectedValue) > 0)
                {
                    Cursor = Cursors.WaitCursor;
                    string sel1 = "select '------' as  payperiod from dual union all select distinct a.payperiod from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid  join pldatta c on a.united=c.compcode and  a.idcardno=c.empid and b.compcode=c.compcode and c.payperiod=a.payperiod where b.compcode='" + combocompcode1.Text + "'";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];
                    comboPayPeriod1.DisplayMember = "payperiod";
                    comboPayPeriod1.ValueMember = "payperiod";
                    comboPayPeriod1.DataSource = dt1;
                    combopayperiodmuster.DisplayMember = "payperiod";
                    combopayperiodmuster.ValueMember = "payperiod";
                    combopayperiodmuster.DataSource = dt1;
                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
                lbltoalrecord.Text = ""; lbltoalrecord.Text = "";
                // combocompcode1.Text = ""; comboPayPeriod1.Text = ""; combocompcode1.SelectedIndex = -1; comboPayPeriod1.SelectedIndex = -1;
            }
            //else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            //{
            //    combocompcode1.Text = ""; comboPayPeriod1.Text = ""; combocompcode1.SelectedIndex = -1; comboPayPeriod1.SelectedIndex = -1;
            //    //gridload();label48.Text = "";
            //    label48.Text = "";
            //}
            else
            {
                //  this.Pdfs.Visible = true; this.Prints.Visible = true; this.Deletes.Visible = false; this.Saves.Visible = false;
                empty(); lbltoalrecord.Text = "";
            }
        }

        private void chkall_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                {

                    if (chkall.Checked == true)
                    {

                        int listCount = listView2.Items.Count;

                        foreach (ListViewItem item in listView2.Items)
                        {
                            item.Checked = true;
                        }


                    }
                    else
                    {
                        foreach (ListViewItem item in listView2.Items)
                        {
                            item.Checked = false;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Invalid Date");
                    this.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Searchs()
        {
           
        }

        public void Searchs(int EditID)
        {
           
        }

        public void Deletes()
        {
            if (checkall.Checked == true)
            {
                string sel1 = "select  b.empid from  gtcompmast a join pldatta b on a.compcode=b.compcode  where  a.compcode='" + combocompcode.Text + "' AND b.payperiod='" + comboPayPeriod.Text + "';";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "pldatta");
                DataTable dt1 = ds1.Tables["pldatta"];
                if (dt1.Rows.Count > 0)
                {

                    var confirmation = MessageBox.Show("Do You want Delete TimeCard Process Entry ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        Cursor = Cursors.WaitCursor;

                        if (dt1.Rows.Count > 0)
                        {
                            progressBar1.Minimum = 0; int tot = 1;
                            progressBar1.Maximum = dt1.Rows.Count;
                            for (int i = 0; i < dt1.Rows.Count; i++)
                            {
                                string del = "delete from pldatta  where compcode='" + combocompcode.Text + "' and payperiod='" + comboPayPeriod.Text + "' and  empid='" + dt1.Rows[i]["empid"].ToString() + "';";
                                Utility.ExecuteNonQuery(del);

                                decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(dt1.Rows.Count)) * (i + 1);
                                lblprogress1.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %" + dt1.Rows[i]["empid"].ToString();
                                lbltoalrecord.Text = "Data Remove from Table" + (per).ToString("N0") + " % " + "EmpID " + dt1.Rows[i]["empid"].ToString();
                                lblprogress1.Refresh(); lbltoalrecord.Refresh();
                                progressBar1.Value = i + 1; tot++;
                            }
                            lblprogress1.Text = ""; progressBar1.Value = 0;
                            Cursor = Cursors.Default; lbltoalrecord.Text = "" + comboPayPeriod.Text + "    Deleted Total Rows " + dt1.Rows.Count;
                            MessageBox.Show("Record Deleted Successfully " + tot.ToString(), "Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                            string sel2 = "select '------' as  payperiod from dual union all select distinct a.payperiod from hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid join pldatta c on c.payperiod=a.payperiod and c.compcode=b.gtcompmastid where b.compcode='" + combocompcode.Text + "'";
                            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                            DataTable dt2 = ds1.Tables["hrpaydetails"];
                            comboPayPeriod.DisplayMember = "payperiod";
                            comboPayPeriod.ValueMember = "payperiod";

                            comboPayPeriod.DataSource = dt2;
                            crystalReportViewer1.ReportSource = null;
                            Cursor = Cursors.Default;

                        }
                        else
                        {
                            Cursor = Cursors.Default;
                            MessageBox.Show("Invalid  Delete", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lbltoalrecord.Text = "";
                        }
                    }

                    Cursor = Cursors.Default;

                }
                else
                {
                    Cursor = Cursors.Default; lbltoalrecord.Text = "";
                    MessageBox.Show("No Data Found  . ", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            else
            {
                Cursor = Cursors.Default; lbltoalrecord.Text = "";
                MessageBox.Show("No Data Found pls Select Check Box . .", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        Pinnacle.PrintFormat.TimeCardProcess rd = new Pinnacle.PrintFormat.TimeCardProcess();

        private void Word_Click(object sender, EventArgs e)
        {
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                DialogResult result = MessageBox.Show("Do You want to Export  PDF Formate??", "PDF", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {

                    string folderLocation = "d:\\TimeCardDownLoad\\";
                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }
                    if (chkall.Checked == true)
                    {
                        Cursor = Cursors.WaitCursor;

                        foreach (ListViewItem item in listView2.Items)
                        {
                            if (item.SubItems.Count > 0)
                            {
                                crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();
                                string sel2 = "select  substr(a.docdate,3,5) as docdate1, CONCAT(date_format(a.docdate,'%d-%m-%Y  '),substr(dayname(a.docdate),1,3)) as docdate,c.compname,c.address, d.cityname,e.statename,f.countryname as country,c.pincode, b.midcard as idcardno , b.empname,b.doj, b.uanno , b.esino , b.fathername, b.united , b.category,  b.department,b.designation,  b.orjpayabledays , b.nhdays , b.payabledays , b.govtdaysalary, b.otwages, b.basicda,  b.basic, b.da, b.hra, b.others, b.ebasic,  b.eda, b.ebasicda, b.ehra, b.eothers, b.payableothrs, b.otamount, b.incentive,  b.govtgross,   b.pfamount, b.esiamount, b.messamount, b.deduction,  b.netamount,   b.bankaccountno , b.bankname , b.ifsccode, b.payperiod , b.fromdate, b.todate,CASE    WHEN a.intime is null  then 'Absent' ELSE  a.intime END intime, CASE    WHEN a.outtime is null  then 'Absent' ELSE  a.outtime END outtime,a.otintime,a.otouttime,  a.shiftcnt*8 totalhrs,ifnull(a.ot,0) as totalothrs,ifnull(a.shiftcnt,0) as present ,b.advance, a.ot, b.othersexp,b.creditdate,(  select COUNT(A.INTIME) totalweeklyoff from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='WO') as totalweeklyoff,(  select sum(aa.totaldays-cc.payabledays)  as totalabsent from( select count(z.empid) totaldays from pldatta z  where z.empid='" + item.SubItems[3].Text + "' and z.compcode='" + item.SubItems[6].Text + "' and z.payperiod='" + item.SubItems[7].Text + "') aa,(   select bb.payabledays from(  SELECT SUM(x.INTIME)  as payabledays from (select y.orjpayabledays as INTIME from hrpaydetails y  where y.idcardno='" + item.SubItems[3].Text + "' and y.united='" + item.SubItems[6].Text + "' and y.payperiod='" + item.SubItems[7].Text + "' union all select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "'and a.intime='WO'  union all  select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='NH'  )x)bb)cc) as totalabsent,b.orjpayabledays as totalpresent  from pldatta a  right join hrpaydetails b on a.payperiod=b.payperiod and a.compcode=b.united and b.idcardno=a.empid Join gtcompmast c on c.compcode=a.compcode and c.gtcompmastid=b.compcode join gtcitymast d on d.gtcitymastid=c.city join gtstatemast e on e.gtstatemastid=c.state join gtcountrymast f on f.gtcountrymastid=c.country where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "';  ";
                                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                                rd.Database.Tables["DataTable1"].SetDataSource(ds2.Tables[0]);
                                rd.Database.Tables["DataTable2"].SetDataSource(ds2.Tables[0]);
                                crystalReportViewer1.ReportSource = rd;
                                rd.ExportToDisk(ExportFormatType.WordForWindows, folderLocation + item.SubItems[4].Text + "-" + item.SubItems[5].Text + " TimeCard.doc");


                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                MessageBox.Show("TimeCard Process Cancelled", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Cursor = Cursors.Default; chkall.Checked = false;
                            }
                        }
                        Cursor = Cursors.Default; chkall.Checked = false;
                    }
                    if (chkall.Checked == false && allip.Items.Count >= 0)
                    {
                        Cursor = Cursors.WaitCursor;
                        foreach (ListViewItem item in allip.Items)
                        {
                            if (item.SubItems.Count > 0)
                            {
                                crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();

                                string sel2 = "select  substr(a.docdate,3,5) as docdate1, CONCAT(date_format(a.docdate,'%d-%m-%Y  '),substr(dayname(a.docdate),1,3)) as docdate,c.compname,c.address, d.cityname,e.statename,f.countryname as country,c.pincode, b.midcard as idcardno , b.empname, b.doj, b.uanno , b.esino , b.fathername, b.united , b.category,  b.department,b.designation,  b.orjpayabledays , b.nhdays , b.payabledays , b.govtdaysalary, b.otwages, b.basicda,  b.basic, b.da, b.hra, b.others, b.ebasic,  b.eda, b.ebasicda, b.ehra, b.eothers, b.payableothrs, b.otamount, b.incentive,  b.govtgross,   b.pfamount, b.esiamount, b.messamount, b.deduction,  b.netamount,   b.bankaccountno , b.bankname , b.ifsccode, b.payperiod , b.fromdate, b.todate,CASE    WHEN a.intime is null  then 'Absent' ELSE  a.intime END intime, CASE    WHEN a.outtime is null  then 'Absent' ELSE  a.outtime END outtime,a.otintime,a.otouttime,  a.shiftcnt*8 totalhrs,ifnull(a.ot,0) as totalothrs,ifnull(a.shiftcnt,0) as present ,b.advance, a.ot, b.othersexp,b.creditdate,(  select COUNT(A.INTIME) totalweeklyoff from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='WO') as totalweeklyoff,(  select sum(aa.totaldays-cc.payabledays)  as totalabsent from( select count(z.empid) totaldays from pldatta z  where z.empid='" + item.SubItems[3].Text + "' and z.compcode='" + item.SubItems[6].Text + "' and z.payperiod='" + item.SubItems[7].Text + "') aa,(   select bb.payabledays from(  SELECT SUM(x.INTIME)  as payabledays from (select y.orjpayabledays as INTIME from hrpaydetails y  where y.idcardno='" + item.SubItems[3].Text + "' and y.united='" + item.SubItems[6].Text + "' and y.payperiod='" + item.SubItems[7].Text + "' union all select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "'and a.intime='WO'  union all  select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='NH'  )x)bb)cc) as totalabsent,b.orjpayabledays as totalpresent  from pldatta a  right join hrpaydetails b on a.payperiod=b.payperiod and a.compcode=b.united and b.idcardno=a.empid Join gtcompmast c on c.compcode=a.compcode and c.gtcompmastid=b.compcode join gtcitymast d on d.gtcitymastid=c.city join gtstatemast e on e.gtstatemastid=c.state join gtcountrymast f on f.gtcountrymastid=c.country where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "';  ";
                                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                                rd.Database.Tables["DataTable1"].SetDataSource(ds2.Tables[0]);
                                rd.Database.Tables["DataTable2"].SetDataSource(ds2.Tables[0]);
                                crystalReportViewer1.ReportSource = rd;
                                rd.ExportToDisk(ExportFormatType.WordForWindows, folderLocation + item.SubItems[4].Text + "-" + item.SubItems[5].Text + " TimeCard.doc");
                            }
                        }

                        foreach (ListViewItem item in listView2.Items)
                        {
                            if (item.Checked == true)
                            {
                                item.Checked = false;
                            }
                        }
                        allip.Items.Clear(); Cursor = Cursors.Default;
                    }
                    //else
                    //{
                    //    MessageBox.Show("pls Select CheckBox ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //    
                    //}
                }
                else
                {

                    Cursor = Cursors.Default; chkall.Checked = false;
                    MessageBox.Show("TimeCard Process Cancelled", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                Cursor = Cursors.Default; chkall.Checked = false;
                MessageBox.Show("Invalid"); this.Dispose();
            }
            Cursor = Cursors.Default; chkall.Checked = false;
        }

        private void checkword_CheckedChanged(object sender, EventArgs e)
        {
            if (checkword.Checked == true)
                GlobalVariables.Pdfs.Text = "Word";
            else
                GlobalVariables.Pdfs.Text = "Pdfs";
        }

        public void ReadOnlys()
        {
            
        }

        public void Imports()
        {

        }

        public void Pdfs()
        {
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                DialogResult result = MessageBox.Show("Do You want to Download??", "DownLoad", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {

                    string folderLocation = "d:\\TimeCardDownLoad\\";
                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }
                    
                    if (chkall.Checked == true)
                    {
                        Cursor = Cursors.WaitCursor;
                        progressBar2.Minimum = 0;
                        progressBar2.Maximum = allip.Items.Count;
                        foreach (ListViewItem item in listView2.Items)
                        {
                            if (item.SubItems.Count > 0)
                            {
                                crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();
                                string sel2 = "select c.tpr as category,c.intime as from1,c.outtime as to1,c.lunchfrom as cityname,c.lunchto as statename,c.breakfrom as country,c.breakto as pincode, substr(a.docdate,3,5) as docdate1, CONCAT(date_format(a.docdate,'%d-%m-%Y  '),substr(dayname(a.docdate),1,3)) as docdate,c.compname,c.address, d.cityname,e.statename,f.countryname as country,c.pincode, b.midcard as idcardno , b.empname,DATE_FORMAT(b.doj,'%d-%m %Y') as doj, b.uanno , b.esino , b.fathername, b.united , b.category,  b.department,b.designation,  b.orjpayabledays , b.nhdays , b.payabledays , b.govtdaysalary, b.otwages, b.basicda,  b.basic, b.da, b.hra, b.others, b.ebasic,  b.eda, b.ebasicda, b.ehra, b.eothers, b.payableothrs, b.otamount, b.incentive,  b.govtgross,   b.pfamount, b.esiamount, b.messamount, b.deduction,  b.netamount,   b.bankaccountno , b.bankname , b.ifsccode, b.payperiod , b.fromdate, b.todate,CASE    WHEN a.intime is null  then 'Absent' ELSE  a.intime END intime, CASE    WHEN a.outtime is null  then 'Absent' ELSE  a.outtime END outtime,a.otintime,a.otouttime,  a.shiftcnt*8 totalhrs,ifnull(a.ot,0) as totalothrs,ifnull(a.shiftcnt,0) as present ,b.advance, a.ot, b.othersexp,b.creditdate,(  select COUNT(A.INTIME) totalweeklyoff from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='WO') as totalweeklyoff,(  select sum(aa.totaldays-cc.payabledays)  as totalabsent from( select count(z.empid) totaldays from pldatta z  where z.empid='" + item.SubItems[3].Text + "' and z.compcode='" + item.SubItems[6].Text + "' and z.payperiod='" + item.SubItems[7].Text + "') aa,(   select bb.payabledays from(  SELECT SUM(x.INTIME)  as payabledays from (select y.orjpayabledays as INTIME from hrpaydetails y  where y.idcardno='" + item.SubItems[3].Text + "' and y.united='" + item.SubItems[6].Text + "' and y.payperiod='" + item.SubItems[7].Text + "' union all select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "'and a.intime='WO'  union all  select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='NH'  )x)bb)cc) as totalabsent,b.orjpayabledays as totalpresent  from pldatta a  right join hrpaydetails b on a.payperiod=b.payperiod and a.compcode=b.united and b.idcardno=a.empid Join gtcompmast c on c.compcode=a.compcode and c.gtcompmastid=b.compcode join gtcitymast d on d.gtcitymastid=c.city join gtstatemast e on e.gtstatemastid=c.state join gtcountrymast f on f.gtcountrymastid=c.country where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' order by 1;  ";
                                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                                rd.Database.Tables["DataTable1"].SetDataSource(ds2.Tables[0]);
                                rd.Database.Tables["DataTable2"].SetDataSource(ds2.Tables[0]);

                                if (checkword.Checked == false)
                                {
                                    rd.ExportToDisk(ExportFormatType.PortableDocFormat, folderLocation + item.SubItems[4].Text + "-" + item.SubItems[5].Text + " TimeCard.pdf");
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(allip.Items.Count)) * (item.Index);
                                    lblprogross4.Text = " Downloading.... : " + (per).ToString("N0") + " %" + " IDCardNo:   " + item.SubItems[4].Text + "          Count  :" + item.Index;
                                    lblprogross4.Refresh();
                                    progressBar2.Value = item.Index + 1;
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.WordForWindows, folderLocation + item.SubItems[4].Text + "-" + item.SubItems[5].Text + " TimeCard.doc");
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(allip.Items.Count)) * (item.Index);
                                    lblprogross4.Text = " Downloading.... : " + (per).ToString("N0") + " %" + " IDCardNo:   " + item.SubItems[4].Text + "           Count  :" + item.Index;
                                    lblprogross4.Refresh();
                                    progressBar2.Value = item.Index + 1;
                                }

                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                MessageBox.Show("TimeCard Process Cancelled", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Cursor = Cursors.Default; chkall.Checked = false; checkword.Checked = false;
                            }
                        }
                        Cursor = Cursors.Default; chkall.Checked = false;
                    }
                    if (chkall.Checked == false && allip.Items.Count > 0)
                    {
                        Cursor = Cursors.WaitCursor; progressBar2.Minimum = 0;
                        progressBar2.Maximum = allip.Items.Count;
                        foreach (ListViewItem item in allip.Items)
                        {
                            if (item.SubItems.Count > 0)
                            {
                                crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();

                                string sel2 = "select c.tpr as category,c.intime as from1,c.outtime as to1,c.lunchfrom as cityname,c.lunchto as statename,c.breakfrom as country,c.breakto as pincode,substr(a.docdate,3,5) as docdate1, CONCAT(date_format(a.docdate,'%d-%m-%Y  '),substr(dayname(a.docdate),1,3)) as docdate,c.compname,c.address, d.cityname,e.statename,f.countryname as country,c.pincode, b.midcard as idcardno , b.empname, b.doj, b.uanno , b.esino , b.fathername, b.united , b.category,  b.department,b.designation,  b.orjpayabledays , b.nhdays , b.payabledays , b.govtdaysalary, b.otwages, b.basicda,  b.basic, b.da, b.hra, b.others, b.ebasic,  b.eda, b.ebasicda, b.ehra, b.eothers, b.payableothrs, b.otamount, b.incentive,  b.govtgross,   b.pfamount, b.esiamount, b.messamount, b.deduction,  b.netamount,   b.bankaccountno , b.bankname , b.ifsccode, b.payperiod , b.fromdate, b.todate,CASE    WHEN a.intime is null  then 'Absent' ELSE  a.intime END intime, CASE    WHEN a.outtime is null  then 'Absent' ELSE  a.outtime END outtime,a.otintime,a.otouttime,  a.shiftcnt*8 totalhrs,ifnull(a.ot,0) as totalothrs,ifnull(a.shiftcnt,0) as present ,b.advance, a.ot, b.othersexp,b.creditdate,(  select COUNT(A.INTIME) totalweeklyoff from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='WO') as totalweeklyoff,(  select sum(aa.totaldays-cc.payabledays)  as totalabsent from( select count(z.empid) totaldays from pldatta z  where z.empid='" + item.SubItems[3].Text + "' and z.compcode='" + item.SubItems[6].Text + "' and z.payperiod='" + item.SubItems[7].Text + "') aa,(   select bb.payabledays from(  SELECT SUM(x.INTIME)  as payabledays from (select y.orjpayabledays as INTIME from hrpaydetails y  where y.idcardno='" + item.SubItems[3].Text + "' and y.united='" + item.SubItems[6].Text + "' and y.payperiod='" + item.SubItems[7].Text + "' union all select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "'and a.intime='WO'  union all  select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + item.SubItems[3].Text + "' and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' and a.intime='NH'  )x)bb)cc) as totalabsent,b.orjpayabledays as totalpresent  from pldatta a  right join hrpaydetails b on a.payperiod=b.payperiod and a.compcode=b.united and b.idcardno=a.empid Join gtcompmast c on c.compcode=a.compcode and c.gtcompmastid=b.compcode join gtcitymast d on d.gtcitymastid=c.city join gtstatemast e on e.gtstatemastid=c.state join gtcountrymast f on f.gtcountrymastid=c.country where a.empid='" + item.SubItems[3].Text + "'  and a.compcode='" + item.SubItems[6].Text + "' and a.payperiod='" + item.SubItems[7].Text + "' order by 1;  ";
                                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                                rd.Database.Tables["DataTable1"].SetDataSource(ds2.Tables[0]);
                                rd.Database.Tables["DataTable2"].SetDataSource(ds2.Tables[0]);
                                if (checkword.Checked == false)
                                {
                                    rd.ExportToDisk(ExportFormatType.PortableDocFormat, folderLocation + item.SubItems[4].Text + "-" + item.SubItems[5].Text + " TimeCard.pdf");
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(allip.Items.Count)) * (item.Index);
                                    lblprogross4.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %" + " IDCardNo :  " + item.SubItems[4].Text + "    Count  :" + item.Index;
                                    lblprogross4.Refresh();
                                    progressBar2.Value = item.Index + 1;

                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.WordForWindows, folderLocation + item.SubItems[4].Text + "-" + item.SubItems[5].Text + " TimeCard.doc");
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(allip.Items.Count)) * (item.Index);
                                    lblprogross4.Text = " Downloading.... : " + (per).ToString("N0") + " %" + " IDCardNo:   " + item.SubItems[4].Text + "        Count  :" + item.Index;
                                    lblprogross4.Refresh();
                                    progressBar2.Value = item.Index + 1;
                                }

                            }
                        }

                        foreach (ListViewItem item in listView2.Items)
                        {
                            if (item.Checked == true)
                            {
                                item.Checked = false;
                            }
                        }
                        allip.Items.Clear(); Cursor = Cursors.Default; checkword.Checked = false; progressBar2.Value = 0; lblprogross4.Text = "";
                    }                   
                }
                else
                {

                    Cursor = Cursors.Default; chkall.Checked = false;
                    MessageBox.Show("TimeCard Process Cancelled", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                Cursor = Cursors.Default; chkall.Checked = false; checkword.Checked = false;
                MessageBox.Show("Invalid"); this.Dispose();
            }
            Cursor = Cursors.Default; chkall.Checked = false; checkword.Checked = false; txtsearch1.Text = "";
        }

        private void combounitsearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combounitsearch.SelectedValue) > 0)
                {
                    Cursor = Cursors.WaitCursor;
                    string sel1 = "select '------' as  payperiod from dual union all select distinct a.payperiod from hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid  where b.compcode='" + combounitsearch.Text + "'";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];
                    comboperiodsearch.DisplayMember = dt1.Rows[0]["payperiod"].ToString();
                    comboperiodsearch.ValueMember = dt1.Rows[0]["payperiod"].ToString();
                    comboperiodsearch.DataSource = dt1; Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void comboperiodsearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combounitsearch.SelectedValue) > 0 && comboperiodsearch.Text != "")
                {
                    listView1.Items.Clear(); listfilter2.Items.Clear(); txtsearch1.Text = ""; Cursor = Cursors.WaitCursor;
                    string sel1 = "select  a.hrpaydetailsid,a.docid ,b.compcode, a.idcardno , a.empname,date_format(a.doj,'%d-%m-%Y') as doj,a.fathername,a.department from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid   where a.payperiod='" + comboperiodsearch.Text + "' and a.united='" + combounitsearch.Text + "' order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRPayDetails");
                    DataTable dt = ds.Tables["HRPayDetails"];
                    if (dt.Rows.Count > 0)
                    {
                        int j = 1;
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list1 = new ListViewItem();
                            list1.SubItems.Add(j.ToString());
                            list1.SubItems.Add(myRow["hrpaydetailsid"].ToString());
                            list1.SubItems.Add(myRow["idcardno"].ToString());
                            list1.SubItems.Add(myRow["empname"].ToString());
                            this.listfilter2.Items.Add((ListViewItem)list1.Clone());
                            if (j % 2 == 0)
                            {
                                list1.BackColor = Color.WhiteSmoke;
                            }
                            else
                            {
                                list1.BackColor = Color.White;
                            }
                            listView1.Items.Add(list1);
                            j++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }
                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        private void compcodeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            companyload();
        }

        private void listView2_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                progressBar2.Minimum = 0;
                if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                {
                    ListViewItem it2 = new ListViewItem();
                    if (e.Item.Checked == true)
                    {
                        e.Item.SubItems[0].Text = "T";
                        e.Item.BackColor = System.Drawing.SystemColors.MenuHighlight;
                        e.Item.ForeColor = Color.White;
                        it2.SubItems.Add(e.Item.SubItems[1].Text);
                        it2.SubItems.Add(e.Item.SubItems[2].Text);
                        it2.SubItems.Add(e.Item.SubItems[3].Text);
                        it2.SubItems.Add(e.Item.SubItems[4].Text);
                        it2.SubItems.Add(e.Item.SubItems[5].Text);
                        it2.SubItems.Add(e.Item.SubItems[6].Text);
                        it2.SubItems.Add(e.Item.SubItems[7].Text);
                        allip.Items.Add(it2);
                        Cursor = Cursors.Default;
                        progressBar2.Maximum = allip.Items.Count;
                    }
                    if (e.Item.Checked == false && e.Item.SubItems[0].Text == "T")
                    {

                        e.Item.SubItems[0].Text = "F"; e.Item.BackColor = System.Drawing.SystemColors.ControlLightLight;
                        e.Item.ForeColor = Color.Black;
                        for (int c = 0; c < allip.Items.Count; c++)
                        {

                            if (e.Item.SubItems[3].Text == allip.Items[c].SubItems[3].Text)
                            {
                                allip.Items[c].Remove();
                                e.Item.SubItems[0].Text = "";
                                c--;
                                progressBar2.Maximum = allip.Items.Count;
                            }
                        }

                    }

                }
                else
                {
                    MessageBox.Show("Invalid"); this.Dispose();
                }
            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

        private void contextMenuStrip5_Opening(object sender, CancelEventArgs e)
        {
            companyload();
        }

        public void ChangePasswords()
        {
            throw new NotImplementedException();
        }

        public void DownLoads()
        {
            
            try
            {
                do
                {
                    int i = 0;
                    for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }

                }
                while (dataGridView1.Rows.Count >= 1);
                if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                {
                    if (Convert.ToInt64(combocompcode.SelectedValue) > 0 && comboPayPeriod.Text != "" && comboseq.Text != "" && comboweekmonth.Text != "")
                    {
                        var confirmation = MessageBox.Show("Do You want Download this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (confirmation == DialogResult.Yes)
                        {
                            Utility.Connect();
                            Cursor.Current = Cursors.WaitCursor;
                            string sel = "call timecardprocess('" + comboPayPeriod.Text + "','" + combocompcode.Text + "','" + comboseq.Text + "','" + comboweekmonth.Text + "');";
                            Utility.ExecuteNonQuery(sel);

                            string sel1 = "select a.idcardno as IDCardNo ,a.empname as EmployeeName , CONCAT(date_format(b.docdate,'%d-%m-%Y  '),substr(dayname(b.docdate),1,3)) as Date, b.intime as InTime,b.outtime as OutTime,b.otintime as OTInTime,b.otouttime as OTOutTime,b.shiftcnt*8 TotalHours ,b.ot as TotalOTHouurs,b.shiftcnt as Present from hrpaydetails a join pldatta b on a.payperiod=b.payperiod and a.idcardno=b.empid where a.payperiod='" + comboPayPeriod.Text + "' and a.united='" + combocompcode.Text + "' order by 1 ;";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "pldatta");
                            DataTable dt1 = ds1.Tables["pldatta"];
                            if (dt1 != null)
                            {
                                progressBar1.Minimum = 0;
                                progressBar1.Maximum = dt1.Rows.Count;
                                dataGridView1.DataSource = dt1;
                                for (int i = 0; i < dt1.Rows.Count; i++)
                                {
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(dt1.Rows.Count)) * (i + 1);
                                    lblprogress1.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %" + dt1.Rows[i]["IDCardNo"].ToString();
                                    lblprogress1.Refresh();
                                    progressBar1.Value = i + 1;
                                }
                                lblprogress1.Text = ""; progressBar1.Value = 0;
                                Cursor = Cursors.Default;
                                lbltoalrecord.Text = "Total Record  :" + dataGridView1.Rows.Count.ToString();
                                Cursor = Cursors.Default; /*comboPayPeriod1_SelectedIndexChanged(sender, e);*/
                                MessageBox.Show("TimeCard Process Successfully Completed...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  
                            }
                            else
                            {
                                MessageBox.Show("TimeCard Process Cancelled", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  

                            }
                        }
                        else
                        {
                            MessageBox.Show("TimeCard Process Cancelled", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please choose .Compcode , PayPeriod & Weekly Field.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  
                    }
                }
                else
                {
                    MessageBox.Show("InvalidDate.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("TimeCard: " + ex.ToString(), " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error); Cursor = Cursors.Default;
            }
        }
        private void comboPayPeriod_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
               
                if (checkall.Checked == true && Convert.ToInt64(combocompcode.SelectedValue) > 0 && comboPayPeriod.Text != "")
                {

                    string sel1 = "select  B.* from  gtcompmast a join pldatta b on a.compcode=b.compcode  where  a.compcode='" + combocompcode.Text + "' AND b.payperiod='" + comboPayPeriod.Text + "';";

                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "pldatta");
                    DataTable dt1 = ds1.Tables["pldatta"];
                    dataGridView1.DataSource = dt1;
                    lbltoalrecord.Text = "" + comboPayPeriod.Text + " Total Record's : " + dataGridView1.Rows.Count;
                    lbltoalrecord.Text = "" + comboPayPeriod.Text + " Total Record's : " + dataGridView1.Rows.Count;
                    lbltotalcount.Refresh();
                    lbltotalcount.Text = "Count :" + dt1.Rows.Count.ToString();
                }
                if (Convert.ToInt64(combocompcode.SelectedValue) > 0 && comboPayPeriod.Text != "------" && checkall.Checked == false)
                {

                    listView2.Items.Clear(); listfilter2.Items.Clear(); txtsearch1.Text = "";
                    string sel1 = "select  a.hrpaydetailsid,a.docid ,b.compcode, a.midcard , a.empname, date_format(a.doj,'%d-%m-%Y') as doj,date_format(a.dol,'%d-%m-%Y') as dol,a.fathername,a.department from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid   where a.payperiod='" + comboPayPeriod.Text + "' and a.united='" + combocompcode.Text + "' order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRPayDetails");
                    DataTable dt = ds.Tables["HRPayDetails"];
                    dataGridView1.DataSource = dt;
                    lbltotalcount.Refresh();
                    lbltotalcount.Text = "Count :" + dt.Rows.Count.ToString();
                    lbltoalrecord.Text = "" + comboPayPeriod.Text + " Total Record's : " + dataGridView1.Rows.Count;
                    lbltoalrecord.Text = "" + comboPayPeriod.Text + " Total Record's : " + dataGridView1.Rows.Count;
                    //}

                    string sel2 = "select a.idseq from   idseqtable a ;";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "idseqtable");
                    DataTable dt2 = ds2.Tables["idseqtable"];
                    comboseq.DisplayMember = "idseq";
                    comboseq.ValueMember = "idseq";
                    comboseq.DataSource = dt2;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            companyload1(); companyloadsearch(); string sel2 = "select a.idseq from   idseqtable a ;";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "idseqtable");
            DataTable dt2 = ds2.Tables["idseqtable"];
            comboseq.DisplayMember = "idseq";
            comboseq.ValueMember = "idseq";
            comboseq.DataSource = dt2;
        }
        private void listView2_ItemActivate(object sender, EventArgs e)
        {
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();
                string sel2 = "select  c.tpr as category, c.intime as from1,c.outtime as to1,c.lunchfrom as cityname,c.lunchto as statename,c.breakfrom as country,c.breakto as pincode,substr(a.docdate,3,5) as docdate1,CONCAT(date_format(a.docdate,'%d-%m-%Y  '),substr(dayname(a.docdate),1,3)) as docdate,c.compname,c.address, d.cityname,e.statename,f.countryname as country,c.pincode, b.midcard as idcardno , b.empname, DATE_FORMAT(b.doj,'%d-%m %Y') as doj, b.uanno , b.esino , b.fathername, b.united , b.category,  b.department,b.designation,  b.orjpayabledays , b.nhdays , b.payabledays , b.govtdaysalary, b.otwages, b.basicda,  b.basic, b.da, b.hra, b.others, b.ebasic,  b.eda, b.ebasicda, b.ehra, b.eothers, b.payableothrs, b.otamount, b.incentive,  b.govtgross,   b.pfamount, b.esiamount, b.messamount, b.deduction,  b.netamount,   b.bankaccountno , b.bankname , b.ifsccode, b.payperiod , b.fromdate, b.todate,CASE    WHEN a.intime is null  then 'Absent' ELSE  a.intime END intime, CASE    WHEN a.outtime is null  then 'Absent' ELSE  a.outtime END outtime,a.otintime,a.otouttime,  a.shiftcnt*8 totalhrs,ifnull(a.ot,0) as totalothrs,ifnull(a.shiftcnt,0) as present ,b.advance, a.ot, b.othersexp,b.creditdate,(  select COUNT(A.INTIME) totalweeklyoff from pldatta a where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "' and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' and a.intime='WO') as totalweeklyoff,(  select sum(aa.totaldays-cc.payabledays)  as totalabsent from( select count(z.empid) totaldays from pldatta z  where z.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "' and z.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and z.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "') aa,(   select bb.payabledays from(  SELECT SUM(x.INTIME)  as payabledays from (select y.orjpayabledays as INTIME from hrpaydetails y  where y.idcardno='" + listView2.SelectedItems[0].SubItems[3].Text + "' and y.united='" + listView2.SelectedItems[0].SubItems[6].Text + "' and y.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' union all select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "'  and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "'and a.intime='WO'  union all  select COUNT(A.INTIME) INTIME from pldatta a where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "' and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' and a.intime='NH'  )x)bb)cc) as totalabsent,b.orjpayabledays as totalpresent  from pldatta a  right join hrpaydetails b on a.payperiod=b.payperiod and a.compcode=b.united and b.idcardno=a.empid Join gtcompmast c on c.compcode=a.compcode and c.gtcompmastid=b.compcode join gtcitymast d on d.gtcitymastid=c.city join gtstatemast e on e.gtstatemastid=c.state join gtcountrymast f on f.gtcountrymastid=c.country where a.empid='" + listView2.SelectedItems[0].SubItems[3].Text + "'  and a.compcode='" + listView2.SelectedItems[0].SubItems[6].Text + "' and a.payperiod='" + listView2.SelectedItems[0].SubItems[7].Text + "' order by 1;  ";

                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                rd.Database.Tables["DataTable1"].SetDataSource(ds2.Tables[0]);
                rd.Database.Tables["DataTable2"].SetDataSource(ds2.Tables[0]);
                crystalReportViewer1.ReportSource = rd;
            }
            else
            {
                this.Dispose();
            }
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();

                string sel1 = "select  a.hrpaydetailsid,a.docid ,b.compcode, a.idcardno , a.empname, date_format(a.doj,'%d-%m-%Y') as doj,a.fathername,a.department from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + Class.Users.HCompcode + "' order by a.hrpaydetailsid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRPayDetails");
                DataTable dt = ds.Tables["HRPayDetails"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["hrpaydetailsid"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["idcardno"].ToString());
                        list.SubItems.Add(myRow["empname"].ToString());
                        list.SubItems.Add(myRow["doj"].ToString());
                        list.SubItems.Add(myRow["fathername"].ToString());
                        list.SubItems.Add(myRow["department"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listView1.Items.Add(list);

                        i++;
                    }

                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void TimeCardProcessEntry_Load(object sender, EventArgs e)
        {

            // fromdatedateTimePicker1.Value = System.DateTime.Now.AddDays(-1);
            //todatedateTimePicker2.Value = System.DateTime.Now.AddDays(30);
            tabControl1.SelectTab(tabPage1);
            txtsearch1.Select();
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {

        }

        void idcardload(string c,string p)
        {
            if (p != "------")
            {
                string sel = "select distinct '0' AS  EMPID,'' IDCARDNO from dual union all select distinct '1' AS EMPID,'ALL' IDCARDNO from dual union all select distinct  B.hrpaydetailsID  AS EMPID,B.midcard as idcardno  from pldatta a JOIN hrpaydetails B ON A.EMPID=B.idcardno  where a.payperiod = '" + p + "' and a.compcode = '" + c + "' order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "pldatta");
                DataTable dt = ds.Tables["pldatta"];
                if (dt != null)
                {
                    comboidcardmuster.ValueMember = "EMPID";
                    comboidcardmuster.DisplayMember = "IDCARDNO";
                    comboidcardmuster.DataSource = dt;
                }
            }
        }
        CommonClass com = new CommonClass();
        private void combopayperiodmuster_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            idcardload(combocompcodemuster.Text, combopayperiodmuster.Text);
            Cursor.Current = Cursors.Default;
        }

        private void comboidcardmuster_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopayperiodmuster.Text != "------" && combocompcodemuster.Text != "------" && comboidcardmuster.Text != "------")
            {
                try
                {
                   
                    DataTable dt1; Cursor.Current = Cursors.WaitCursor;
                    if (comboidcardmuster.Text == "ALL")
                    {
                        dt1 = com.select("select distinct b.midcard AS IDCARDNO from pldatta a JOIN hrpaydetails B ON A.EMPID=B.idcardno where a.payperiod = '" + combopayperiodmuster.Text + "' and a.compcode = '" + combocompcodemuster.Text + "'   order by 1", "pldatta");//and A.empid='" + comboidcardmuster.Text + "'
                    }
                    else
                    {
                        dt1 = com.select("select distinct b.midcard AS IDCARDNO from pldatta a JOIN hrpaydetails B ON A.EMPID=B.idcardno where a.payperiod = '" + combopayperiodmuster.Text + "' and a.compcode = '" + combocompcodemuster.Text + "'   and b.midcard='" + comboidcardmuster.Text + "' order by 1", "pldatta");

                    }
                    int i = 0; DataTable dts = new DataTable();
                    int j = 0, k = 0, l = 0; int cnt = 0; int cnt1 = 0;
                    int dat = 0; 
                    if (dt1.Rows.Count > 0)
                    {
                        dts.Rows.Clear();
                        for (i = 0; i < dt1.Rows.Count; i++)
                        {
                            DataTable dt2 = com.select("select distinct b.midcard AS idcardno,a.deptname AS department,a.designation,a.docdate,a.muster,B.empname,b.doj,c.compname,'' as payperiod,'' idcard0,'' idcard1,''idcard2,'' idcard3,'' idcard4,'' idcard5,'' idcard6,'' idcard7,''idcard8,'' idcard9,'' idcard10,'' idcard11,'' idcard12,'' idcard13,'' idcard14,'' idcard15,'' idcard16,'' idcard17,'' idcard18,'' idcard19,'' idcard20,'' idcard21,'' idcard22,'' idcard23,'' idcard24,'' idcard25,'' idcard26,'' idcard27,'' idcard28,'' idcard29,'' idcard30,'' idcard31,'' date0,'' date1,''date2,'' date3,'' date4,'' date5,'' date6,'' date7,''date8,'' date9,'' date10,'' date11,'' date12,'' date13,'' date14,'' date15,'' date16,'' date17,'' date18,'' date19,'' date20,'' date21,'' date22,'' date23,'' date24,'' date25,'' date26,'' date27,'' date28,'' date29,'' date30,'' in0,'' in1,''in2,'' in3,'' in4,'' in5,'' in6,'' in7,''in8,'' in9,'' in10,'' in11,'' in12,'' in13,'' in14,'' in15,'' in16,'' in17,'' in18,'' in19,'' in20,'' in21,'' in22,'' in23,'' in24,'' in25,'' in26,'' in27,'' in28,'' in29,'' in30,'' in31 from pldatta A join hrpaydetails B on  B.idcardno=A.empid  JOIN gtcompmast C ON C.gtcompmastid=B.compcode   where  A.compcode='" + combocompcodemuster.Text + "' AND A.payperiod='" + combopayperiodmuster.Text + "' and b.midcard='" + dt1.Rows[i]["IDCARDNO"].ToString() + "'  order by 4", "pldatta");

                            if (dts.Rows.Count <= 0)
                            {
                                dts = dt2.Clone();

                            }

                            if (dt2.Rows.Count > 0)
                            {
                                dts.Rows.Add();

                                for (j = 0; j < dt2.Rows.Count; j++)
                                {

                                    dts.Rows[i]["idcard" + j] = dt2.Rows[j]["docdate"].ToString().Substring(0, 2);
                                    dts.Rows[i]["idcardno"] = dt2.Rows[j]["idcardno"].ToString();
                                    dts.Rows[i]["date" + j] = dt2.Rows[j]["docdate"].ToString().Substring(0, 2);
                                    dts.Rows[i]["in" + j] = dt2.Rows[j]["muster"].ToString();
                                    dts.Rows[i]["empname"] = dt2.Rows[j]["empname"].ToString();
                                    dts.Rows[i]["doj"] = dt2.Rows[j]["doj"].ToString().Substring(0, 10);
                                    dts.Rows[i]["department"] = dt2.Rows[j]["department"].ToString();
                                    dts.Rows[i]["designation"] = dt2.Rows[j]["designation"].ToString();
                                    dts.Rows[i]["payperiod"] = combopayperiodmuster.Text;
                                    dts.Rows[i]["compname"] = dt2.Rows[j]["compname"].ToString();
                                }

                            }
                        }
                        musterrd3.Database.Tables["DataTable1"].SetDataSource(dts);
                        crystalReportViewer2.ReportSource = null;
                        crystalReportViewer2.ReportSource = musterrd3;
                        crystalReportViewer2.Refresh(); crystalReportViewer2.Zoom(100);
                        Cursor.Current = Cursors.Default;
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); Cursor.Current = Cursors.Default; }

                Cursor.Current = Cursors.Default;
            }
        }
        Pinnacle.Report.SKL.MusterReport musterrd3 = new Pinnacle.Report.SKL.MusterReport();

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                DialogResult result = MessageBox.Show("Do You want to Download??", "DownLoad", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    string folderLocation = "d:\\TimeCardDownLoad\\";
                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }

                    musterrd3.ExportToDisk(ExportFormatType.WordForWindows, folderLocation + combopayperiodmuster.Text + " MusterRoll.doc");
                }
            }
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkexcel_CheckedChanged(object sender, EventArgs e)
        {
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                DialogResult result = MessageBox.Show("Do You want to Download??", "DownLoad", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    string folderLocation = "d:\\TimeCardDownLoad\\";
                    if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }
                    musterrd3.ExportToDisk(ExportFormatType.ExcelWorkbook, folderLocation + combopayperiodmuster.Text + " MusterRoll.xls");
                }
            }
        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            combocompcode_SelectedIndexChanged(sender,e);

        }

        private void combocompcodemuster_SelectedIndexChanged(object sender, EventArgs e)
        {
         //   combocompcode_SelectedIndexChanged(sender, e);
        }

        private void contextMenuStrip3_Opening(object sender, CancelEventArgs e)
        {

        }
    }
}
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

namespace Pinnacle.Master.Bank
{
    public partial class PartyMaster : Form,ToolStripAccess
    {
        private static PartyMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        public static PartyMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PartyMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public PartyMaster()
        {
            InitializeComponent();          
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        void empty()
        {
            this.BackColor = Class.Users.BackColors;       
            this.Font = Class.Users.FontName;
            butheader.BackColor = Class.Users.BackColors;
           Class.Users.UserTime = 0;

            txtcontactno.Text = "";
            progressBar3.Value = 0; lblprogress3.Text = "";
            do
            {
                int i = 0;txtcode.Text = "";
                for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }

            }
            while (dataGridView1.Rows.Count > 0);
            txtPartycode.Text = ""; txtpartyname.Text = ""; txtaddress.Text = ""; 
          txtpincode.Text = ""; txtpanno.Text = ""; txttinno.Text = ""; txtgstno.Text = "";
            txtgstdate.Text = ""; txtphoneno.Text = ""; txtemail.Text = ""; txtaccholdername.Text = "";
            txtwebsite.Text = ""; txtcontact.Text = "";checkactive.Checked = true; txtaccno.Text = "";
            txtpartyid.Text = "";txtpartyid1.Text = "";
            combocity.SelectedIndex = -1;
            combostate.SelectedIndex = -1; combostate.Text = "";
            combocountry.SelectedIndex = -1;
            combobank.SelectedIndex = -1;
            combobranch.SelectedIndex = -1;
            comboIFSCE.SelectedIndex = -1; combocountry.Text = "";
            combobank.Text = "";
            combobranch.Text = "";
            comboIFSCE.Text = "";txtPartycode.Focus();
           
        }      
        private void PartyMaster_Load(object sender, EventArgs e)
        {
            
        }
        protected override bool ProcessCmdKey(ref Message message, Keys keys)
        {
            switch (keys)
            {
                case Keys.N | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    News();
                    return true; // signal that we've processed this key
                case Keys.S | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Saves();
                    return true; // signal that we've processed this key
                //case Keys.E | Keys.Control:
                //    // ... Process Shift+Ctrl+Alt+B ...
                //    updating = true;
                //    adding = false;
                //    EnableText();
                //    return true; // signal that we've processed this key
                case Keys.D | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Deletes();
                    return true; // signal that we've processed this key
            }
            // run base implementation
            return base.ProcessCmdKey(ref message, keys);
        }

        //public void compid(string s)
        //{
        //    string sel = "select max(a.asptblpartymasid1) as id  from  asptblpartymas a  where a.PartyCode='" + s.ToUpper() + "' ;";
        //    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
        //    DataTable dt = ds.Tables["asptblpartymas"];
        //    Int64 partycount = 1;
        //    if (dt.Rows.Count >= 1)
        //    {
        //        txtpartyid1.Text = "";
        //        Int64 cc = Convert.ToInt64("0" + dt.Rows[0]["id"].ToString()) + partycount;
        //        txtpartyid1.Text = cc.ToString();
        //    }
        //    else
        //    {
        //        txtpartyid1.Text = "1";
        //    }
        //}
        DataTable dt = new DataTable();
        public DataTable ifce(string s, string ss)
        {
            if (ss != "")
            {
                sb.Clear();
                sb.Append("select distinct  a.ifsc,a.branch, a.asptblifscmasid,b.asptblbanmasid from asptblifscmas a join asptblbanmas b on b.asptblbanmasid=a.bankname where a.active='T' and b.bankname='" + s.Trim() + "' and  a.branch='" + ss.Trim() + "' ");
                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblifscmas");
                dt = ds.Tables["asptblifscmas"];
            }
            return dt;
        }
        public DataTable branch(string s)
        {
            if (s != "")
            {
                sb.Clear();
                sb.Append("select distinct a.asptblifscmasid, a.branch from  asptblifscmas a join asptblbanmas b on b.asptblbanmasid=a.bankname   where a.active='T' and b.bankname='" + s + "' ");
                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblifscmas");
                dt = ds.Tables["asptblifscmas"];
            }
            return dt;
        }
        public void bankname()
        {
            sb.Clear();
            sb.Append("select distinct  a.asptblbanmasid,a.bankname from asptblbanmas a join asptblifscmas b on a.asptblbanmasid=b.bankname where a.active='T' ");
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblbanmas");
            DataTable dt = ds.Tables["asptblbanmas"];
            combobank.DataSource = dt;
            combobank.DisplayMember = "bankname";
            combobank.ValueMember = "asptblbanmasid";

        }


        public void GridLoad()
        {
           
            try
            {


                Class.Users.SearchQuery = "select a.asptblpartymasid as ID,a.PartyName,a.accno as AccountNumber,E.bankname AS BankName,f.branch as Branch,f.ifsc as IFSC, a.accountholdername as  AccountHolderName,a.active  from asptblpartymas a join gtcitymast b on a.city = b.gtcitymastid  join gtstatemast c on a.state = c.gtstatemastid join gtcountrymast d on a.country = d.gtcountrymastid  JOIN asptblbanmas E ON E.asptblbanmasid=A.bankname join asptblifscmas f on f.bankname=a.bankname and f.asptblifscmasid=a.branch  order by a.asptblpartymasid desc";
                Class.Users.HideCols = new string[] { "ID" };
                uccListView1.Load_Details();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public DataTable state(Int64 s)
        {
            string sel = "select distinct b.gtstatemastid,b.statename from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where b.active='T' and a.gtcitymastid='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];
            return dt;
        }
        public DataTable city(string s)
        {
            string sel = "select distinct a.gtcitymastid,a.cityname from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where b.active='T' and a.cityname='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];
            return dt;
        }
        public DataTable state(string s)
        {
            string sel = "select distinct  b.gtstatemastid,b.statename from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where b.active='T' and b.statename='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];
            return dt;
        }
        public void  cityLoad()
        {
            string sel = "select distinct  a.gtcitymastid,a.cityname from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcitymast");
            DataTable dt = ds.Tables["gtcitymast"];

            combocity.DisplayMember = "cityname";
            combocity.ValueMember = "gtcitymastid";
            combocity.DataSource = dt;
           
        }
        public DataTable country(Int64 s)
        {
            string sel1 = "SELECT distinct   b.gtcountrymastid,b.countryname from gtstatemast a join gtcountrymast b on a.country=b.gtcountrymastid where a.gtstatemastid='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];

            return dt;
        }

        public DataTable country(string s)
        {
            string sel1 = "SELECT  distinct  b.gtcountrymastid,b.countryname from gtstatemast a join gtcountrymast b on a.country=b.gtcountrymastid where b.countryname='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];

            return dt;
        }
      
        private void PartyMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       


        private void listView1_ItemActivate(object sender, EventArgs e)
        {

        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void Deletes_Click(object sender, EventArgs e)
        {

        }

        public void News()
        {
            empty();
            bankname(); Class.Users.UserTime = 0;
            cityLoad(); GridLoad();
        }

        public void Saves()
        {
            try
            {
                sb.Clear();


                Class.Users.UserTime = 0;

                if (dataGridView1.Rows.Count > 0 && Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                {
                    int cc = 0,i=0;
                    cc = dataGridView1.Rows.Count;
                    if (cc >= 0)
                    {
                        progressBar3.Minimum = 0; progressBar3.Refresh();
                        progressBar3.Maximum = dataGridView1.Rows.Count;

                        for (i = 0; i < cc; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value.ToString() != "")
                            {
                                Models.Party p = new Models.Party();
                                sb.Clear();
                                p.Active = "T";
                                p.partycode = dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper();
                                p.partyname = dataGridView1.Rows[i].Cells[1].Value.ToString().ToUpper();
                                
                                if (dataGridView1.Rows[i].Cells[2].FormattedValue.ToString() != "")
                                {
                                    DataTable dt0 = city(dataGridView1.Rows[i].Cells[2].FormattedValue.ToString().Trim());

                                    p.City = dt0.Rows[0]["gtcitymastid"].ToString();
                                }
                                if (dataGridView1.Rows[i].Cells[3].FormattedValue.ToString() != "")
                                {
                                    DataTable dt0 = state(Convert.ToString(dataGridView1.Rows[i].Cells[3].FormattedValue.ToString().Trim()));

                                    p.State = dt0.Rows[0]["gtstatemastid"].ToString();
                                }
                                if (dataGridView1.Rows[i].Cells[4].FormattedValue.ToString() != "")
                                {
                                    DataTable dt1 = country(dataGridView1.Rows[i].Cells[4].FormattedValue.ToString().Trim());

                                    p.Country = dt1.Rows[0]["gtcountrymastid"].ToString();
                                }
                                p.Address = dataGridView1.Rows[i].Cells[5].Value.ToString().ToUpper();
                                p.PinCode = Convert.ToInt64(dataGridView1.Rows[i].Cells[6].Value.ToString());
                                p.PanNo = dataGridView1.Rows[i].Cells[7].Value.ToString().ToUpper();
                                p.TinNo = dataGridView1.Rows[i].Cells[8].Value.ToString().ToUpper();
                                p.GstNo = dataGridView1.Rows[i].Cells[9].Value.ToString().ToUpper();
                                p.GstDate = dataGridView1.Rows[i].Cells[10].Value.ToString();
                                p.PhoneNo = dataGridView1.Rows[i].Cells[11].Value.ToString();
                                p.Email = dataGridView1.Rows[i].Cells[12].Value.ToString().ToLower();
                                p.WebSite = dataGridView1.Rows[i].Cells[13].Value.ToString().ToLower();
                                p.contactname = dataGridView1.Rows[i].Cells[14].Value.ToString().ToUpper();
                                p.accno = dataGridView1.Rows[i].Cells[15].Value.ToString();
                                if (dataGridView1.Rows[i].Cells[16].Value.ToString() != "")
                                {
                                    //DataTable dt4 = branch(dataGridView1.Rows[i].Cells[16].Value.ToString());
                                    //  p.bankname = dt4.Rows[0]["asptblbanmasid"].ToString();
                                    DataTable dt5 = ifce(dataGridView1.Rows[i].Cells[16].Value.ToString(), dataGridView1.Rows[i].Cells[19].Value.ToString());
                                    p.bankname = dt5.Rows[0]["asptblbanmasid"].ToString();
                                    p.branch = dt5.Rows[0]["asptblifscmasid"].ToString();
                                    p.ifsc = dt5.Rows[0]["asptblifscmasid"].ToString();
                                }

                                p.accountholdername = dataGridView1.Rows[i].Cells[17].Value.ToString().ToUpper();
                                p.CodeNo = dataGridView1.Rows[i].Cells[18].Value.ToString();
                                p.CreatedBy = Convert.ToString(Class.Users.CREATED);
                                p.CreatedOn = Convert.ToString(Class.Users.CREATED);
                                p.ModifiedOn = Convert.ToString(Class.Users.HUserName);
                                p.IpAddress = Class.Users.IPADDRESS;
                               
                                sb.Clear();
                                sb.Append("select a.asptblpartymasid    from  asptblpartymas a    WHERE partycode='" + p.partycode + "'and partyname='" + p.partyname + "' and address='" + p.Address + "'and city='" + p.City + "' and state='" + p.State + "' and country='" + p.Country + "' and pincode='" + p.PinCode + "' and panno='" + p.PanNo + "' and tinno='" + p.TinNo + "' and gstno='" + p.GstNo + "' and gstdate='" + p.GstDate + "' and phoneno='" + p.PhoneNo + "' and email='" + p.Email + "' and website='" + p.WebSite + "' and contactname='" + p.contactname + "' and active='" + p.Active + "' and accno='" + p.accno + "' and ifsc='" + p.ifsc + "' and branch='" + p.branch + "' and bankname='" + p.bankname + "'   and accountholdername='" + p.accountholdername + "'");
                                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblpartymas");
                                DataTable dt = ds.Tables["asptblpartymas"];
                                if (dt.Rows.Count != 0)
                                {
                                    sb.Clear();
                                }
                                else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtpartyid) == 0 || Convert.ToInt32("0" + txtpartyid.Text) == 0)
                                {
                                    sb.Clear();
                                    sb.Append("insert into asptblpartymas(partycode,partyname,address,city,state,country,pincode,panno,tinno,gstno,gstdate,phoneno,email,website,contactname,active,accno,bankname,branch,ifsc,accountholdername,createdon, createdby,modifiedby,ipaddress,codeno)  VALUES('" + p.partycode + "','" + p.partyname + "','" + p.Address + "','" + p.City + "','" + p.State + "','" + p.Country + "','" + p.PinCode + "','" + p.PanNo + "','" + p.TinNo + "','" + p.GstNo + "','" + p.GstDate + "','" + p.PhoneNo + "','" + p.Email + "','" + p.WebSite + "','" + p.contactname + "','" + p.Active + "','" + p.accno + "','" + p.bankname + "','" + p.branch + "','" + p.ifsc + "','" + p.accountholdername + "','" + Convert.ToDateTime(System.DateTime.Now.ToString()).ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "','" + p.CodeNo + "')");
                                    Utility.ExecuteNonQuery(sb.ToString());
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(dataGridView1.Rows.Count)) * (i + 1);
                                    lblprogress3.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %";
                                    label48.Text = " Total Rows : " + i.ToString() + "  " + (per).ToString("N0") + " %" + "Partycode" + p.partycode + " Name : " + p.partyname;
                                    lblprogress3.Refresh(); label48.Refresh();
                                    progressBar3.Value = i + 1;

                                }
                                else
                                {
                                    sb.Clear();
                                    sb.Append("update  asptblpartymas  set partycode='" + p.partycode + "', partyname='" + p.partyname + "' , address='" + p.Address + "', city='" + p.City + "' , state='" + p.State + "' , country='" + p.Country + "' , pincode='" + p.PinCode + "' , panno='" + p.PanNo + "' , tinno='" + p.TinNo + "' , gstno='" + p.GstNo + "' , gstdate='" + p.GstDate + "' , phoneno='" + p.PhoneNo + "' , email='" + p.Email + "' , website='" + p.WebSite + "' , contactname='" + p.contactname + "' , active='" + p.Active + "' , accno='" + p.accno + "' , ifsc='" + p.ifsc + "' , branch='" + p.branch + "' , bankname='" + p.bankname + "'   , accountholdername='" + p.accountholdername + "', createdon ='" + p.CreatedOn + "',createdby='" + Class.Users.HUserName + "',modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "',codeno='" + p.CodeNo + "' where asptblpartymasid='" + txtpartyid.Text + "'");
                                    Utility.ExecuteNonQuery(sb.ToString());
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(dataGridView1.Rows.Count)) * (i + 1);
                                    lblprogress3.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %";
                                    label48.Text = " Total Rows : " + i.ToString() + "  " + (per).ToString("N0") + " %" + "PartyCode" + p.partycode + " Name : " + p.partyname;
                                    lblprogress3.Refresh(); label48.Refresh();

                                    progressBar3.Value = i + 1;

                                }
                            }
                        }
                    }
                    if (i == cc)
                    {
                        MessageBox.Show("Record Saved Successfully ", "Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty(); progressBar3.Value = 0; lblprogress3.Text = "";
                    }
                }
                else
                {
                    if (txtPartycode.Text == "")
                    {
                        MessageBox.Show("'PartyCode  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.txtPartycode.Focus();
                        return;
                    }
                    if (txtpartyname.Text == "")
                    {
                        MessageBox.Show("'PartyName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtpartyname.Focus();
                        return;
                    }

                    if (combocity.Text == "")
                    {
                        MessageBox.Show("'CityName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combocity.Focus();
                        return;
                    }
                    if (combostate.Text == "")
                    {
                        MessageBox.Show("'StateName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combostate.Focus();
                        return;
                    }
                    if (txtaddress.Text == "")
                    {
                        MessageBox.Show("'Address  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtaddress.Focus();
                        return;
                    }
                    if (txtpincode.Text == "")
                    {
                        MessageBox.Show("'PinCode  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtpincode.Focus();
                        return;
                    }
                    if (txtgstno.Text == "")
                    {
                        MessageBox.Show("'GSTNO  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtgstno.Focus();
                        return;
                    }
                    if (txtphoneno.Text == "")
                    {
                        MessageBox.Show("'Phone No  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtphoneno.Focus();
                        return;
                    }
                    if (txtaccholdername.Text == "")
                    {
                        MessageBox.Show("'Account Holder Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtaccholdername.Focus();
                        return;
                    }
                    if (txtaccno.Text == "")
                    {
                        MessageBox.Show("'Account Number  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtaccno.Focus();
                        return;
                    }

                    if (combobank.Text == "")
                    {
                        MessageBox.Show("'Bank Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combobank.Focus();
                        return;
                    }
                    else
                    {
                        string chk = ""; sb.Clear();
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        sb.Append("select a.asptblpartymasid    from  asptblpartymas a    WHERE partycode='" + txtPartycode.Text + "'and partyname='" + txtpartyname.Text + "' and address='" + txtaddress.Text.ToUpper() + "'and city='" + combocity.SelectedValue + "' and state='" + combostate.SelectedValue + "' and country='" + combocountry.SelectedValue + "' and pincode='" + txtpincode.Text + "' and panno='" + txtpanno.Text.ToUpper() + "' and tinno='" + txttinno.Text + "' and gstno='" + txtgstno.Text.ToUpper() + "' and gstdate='" + txtgstdate.Text + "' and phoneno='" + txtphoneno.Text + "' and email='" + txtemail.Text + "' and website='" + txtwebsite.Text + "' and contactname='" + txtcontact.Text + "' and active='" + chk + "' and accno='" + txtaccno.Text + "' and ifsc='" + comboIFSCE.SelectedValue + "' and branch='" + combobranch.SelectedValue + "' and bankname='" + combobank.SelectedValue + "'  and active='" + chk + "' and accountholdername='" + txtaccholdername.Text.ToUpper().Trim() + "' and contactno='" + txtcontactno.Text + "'");
                        DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblpartymas");
                        DataTable dt = ds.Tables["asptblpartymas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + comboIFSCE.SelectedValue, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + comboIFSCE.SelectedValue) == 0 || Convert.ToInt32("0" + txtpartyid.Text) == 0)
                        {
                            sb.Clear();
                            sb.Append("insert into asptblpartymas(partycode,partyname,address,city,state,country,pincode,panno,tinno,gstno,gstdate,phoneno,email,website,contactname,active,accno,bankname,branch,ifsc,createdon, createdby,modifiedby,ipaddress,accountholdername,contactno)  VALUES('" + txtPartycode.Text.Trim() + "','" + txtpartyname.Text.Trim() + "','" + txtaddress.Text.ToUpper().Trim() + "','" + combocity.SelectedValue + "','" + combostate.SelectedValue + "','" + combocountry.SelectedValue + "','" + Convert.ToInt32("0" + txtpincode.Text).ToString() + "','" + txtpanno.Text.ToUpper() + "','" + txttinno.Text + "','" + txtgstno.Text.ToUpper() + "','" + txtgstdate.Text + "','" + txtphoneno.Text + "','" + txtemail.Text + "','" + txtwebsite.Text + "','" + txtcontact.Text.ToUpper() + "','" + chk + "','" + txtaccno.Text + "','" + combobank.SelectedValue + "','" + combobranch.SelectedValue + "','" + comboIFSCE.SelectedValue + "','" + Convert.ToDateTime(System.DateTime.Now.ToString()).ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "','" + txtaccholdername.Text.ToUpper().Trim() + "','"+txtcontactno.Text+"')");
                            Utility.ExecuteNonQuery(sb.ToString());
                            MessageBox.Show("Record Saved Successfully " + txtpartyname.Text.ToUpper(), " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            sb.Clear();
                            sb.Append("update  asptblpartymas  set partycode='" + txtPartycode.Text.Trim() + "',partyname='" + txtpartyname.Text.Trim() + "',address='" + txtaddress.Text.ToUpper().Trim() + "',city='" + combocity.SelectedValue + "',state='" + combostate.SelectedValue + "',country='" + combocountry.SelectedValue + "',pincode='" + Convert.ToInt32("0" + txtpincode.Text).ToString() + "',panno='" + txtpanno.Text.ToUpper() + "',tinno='" + txttinno.Text + "',gstno='" + txtgstno.Text.ToUpper() + "',gstdate='" + txtgstdate.Text + "',phoneno='" + txtphoneno.Text + "',email='" + txtemail.Text + "',website='" + txtwebsite.Text + "',contactname='" + txtcontact.Text + "',active='" + chk + "',accno='" + txtaccno.Text + "',ifsc='" + comboIFSCE.SelectedValue + "', branch='" + combobranch.SelectedValue + "',bankname='" + combobank.SelectedValue + "' , active='" + chk + "' , createdon='" + Convert.ToDateTime(System.DateTime.Now.ToString()).ToString() + "',createdby='" + Class.Users.HUserName + "',modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "',accountholdername='" + txtaccholdername.Text.ToUpper().Trim() + "',contactno='"+txtcontactno.Text+"' where asptblpartymasid='" + txtpartyid.Text + "'");
                            Utility.ExecuteNonQuery(sb.ToString());
                            MessageBox.Show("Record Updated Successfully " + txtpartyname.Text.ToUpper(), "Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }
                    }
                }
               

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        public void Prints()
        {
           
        }

        public void Searchs()
        {
           
        }

        public void Searchs(int EditID)
        {
            try
            {
                if (EditID > 0)
                {

                    txtpartyid.Text = EditID.ToString();
                    DataTable dt = Utility.SQLQuery("select a.asptblpartymasid, a.partycode,a.partyname,a.address,b.cityname as city," +
                        "c.statename as  state,d.countryname as country,a.pincode,a.panno,a.tinno,a.gstno,a.gstdate," +
                        "a.phoneno,a.email,a.website,a.contactname,  a.active, a.accno, e.bankname, f.branch, f.ifsc, " +
                        " a.accountholdername,a.codeno,a.contactno  from  asptblpartymas a  join gtcitymast b on b.gtcitymastid=a.city join GTSTATEMAST c on a.state=c.GTSTATEMASTID   join GTCOUNTRYMAST d on b.country=d.GTCOUNTRYMASTID  join asptblbanmas e on e.asptblbanmasid=a.bankname  join asptblifscmas f on f.bankname=e.asptblbanmasid and f.bankname=a.bankname and f.asptblifscmasid=a.branch where a.asptblpartymasid=" + txtpartyid.Text + "");
                    if (dt.Rows.Count > 0)
                    {
                        txtpartyid.Text = Convert.ToString(dt.Rows[0]["asptblpartymasid"].ToString());
                        txtPartycode.Text= Convert.ToString(dt.Rows[0]["partycode"].ToString());
                        txtpartyname.Text = Convert.ToString(dt.Rows[0]["partyname"].ToString());
                        txtaddress.Text = Convert.ToString(dt.Rows[0]["address"].ToString());
                        combocity.Text = Convert.ToString(dt.Rows[0]["city"].ToString());
                        combostate.Text = Convert.ToString(dt.Rows[0]["state"].ToString());
                        combocountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString()); 
                        txtpincode.Text = Convert.ToString(dt.Rows[0]["pincode"].ToString()); 
                        txtpanno.Text = Convert.ToString(dt.Rows[0]["panno"].ToString()); 
                        txttinno.Text = Convert.ToString(dt.Rows[0]["tinno"].ToString());
                        txtgstno.Text = Convert.ToString(dt.Rows[0]["gstno"].ToString());
                        txtgstdate.Text = Convert.ToString(dt.Rows[0]["gstdate"].ToString()); 
                        txtphoneno.Text = Convert.ToString(dt.Rows[0]["phoneno"].ToString());
                        txtcontactno.Text = Convert.ToString(dt.Rows[0]["contactno"].ToString());
                        txtemail.Text = Convert.ToString(dt.Rows[0]["email"].ToString()); 
                        txtwebsite.Text = Convert.ToString(dt.Rows[0]["website"].ToString());
                        txtcontact.Text = Convert.ToString(dt.Rows[0]["contactname"].ToString());
                        txtaccno.Text = Convert.ToString(dt.Rows[0]["accno"].ToString());
                        combobank.Text = Convert.ToString(dt.Rows[0]["bankname"].ToString());
                        combobranch.Text = Convert.ToString(dt.Rows[0]["branch"].ToString());
                        comboIFSCE.Text = Convert.ToString(dt.Rows[0]["ifsc"].ToString());
                        txtaccholdername.Text=Convert.ToString(dt.Rows[0]["accountholdername"].ToString());
                        txtcode.Text = Convert.ToString(dt.Rows[0]["codeno"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        tabControl1.SelectTab(tabPage1);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                string filePath = string.Empty; dataGridView1.AllowUserToAddRows = false;
                string fileExt = string.Empty; 
                OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
                if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
                {
                    filePath = file.FileName; //get the path of the file  
                    fileExt = Path.GetExtension(filePath); //get the file extension  
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        try
                        {
                            DataTable dtExcel = new DataTable();
                            dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                            dataGridView1.Visible = true;
                            dataGridView1.DataSource = dtExcel;

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                    }
                }
              
                int cnt = dataGridView1.Rows.Count - 1;
                label48.Text = "Total Count  :" + cnt.ToString();
            }
            else
            {
                MessageBox.Show("pls Contact your Administrator." + Class.Users.Log.ToString(), "Register Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Dispose();
            }
            CommonFunctions.SetRowNumber(dataGridView1);
        }
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
          //if (fileExt.CompareTo(".xls") == 0)
          //    path = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
          //else
          //    path = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  


            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES';"; //for above excel 2007  
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
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void combocity_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = state(Convert.ToInt64(combocity.SelectedValue));
                combostate.DisplayMember = "statename";
                combostate.ValueMember = "gtstatemastid";
                combostate.DataSource = dt;
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }
        }

        private void combostate_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = country(Convert.ToInt64(combostate.SelectedValue));

                combocountry.DisplayMember = "countryname";
                combocountry.ValueMember = "gtcountrymastid";
                combocountry.DataSource = dt;
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }
        }

        private void txtpincode_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        }

        private void txtphoneno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        }

        private void txtaccno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        }

        private void txtgstdate_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == '/' || e.KeyChar == (char)Keys.Back);

        }

        private void uccListView1_Click(object sender, EventArgs e)
        {
            
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void txtaccholdername_KeyPress(object sender, KeyPressEventArgs e)
        {
           // e.Handled = !(char.IsLetter(e.KeyChar) || (char.IsWhiteSpace(e.KeyChar)) || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back);

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void txtphoneno_TextChanged(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void txtaccholdername_TextChanged(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void combobranch_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (combobranch.SelectedValue.ToString() != "System.Data.DataRowView" && combobranch.SelectedValue != null)
            //{
            //    sb.Clear();
            //    sb.Append("select  a.asptblifscmasid,a.ifsc from asptblifscmas a join asptblbanmas b on a.bankname=b.asptblbanmasid where a.active='T' and a.branch='" + combobranch.SelectedValue + "' ");
            //    DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblifscmas");
            //    DataTable dt = ds.Tables["asptblifscmas"];
            //    comboIFSCE.DataSource = dt;
            //    comboIFSCE.DisplayMember = "ifsc";
            //    comboIFSCE.ValueMember = "asptblifscmasid";
            //}
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bankname();
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void txtwebsite_TextChanged(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void txtemail_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void combobank_SelectedValueChanged(object sender, EventArgs e)
        {
            if (combobank.Text.ToString() != "System.Data.DataRowView" && combobank.Text != "")
            {
                DataTable dt = branch(combobank.Text.ToString());    
                combobranch.DataSource = dt;
                combobranch.DisplayMember = "branch";
                combobranch.ValueMember = "asptblifscmasid";
            }
        }
        private void combobranch_SelectedValueChanged(object sender, EventArgs e)
        {
            if (combobranch.Text.ToString() != "System.Data.DataRowView" && combobranch.Text != "")
            {
                DataTable dt = ifce(combobank.Text,combobranch.Text.ToString());
                comboIFSCE.DataSource = dt;
                comboIFSCE.DisplayMember = "ifsc";
                comboIFSCE.ValueMember = "asptblifscmasid";
            }
        }

        private void combobank_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtpartyname_TextChanged(object sender, EventArgs e)
        {
            if(txtpartyname.Text != "")
            {
                txtaccholdername.Text = txtpartyname.Text;
            }
        }
    }
}

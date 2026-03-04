using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class BuyerMaster : Form,ToolStripAccess
    {

        private static BuyerMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
   
        public static BuyerMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new BuyerMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public BuyerMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
  
        }

        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    //for (int r = 0; r < dt1.Rows.Count; r++)
                    //{

                    //    if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                    //    if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                    //    if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                    //    if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                    //    if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                    //    if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                    //    if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                    //    if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                    //    if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                    //    if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                    //    if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                    //    if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }

                    //}
                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        private void BuyerMaster_Load(object sender, EventArgs e)
        {
            News();GridLoad();
        }

        public void compid(string s)
        {
            string sel = "select max(a.ASPTBLBUYMASID1) as id  from  ASPTBLBUYMAS a  where a.BUYERCODE='" + s.ToUpper() + "' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLBUYMAS");
            DataTable dt = ds.Tables["ASPTBLBUYMAS"];
            Int64 partycount = 1;
            if (dt.Rows.Count >= 1)
            {
                txtbuyid1.Text = "";
                Int64 cc = Convert.ToInt64("0" + dt.Rows[0]["id"].ToString()) + partycount;
                txtbuyid1.Text = cc.ToString();
            }
            else
            {
                txtbuyid1.Text = "1";
            }
        }
        public DataTable compcode()
        {

            DataTable dt = mas.aspcomcode();
            combocompcode.DataSource = dt;
            combocompcode.DisplayMember = "compcode";
            combocompcode.ValueMember = "GTCOMPMASTID";
            return dt;
        }
        public void state()
        {

            string sel = "select b.GTSTATEMASTID,b.statename from   join GTSTATEMAST b  where b.active='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTSTATEMAST");
            DataTable dt = ds.Tables["GTSTATEMAST"];
            combostate.DataSource = dt;
            combostate.DisplayMember = "statename";
            combostate.ValueMember = "GTSTATEMASTID";
           
        }
        public DataTable state(Int64 s)
        {
            string sel = "select b.GTSTATEMASTID,b.statename from  GTCITYMAST a  join GTSTATEMAST b on a.state=b.GTSTATEMASTID   join GTCOUNTRYMAST c on b.country=c.GTCOUNTRYMASTID   where b.active='T' and a.GTCITYMASTID='" + s + "' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTSTATEMAST");
            DataTable dt = ds.Tables["GTSTATEMAST"];
            return dt;
        }
        public DataTable city()
        {
            string sel = "select a.GTCITYMASTID,a.cityname from  GTCITYMAST a  join GTSTATEMAST b on a.state=b.GTSTATEMASTID   join GTCOUNTRYMAST c on b.country=c.GTCOUNTRYMASTID   where a.active='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCITYMAST");
            DataTable dt = ds.Tables["GTCITYMAST"];
            combocity.DataSource = dt;
            combocity.DisplayMember = "cityname";
            combocity.ValueMember = "GTCITYMASTID";
            return dt;
        }
        public DataTable country(Int64 s)
        {
            string sel1 = "SELECT  b.GTCOUNTRYMASTID,b.countryname from GTSTATEMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID where a.GTSTATEMASTID='" + s + "' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTSTATEMAST");
            DataTable dt = ds.Tables["GTSTATEMAST"];
            return dt;
        }

        public void Saves()
        {
            try
            {

                if (combocompcode.Text == "")
                {
                    MessageBox.Show("'CompCode  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combocompcode.Focus();
                    return;
                }
                if (combocompname.Text == "")
                {
                    MessageBox.Show("'CompName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combocompname.Focus();
                    return;
                }
                if (combocity.SelectedValue == null)
                {
                    MessageBox.Show("'City Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combocity.Focus();
                    return;
                }
                if (combostate.SelectedValue == null)
                {
                    MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combostate.Focus();
                    return;
                }
                if (combocountry.SelectedValue == null)
                {
                    MessageBox.Show("'Contry Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combocountry.Focus(); return;
                }
                if (txtaddress.Text == "")
                {
                    MessageBox.Show("'Address  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtaddress.Focus(); return;
                }

                else
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsStringNumbericslacehyper(combocompcode.Text) == true)
                    {
                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select ASPTBLBUYMASID    from  ASPTBLBUYMAS   WHERE compcode='" + combocompcode.SelectedValue + "' and compname='" + combocompname.SelectedValue + "' and buyercode='" + txtbuycode.Text + "' and buyername='" + txtbuyname.Text + "' and buyingagent='" + txtbuyagent.Text.ToUpper() + "' and city='" + combocity.SelectedValue + "' and state='" + combostate.SelectedValue + "' and country='" + combocountry.SelectedValue + "' and address='" + txtaddress.Text + "' and pincode='" + txtpincode.Text + "' and phoneno='" + txtphoneno.Text + "' and email='" + txtemail.Text + "' and website='" + txtwebsite.Text + "'  and contactname='" + txtcontact.Text + "'  and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLBUYMAS");
                        DataTable dt = ds.Tables["ASPTBLBUYMAS"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtbuyid.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtbuyid.Text) == 0 || Convert.ToInt32("0" + txtbuyid.Text) == 0)
                        {
                            string ins = "insert into ASPTBLBUYMAS(compcode,compname,buyercode,buyername,buyingagent,city,state,country,address,pincode,phoneno,email,website,contactname,active,createdby,modifiedby,ipaddress)  VALUES('" + combocompcode.SelectedValue + "','" + combocompname.SelectedValue + "', '" + txtbuycode.Text.ToUpper() + "','" + txtbuyname.Text.ToUpper() + "','" + txtbuyagent.Text.ToUpper() + "', '" + combocity.SelectedValue + "', '" + combostate.SelectedValue + "','" + combocountry.SelectedValue + "' ,'" + txtaddress.Text.ToUpper() + "','" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "','" + txtphoneno.Text + "','" + txtemail.Text.ToUpper() + "','" + txtwebsite.Text.ToUpper() + "','" + txtcontact.Text.ToUpper() + "','" + chk + "', '" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "')";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtbuyid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  ASPTBLBUYMAS  set  ASPTBLBUYMASID1='" + txtbuyid.Text + "' , compcode='" + combocompcode.SelectedValue + "',compname='" + combocompname.SelectedValue + "',buyercode='" + txtbuycode.Text + "',buyername='" + txtbuyname.Text + "',buyingagent='" + txtbuyagent.Text.ToUpper() + "', city='" + combocity.SelectedValue + "', state='" + combostate.SelectedValue + "',country='" + combocountry.SelectedValue + "' ,address='" + txtaddress.Text.ToUpper() + "',pincode='" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "',phoneno='" + txtphoneno.Text + "',email='" + txtemail.Text.ToUpper() + "',website='" + txtwebsite.Text.ToUpper() + "',contactname='" + txtcontact.Text.ToUpper() + "', active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where  ASPTBLBUYMASID='" + txtbuyid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtbuyid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid Data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;

                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("Error " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void BuyerMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        public void News()
        {

            GridLoad(); compcode(); state();
            city(); empty();
        }

        private void empty()
        {
            txtbuyid.Text = "";
            txtbuyid1.Text = ""; combocompcode.Select();
            combostate.SelectedIndex = -1; combostate.Text = "";
            txtbuycode.Text = ""; txtbuyname.Text = "";
            combocountry.SelectedIndex = -1; combocountry.Text = ""; txtsearch.Text = "";
            checkactive.Checked = true;
            combocompcode.Text = "";
            combocompname.Text = "";
            txtaddress.Text = "";
            txtpincode.Text = "";
         
            txtphoneno.Text = "";
            txtemail.Text = "";
            txtwebsite.Text = "";
            txtcontact.Text = "";txtbuyagent.Text = "";combocity.Text = "";
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.ASPTBLBUYMASID,e.compcode,a.BUYERCODE,a.buyername,b.cityname as city,c.statename as state,D.countryname as  country,a.address,  a.pincode, a.phoneno, a.active from ASPTBLBUYMAS a join GTCITYMAST b on a.city = b.GTCITYMASTID  join GTSTATEMAST c on a.state = c.GTSTATEMASTID join GTCOUNTRYMAST d on a.country = d.GTCOUNTRYMASTID  join gtcompmast e on e.gtcompmastid=a.compcode order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLBUYMAS");
                DataTable dt = ds.Tables["ASPTBLBUYMAS"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLBUYMASID"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["BUYERCODE"].ToString());
                        list.SubItems.Add(myRow["BUYERNAME"].ToString());
                        list.SubItems.Add(myRow["city"].ToString());
                        list.SubItems.Add(myRow["state"].ToString());
                        list.SubItems.Add(myRow["country"].ToString());
                        list.SubItems.Add(myRow["pincode"].ToString());
                        list.SubItems.Add(myRow["phoneno"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtbuyid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                   

                    string sel1 = "select a.ASPTBLBUYMASID,a.compcode,a.compname,a.buyercode,a.buyername,a.buyingagent, b.cityname as city,c.statename as state,D.countryname as  country,a.address,  a.pincode, a.phoneno, a.email,  a.website,  a.active from ASPTBLBUYMAS a join GTCITYMAST b on a.city = b.GTCITYMASTID  join GTSTATEMAST c on a.state = c.GTSTATEMASTID join GTCOUNTRYMAST d on a.country = d.GTCOUNTRYMASTID   where a.ASPTBLBUYMASID='" + txtbuyid.Text + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLBUYMAS");
                    DataTable dt = ds.Tables["ASPTBLBUYMAS"];

                    if (dt.Rows.Count > 0)
                    {
                        // txtcompid.Text = "";
                        txtbuyid.Text = Convert.ToString(dt.Rows[0]["ASPTBLBUYMASID"].ToString());
                        txtbuyid1.Text = Convert.ToString(dt.Rows[0]["ASPTBLBUYMASID"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        txtbuycode.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                        txtbuyname.Text = Convert.ToString(dt.Rows[0]["buyername"].ToString());
                        txtbuyagent.Text = Convert.ToString(dt.Rows[0]["buyingagent"].ToString());
                        combocity.Text = Convert.ToString(dt.Rows[0]["city"].ToString());
                        combostate.Text = Convert.ToString(dt.Rows[0]["state"].ToString());
                        combocountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
                        txtaddress.Text = Convert.ToString(dt.Rows[0]["address"].ToString());
                        txtpincode.Text = Convert.ToString(dt.Rows[0]["pincode"].ToString());
                       
                        txtphoneno.Text = Convert.ToString(dt.Rows[0]["phoneno"].ToString());
                        txtemail.Text = Convert.ToString(dt.Rows[0]["email"].ToString());
                        txtwebsite.Text = Convert.ToString(dt.Rows[0]["website"].ToString());

                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }


                        Combocity_SelectedIndexChanged(sender, e);
                        Combostate_SelectedIndexChanged(sender, e);

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            tabControl1.SelectTab(tabPage1);
            txtbuycode.Select();
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int i = 1;
                if (txtsearch.Text.Length > 1)
                {
                    listView1.Items.Clear();i = 1;
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text) || item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text))
                        {
                          
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.SubItems.Add(item.SubItems[7].Text);
                            list.SubItems.Add(item.SubItems[8].Text);
                            list.SubItems.Add(item.SubItems[9].Text);
                            list.SubItems.Add(item.SubItems[10].Text);
                            list.SubItems.Add(item.SubItems[11].Text);    
                            
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                        }
                       i++;
                    }
                }
                else
                {

                    listView1.Items.Clear();i = 1;
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.SubItems.Add(item.SubItems[5].Text);
                        list.SubItems.Add(item.SubItems[6].Text);
                        list.SubItems.Add(item.SubItems[7].Text);
                        list.SubItems.Add(item.SubItems[8].Text);
                        list.SubItems.Add(item.SubItems[9].Text);
                        list.SubItems.Add(item.SubItems[10].Text);
                        list.SubItems.Add(item.SubItems[11].Text);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("---" + ex.ToString());
            }
        }

        private void Combostate_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = country(Convert.ToInt64(combostate.SelectedValue));
                combocountry.DataSource = dt;
                combocountry.DisplayMember = "countryname";
                combocountry.ValueMember = "GTCOUNTRYMASTID";
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }



        }
        private void StateRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Combostate_SelectedIndexChanged(sender, e);
        }

        private void Combocity_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = state(Convert.ToInt64(combocity.SelectedValue));
                combostate.DataSource = dt;
                combostate.DisplayMember = "statename";
                combostate.ValueMember = "GTSTATEMASTID";
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }
        }

        private void Combostate_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void Txtcompname_TextChanged(object sender, EventArgs e)
        {

        }

        private void Txtcompcode_TextChanged(object sender, EventArgs e)
        {
            if (txtbuyid.Text == "")
            {
                if (combocompcode.Text.Length >= 3)
                {
                    compid(combocompcode.Text);
                }
            }
        }

        public void Deletes()
        {
            if (txtbuyid.Text != "")
            {
                string sel1 = "select a.ASPTBLBUYMASID from ASPTBLBUYMAS a join ASPTBLPUR b on a.ASPTBLBUYMASID=b.buyer where a.ASPTBLBUYMASID='" + txtbuyid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLBUYMAS");
                DataTable dt = ds.Tables["ASPTBLBUYMAS"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + combocompcode.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from ASPTBLBUYMAS where ASPTBLBUYMASID='" + Convert.ToInt64("0" + txtbuyid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + combocompcode.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

     
       

      

        private void combocity_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combostate.Select();
            }
        }

        private void combostate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtaddress.Select();
            }
        }

        private void txtaddress_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtpincode.Select();
            }
        }

        
       

        private void txtgstdate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtphoneno.Select();
            }
        }

        private void txtphoneno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtemail.Select();
            }
        }

        private void txtemail_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtwebsite.Select();
            }
        }

        private void txtwebsite_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtcontact.Select();

            }
        }

        private void txtcompcode_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtcompname_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        public void Prints()
        {
            throw new NotImplementedException();
        }

        public void Searchs()
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

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = mas.aspcomcode(combocompcode.Text);
            combocompname.DataSource = dt;
            combocompname.DisplayMember = "compname";
            combocompname.ValueMember = "GTCOMPMASTID";
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class CompanyMaster : Form,ToolStripAccess
    {
        private static CompanyMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        ListView listfilter = new ListView();
      
        public static CompanyMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new CompanyMaster();
                return _instance;
            }
        }
        public CompanyMaster()
        {
            InitializeComponent();
           // usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
         
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            //DGBankDtl.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            //panel2.BackColor = Class.Users.BackColors;
            //panel3.BackColor = Class.Users.BackColors;
            //butheader.BackColor = Class.Users.BackColors;
            //this.BackColor = Class.Users.BackColors;
            
        }
     
       
        private void CompanyMaster_Load(object sender, EventArgs e)
        {
            News();empty();
        }
        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {

                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        public void compid(string s)
        {
            string sel = "select max(a.gtcompmastid1) as id  from  gtcompmast a  where a.compcode='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            Int64 partycount = 1;
            if (dt.Rows[0]["id"].ToString() == "")
            {

                txtcompid1.Text = "1";

            }
            else
            {
                Int64 cc = Convert.ToInt64("0" + dt.Rows[0]["id"].ToString()) + partycount;
                txtcompid1.Text = cc.ToString();
            }
        }
        public DataTable state(Int64 s)
        {
            string sel = "select b.gtstatemastid,b.statename from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where b.active='T' and a.gtcitymastid='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];          
            return dt;
        }
        public DataTable city()
        {
            string sel = "select a.gtcitymastid,a.cityname from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcitymast");
            DataTable dt = ds.Tables["gtcitymast"];
            combocity.DataSource = dt;
            combocity.DisplayMember = "cityname";
            combocity.ValueMember = "gtcitymastid";
            return dt;
        }
        public DataTable country(Int64 s)
        {
            string sel1 = "SELECT  b.gtcountrymastid,b.countryname from gtstatemast a join gtcountrymast b on a.country=b.gtcountrymastid where a.gtstatemastid='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];

            return dt;
        }
        public void Saves()
        {
            try
            {
               
                if (txtcompcode.Text == "")
                {
                    MessageBox.Show("pls Enter CompCode"); txtcompcode.Focus(); return;
                }
                if (txtcompname.Text == "")
                {
                    MessageBox.Show("pls Enter Compname"); txtcompname.Focus(); return;
                }
                    if (combocity.Text == "")
                    {
                        MessageBox.Show("pls select City"); combocity.Focus(); return;
                    }
                    if (combostate.Text == "")
                    {
                        MessageBox.Show("pls select State"); combostate.Focus(); return;
                    }
                    if (txtaddress.Text == "")
                    {
                        MessageBox.Show("pls Enter Address"); txtaddress.Focus(); return;
                    }
                    if (txtgstno.Text == "")
                    {
                        MessageBox.Show("pls Enter GstNo"); txtgstno.Focus(); return;
                    }
                   

                  


                    if (txtcompcode.Text != "" && txtcompname.Text != "")
                    {
                        string chk = "", chkcustomer = "";



                        DateTime mteain = DateTime.Parse(txtbreakfrom.Text);
                        DateTime mteaout = DateTime.Parse(txtbreakto.Text);

                        DateTime lin = DateTime.Parse(txtlunchout.Text);
                        DateTime lout = DateTime.Parse(txtlunchin.Text);

                        DateTime eteain = DateTime.Parse(txteveteafrom.Text);
                        DateTime eteaout = DateTime.Parse(txteveteato.Text);

                        TimeSpan mdiffer = mteaout.Subtract(mteain);
                        TimeSpan ldiffer = lin.Subtract(lout);
                        TimeSpan ddiffer = eteaout.Subtract(eteain);
                        TimeSpan adj = TimeSpan.Parse("0" + textBox1.Text);
                        TimeSpan totdiffer = mdiffer + ldiffer + ddiffer + adj;
                        txttothrs.Text = "";
                        txttothrs.Text = totdiffer.ToString();
                        MySqlCommand cmd;
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }

                       
                        string sel = "select gtcompmastid    from  gtcompmast   WHERE compcode='" + txtcompcode.Text + "' and compname='" + txtcompname.Text + "' and displayname='"+txtdisplayname.Text.ToUpper().Trim()+"' and city='" + combocity.SelectedValue + "' and state='" + combostate.SelectedValue + "' and country='" + combocountry.SelectedValue + "' and address='" + txtaddress.Text + "' and pincode='" + txtpincode.Text + "' and panno='" + txtpanno.Text + "' and tinno='" + txttinno.Text + "' and gstno='" + txtgstno.Text + "' and gstdate='" + txtgstdate.Text + "' and phoneno='" + txtphoneno.Text + "' and email='" + txtemail.Text + "' and website='" + txtwebsite.Text + "'  and contactname='" + txtcontact.Text + "' and ptransaction='" + chkcustomer + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                        DataTable dt = ds.Tables["gtcompmast"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtcompid.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtcompid.Text) == 0 || Convert.ToInt32("0" + txtcompid.Text) == 0)
                        {
                            compid(txtcompcode.Text);
                            string ins = "insert into gtcompmast(gtcompmastid1,compcode,compname,displayname,city,state,country,address,pincode,panno,tinno,gstno,gstdate,phoneno,email,website,contactname,active,ptransaction,createdby,modifiedby,ipaddress)  VALUES('" + txtcompid1.Text + "','" + txtcompcode.Text.ToUpper() + "','" + txtcompname.Text.ToUpper() + "','"+txtdisplayname.Text.ToUpper().Trim()+"', '" + combocity.SelectedValue + "', '" + combostate.SelectedValue + "','" + combocountry.SelectedValue + "' ,'" + txtaddress.Text.ToUpper() + "','" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "','" + txtpanno.Text + "','" + txttinno.Text.ToUpper() + "','" + txtgstno.Text.ToUpper() + "','" + txtgstdate.Text + "','" + txtphoneno.Text + "','" + txtemail.Text.ToUpper() + "','" + txtwebsite.Text.ToUpper() + "','" + txtcontact.Text.ToUpper() + "','" + chk + "','" + chkcustomer + "' ,'" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "')";
                            Utility.ExecuteNonQuery(ins);
                            string sel1 = "select MAX(gtcompmastid) ID    from  gtcompmast   WHERE compcode='" + txtcompcode.Text + "' and gstno='" + txtgstno.Text + "'";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                            DataTable dt1 = ds1.Tables["gtcompmast"];
                            if (dt1.Rows.Count > 0 && stdbytes != null)
                            {
                                string ins1 = "UPDATE  gtcompmast SET companylogo=@companylogo where  gtcompmastid='" + dt1.Rows[0]["ID"].ToString() + "'";
                                cmd = new MySqlCommand(ins1, Utility.Connect());
                                cmd.Parameters.Add(new MySqlParameter("@companylogo", stdbytes));
                                cmd.ExecuteNonQuery();
                            }
                            MessageBox.Show("Record Saved Successfully " + txtcompid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {

                            string up = "update  gtcompmast  set  gtcompmastid1='" + txtcompid1.Text + "',compcode='" + txtcompcode.Text.ToUpper() + "',compname='" + txtcompname.Text.ToUpper() + "',displayname='" + txtdisplayname.Text.ToUpper().Trim()+"', city='" + combocity.SelectedValue + "', state='" + combostate.SelectedValue + "',country='" + combocountry.SelectedValue + "' ,address='" + txtaddress.Text.ToUpper() + "',pincode='" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "',panno='" + txtpanno.Text.ToUpper() + "',tinno='" + txttinno.Text.ToUpper() + "',gstno='" + txtgstno.Text.ToUpper() + "',gstdate='" + txtgstdate.Text + "',phoneno='" + txtphoneno.Text + "',email='" + txtemail.Text.ToUpper() + "',website='" + txtwebsite.Text.ToUpper() + "',contactname='" + txtcontact.Text.ToUpper() + "', active='" + chk + "', modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where  gtcompmastid='" + txtcompid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            if (txtcompid.Text != "" && stdbytes != null)
                            {
                                string up1 = "UPDATE  gtcompmast SET  companylogo=@companylogo where  gtcompmastid='" + txtcompid.Text + "'";
                                cmd = new MySqlCommand(up1, Utility.Connect());
                                cmd.Parameters.Add(new MySqlParameter("@companylogo", stdbytes));
                                cmd.ExecuteNonQuery();

                            }

                            MessageBox.Show("Record Updated Successfully " + txtcompid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            catch (Exception ex)
            {

                MessageBox.Show("Error " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 4)
            {
                DateTime mteain = DateTime.Parse(txtbreakfrom.Text);
                DateTime mteaout = DateTime.Parse(txtbreakto.Text);

                DateTime lin = DateTime.Parse(txtlunchout.Text);
                DateTime lout = DateTime.Parse(txtlunchin.Text);

                DateTime eteain = DateTime.Parse(txteveteafrom.Text);
                DateTime eteaout = DateTime.Parse(txteveteato.Text);

                TimeSpan mdiffer = mteaout.Subtract(mteain);
                TimeSpan ldiffer = lin.Subtract(lout);
                TimeSpan ddiffer = eteaout.Subtract(eteain);
                TimeSpan adj = TimeSpan.Parse(textBox1.Text);
                TimeSpan totdiffer = mdiffer + ldiffer + ddiffer + adj;
                txttothrs.Text = "";
                txttothrs.Text = totdiffer.ToString();
            }
           
        }
        private void CompanyMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       
        public void News()
        {
                    GridLoad(); city(); bankname(); empty();
        }
        private void empty()
        {
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            GlobalVariables.CurrentForm = this;
            Class.Users.UserTime = 0;
            txtcompid.Text = "";txtdisplayname.Text = "";
            txtcompid1.Text = "";txtcompcode.Select(); listfilter.Items.Clear();
            combocity.SelectedIndex = -1; combostate.SelectedIndex = -1; combostate.Text = "";
            combocountry.SelectedIndex = -1; combocountry.Text = ""; txtsearch.Text = "";
            checkactive.Checked = true;           
            txtcompcode.Text = "";txtbranch.Text = "";
            txtcompname.Text = "";
            txtaddress.Text = "";
            txtpincode.Text = "";
            txtpanno.Text = "";
            txttinno.Text = "";
            txtgstno.Text = "";
            txtgstdate.Text = "";
            txtphoneno.Text = "";
            txtemail.Text = "";
            txtwebsite.Text = "";
            txtcontact.Text = "";
            //txtcontact.Text = ""; txttothrs.Text = "";txtfemaleout.Text = "";txttpr.Text = "";
            //txtintime.Text = ""; txtouttime.Text = ""; txteveteafrom.Text = ""; txteveteato.Text = "";
            //txtbreakfrom.Text = ""; txtbreakto.Text = ""; txtlunchout.Text = ""; txtlunchin.Text = "";
            txtaccno.Text = "";txtifsc.Text = "";combobank.SelectedIndex = -1;
          
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
           
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            pictureBox1.Image = null;
            tabControl1.SelectTab(tabPage1);
            txtcompcode.Select();
        }
        void bankname()
        {
            //string ins= "select distinct  a.asptblbanmasid,a.bankname from asptblbanmas a join asptblifscmas b on a.asptblbanmasid=b.bankname where b.active='T' ";
            //DataSet ds = Utility.ExecuteSelectQuery(ins, "asptblbanmas");
            //DataTable dt = ds.Tables["asptblbanmas"];
            //combobank.DataSource = dt;
            //combobank.DisplayMember = "bankname";
            //combobank.ValueMember = "asptblbanmasid";
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.gtcompmastid,a.compcode,a.compname,b.cityname as city,c.statename as state, a.active  from gtcompmast a join gtcitymast b on a.city = b.gtcitymastid  join gtstatemast c on a.state = c.gtstatemastid join gtcountrymast d on a.country = d.gtcountrymastid    order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["gtcompmastid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["compname"].ToString());
                        list.SubItems.Add(myRow["city"].ToString());
                        list.SubItems.Add(myRow["state"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        list.SubItems.Add("");
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
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
                if (listView1.SelectedItems.Count == 0)
                    return;

                txtcompid.Text = listView1.SelectedItems[0].SubItems[2].Text;

                string query = @" SELECT a.gtcompmastid1, a.gtcompmastid, a.compcode, a.compname, a.displayname, b.cityname AS city,            c.statename AS state, d.countryname AS country,            a.address, a.pincode, a.panno, a.tinno,            a.gstno, '' AS gstdate, a.phoneno, a.email,            a.website, a.active, a.ptransaction, a.companylogo        FROM gtcompmast a        JOIN gtcitymast b ON a.city = b.gtcitymastid        JOIN gtstatemast c ON a.state = c.gtstatemastid        JOIN gtcountrymast d ON a.country = d.gtcountrymastid        WHERE a.gtcompmastid='"+ txtcompid.Text + "'";

                DataSet ds = Utility.ExecuteSelectQuery(query,  "gtcompmast");

                DataTable dt = ds.Tables["gtcompmast"];
                if (dt == null || dt.Rows.Count == 0)
                    return;

                DataRow row = dt.Rows[0];

                txtcompid.Text = Convert.ToString(row["gtcompmastid"]);
                txtcompid1.Text = Convert.ToString(row["gtcompmastid1"]);
                txtcompcode.Text = Convert.ToString(row["compcode"]);
                txtcompname.Text = Convert.ToString(row["compname"]);
                txtdisplayname.Text = Convert.ToString(row["displayname"]);

                combocity.Text = Convert.ToString(row["city"]);
                combostate.Text = Convert.ToString(row["state"]);
                combocountry.Text = Convert.ToString(row["country"]);

                txtaddress.Text = Convert.ToString(row["address"]);
                txtpincode.Text = Convert.ToString(row["pincode"]);
                txtpanno.Text = Convert.ToString(row["panno"]);
                txttinno.Text = Convert.ToString(row["tinno"]);
                txtgstno.Text = Convert.ToString(row["gstno"]);
                txtgstdate.Text = Convert.ToString(row["gstdate"]);
                txtphoneno.Text = Convert.ToString(row["phoneno"]);
                txtemail.Text = Convert.ToString(row["email"]);
                txtwebsite.Text = Convert.ToString(row["website"]);

                checkactive.Checked = Convert.ToString(row["active"]) == "T";

                if (row["companylogo"] != DBNull.Value)
                {
                    byte[] bytes = (byte[])row["companylogo"];
                    pictureBox1.Image = Models.Device.ByteArrayToImage(bytes);
                }
                else
                {
                    pictureBox1.Image = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            tabControl1.SelectTab(tabPage1);

            //try
            //{
            //    if (listView1.Items.Count > 0)
            //    {

            //        txtcompid.Text = listView1.SelectedItems[0].SubItems[2].Text;
            //        string sel1 = "select a.gtcompmastid1, a.gtcompmastid,a.compcode,a.compname,a.displayname, b.cityname as city,c.statename as state,D.countryname as  country,a.address,  a.pincode,a.panno,a.tinno,  a.gstno,'' as gstdate,  a.phoneno, a.email,  a.website,  a.active,a.ptransaction,a.companylogo from gtcompmast a join gtcitymast b on a.city = b.gtcitymastid  join gtstatemast c on a.state = c.gtstatemastid join gtcountrymast d on a.country = d.gtcountrymastid   where a.gtcompmastid='" + txtcompid.Text + "'";

            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
            //        DataTable dt = ds.Tables["gtcompmast"];

            //        if (dt != null)
            //        {

            //            txtcompid.Text = Convert.ToString(dt.Rows[0]["gtcompmastid"].ToString());
            //            txtcompid1.Text = Convert.ToString(dt.Rows[0]["gtcompmastid1"].ToString());
            //            txtcompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
            //            txtcompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
            //            txtdisplayname.Text = Convert.ToString(dt.Rows[0]["displayname"].ToString());
            //            combocity.Text = Convert.ToString(dt.Rows[0]["city"].ToString());
            //            Combocity_SelectedIndexChanged(sender, e);
            //            combostate.Text = Convert.ToString(dt.Rows[0]["state"].ToString());
            //            combocountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
            //            txtaddress.Text = Convert.ToString(dt.Rows[0]["address"].ToString());
            //            txtpincode.Text = Convert.ToString(dt.Rows[0]["pincode"].ToString());
            //            txtpanno.Text = Convert.ToString(dt.Rows[0]["panno"].ToString());
            //            txttinno.Text = Convert.ToString(dt.Rows[0]["tinno"].ToString());
            //            txtgstno.Text = Convert.ToString(dt.Rows[0]["gstno"].ToString());
            //            txtgstdate.Text = Convert.ToString(dt.Rows[0]["gstdate"].ToString());
            //            txtphoneno.Text = Convert.ToString(dt.Rows[0]["phoneno"].ToString());
            //            txtemail.Text = Convert.ToString(dt.Rows[0]["email"].ToString());
            //            txtwebsite.Text = Convert.ToString(dt.Rows[0]["website"].ToString());


            //            if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }


            //            if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
            //            if(dt.Rows[0]["companylogo"].ToString() != "")
            //            {                           
            //                pictureBox1.Image = null; stdbytes = null;
            //                stdbytes = (byte[])dt.Rows[0]["companylogo"];
            //                Image img = Models.Device.ByteArrayToImage(stdbytes);
            //                pictureBox1.Image = img;
            //            }



            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
            //tabControl1.SelectTab(tabPage1);
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    listView1.Items.Clear();

            //    int i = 1;
            //    string search = txtsearch.Text.Trim().ToUpper();

            //    foreach (ListViewItem item in listfilter.Items)
            //    {
            //        bool isMatch = string.IsNullOrEmpty(search) ||
            //                       item.SubItems[3].Text.Contains(search) ||
            //                       item.SubItems[4].Text.Contains(search);

            //        if (!isMatch)
            //            continue;

            //        ListViewItem list = new ListViewItem(i.ToString());

            //        list.SubItems.Add(item.SubItems[2].Text);
            //        list.SubItems.Add(item.SubItems[3].Text);
            //        list.SubItems.Add(item.SubItems[4].Text);
            //        list.SubItems.Add(item.SubItems[5].Text);
            //        list.SubItems.Add(item.SubItems[6].Text);
            //        list.SubItems.Add(item.SubItems[7].Text);
            //        list.SubItems.Add(item.SubItems[8].Text);

            //        list.BackColor = i % 2 == 0
            //            ? Class.Users.Color1
            //            : Class.Users.Color2;

            //        listView1.Items.Add(list);
            //        i++;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

            try
            {
                int item0 = 0; int i = 1;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text) || item.SubItems[4].ToString().Contains(txtsearch.Text))
                        {
                            list.SubItems.Add(i.ToString());

                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.SubItems.Add(item.SubItems[7].Text);
                            list.SubItems.Add(item.SubItems[8].Text);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {

                    listView1.Items.Clear(); item0 = 0;
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
                        list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);

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
              
                combocountry.DisplayMember = "countryname";
                combocountry.ValueMember = "gtcountrymastid";
                combocountry.DataSource = dt;
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
              
                combostate.DisplayMember = "statename";
                combostate.ValueMember = "gtstatemastid";
                combostate.DataSource = dt;
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }
        }

        private void Combostate_SelectedIndexChanged_1(object sender, EventArgs e)
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

        private void Txtcompname_TextChanged(object sender, EventArgs e)
        {

        }

        private void Txtcompcode_TextChanged(object sender, EventArgs e)
        {
            //if (txtcompid.Text == "")
            //{
               
            //        compid(txtcompcode.Text);
                
            //}
        }

        private void Deletes_Click(object sender, EventArgs e)
        {
            if (txtcompid.Text != "")
            {
                string sel1 = "select a.gtcompmastid from gtcompmast a join HRPayDetails b on a.gtcompmastid=b.compname where a.gtcompmastid='" + txtcompid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcompcode.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from gtcompmast where gtcompmastid='" + Convert.ToInt64("0" + txtcompid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtcompcode.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void txtcompcode_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    txtcompname.Focus();
            //}
        }

        private void txtcompname_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    combodivision.Select();
            //}
        }

        private void combodivision_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combocity.Select();
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

        private void txtpincode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtpanno.Select();
            }
        }

        private void txtpanno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txttinno.Select();
            }
        }

        private void txttinno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtgstno.Select();
            }
        }

        private void txtgstno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtgstdate.Select();
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
            if (e.KeyChar >= 'a' && e.KeyChar <= 'z' || e.KeyChar >= 'A' && e.KeyChar <= 'Z' || e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtcompname_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private void txtintime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);

        }

        private void txtouttime_KeyPress(object sender, KeyPressEventArgs e)
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        byte[] stdbytes; Int64 std;
         OpenFileDialog open = new OpenFileDialog();
        private void butlogo_Click(object sender, EventArgs e)
        {
            try
            {
                stdbytes = null;
              
               
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.png;*.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {

                        pictureBox1.Image = new Bitmap(open.FileName);
                        stdbytes = Models.Device.ImageToByteArray(pictureBox1);
                        System.Text.Encoding enc = System.Text.Encoding.ASCII;
                        std = Convert.ToInt64("0" + stdbytes.Length);

                    }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
              
                panel2.BackColor = Class.Users.BackColors;
                panel3.BackColor = Class.Users.BackColors;
                butheader.BackColor = Class.Users.BackColors;
                this.BackColor = Class.Users.BackColors;
                butheader.Text = Class.Users.ScreenName;
            }
            else
            {
                panel2.BackColor = Class.Users.BackColors;
                panel3.BackColor = Class.Users.BackColors;
                butheader.BackColor = Class.Users.BackColors;
                this.BackColor = Class.Users.BackColors;
            }
        }

        private void txtphoneno_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtphoneno.Lines.Length > 15)
            {
                txtphoneno.ScrollBars = ScrollBars.Vertical;
            }
            else
            {
                txtphoneno.ScrollBars = ScrollBars.None;
            }
        }

        private void txtaddress_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtaddress.Lines.Length > 3)
            {
                txtaddress.ScrollBars = ScrollBars.Vertical;
            }
            else
            {
                txtaddress.ScrollBars = ScrollBars.None;
            }
        }

        private void ucDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DGBankDtl_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
           
            //if (DGBankDtl.Columns[e.ColumnIndex].Name == "BankName" || DGBankDtl.Columns[e.ColumnIndex].Name == "AccountNo" || DGBankDtl.Columns[e.ColumnIndex].Name == "IFSCCode" || DGBankDtl.Columns[e.ColumnIndex].Name == "Branch" || DGBankDtl.Columns[e.ColumnIndex].Name == "OpeningBal")
            //{
            //    CommonFunctions.Enter_Cells(DGBankDtl, TxtCell);
            //}
            //else if (DGBankDtl.Columns[e.ColumnIndex].Name == "AsOnDate")
            //{
            //    CommonFunctions.Enter_Cells(DGBankDtl, DtpCell);
            //}
        }

        private void TxtCell_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    //if (sender == checkactive)
            //    //{
            //    //    tabControl1.SelectedTab = tabPage1;
            //    //}
            //    if (sender == TxtCell)
            //    {
            //        if (DGBankDtl.Columns[DGBankDtl.CurrentCell.ColumnIndex].Name == "BankName" && CommonFunctions.Chk_GridDuplicate(DGBankDtl, TxtCell.Text.Trim(), "AccountNo") == true)
            //        {
            //            if (TxtCell.Text.Trim() != "")
            //            {
            //                CommonFunctions.Move_GridCtrl(sender, DGBankDtl, "AccountNo");
            //            }
            //            else if (DGBankDtl.CurrentCell.RowIndex != 0)
            //            {
                           
            //            }
            //        }
            //        else if (DGBankDtl.Columns[DGBankDtl.CurrentCell.ColumnIndex].Name == "AccountNo" && TxtCell.Text.Trim() != "" && CommonFunctions.Chk_GridDuplicate(DGBankDtl, TxtCell.Text.Trim(), "AccountNo") == true)
            //        {
            //            CommonFunctions.Move_GridCtrl(sender, DGBankDtl, "IFSCCode");
            //        }
            //        else if (DGBankDtl.Columns[DGBankDtl.CurrentCell.ColumnIndex].Name == "IFSCCode" && TxtCell.Text.Trim() != "" && CommonFunctions.Chk_GridDuplicate(DGBankDtl, TxtCell.Text.Trim(), "AccountNo") == true)
            //        {
            //            CommonFunctions.Move_GridCtrl(sender, DGBankDtl, "Branch");
            //        }
            //        else if (DGBankDtl.Columns[DGBankDtl.CurrentCell.ColumnIndex].Name == "Branch" && TxtCell.Text.Trim() != "" && CommonFunctions.Chk_GridDuplicate(DGBankDtl, TxtCell.Text.Trim(), "AccountNo") == true)
            //        {
            //            CommonFunctions.Move_GridCtrl(sender, DGBankDtl, "OpeningBal");
            //        }
            //        else if (DGBankDtl.Columns[DGBankDtl.CurrentCell.ColumnIndex].Name == "OpeningBal" && TxtCell.Text.Trim() != "" && CommonFunctions.Chk_GridDuplicate(DGBankDtl, TxtCell.Text.Trim(), "AccountNo") == true)
            //        {
            //            CommonFunctions.Move_GridCtrl(sender, DGBankDtl, "AsOnDate");
            //        }
            //    }
            //    else if (sender == DtpCell)
            //    {
            //        if (DGBankDtl.Columns[DGBankDtl.CurrentCell.ColumnIndex].Name == "AsOnDate" && DtpCell.Text.Trim() != "" && CommonFunctions.Chk_GridDuplicate(DGBankDtl, TxtCell.Text.Trim(), "AccountNo") == true)
            //        {
            //            CommonFunctions.Move_GridCtrl(sender, DGBankDtl, "BankName", true);
            //        }
            //    }
            //}
        }

        private void DGBankDtl_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center

            };
            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
          
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void txteveteato_TextChanged(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }
       
        private void txtintime_TextChanged(object sender, EventArgs e)
        {
         mas.DateTimeChange(txtintime);
        }

        private void txtintime_KeyPress_1(object sender, KeyPressEventArgs e)
        {

        }

        private void txtouttime_TextChanged(object sender, EventArgs e)
        {
            mas.DateTimeChange(txtouttime);
        }

        private void txtouttime_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);

        }

        private void txtlunchin_TextChanged(object sender, EventArgs e)
        {
            mas.DateTimeChange(txtlunchin);

        }

        private void txtlunchin_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);

        }

        private void txtlunchout_TextChanged(object sender, EventArgs e)
        {
            mas.DateTimeChange(txtlunchout);

        }

        private void txtlunchout_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);

        }

        private void combobank_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combobank.SelectedValue != null && combobank.SelectedValue.ToString() != "System.Data.DataRowView" && combobank.Text != "")
            {

                string sel1 = "select distinct a.ifsc,a.branch from asptblifscmas a  where a.bankname='" + combobank.SelectedValue + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblifscmas");
                DataTable dt = ds.Tables["asptblifscmas"];
                if (dt.Rows.Count > 0)
                {
                    txtifsc.Text = dt.Rows[0]["ifsc"].ToString();
                    txtbranch.Text= dt.Rows[0]["branch"].ToString();
                }

            }
        }

        private void txtgstdate_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == '/' || e.KeyChar == (char)Keys.Back);
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void bankToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bankname();
        }
    }
}

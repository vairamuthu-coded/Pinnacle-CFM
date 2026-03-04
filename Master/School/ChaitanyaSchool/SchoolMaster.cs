using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master.School.ChaitanyaSchool
{
    public partial class SchoolMaster : Form, ToolStripAccess
    {
        private static SchoolMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Employee em = new Models.Employee();
        Models.UserRights sm = new Models.UserRights();
        byte[] bytes; ListView listfilter = new ListView();
        OpenFileDialog open = new OpenFileDialog(); Int64 myString = 0; int i = 0;
        public static SchoolMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new SchoolMaster();
                return _instance;
            }
        }
        public SchoolMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }

        private void SchoolMaster_Load(object sender, EventArgs e)
        {
           
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
            string sel = "select max(a.asptblschoolmasid) as id  from  asptblschoolmas a  where a.compcode='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblschoolmas");
            DataTable dt = ds.Tables["asptblschoolmas"];
            Int64 partycount = 1;
            if (dt.Rows[0]["id"].ToString() == "")
            {

                txtschoolid1.Text = "1";

            }
            else
            {
                Int64 cc = Convert.ToInt64("0" + dt.Rows[0]["id"].ToString()) + partycount;
                txtschoolid1.Text = cc.ToString();
            }
        }
        public DataTable state(Int64 s)
        {
            string sel = "select  '0' gtstatemastid,'' statename from  DUAL  UNION ALL select b.gtstatemastid,b.statename from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where b.active='T' and a.gtcitymastid='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];
            return dt;
        }
        public DataTable city()
        {
            string sel = "select  '0' gtcitymastid,'' cityname from  DUAL  UNION ALL select a.gtcitymastid,a.cityname from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where a.active='T';";
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
                    MessageBox.Show("'CompCode  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtcompcode.Focus();
                    return;
                }
                if (txtcompname.Text == "")
                {
                    MessageBox.Show("'CompName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtcompname.Focus();
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
                if (txtintime.Text == "")
                {
                    MessageBox.Show("'InTime  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtintime.Focus(); return;
                }
                if (txtouttime.Text == "")
                {
                    MessageBox.Show("'OutTime  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtouttime.Focus(); return;
                }
                else
                {
                    Models.Validate va = new Models.Validate();
                    if (txtcompcode.Text != "" && txtcompname.Text != "")
                    {
                        string chk = "", chkcustomer = "";
                        //byte[] bytes1 = bytes;
                        myString = Convert.ToInt64("0" + bytes.Length);
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        if (checkcustomer.Checked == true) { chkcustomer = "PARTY"; } else { chkcustomer = "SCHOOL"; checkcustomer.Checked = false; }
                        DateTime mteain = DateTime.Parse(txtbreakfrom.Text);
                        DateTime mteaout = DateTime.Parse(txtbreakto.Text);

                        DateTime lin = DateTime.Parse(txtlunchout.Text);
                        DateTime lout = DateTime.Parse(txtlunchin.Text);

                        DateTime eteain = DateTime.Parse(txteveteafrom.Text);
                        DateTime eteaout = DateTime.Parse(txteveteato.Text);

                        TimeSpan mdiffer = mteaout.Subtract(mteain);
                        TimeSpan ldiffer = lin.Subtract(lout);
                        TimeSpan ddiffer = eteaout.Subtract(eteain);

                        TimeSpan totdiffer = mdiffer + ldiffer + ddiffer;
                        txttothrs.Text = totdiffer.ToString();
                        string sel = "select asptblschoolmasid    from  asptblschoolmas   WHERE compcode='" + txtcompcode.Text + "' and compname='" + txtcompname.Text + "' and city='" + combocity.SelectedValue + "' and state='" + combostate.SelectedValue + "' and country='" + combocountry.SelectedValue + "' and address='" + txtaddress.Text + "' and pincode='" + txtpincode.Text + "' and panno='" + txtpanno.Text + "' and tinno='" + txttinno.Text + "' and gstno='" + txtgstno.Text + "' and gstdate='" + txtgstdate.Text + "' and phoneno='" + txtphoneno.Text + "' and email='" + txtemail.Text + "' and website='" + txtwebsite.Text + "'  and contactname='" + txtcontact.Text + "' and ptransaction='" + chkcustomer + "' and active='" + chk + "' and intime='" + txtintime.Text + "' and outtime='" + txtouttime.Text + "'  and lunchfrom='" + txtlunchin.Text + "' and lunchto='" + txtlunchout.Text + "' and breakfrom='" + txtbreakfrom.Text + "' and breakto='" + txtbreakto.Text + "' and tpr='" + txttpr.Text + "' and eventeafrom='" + txteveteafrom.Text + "' and eventeato='" + txteveteato.Text + "' AND PERMISSIONHRS='" + txttothrs.Text + "' and IMAGEBYTES='"+myString+"'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblschoolmas");
                        DataTable dt = ds.Tables["asptblschoolmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtschoolid.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtschoolid.Text) == 0 || Convert.ToInt32("0" + txtschoolid.Text) == 0)
                        {
                           
                            string ins = "insert into asptblschoolmas(EMPIMAGE,compcode,compname,city,state,country,address,pincode,panno,tinno,gstno,gstdate,phoneno,email,website,contactname,active,ptransaction,intime,outtime,createdby,modifiedby,ipaddress,lunchfrom,lunchto,breakfrom,breakto,tpr,eventeafrom,eventeato,PERMISSIONHRS,EMPIMAGE,IMAGEBYTES)  VALUES(@EMPIMAGE,'" + txtcompcode.Text.ToUpper() + "','" + txtcompname.Text.ToUpper() + "', '" + combocity.SelectedValue + "', '" + combostate.SelectedValue + "','" + combocountry.SelectedValue + "' ,'" + txtaddress.Text.ToUpper() + "','" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "','" + txtpanno.Text + "','" + txttinno.Text.ToUpper() + "','" + txtgstno.Text.ToUpper() + "','" + txtgstdate.Text + "','" + txtphoneno.Text + "','" + txtemail.Text.ToUpper() + "','" + txtwebsite.Text.ToUpper() + "','" + txtcontact.Text.ToUpper() + "','" + chk + "', '" + chkcustomer + "' , '" + txtintime.Text + "' ,'" + txtouttime.Text + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "','" + txtlunchin.Text + "' ,'" + txtlunchout.Text + "' ,'" + txtbreakfrom.Text + "' ,'" + txtbreakto.Text + "','" + txttpr.Text + "','" + txteveteafrom.Text + "','" + txteveteato.Text + "','" + txttothrs.Text + "','" + myString + "');";
                            Utility.ExecuteNonQuery(ins, bytes);
                            MessageBox.Show("Record Saved Successfully " + txtschoolid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                      
                            string up = "update  asptblschoolmas  set  EMPIMAGE=@EMPIMAGE,compcode='" + txtcompcode.Text.ToUpper() + "',compname='" + txtcompname.Text.ToUpper() + "', city='" + combocity.SelectedValue + "', state='" + combostate.SelectedValue + "',country='" + combocountry.SelectedValue + "' ,address='" + txtaddress.Text.ToUpper() + "',pincode='" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "',panno='" + txtpanno.Text.ToUpper() + "',tinno='" + txttinno.Text.ToUpper() + "',gstno='" + txtgstno.Text.ToUpper() + "',gstdate='" + txtgstdate.Text + "',phoneno='" + txtphoneno.Text + "',email='" + txtemail.Text.ToUpper() + "',website='" + txtwebsite.Text.ToUpper() + "',contactname='" + txtcontact.Text.ToUpper() + "',ptransaction='" + chkcustomer + "', active='" + chk + "' , intime='" + txtintime.Text + "' , outtime='" + txtouttime.Text + "', modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' , lunchfrom='" + txtlunchin.Text + "' , lunchto='" + txtlunchout.Text + "' , breakfrom='" + txtbreakfrom.Text + "' , breakto='" + txtbreakto.Text + "', tpr='" + txttpr.Text + "' , eventeafrom='" + txteveteafrom.Text + "' , eventeato='" + txteveteato.Text + "' ,  PERMISSIONHRS='" + txttothrs.Text + "' ,IMAGEBYTES='" + myString + "' where  asptblschoolmasid='" + txtschoolid.Text + "';";
                             Utility.ExecuteNonQuery(up, bytes);
                            

                            MessageBox.Show("Record Updated Successfully " + txtschoolid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void CompanyMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       
        public void News()
        {

            GridLoad(); city(); empty();
        }
        private void empty()
        {
            txtschoolid.Text = "";
            txtschoolid1.Text = ""; txtcompcode.Select();pictureBox1.Image = null;
            combostate.SelectedIndex = -1; combostate.Text = "";
            combocountry.SelectedIndex = -1; combocountry.Text = ""; txtsearch.Text = "";
            checkactive.Checked = true;
            checkcustomer.Checked = false;
            txtcompcode.Text = ""; txttpr.Text = "";
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
            txtcontact.Text = ""; txttothrs.Text = "";
            txtintime.Text = ""; txtouttime.Text = ""; txteveteafrom.Text = ""; txteveteato.Text = "";
            txtbreakfrom.Text = ""; txtbreakto.Text = ""; txtlunchout.Text = ""; txtlunchin.Text = "";
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors; panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "select a.asptblschoolmasid,a.compcode,a.compname,b.cityname as city,c.statename as state, a.active,a.ptransaction,a.intime,a.outtime  from asptblschoolmas a join gtcitymast b on a.city = b.gtcitymastid  join gtstatemast c on a.state = c.gtstatemastid join gtcountrymast d on a.country = d.gtcountrymastid   order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblschoolmas");
                DataTable dt = ds.Tables["asptblschoolmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblschoolmasid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["compname"].ToString());
                        list.SubItems.Add(myRow["city"].ToString());
                        list.SubItems.Add(myRow["state"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        list.SubItems.Add(myRow["ptransaction"].ToString());
                        list.SubItems.Add(myRow["intime"].ToString());
                        list.SubItems.Add(myRow["outtime"].ToString());
                        listView1.Items.Add(list);
                        list.BackColor = i  % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
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

                    txtschoolid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    //txtcompid1.Text = listView1.SelectedItems[0].SubItems[2].Text;

                    string sel1 = "select a.tpr, a.asptblschoolmasid,a.compcode,a.compname,b.cityname as city,c.statename as state,D.countryname as  country,a.address,  a.pincode,a.panno,a.tinno,  a.gstno,'' as gstdate,  a.phoneno, a.email,  a.website,  a.active,a.ptransaction,a.intime,a.outtime,a.lunchfrom,a.lunchto,a.breakfrom,a.breakto,a.eventeafrom,a.eventeato,A.PERMISSIONHRS,A.EMPIMAGE from asptblschoolmas a join gtcitymast b on a.city = b.gtcitymastid  join gtstatemast c on a.state = c.gtstatemastid join gtcountrymast d on a.country = d.gtcountrymastid   where a.asptblschoolmasid='" + txtschoolid.Text + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblschoolmas");
                    DataTable dt = ds.Tables["asptblschoolmas"];

                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow myRow in dt.Rows)
                        {
                            txttpr.Text = Convert.ToString(myRow["tpr"].ToString());
                            txtschoolid.Text = Convert.ToString(myRow["asptblschoolmasid"].ToString());
                            txtschoolid1.Text = Convert.ToString(myRow["asptblschoolmasid"].ToString());
                            txtcompcode.Text = Convert.ToString(myRow["compcode"].ToString());
                            txtcompname.Text = Convert.ToString(myRow["compname"].ToString());
                            combocity.Text = Convert.ToString(myRow["city"].ToString());
                            Combocity_SelectedIndexChanged(sender, e);
                            combostate.Text = Convert.ToString(myRow["state"].ToString());
                            combocountry.Text = Convert.ToString(myRow["country"].ToString());
                            txtaddress.Text = Convert.ToString(myRow["address"].ToString());
                            txtpincode.Text = Convert.ToString(myRow["pincode"].ToString());
                            txtpanno.Text = Convert.ToString(myRow["panno"].ToString());
                            txttinno.Text = Convert.ToString(myRow["tinno"].ToString());
                            txtgstno.Text = Convert.ToString(myRow["gstno"].ToString());
                            txtgstdate.Text = Convert.ToString(myRow["gstdate"].ToString());
                            txtphoneno.Text = Convert.ToString(myRow["phoneno"].ToString());
                            txtemail.Text = Convert.ToString(myRow["email"].ToString());
                            txtwebsite.Text = Convert.ToString(myRow["website"].ToString());
                            txtintime.Text = Convert.ToString(myRow["intime"].ToString());
                            txtouttime.Text = Convert.ToString(myRow["outtime"].ToString());

                            if (myRow["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                            if (myRow["ptransaction"].ToString() == "PARTY") { checkcustomer.Checked = true; } else { checkcustomer.Checked = false; }

                            txtlunchin.Text = Convert.ToString(myRow["lunchfrom"].ToString());
                            txtlunchout.Text = Convert.ToString(myRow["lunchto"].ToString());
                            txtbreakfrom.Text = Convert.ToString(myRow["breakfrom"].ToString());
                            txtbreakto.Text = Convert.ToString(myRow["breakto"].ToString());
                            txteveteafrom.Text = Convert.ToString(myRow["eventeafrom"].ToString());
                            txteveteato.Text = Convert.ToString(myRow["eventeato"].ToString());
                            txttothrs.Text = Convert.ToString(myRow["PERMISSIONHRS"].ToString());
                           
                            //Combostate_SelectedIndexChanged(sender, e);
                            if (myRow["EMPIMAGE"].ToString() != "")
                            {

                                bytes = (byte[])myRow["EMPIMAGE"];
                                Image img = Models.Device.ByteArrayToImage(bytes);
                                pictureBox1.Image = img;


                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            txtcompcode.Select();
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtsearch.Text.ToUpper() != "")
                {
                    listView1.Items.Clear(); int iGLCount = 1;
                    string sel1 = "select a.asptblschoolmasid,a.compcode,a.compname,b.cityname as city,c.statename as state, a.active,a.ptransaction,a.intime,a.outtime   from  asptblschoolmas a  join gtcitymast b on a.city=b.gtcitymastid  join gtstatemast c on a.state=c.gtstatemastid   join gtcountrymast d on c.country=d.gtcountrymastid where a.compcode LIKE'%" + txtsearch.Text.ToUpper() + "%'  || a.compname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.ptransaction LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.pincode LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.panno LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.gstno LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.phoneno LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.contactname LIKE'%" + txtsearch.Text.ToUpper() + "%';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblschoolmas");
                    DataTable dt = ds.Tables["asptblschoolmas"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(iGLCount.ToString());
                            list.SubItems.Add(myRow["asptblschoolmasid"].ToString());
                            list.SubItems.Add(myRow["compcode"].ToString());
                            list.SubItems.Add(myRow["compname"].ToString());
                            list.SubItems.Add(myRow["city"].ToString());
                            list.SubItems.Add(myRow["state"].ToString());
                            list.SubItems.Add(myRow["active"].ToString());
                            list.SubItems.Add(myRow["ptransaction"].ToString());
                            list.SubItems.Add(myRow["intime"].ToString());
                            list.SubItems.Add(myRow["outtime"].ToString());
                            list.BackColor = iGLCount % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);
                            iGLCount++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }
                    else
                    {
                        listView1.Items.Clear();
                    }
                }
                else
                {

                    listView1.Items.Clear();
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Combostate_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = country(Convert.ToInt64(combostate.SelectedValue));
                combocountry.DataSource = dt;
                combocountry.DisplayMember = "countryname";
                combocountry.ValueMember = "gtcountrymastid";
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
                combostate.ValueMember = "gtstatemastid";
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
            if (txtschoolid.Text == "")
            {

                compid(txtcompcode.Text);

            }
        }

        private void Deletes_Click(object sender, EventArgs e)
        {
            if (txtschoolid.Text != "")
            {
                string sel1 = "select a.asptblschoolmasid from asptblschoolmas a join HRPayDetails b on a.asptblschoolmasid=b.compname where a.asptblschoolmasid='" + txtschoolid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblschoolmas");
                DataTable dt = ds.Tables["asptblschoolmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcompcode.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblschoolmas where asptblschoolmasid='" + Convert.ToInt64("0" + txtschoolid.Text) + "'";
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
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtouttime_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
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
            try
            {
                bytes = null;
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {

                        p.Image = new Bitmap(open.FileName);
                        bytes = Models.Device.ImageToByteArray(p);
                        System.Text.Encoding enc = System.Text.Encoding.ASCII;
                        myString = Convert.ToInt64("0" + bytes.Length);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void combocountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            Combostate_SelectedIndexChanged(sender,e);
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}
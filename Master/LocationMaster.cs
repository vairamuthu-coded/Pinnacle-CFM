using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class LocationMaster : Form, ToolStripAccess
    {
        public LocationMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
        }

        private static LocationMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights(); ListView listfilter = new ListView();
        public static LocationMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new LocationMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
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
        private void locationMaster_Load(object sender, EventArgs e)
        {
            News();
        }
        public void compcodeload()
        {
            try
            {
                string sel = "select gtcompmastid,compcode from  gtcompmast where active='T'  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;

            }
            catch (Exception EX)
            { }
        }
        public void compcodeload(string s)
        {
            try
            {
                string sel = "select a.gtcompmastid,a.compname from  gtcompmast a where a.active='T' and a.compcode='"+s+"' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                txtcompname.Text = dt.Rows[0]["compname"].ToString();

            }
            catch (Exception EX)
            { }
        }
        public void Saves()
        {
            try
            {
                if (combocompcode.Text == "")
                {
                    MessageBox.Show("CompCode is empty " + " Alert " + combocompcode.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    combocompcode.Select();
                    return;
                }
                if (txtunit.Text == "")
                {
                    MessageBox.Show("location Name is empty " + " Alert " + txtunit.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtunit.Select();
                    return;
                }
                if (txtlocation.Text == "")
                {
                    MessageBox.Show("location Name is empty " + " Alert " + txtunit.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtlocation.Select();
                    return;
                }
                if (txtunit.Text != "" && txtlocation.Text != "")
                {


                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptbllocmasid    from  asptbllocmas    WHERE   compcode='" + combocompcode.SelectedValue + "' and unit='" + txtunit.Text + "' and location='" + txtlocation.Text + "' and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllocmas");
                    DataTable dt = ds.Tables["asptbllocmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtunit.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtlocationid.Text) == 0 || Convert.ToInt32("0" + txtlocationid.Text) == 0)
                    {
                        string ins = "insert into asptbllocmas(compcode,unit,location,active,compcode1,username,createdby,modifiedby,ipaddress)  VALUES('" + combocompcode.SelectedValue + "','" + txtunit.Text.ToUpper() + "','" + txtlocation.Text.ToUpper() + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtunit.Text.ToUpper(), " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptbllocmas  set   compcode='" + combocompcode.SelectedValue + "' , unit='" + txtunit.Text.ToUpper() + "',location='" + txtlocation.Text.ToUpper() + "',active='" + chk + "' , modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbllocmasid='" + txtlocationid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtunit.Text.ToUpper(), " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty();
                    }

                }
                else
                {
                    MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("location " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void locationMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        public void News()
        {

            empty(); GridLoad(); compcodeload();
        }
        private void empty()
        {
            txtlocationid.Text = "";
            txtunit.Text = "";
            combocompcode.Text = "";
            txtlocation.Text = "";
            checkactive.Checked = false;
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
                listView1.Items.Clear();
                string sel1 = "SELECT a.asptbllocmasid,b.compcode, a.location ,a.unit, a.active  FROM  asptbllocmas a  join gtcompmast b on b.gtcompmastid=a.compcode  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllocmas");
                DataTable dt = ds.Tables["asptbllocmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbllocmasid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["location"].ToString());                       
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count    :" + listView1.Items.Count;
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

                    txtlocationid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptbllocmasid,b.compcode, a.location ,a.unit, a.active    from  asptbllocmas a JOIN gtcompmast b on a.compcode=b.gtcompmastid   where a.asptbllocmasid=" + txtlocationid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllocmas");
                    DataTable dt = ds.Tables["asptbllocmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtlocationid.Text = Convert.ToString(dt.Rows[0]["asptbllocmasid"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        txtunit.Text = Convert.ToString(dt.Rows[0]["unit"].ToString());
                        txtlocation.Text = Convert.ToString(dt.Rows[0]["location"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }


                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int i = 1;
                if (txtsearch.Text.ToUpper() != "")
                {

                    int item0 = 0; listView1.Items.Clear();
                    if (txtsearch.Text.Length > 1)
                    {
                        i = 1;
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            if (item.SubItems[4].ToString().Contains(txtsearch.Text))
                            {


                             
                                list.SubItems.Add(i.ToString());
                                list.SubItems.Add(item.SubItems[2].Text);
                                list.SubItems.Add(item.SubItems[3].Text);
                                list.SubItems.Add(item.SubItems[4].Text);
                                list.SubItems.Add(item.SubItems[5].Text);
                               
                                list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                                listView1.Items.Add(list);


                            }
                            i++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }
                    else
                    {

                        ListView ll = new ListView();
                        i = 1;
                        listView1.Items.Clear(); item0 = listfilter.Items.Count;
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                           
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                            i++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
        }

        public void Deletes()
        {
            if (txtlocationid.Text != "")
            {
                string sel1 = "select a.asptbllocmasid from asptbllocmas a join gtstatemast b on a.asptbllocmasid=b.country where a.asptbllocmasid='" + txtlocationid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllocmas");
                DataTable dt = ds.Tables["asptbllocmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtunit.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptbllocmas where asptbllocmasid='" + Convert.ToInt64("0" + txtlocationid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtunit.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void refeshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            GridLoad();
        }

        private void txtunit_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= 'a' && e.KeyChar <= 'z' || e.KeyChar >= 'A' && e.KeyChar <= 'Z' || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtunit_TextChanged(object sender, EventArgs e)
        {

            txtlocation.Text = combocompcode.Text + "-" + txtunit.Text;
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            compcodeload(combocompcode.Text);
        }
    }
}

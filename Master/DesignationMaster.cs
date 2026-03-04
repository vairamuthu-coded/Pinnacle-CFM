using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class DesignationMaster : Form,ToolStripAccess
    {
        private static DesignationMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        ListView listfilter = new ListView();

        public static DesignationMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DesignationMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public DesignationMaster()
        {
            InitializeComponent();
            
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
        }
        public void usercheck(string s, string ss, string sss)
        {

           

        }
        private void DesignationMaster_Load(object sender, EventArgs e)
        {
            GridLoad();empty(); txtdesignation.Select();
        }
        private bool ValidateContact(TextBox s)
        {
            if (!Regex.Match(s.Text, "^[A-Z][a-zA-Z]*$").Success)
            {

                s.Text.Remove(s.Text.Length - 1);

            }
            return true;
        }

        public void Saves()
        {
            try
            {
                if (txtdesignation.Text != "")
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsStringSpace(txtdesignation.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select asptbldesigmasid    from  asptbldesigmas    WHERE designation='" + txtdesignation.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbldesigmas");
                        DataTable dt = ds.Tables["asptbldesigmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtdesignation.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtdesignationid.Text) == 0 || Convert.ToInt32("0" + txtdesignationid.Text) == 0)
                        {
                            string ins = "insert into asptbldesigmas(designation,active,createdby,modifiedby,ipaddress)  VALUES('" + txtdesignation.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtdesignation.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptbldesigmas  set   designation='" + txtdesignation.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbldesigmasid='" + txtdesignationid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtdesignation.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("'country  is Wrong'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("country " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void DesignationMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

 

        public void News()
        {
            GridLoad();
            empty();
        }
        private void empty()
        {
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            txtdesignationid.Text = "";
            txtdesignation.Text = "";
            txtdesignation.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = " select a.asptbldesigmasid, a.designation , a.active  from  asptbldesigmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldesigmas");
                DataTable dt = ds.Tables["asptbldesigmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldesigmasid"].ToString());
                        list.SubItems.Add(myRow["designation"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        if (i % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }
                        this.listfilter.Items.Add((ListViewItem)list.Clone());             
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

                    txtdesignationid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptbldesigmasid, a.designation as country , a.active    from  asptbldesigmas a    where a.asptbldesigmasid=" + txtdesignationid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldesigmas");
                    DataTable dt = ds.Tables["asptbldesigmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtdesignationid.Text = Convert.ToString(dt.Rows[0]["asptbldesigmasid"].ToString());
                        txtdesignation.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
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
                if (txtsearch.Text.ToUpper() != "")
                {

                    int item0 = 0; listView1.Items.Clear();
                    if (txtsearch.Text.Length > 1)
                    {

                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text.ToUpper()))
                            {


                                list.Text = listfilter.Items[item0].SubItems[0].Text;
                                list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                                if (item0 % 2 == 0)
                                {
                                    list.BackColor = System.Drawing.Color.White;

                                }
                                else
                                {
                                    list.BackColor = System.Drawing.Color.WhiteSmoke;
                                }
                                listView1.Items.Add(list);


                            }
                            item0++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }
                    else
                    {

                      ListView ll = new ListView();

                        listView1.Items.Clear(); item0 = listfilter.Items.Count;
                        foreach (ListViewItem item in listfilter.Items)
                        {

                            this.listView1.Items.Add((ListViewItem)item.Clone());
                            if (item0 % 2 == 0)
                            {
                                item.BackColor = System.Drawing.Color.White;

                            }
                            else
                            {
                                item.BackColor = System.Drawing.Color.WhiteSmoke;
                            }

                            item0++;
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
            if (txtdesignationid.Text != "")
            {
                string sel1 = "select a.asptbldesigmasid from asptbldesigmas a join gtstatemast b on a.asptbldesigmasid=b.country where a.asptbldesigmasid='" + txtdesignationid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldesigmas");
                DataTable dt = ds.Tables["asptbldesigmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtdesignation.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptbldesigmas where asptbldesigmasid='" + Convert.ToInt64("0" + txtdesignationid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtdesignation.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
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

       
    }
}

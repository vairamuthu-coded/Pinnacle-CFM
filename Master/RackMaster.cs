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
    public partial class RackMaster : Form, ToolStripAccess
    {
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        private static RackMaster _instance; ListView listfilter = new ListView();
        public static RackMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new RackMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public RackMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.UserTime = 0;
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
           
           
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

            }

        }


        private void RackMaster_Load(object sender, EventArgs e)
        {
            News();
        }


        public void Saves()
        {
            try
            {
                if (txtrack.Text == "")
                {
                    MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtrack.Focus();
                    return;
                }

                else
                {
                    if (txtrack.Text != "")
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.ASPTBLRACKMASID    from  ASPTBLRACKMAS a     WHERE a.RACK='" + txtrack.Text + "' and a.active='" + chk + "' and a.ASPTBLRACKMASID='" + txtrackid.Text + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLRACKMAS");
                        DataTable dt = ds.Tables["ASPTBLRACKMAS"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtrack.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtrackid.Text) == 0 || Convert.ToInt32("0" + txtrackid.Text) == 0)
                        {
                            string ins = "insert into ASPTBLRACKMAS(RACK,active,createdby,modifiedby,ipaddress)  VALUES('" + txtrack.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtrack.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  ASPTBLRACKMAS  set RACK='" + txtrack.Text.ToUpper() + "',  active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where ASPTBLRACKMASID='" + txtrackid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtrack.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("RACK " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void RackMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }

        public void News()
        {

            GridLoad(); empty();
        }
        private void empty()
        {
            txtrackid.Text = "";
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            Class.Users.UserTime = 0;
            txtrack.Text = ""; txtrack.Select();
            txtsearch.Text = "";
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.ASPTBLRACKMASID,a.RACK , a.active    from  ASPTBLRACKMAS a     order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLRACKMAS");
                DataTable dt = ds.Tables["ASPTBLRACKMAS"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLRACKMASID"].ToString());                      
                        list.SubItems.Add(myRow["RACK"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listView1.Items.Add(list);
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
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
                Class.Users.UserTime = 0;
                if (listView1.Items.Count > 0)
                {

                    txtrackid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.ASPTBLRACKMASID,a.RACK , a.active    from  ASPTBLRACKMAS a     where a.ASPTBLRACKMASID=" + txtrackid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLRACKMAS");
                    DataTable dt = ds.Tables["ASPTBLRACKMAS"];

                    if (dt.Rows.Count > 0)
                    {
                        txtrackid.Text = Convert.ToString(dt.Rows[0]["ASPTBLRACKMASID"].ToString());
                        txtrack.Text = Convert.ToString(dt.Rows[0]["RACK"].ToString());
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
                int item0 = 0;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text))
                        {
                            list.SubItems.Add(item.SubItems[0].Text);
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {
                    try
                    {
                        listView1.Items.Clear(); item0 = 0;
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(item.SubItems[0].Text);
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                            item0++;
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("---" + ex.ToString());
            }
        }

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        public void Deletes()
        {
            if (txtrackid.Text != "")
            {
                string sel1 = "select a.ASPTBLRACKMASID from ASPTBLRACKMAS a  join asptblbuysaminw b on a.ASPTBLRACKMASID=b.rack  where a.ASPTBLRACKMASID='" + txtrackid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLRACKMAS");
                DataTable dt = ds.Tables["ASPTBLRACKMAS"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtrack.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
                else
                {
                    string del = "delete from ASPTBLRACKMAS where ASPTBLRACKMASID='" + Convert.ToInt64("0" + txtrackid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtrack.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty(); return;
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

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad(); usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void Searchs(int id)
        {
            throw new NotImplementedException();
        }

        public void Deletes(int id)
        {
            throw new NotImplementedException();
        }
    }
}

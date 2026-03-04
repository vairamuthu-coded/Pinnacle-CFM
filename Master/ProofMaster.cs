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
    public partial class ProofMaster : Form, ToolStripAccess
    {
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        private static ProofMaster _instance; ListView listfilter = new ListView();
        public static ProofMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ProofMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public ProofMaster()
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


        private void ProofMaster_Load(object sender, EventArgs e)
        {
            News();
        }


        public void Saves()
        {
            try
            {
                if (txtproof.Text == "")
                {
                    MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtproof.Focus();
                    return;
                }

                else
                {
                    if (txtproof.Text != "")
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.asptblproofmasID    from  asptblproofmas a     WHERE a.ProofName='" + txtproof.Text + "' and a.active='" + chk + "' and a.asptblproofmasID='" + txtproofid.Text + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproofmas");
                        DataTable dt = ds.Tables["asptblproofmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtproof.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtproofid.Text) == 0 || Convert.ToInt32("0" + txtproofid.Text) == 0)
                        {
                            string ins = "insert into asptblproofmas(ProofName,active,createdby,modifiedby,ipaddress)  VALUES('" + txtproof.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtproof.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblproofmas  set ProofName='" + txtproof.Text.ToUpper() + "',  active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblproofmasID='" + txtproofid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtproof.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("ProofName " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void ProofMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            txtproofid.Text = "";
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            Class.Users.UserTime = 0;
            txtproof.Text = ""; txtproof.Select();
            txtsearch.Text = "";
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.asptblproofmasID,a.ProofName , a.active    from  asptblproofmas a     order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproofmas");
                DataTable dt = ds.Tables["asptblproofmas"];
                if (dt!= null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblproofmasID"].ToString());                      
                        list.SubItems.Add(myRow["ProofName"].ToString());
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

                    txtproofid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.asptblproofmasID,a.ProofName , a.active    from  asptblproofmas a     where a.asptblproofmasID=" + txtproofid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproofmas");
                    DataTable dt = ds.Tables["asptblproofmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtproofid.Text = Convert.ToString(dt.Rows[0]["asptblproofmasID"].ToString());
                        txtproof.Text = Convert.ToString(dt.Rows[0]["ProofName"].ToString());
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
                int item0 = 0, i = 1; ;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            if (item0 % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {
                    GridLoad();
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
            if (txtproofid.Text != "")
            {
                string sel1 = "select a.asptblproofmasID from asptblproofmas a  join asptblbuysaminw b on a.asptblproofmasID=b.ProofName  where a.asptblproofmasID='" + txtproofid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproofmas");
                DataTable dt = ds.Tables["asptblproofmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtproof.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
                else
                {
                    string del = "delete from asptblproofmas where asptblproofmasID='" + Convert.ToInt64("0" + txtproofid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtproof.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

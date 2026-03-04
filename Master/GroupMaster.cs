using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Pinnacle.UserControls;


namespace Pinnacle.Master
{
    public partial class GroupMaster : Form,ToolStripAccess
    {
        private static GroupMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        Pinnacle.UserControls.UCCListView ucclist=new UCCListView();
        public static GroupMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new GroupMaster();

                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public GroupMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
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
        private void GroupMaster_Load(object sender, EventArgs e)
        {
            GridLoad(); txtgroup.Select();
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
                if (txtgroup.Text != "")
                {
                   
                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select asptblgromasID    from  asptblgromas    WHERE groupname='" + txtgroup.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblgromas");
                        DataTable dt = ds.Tables["asptblgromas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtgroup.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtgroupid.Text) == 0 || Convert.ToInt32("0" + txtgroupid.Text) == 0)
                        {
                            string ins = "insert into asptblgromas(groupname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtgroup.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtgroup.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblgromas  set   groupname='" + txtgroup.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblgromasID='" + txtgroupid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtgroup.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                   
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("groupname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void GroupMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            this.Font = Class.Users.FontName;
            txtgroupid.Text = "";
            txtgroup.Text = "";
            txtgroup.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptblgromasID, a.groupname , a.active  from  asptblgromas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgromas");
                DataTable dt = ds.Tables["asptblgromas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblgromasID"].ToString());
                        list.SubItems.Add(myRow["groupname"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listView1.Items.Add(list);
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

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

                    txtgroupid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblgromasID, a.groupname , a.active    from  asptblgromas a    where a.asptblgromasID=" + txtgroupid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgromas");
                    DataTable dt = ds.Tables["asptblgromas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtgroupid.Text = Convert.ToString(dt.Rows[0]["asptblgromasID"].ToString());
                        txtgroup.Text = Convert.ToString(dt.Rows[0]["groupname"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Search()
        {
            try
            {
                if (txtsearch.Text.ToUpper() != "")
                {
                    listView1.Items.Clear(); int iGLCount = 1;
                    string sel1 = "  SELECT  a.asptblgromasID,a.groupname as groupname,a.active from asptblgromas a  where a.groupname LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgromas");
                    DataTable dt = ds.Tables["asptblgromas"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["asptblgromasID"].ToString());
                            list.SubItems.Add(myRow["groupname"].ToString());
                            list.SubItems.Add(myRow["active"].ToString());
                            listView1.Items.Add(list);
                            iGLCount++;
                        }
                        lbltotal.Text = "Total Count    :" + listView1.Items.Count;
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

        public void Deletes()
        {
            if (txtgroupid.Text != "")
            {
                string sel1 = "select a.asptblgromasID from asptblgromas a join gtstatemast b on a.asptblgromasID=b.groupname where a.asptblgromasID='" + txtgroupid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgromas");
                DataTable dt = ds.Tables["asptblgromas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtgroup.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblgromas where asptblgromasID='" + Convert.ToInt64("0" + txtgroupid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtgroup.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void GroupMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            _instance = null;

        }
    }
}

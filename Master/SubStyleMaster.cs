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
    public partial class SubStyleMaster : Form,ToolStripAccess
    {
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        private static SubStyleMaster _instance;
        ListView listfilter = new ListView();
        public static SubStyleMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new SubStyleMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public SubStyleMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

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


        private void SubStyleMaster_Load(object sender, EventArgs e)
        {
            News();
        }

        public DataTable STYLENAME()
        {
            string sel = " select a.ASPTBLSTYMASID, a.STYLENAME  , a.active from  ASPTBLSTYMAS a    where a.active='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLSTYMAS");
            DataTable dt = ds.Tables["ASPTBLSTYMAS"];
            combostyle.DataSource = dt;
            combostyle.DisplayMember = "STYLENAME";
            combostyle.ValueMember = "ASPTBLSTYMASID";
            return dt;
        }
        public void Saves()
        {
            try
            {
                if (txtsubstyle.Text == "")
                {
                    MessageBox.Show("'substyle Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtsubstyle.Focus();
                    return;
                }
                if (combostyle.Text == "")
                {
                    MessageBox.Show("'Style Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combostyle.Focus(); return;
                }
                else
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsString(txtsubstyle.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.ASPTBLSUBSTYMASID    from  ASPTBLSUBSTYMAS a join ASPTBLSTYMAS b on a.STYLENAME=b.ASPTBLSTYMASID   WHERE a.SUBSTYLE='" + txtsubstyle.Text + "' and b.STYLENAME='" + combostyle.Text + "' and a.active='" + chk + "' and a.ASPTBLSUBSTYMASID='" + txtsubstyleid.Text + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLSUBSTYMAS");
                        DataTable dt = ds.Tables["ASPTBLSUBSTYMAS"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtsubstyle.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtsubstyleid.Text) == 0 || Convert.ToInt32("0" + txtsubstyleid.Text) == 0)
                        {
                            string ins = "insert into ASPTBLSUBSTYMAS(SUBSTYLE,STYLENAME,active,createdby,modifiedby,ipaddress)  VALUES('" + txtsubstyle.Text.ToUpper() + "','" + combostyle.SelectedValue + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtsubstyle.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  ASPTBLSUBSTYMAS  set SUBSTYLE='" + txtsubstyle.Text.ToUpper() + "', STYLENAME='" + combostyle.SelectedValue + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where ASPTBLSUBSTYMASID='" + txtsubstyleid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtsubstyle.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("STYLENAME " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void SubStyleMaster_FormClosed(object sender, FormClosedEventArgs e)
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

            STYLENAME(); GridLoad(); empty();
        }
        private void empty()
        {
            txtsubstyleid.Text = "";
            txtsubstyle.Text = ""; txtsubstyle.Select(); combostyle.Text = "";
            combostyle.SelectedIndex = -1;
            txtsearch.Text = "";
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            txtsearch.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.ASPTBLSUBSTYMASID,a.SUBSTYLE , b.STYLENAME  , a.active    from  ASPTBLSUBSTYMAS a join ASPTBLSTYMAS b on a.STYLENAME=b.ASPTBLSTYMASID    order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSUBSTYMAS");
                DataTable dt = ds.Tables["ASPTBLSUBSTYMAS"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["ASPTBLSUBSTYMASID"].ToString());
                        list.SubItems.Add(myRow["substyle"].ToString());
                        list.SubItems.Add(myRow["STYLENAME"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listView1.Items.Add(list);
                        listfilter.Items.Add((ListViewItem)list.Clone());
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

                    txtsubstyleid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.ASPTBLSUBSTYMASID,a.SUBSTYLE , b.STYLENAME , a.active    from  ASPTBLSUBSTYMAS a join ASPTBLSTYMAS b on a.STYLENAME=b.ASPTBLSTYMASID    where a.ASPTBLSUBSTYMASID=" + txtsubstyleid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSUBSTYMAS");
                    DataTable dt = ds.Tables["ASPTBLSUBSTYMAS"];

                    if (dt.Rows.Count > 0)
                    {
                        txtsubstyleid.Text = Convert.ToString(dt.Rows[0]["ASPTBLSUBSTYMASID"].ToString());
                        txtsubstyle.Text = Convert.ToString(dt.Rows[0]["SUBSTYLE"].ToString());
                        combostyle.Text = Convert.ToString(dt.Rows[0]["STYLENAME"].ToString());
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
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
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
                    try
                    {
                        listView1.Items.Clear(); item0 = 0;
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            this.listView1.Items.Add((ListViewItem)item.Clone());
                            if (item0 % 2 == 0)
                            {
                                item.BackColor = Color.White;
                            }
                            else
                            {
                                item.BackColor = Color.WhiteSmoke;
                            }
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
            STYLENAME();
        }

        public void Deletes()
        {
            if (txtsubstyleid.Text != "")
            {
                string sel1 = "select a.ASPTBLSUBSTYMASID from ASPTBLSUBSTYMAS a join ASPTBLSTYMAS b on a.STYLENAME=b.ASPTBLSTYMASID where a.ASPTBLSUBSTYMASID='" + txtsubstyleid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSUBSTYMAS");
                DataTable dt = ds.Tables["ASPTBLSUBSTYMAS"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtsubstyle.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from ASPTBLSUBSTYMAS where ASPTBLSUBSTYMASID='" + Convert.ToInt64("0" + txtsubstyleid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtsubstyle.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

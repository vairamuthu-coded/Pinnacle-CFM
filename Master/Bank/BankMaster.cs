using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master.Bank
{
    public partial class BankMaster : Form, ToolStripAccess
    {
        private static BankMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights(); ListView listfilter = new ListView();
        public static BankMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new BankMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }

        }
        public BankMaster()
        {
            InitializeComponent();
            // usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());

            // panel2.BackColor = Class.Users.BackColors;
            //panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;

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


        //public void usercheck(string s, string ss, string sss)
        //{

        //    DataTable dt1 = sm.headerdropdowns(s, ss, sss);
        //    if (dt1.Rows.Count > 0)
        //    {
        //        if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
        //        {
        //            //for (int r = 0; r < dt1.Rows.Count; r++)
        //            //{

        //            //    if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
        //            //    if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
        //            //    if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
        //            //    if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
        //            //    if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
        //            //    if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
        //            //    if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
        //            //    if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
        //            //    if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
        //            //    if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
        //            //    if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
        //            //    if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
        //            //    if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
        //            //    if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }

        //            //}
        //        }


        //    }
        //    else
        //    {

        //    }

        //}
        private void BankMaster_Load(object sender, EventArgs e)
        {

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
                if (txtbank.Text != "")
                {


                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblbanmasID    from  asptblbanmas    WHERE bankname='" + txtbank.Text + "' and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbanmas");
                    DataTable dt = ds.Tables["asptblbanmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtbank.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtbankid.Text) == 0 || Convert.ToInt32("0" + txtbankid.Text) == 0)
                    {
                        string ins = "insert into asptblbanmas(bankname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtbank.Text.ToUpper().Trim() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtbank.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblbanmas  set   bankname='" + txtbank.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblbanmasID='" + txtbankid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtbank.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty();
                    }


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("bankname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void BankMaster_FormClosed(object sender, FormClosedEventArgs e)
        {

            _instance = null;

        }



        public void News()
        {

            empty();
            GridLoad();

        }
        private void empty()
        {
            txtbankid.Text = "";
            txtbank.Text = ""; Class.Users.UserTime = 0;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            txtbank.Select();
        }
        public void Searchs(int EditID)
        {
            try
            {
                if (EditID > 0)
                {

                    txtbankid.Text = EditID.ToString();
                    DataTable dt = Utility.SQLQuery("select a.asptblbanmasID, a.bankname  , a.active    from  asptblbanmas a    where a.asptblbanmasID=" + txtbankid.Text + "");
                    if (dt.Rows.Count > 0)
                    {
                        txtbankid.Text = Convert.ToString(dt.Rows[0]["asptblbanmasID"].ToString());
                        txtbank.Text = Convert.ToString(dt.Rows[0]["bankname"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void GridLoad()
        {
            Class.Users.SearchQuery = "select a.asptblbanmasID AS ID, a.bankname as BankName , a.ACTIVE as Active  from  asptblbanmas a   order by 1";
            Class.Users.HideCols = new string[] { "ID" };
            listView1.Load_Details();

        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            //try
            //{
            //    if (listView1.Items.Count > 0)
            //    {

            //        txtbankid.Text = listView1.SelectedItems[0].SubItems[1].Text;
            //        string sel1 = " select a.asptblbanmasID, a.bankname as country , a.active    from  asptblbanmas a    where a.asptblbanmasID=" + txtcountryid.Text;
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbanmas");
            //        DataTable dt = ds.Tables["asptblbanmas"];

            //        if (dt.Rows.Count > 0)
            //        {
            //            txtbankid.Text = Convert.ToString(dt.Rows[0]["asptblbanmasID"].ToString());
            //            txtbank.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
            //            if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        public void Search()
        {
            //try
            //{
            //    if (txtsearch.Text.ToUpper() != "")
            //    {
            //        listView1.Items.Clear(); int iGLCount = 1;
            //        string sel1 = "  SELECT  a.asptblbanmasID,a.bankname as country,a.active from asptblbanmas a  where a.bankname LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbanmas");
            //        DataTable dt = ds.Tables["asptblbanmas"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptblbanmasID"].ToString());
            //                list.SubItems.Add(myRow["country"].ToString());                          
            //                list.SubItems.Add(myRow["active"].ToString());
            //                listView1.Items.Add(list);
            //                iGLCount++;
            //            }
            //            lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            //        }
            //        else
            //        {
            //            listView1.Items.Clear();
            //        }
            //    }
            //    else
            //    {

            //        listView1.Items.Clear();
            //        GridLoad();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        public void Deletes()
        {
            if (txtbankid.Text != "")
            {
                string sel1 = "select a.asptblbanmasID from asptblbanmas a join asptblifscmas b on a.asptblbanmasID=b.bankname where a.asptblbanmasID='" + txtbankid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbanmas");
                DataTable dt = ds.Tables["asptblbanmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtbank.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblbanmas where asptblbanmasID='" + Convert.ToInt64("0" + txtbankid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtbank.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        public void Prints()
        {

        }

        public void Searchs()
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }
        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int item0 = 0; int i = 1;
            //    if (txtsearch.Text.Length > 0)
            //    {
            //        listView1.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {
            //            ListViewItem list = new ListViewItem();

            //            if (item.SubItems[2].ToString().Contains(txtsearch.Text.ToUpper()))
            //            {
            //                list.Text = i.ToString();
            //                list.SubItems.Add(item.SubItems[1].Text);
            //                list.SubItems.Add(item.SubItems[2].Text);
            //                list.SubItems.Add(item.SubItems[3].Text);                            
            //                list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

            //                listView1.Items.Add(list);
            //            }
            //            i++;
            //            item0++;
            //        }
            //    }
            //    else
            //    {
            //        try
            //        {
            //            listView1.Items.Clear(); item0 = 0;
            //            foreach (ListViewItem item in listfilter.Items)
            //            {

            //                ListViewItem list = new ListViewItem();
            //                list.Text = i.ToString();
            //                list.SubItems.Add(item.SubItems[1].Text);
            //                list.SubItems.Add(item.SubItems[2].Text);
            //                list.SubItems.Add(item.SubItems[3].Text);
            //                list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

            //                listView1.Items.Add(list);
            //                item0++;i++;
            //            }
            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    // MessageBox.Show("---" + ex.ToString());
            //}
        }

        private void uccListView1_Click(object sender, EventArgs e)
        {

        }

        private void txtbank_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == '(' || e.KeyChar == ')' || (char.IsWhiteSpace(e.KeyChar)) || e.KeyChar == (char)Keys.Back);

        }

        private void txtbank_TextChanged(object sender, EventArgs e)
        {
            if (txtbankid.Text == "" && txtbank.Text.Length >= 6)
            {
               
                string sel = "select asptblbanmasID    from  asptblbanmas    WHERE bankname='" + txtbank.Text.ToUpper().Trim() + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbanmas");
                DataTable dt = ds.Tables["asptblbanmas"];
                if (dt.Rows.Count > 0)
                {
                    txtbank.Text = "";
                    MessageBox.Show("Child Recrod Found");
                }
            }
        }
    }
}

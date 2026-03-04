using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql;

namespace Pinnacle.Master.CFM
{
    public partial class User : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        public User()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static User _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        public static User Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new User();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
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

        private void VarietyItemMaster_Load(object sender, EventArgs e)
        {
            try
            {

                string sel1 = "select a.userid,a.username from asptblusermas  a where NOT a.username='VAIRAM'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblusermas");
               DataTable dt = ds.Tables["asptblusermas"];
                if (dt.Rows.Count > 0)
                {
                    combouser.DisplayMember = "username";
                    combouser.ValueMember = "userid";
                    combouser.DataSource = dt;
                }              


             
                combouser.SelectedIndex = -1;

            }

            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
        }

        public void Saves()
        {
            try
            {
                Models.Validate va = new Models.Validate();
                if (combouser.Text == "VAIRAM") { combouser.Text = ""; }
                string username = combouser.Text.Trim();
                // ---------- Validation ----------
                if (!va.IsStringNumbericspacehypen(combouser.Text))
                {
                    MessageBox.Show(
                        "User Name is invalid or empty",
                        "Alert",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

                    combouser.Focus();
                    return;
                }
                if (!va.IsStringNumbericspacehypen(txtpassword.Text))
                {
                    MessageBox.Show(
                        "Password is invalid or empty",
                        "Alert",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

                    txtpassword.Focus();
                    return;
                }
               
                string password = sm.Encrypt(txtpassword.Text.Trim());
                string activeFlag = checkactive.Checked ? "T" : "F";
                int itemId = Convert.ToInt32("0" + txtuserid.Text);

                // ---------- Duplicate Check ----------

                string sel = "select userid    from  asptblusermas    WHERE  username='" + username + "' AND pasword='" + password + "' and active='T' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                bool recordExists = ds.Tables[0].Rows.Count > 0;
                if (recordExists)
                {
                    MessageBox.Show("Record already exists", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // ---------- INSERT ----------
                else if (itemId == 0)
                {
                   

                    //string ins = "insert into asptblusermas(username,pasword,active,ipaddres,createdon)  VALUES('" + username + "','" + password + "','" + activeFlag + "','" + Class.Users.IPADDRESS + "','" + Class.Users.CREATED + "')";
                    //Utility.ExecuteNonQuery(ins);
                    MessageBox.Show("Invalid Record","Information", MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
                // ---------- UPDATE ----------
                else
                {
                    string up = "update  asptblusermas  set   username='" + username + "' ,pasword='" + password + "', active='" + activeFlag + "' ,ipaddres='" + Class.Users.IPADDRESS + "' where userid='" + txtuserid.Text + "';";
                    Utility.ExecuteNonQuery(up);
                    MessageBox.Show("Record Updated Successfully",  "Update",  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"Error",  MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                GridLoad();
                empty();
            }

            
        }

        private void VarietyItemMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

    
        public void News()
        {

            empty(); GridLoad();
        }
        private void empty()
        {
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            Class.Users.UserTime = 0;
            txtuserid.Text = "";
            combouser.Text = "";txtpassword.Text = "";
            combouser.Select();
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
        }
        public void GridLoad()
        {
            try
            {
                listView1.BeginUpdate();
                listView1.Items.Clear();
                listfilter.Items.Clear();
                string query = @"SELECT  userid, username,pasword, active  FROM asptblusermas where active='T' AND NOT username='VAIRAM' ORDER BY userid DESC";
                DataSet ds = Utility.ExecuteSelectQuery(query, "asptblusermas");
                DataTable dt = ds.Tables["asptblusermas"];
                if (dt.Rows.Count == 0)
                    return;

                int rowIndex = 1;
                foreach (DataRow row in dt.Rows)
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems.Add(rowIndex.ToString());
                    item.SubItems.Add(row["userid"].ToString());
                    item.SubItems.Add(row["username"].ToString());
                    item.SubItems.Add(row["pasword"].ToString());
                    item.SubItems.Add(row["active"].ToString());

                    item.BackColor = rowIndex % 2 == 0
                        ? Class.Users.Color1
                        : Class.Users.Color2;

                    listView1.Items.Add(item);
                    listfilter.Items.Add((ListViewItem)item.Clone());

                    rowIndex++;
                }
                lbltotal.Text = $"Total Count: {listView1.Items.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                listView1.EndUpdate();
            }


            
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtuserid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.userid, a.username ,a.pasword, a.active    from  asptblusermas a    where a.userid=" + txtuserid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblusermas");
                    DataTable dt = ds.Tables["asptblusermas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtuserid.Text = Convert.ToString(dt.Rows[0]["userid"].ToString());
                        combouser.Text = Convert.ToString(dt.Rows[0]["username"].ToString()); combouser.Enabled = false;
                        txtpassword.Text = Convert.ToString(dt.Rows[0]["pasword"].ToString());
                        if (combouser.Text == "VAIRAM") { txtpassword.Visible = false; }
                        else
                        {
                            txtpassword.Text = sm.Decrypt(dt.Rows[0]["pasword"].ToString()); txtpassword.Visible = true;
                        }
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
                listView1.BeginUpdate();
                listView1.Items.Clear();

                string searchText = txtsearch.Text.Trim();
                int rowIndex = 0;

                if (searchText.Length > 1)
                {
                    foreach (ListViewItem src in listfilter.Items)
                    {
                        if (src.SubItems[3].Text.Contains(searchText))
                        {
                            ListViewItem item = (ListViewItem)src.Clone();

                            item.BackColor = rowIndex % 2 == 0
                                ? Class.Users.Color1
                                : Class.Users.Color2;

                            listView1.Items.Add(item);
                            rowIndex++;
                        }
                    }
                }
                else
                {
                    GridLoad();
                }

                lbltotal.Text = $"Total Count: {listView1.Items.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                listView1.EndUpdate();
            }


            

        }

        private void txtitem_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Enter)
            {
                Saves();
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        public void Deletes()
        {
            throw new NotImplementedException();
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

        private void combouser_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                txtpassword.Focus();
            }
        }

        private void txtpassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                
                Saves();
            }
        }
    }
}

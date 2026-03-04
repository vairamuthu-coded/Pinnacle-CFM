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
    public partial class WheatVarietyMaster : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        public WheatVarietyMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static WheatVarietyMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        public static WheatVarietyMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new WheatVarietyMaster();
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
            
        }

        public void Saves()
        {
            try
            {
                Models.Validate va = new Models.Validate();

                // ---------- Validation ----------
                if (!va.IsStringNumbericspacehypen(txtitem.Text))
                {
                    MessageBox.Show(
                        "Item Name is invalid or empty",
                        "Alert",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);

                    txtitem.Focus();
                    return;
                }

                string itemName = txtitem.Text.Trim().ToUpper();
                string activeFlag = checkactive.Checked ? "T" : "F";
                int itemId = Convert.ToInt32("0" + txtitemid.Text);

                // ---------- Duplicate Check ----------

                string sel = "select asptblitemmastid    from  asptblitemmast    WHERE itemname='" + txtitem.Text + "' and active='" + activeFlag + "' ";
               DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblitemmast");
                bool recordExists = ds.Tables[0].Rows.Count > 0;
                if (recordExists)
                {
                    MessageBox.Show("Record already exists", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // ---------- INSERT ----------
                else if (itemId == 0)
                {
                   

                    string ins = "insert into asptblitemmast(itemname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtitem.Text.ToUpper() + "','" + activeFlag + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                    Utility.ExecuteNonQuery(ins);
                    MessageBox.Show("Record Saved Successfully","Success", MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
                // ---------- UPDATE ----------
                else
                {
                    string up = "update  asptblitemmast  set   itemname='" + txtitem.Text.ToUpper() + "' , active='" + activeFlag + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblitemmastid='" + txtitemid.Text + "';";
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

            //try
            //{
            //    Models.Validate va = new Models.Validate();
            //    if (va.IsStringNumbericspacehypen(txtitem.Text) == false)
            //    {
            //        MessageBox.Show("Wheat Variety/ItemName is empty " + " Alert " + txtitem.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        return;
            //    }
            //    else
            //    {
            //        if (va.IsStringNumbericspacehypen(txtitem.Text) == true)
            //        {

            //            string chk = "";
            //            if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
            //            string sel = "select asptblitemmastid    from  asptblitemmast    WHERE itemname='" + txtitem.Text + "' and active='" + chk + "' ";
            //            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblitemmast");
            //            DataTable dt = ds.Tables["asptblitemmast"];
            //            if (dt.Rows.Count != 0)
            //            {
            //                MessageBox.Show("Child Record Found " + " Alert " + txtitem.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
            //            }
            //            else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtitemid.Text) == 0 || Convert.ToInt32("0" + txtitemid.Text) == 0)
            //            {
            //                string ins = "insert into asptblitemmast(itemname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtitem.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
            //                Utility.ExecuteNonQuery(ins);
            //                MessageBox.Show("Record Saved Successfully " + txtitem.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                GridLoad(); empty();
            //            }
            //            else
            //            {
            //                string up = "update  asptblitemmast  set   itemname='" + txtitem.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblitemmastid='" + txtitemid.Text + "';";
            //                Utility.ExecuteNonQuery(up);
            //                MessageBox.Show("Record Updated Successfully " + txtitem.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                GridLoad();
            //                empty();
            //            }

            //        }
            //        else
            //        {
            //            MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show("Saves_Click " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //}
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
            txtitemid.Text = "";
            txtitem.Text = "";
            txtitem.Select();
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
                string query = @"SELECT  asptblitemmastid, itemname, active  FROM asptblitemmast ORDER BY asptblitemmastid DESC";
                DataSet ds = Utility.ExecuteSelectQuery(query, "asptblitemmast");
                DataTable dt = ds.Tables["asptblitemmast"];
                if (dt.Rows.Count == 0)
                    return;

                int rowIndex = 1;
                foreach (DataRow row in dt.Rows)
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems.Add(rowIndex.ToString());
                    item.SubItems.Add(row["asptblitemmastid"].ToString());
                    item.SubItems.Add(row["itemname"].ToString());
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


            //try
            //{
            //    listView1.Items.Clear();listfilter.Items.Clear();
            //    string sel1 = " select a.asptblitemmastid, a.itemname , a.active  FROM  asptblitemmast a   order by  1;";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblitemmast");
            //    DataTable dt = ds.Tables["asptblitemmast"];
            //    if (dt.Rows.Count > 0)
            //    {
            //        int i = 1;
            //        foreach (DataRow myRow in dt.Rows)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            list.SubItems.Add(i.ToString());
            //            list.SubItems.Add(myRow["asptblitemmastid"].ToString());
            //            list.SubItems.Add(myRow["itemname"].ToString());
            //            list.SubItems.Add(myRow["active"].ToString());

            //            listfilter.Items.Add((ListViewItem)list.Clone());
            //            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
            //            listView1.Items.Add(list);
            //            i++;
            //        }
            //        lbltotal.Text = "Total Count: " + listView1.Items.Count;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtitemid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblitemmastid, a.itemname , a.active    from  asptblitemmast a    where a.asptblitemmastid=" + txtitemid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblitemmast");
                    DataTable dt = ds.Tables["asptblitemmast"];

                    if (dt.Rows.Count > 0)
                    {
                        txtitemid.Text = Convert.ToString(dt.Rows[0]["asptblitemmastid"].ToString());
                        txtitem.Text = Convert.ToString(dt.Rows[0]["itemname"].ToString());
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

                string searchText = txtsearch.Text.Trim().ToUpper();
                int rowIndex = 0;

                if (searchText.Length > 1)
                {
                    foreach (ListViewItem src in listfilter.Items)
                    {
                        if (src.SubItems[3].Text.ToUpper().Contains(searchText))
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


            //try
            //{
            //    int item0 = 0;
            //    if (txtsearch.Text.Length > 1)
            //    {
            //        listView1.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (item.SubItems[3].ToString().Contains(txtsearch.Text.ToUpper()))
            //            {
            //                list.Text = item.SubItems[0].Text;
            //                list.SubItems.Add(item.SubItems[1].Text);
            //                list.SubItems.Add(item.SubItems[2].Text);
            //                list.SubItems.Add(item.SubItems[3].Text);
            //                list.SubItems.Add(item.SubItems[4].Text);
            //                list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

            //                listView1.Items.Add(list);

            //                item0++;
            //            }
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
            //lbltotal.Text = "";
            //lbltotal.Text = "Total Count: " + listView1.Items.Count;

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
            if (txtitemid.Text != "")
            {
                string sel1 = "select a.asptblitemmastid from asptblitemmast a join asptblrawmaterial b on a.asptblitemmastid=b.itemnamevarity where a.asptblitemmastid='" + txtitemid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblitemmast");
                DataTable dt = ds.Tables["asptblitemmast"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtitem.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblitemmast where asptblitemmastid='" + Convert.ToInt64("0" + txtitemid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtitem.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

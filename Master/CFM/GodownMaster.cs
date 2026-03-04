using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master.CFM
{
    public partial class GodownMaster : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        public GodownMaster()
        {
            InitializeComponent();
         
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static GodownMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        public static GodownMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new GodownMaster();
               return _instance;
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


        private void GodownMaster_Load(object sender, EventArgs e)
        {
           
        }
      
        public void Saves()
        {
            try
            {
                if (txtgodown.Text == "")
                {
                    MessageBox.Show("Godown Name is empty " + " Alert " + txtgodown.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtgodown.Text != "")
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsStringNumbericspacehypen(txtgodown.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select asptblgodwonmasid    from  asptblgodwonmas    WHERE godownname='" + txtgodown.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblgodwonmas");
                        DataTable dt = ds.Tables["asptblgodwonmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtgodown.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtgodownid.Text) == 0 || Convert.ToInt32("0" + txtgodownid.Text) == 0)
                        {
                            string ins = "insert into asptblgodwonmas(godownname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtgodown.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtgodown.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblgodwonmas  set   godownname='" + txtgodown.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblgodwonmasid='" + txtgodownid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtgodown.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("godownname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void GodownMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            txtgodownid.Text = "";
            txtgodown.Text = ""; txtgodown.Select();
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
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "   SELECT A.asptblgodwonmasid, A.godownname , a.active  FROM  asptblgodwonmas A   order by A.asptblgodwonmasid DESC";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgodwonmas");
                DataTable dt = ds.Tables["asptblgodwonmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["asptblgodwonmasid"].ToString());
                        list.SubItems.Add(myRow["godownname"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listView1.Items.Add(list);
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
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

                    txtgodownid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.asptblgodwonmasid, a.godownname , a.active    from  asptblgodwonmas a    where a.asptblgodwonmasid=" + txtgodownid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgodwonmas");
                    DataTable dt = ds.Tables["asptblgodwonmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtgodownid.Text = Convert.ToString(dt.Rows[0]["asptblgodwonmasid"].ToString());
                        txtgodown.Text = Convert.ToString(dt.Rows[0]["godownname"].ToString());
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
                if (txtsearch.Text.Length > 1)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text.ToUpper()))
                        {
                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);

                            item0++;
                        }
                    }
                }
                else
                {
                    listView1.Items.Clear();
                    GridLoad();
                }
                lbltotal.Refresh();
                lbltotal.Text = "Total " + listView1.Items.Count.ToString();
                //if (txtsearch.Text.ToUpper() != "")
                //{
                //    listView1.Items.Clear(); int iGLCount = 1;
                //    string sel1 = "  SELECT  a.asptblgodwonmasid,a.godownname,a.active from asptblgodwonmas a  where a.godownname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
                //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgodwonmas");
                //    DataTable dt = ds.Tables["asptblgodwonmas"];
                //    if (dt.Rows.Count > 0)
                //    {

                //        foreach (DataRow myRow in dt.Rows)
                //        {
                //            ListViewItem list = new ListViewItem();
                //            list.Text = iGLCount.ToString();
                //            list.SubItems.Add(myRow["asptblgodwonmasid"].ToString());
                //            list.SubItems.Add(myRow["godownname"].ToString());
                //            list.SubItems.Add(myRow["active"].ToString());
                //            listView1.Items.Add(list);
                //            iGLCount++;
                //        }
                //        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                //    }
                //    else
                //    {
                //        listView1.Items.Clear();
                //    }
                //}
                //else
                //{

                //    listView1.Items.Clear();
                //    GridLoad();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Deletes()
        {
            if (txtgodownid.Text != "")
            {
                string sel1 = "select a.asptblgodwonmasid from asptblgodwonmas a join asptblrawmaterial b on a.asptblgodwonmasid=b.godownname where a.asptblgodwonmasid='" + txtgodownid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblgodwonmas");
                DataTable dt = ds.Tables["asptblgodwonmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtgodown.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblgodwonmas where asptblgodwonmasid='" + Convert.ToInt64("0" + txtgodownid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtgodown.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        public void Searchs(int EditID)
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


    }
}

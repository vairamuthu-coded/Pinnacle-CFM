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
    public partial class HSNMaster : Form,ToolStripAccess
    {
        public HSNMaster()
        {
            InitializeComponent();
          
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static HSNMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights(); ListView listfilter = new ListView();
        public static HSNMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new HSNMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


   
        private void HSNMaster_Load(object sender, EventArgs e)
        {
            News();GridLoad();
        }

        public void Saves()
        {
            try
            {
                if (combohsncode.Text == "")
                {
                    MessageBox.Show("Code Name is empty " + " Alert " + txthsn.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txthsn.Text == "")
                {
                    MessageBox.Show("HSN  is empty " + " Alert " + txthsn.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txthsn.Text != "" && combohsncode.Text != "")
                {


                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblhsnmasid    from  asptblhsnmas    WHERE hsncode='" + combohsncode.Text + "' and hsn='" + txthsn.Text + "'and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblhsnmas");
                    DataTable dt = ds.Tables["asptblhsnmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txthsn.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txthsnid.Text) == 0 || Convert.ToInt32("0" + txthsnid.Text) == 0)
                    {
                        string ins = "insert into asptblhsnmas(hsncode,hsn,active,createdby,modifiedby,ipaddress)  VALUES('" + combohsncode.Text + "','" + txthsn.Text + "','" + chk + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txthsn.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblhsnmas  set   hsncode='" + combohsncode.Text + "' , hsn='" + txthsn.Text + "' ,active='" + chk + "' , modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblhsnmasid='" + txthsnid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txthsn.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("hsncode " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void HSNMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void News()
        {

            empty(); GridLoad();
        }
        private void empty()
        {
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            this.Font = Class.Users.FontName;
            txthsnid.Text = "";
            txthsn.Text = "";
            combohsncode.Text = "";
            checkactive.Checked = false;
            txthsn.Select();
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "   SELECT a.asptblhsnmasid, a.hsncode ,a.hsn, a.active  FROM  asptblhsnmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhsnmas");
                DataTable dt = ds.Tables["asptblhsnmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblhsnmasid"].ToString());
                        list.SubItems.Add(myRow["hsncode"].ToString());
                        list.SubItems.Add(myRow["hsn"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
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

                    txthsnid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblhsnmasid, a.hsncode ,a.hsn, a.active    from  asptblhsnmas a    where a.asptblhsnmasid=" + txthsnid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhsnmas");
                    DataTable dt = ds.Tables["asptblhsnmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txthsnid.Text = Convert.ToString(dt.Rows[0]["asptblhsnmasid"].ToString());
                        combohsncode.Text = Convert.ToString(dt.Rows[0]["hsncode"].ToString());
                        txthsn.Text = Convert.ToString(dt.Rows[0]["hsn"].ToString());
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
                int item0 = 0; listView1.Items.Clear();
                    if (txtsearch.Text.Length > 1)
                    {
                   
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            if (item.SubItems[3].ToString().Contains(txtsearch.Text) || item.SubItems[4].ToString().Contains(txtsearch.Text))
                            {


                                list.Text = item.SubItems[0].Text;
                                list.SubItems.Add(item.SubItems[1].Text);
                                list.SubItems.Add(item.SubItems[2].Text);
                                list.SubItems.Add(item.SubItems[3].Text);
                                list.SubItems.Add(item.SubItems[4].Text);
                                list.SubItems.Add(item.SubItems[5].Text);

                                listView1.Items.Add(list);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        }
                            item0++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }
                    else
                    {

                        ListView ll = new ListView();

                        listView1.Items.Clear(); 
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                       
                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);

                            listView1.Items.Add(list);

                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }


               
            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
        }

        public void Deletes()
        {
            if (txthsnid.Text != "")
            {
                string sel1 = "select a.asptblhsnmasid from asptblhsnmas a join gtstatemast b on a.asptblhsnmasid=b.country where a.asptblhsnmasid='" + txthsnid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhsnmas");
                DataTable dt = ds.Tables["asptblhsnmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txthsn.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblhsnmas where asptblhsnmasid='" + Convert.ToInt64("0" + txthsnid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txthsn.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

       

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

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
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }

      
    }
}

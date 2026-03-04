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

namespace Pinnacle.Master.Hospital
{
    public partial class LabTestItemMaster : Form,ToolStripAccess
    {
        private static LabTestItemMaster _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
        public static LabTestItemMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new LabTestItemMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public LabTestItemMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;

           
        }
        private void LabTestItemMaster_Load(object sender, EventArgs e)
        {
            News();
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
                if (txtlabtestitem.Text != "")
                {



                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "SELECT a.asptbllabtestitemmasid FROM asptbllabtestitemmas a    WHERE labtest='" + combolabtest.SelectedValue + "' and labtestitem='" + txtlabtestitem.Text + "' and active='" + chk + "' and rate='" + txtrate.Text + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllabtestitemmas");
                    DataTable dt = ds.Tables["asptbllabtestitemmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtlabtestitem.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtlabtestid.Text) == 0 || Convert.ToInt32("0" + txtlabtestid.Text) == 0)
                    {
                        string ins = "insert into asptbllabtestitemmas(labtest,labtestitem,active,rate,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + combolabtest.SelectedValue + "','" + txtlabtestitem.Text.ToUpper() + "','" + chk + "','" + txtrate.Text + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtlabtestitem.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptbllabtestitemmas  set labtest='" + combolabtest.SelectedValue + "' ,  labtestitem='" + txtlabtestitem.Text.ToUpper() + "' , active='" + chk + "' ,  rate='" + txtrate.Text + "' ,modifiedby='" + Class.Users.USERID + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbllabtestitemmasID='" + txtlabtestid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtlabtestitem.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty();
                    }


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void LabTestItemMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        public void News()
        {
            GridLoad(); labtestload();
            empty();
        }
        private void empty()
        {
            txtlabtestid.Text = ""; txtrate.Text = "";
            txtlabtestitem.Text = ""; combolabtest.Text = "";
            this.BackColor = Class.Users.BackColors;

            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            txtlabtestitem.Select();
        }
        void labtestload()
        {
            string sel1 = "select a.asptbllabtestmasid,a.labtest from asptbllabtestmas a where a.active='T' order by 1;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestmas");
            DataTable dt = ds.Tables["asptbllabtestmas"];
            combolabtest.DisplayMember = "labtest";
            combolabtest.ValueMember = "asptbllabtestmasid"; combolabtest.DataSource = dt;
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.asptbllabtestitemmasID, b.labtest,a.labtestitem,a.rate, a.active  from  asptbllabtestitemmas a  join asptbllabtestmas b on a.labtest=b.asptbllabtestmasid   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
                DataTable dt = ds.Tables["asptbllabtestitemmas"];
                if (dt.Rows.Count > 0 || dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbllabtestitemmasID"].ToString());
                        list.SubItems.Add(myRow["labtest"].ToString());
                        list.SubItems.Add(myRow["labtestitem"].ToString());
                        list.SubItems.Add(myRow["rate"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                        listView1.Items.Add(list);
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

                    txtlabtestid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptbllabtestitemmasID, b.labtest ,a.labtestitem,a.rate, a.active    from  asptbllabtestitemmas a join asptbllabtestmas b on a.labtest=b.asptbllabtestmasid   where a.asptbllabtestitemmasID=" + txtlabtestid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
                    DataTable dt = ds.Tables["asptbllabtestitemmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtlabtestid.Text = Convert.ToString(dt.Rows[0]["asptbllabtestitemmasID"].ToString());
                        combolabtest.Text = Convert.ToString(dt.Rows[0]["labtest"].ToString());
                        txtlabtestitem.Text = Convert.ToString(dt.Rows[0]["labtestitem"].ToString());
                        txtrate.Text = Convert.ToString(dt.Rows[0]["rate"].ToString());
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
                    string sel1 = "  SELECT  a.asptbllabtestitemmasID,a.labtest,a.rate,a.active from asptbllabtestitemmas a  where a.labtest LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
                    DataTable dt = ds.Tables["asptbllabtestitemmas"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["asptbllabtestitemmasID"].ToString());
                            list.SubItems.Add(myRow["labtest"].ToString());
                            list.SubItems.Add(myRow["rate"].ToString());
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
            if (txtlabtestid.Text != "")
            {
                string sel1 = "select a.asptbllabtestitemmasID from asptbllabtestitemmas a join asptbldiagnosismas b on a.asptbllabtestitemmasID=b.labtest where a.asptbllabtestitemmasID='" + txtlabtestid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
                DataTable dt = ds.Tables["asptbllabtestitemmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtlabtestitem.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptbllabtestitemmas where asptbllabtestitemmasID='" + Convert.ToInt64("0" + txtlabtestid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtlabtestitem.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void txtlabtest_TextChanged(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            labtestload();GridLoad();empty();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

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
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
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
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());

                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            if (i % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            i++;
                            listView1.Items.Add(list);
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

        private void txtrate_KeyDown(object sender, KeyEventArgs e)
        {
           // e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        }

        private void txtrate_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        }

        private void combolabtest_Leave(object sender, EventArgs e)
        {
            combolabtest.BackColor = Class.Users.Color1;
        }

        private void combolabtest_Enter(object sender, EventArgs e)
        {
            combolabtest.BackColor = Class.Users.Color2;
        }

        private void txtlabtestitem_Enter(object sender, EventArgs e)
        {
            txtlabtestitem.BackColor = Class.Users.Color2;
        }

        private void txtlabtestitem_Leave(object sender, EventArgs e)
        {
            txtlabtestitem.BackColor = Class.Users.Color1;
        }

        private void txtrate_Leave(object sender, EventArgs e)
        {
            txtrate.BackColor = Class.Users.Color1;
        }

        private void txtrate_Enter(object sender, EventArgs e)
        {
            txtrate.BackColor = Class.Users.Color2;
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}


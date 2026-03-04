using Pinnacle.UserControls;
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

namespace Pinnacle.Master.Lyla
{
    public partial class RemarksMaster : Form,ToolStripAccess
    {
        private static RemarksMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        Pinnacle.UserControls.UCCListView ucclist = new UCCListView();
        public static RemarksMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new RemarksMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public RemarksMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());

        }
 
        private void RemarksMaster_Load(object sender, EventArgs e)
        {
            txtremarks.Select();
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
                if (txtremarks.Text != "")
                {
                    //Models.Validate va = new Models.Validate();
                    //if (va.IsString(txtremarks.Text) == true)
                    //{

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select asptblremmasID    from  asptblremmas    WHERE Remarks='" + txtremarks.Text + "' and active='" + chk + "' and compcode='" + Class.Users.COMPCODE + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblremmas");
                        DataTable dt = ds.Tables["asptblremmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtremarks.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtremarksid.Text) == 0 || Convert.ToInt32("0" + txtremarksid.Text) == 0)
                        {
                            string ins = "insert into asptblremmas(Remarks,active,compcode,createdby,modifiedby,ipaddress)  VALUES('" + txtremarks.Text.ToUpper() + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtremarks.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblremmas  set   Remarks='" + txtremarks.Text.ToUpper() + "' , active='" + chk + "' ,compcode='" + Class.Users.COMPCODE + "',modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblremmasID='" + txtremarksid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtremarks.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    //}
                    //else
                    //{
                    //    MessageBox.Show("'Remarks  is Wrong'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                    //}
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Remarks " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void RemarksMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            txtremarksid.Text = "";
            txtremarks.Text = "";
            txtremarks.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptblremmasID, a.Remarks as  Remarks , a.active  from  asptblremmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblremmas");
                DataTable dt = ds.Tables["asptblremmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["asptblremmasID"].ToString());
                        list.SubItems.Add(myRow["Remarks"].ToString());
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

                    txtremarksid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.asptblremmasID, a.Remarks as Remarks , a.active    from  asptblremmas a    where a.asptblremmasID=" + txtremarksid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblremmas");
                    DataTable dt = ds.Tables["asptblremmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtremarksid.Text = Convert.ToString(dt.Rows[0]["asptblremmasID"].ToString());
                        txtremarks.Text = Convert.ToString(dt.Rows[0]["Remarks"].ToString());
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
                    string sel1 = "  SELECT  a.asptblremmasID,a.Remarks as Remarks,a.active from asptblremmas a  where a.Remarks LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblremmas");
                    DataTable dt = ds.Tables["asptblremmas"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["asptblremmasID"].ToString());
                            list.SubItems.Add(myRow["Remarks"].ToString());
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
            if (txtremarksid.Text != "")
            {
                string sel1 = "select a.asptblremmasID from asptblremmas a join gtstatemast b on a.asptblremmasID=b.Remarks where a.asptblremmasID='" + txtremarksid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblremmas");
                DataTable dt = ds.Tables["asptblremmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtremarks.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblremmas where asptblremmasID='" + Convert.ToInt64("0" + txtremarksid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtremarks.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                            list.Text = item.SubItems[0].Text;
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

                            list.Text = item.SubItems[0].Text;
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
    }
}

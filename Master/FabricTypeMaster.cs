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
    public partial class FabricTypeMaster : Form,ToolStripAccess
    {
        public FabricTypeMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
         
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static FabricTypeMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        public static FabricTypeMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new FabricTypeMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void usercheck(string s, string ss, string sss)
        {

          

        }
        private void FabricTypeMaster_Load(object sender, EventArgs e)
        {
            News();GridLoad();
        }

        public void Saves()
        {
            try
            {
                if (txtfabrictype.Text == "")
                {
                    MessageBox.Show("fabrictype Name is empty " + " Alert " + txtfabrictype.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtfabrictype.Text != "")
                {


                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblfabrictypemasid    from  asptblfabrictypemas    WHERE fabrictype='" + txtfabrictype.Text + "' and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblfabrictypemas");
                    DataTable dt = ds.Tables["asptblfabrictypemas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtfabrictype.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtfabrictypeid.Text) == 0 || Convert.ToInt32("0" + txtfabrictypeid.Text) == 0)
                    {
                        string ins = "insert into asptblfabrictypemas(fabrictype,active,createdby,modifiedby,ipaddress)  VALUES('" + txtfabrictype.Text.ToUpper() + "','" + chk + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtfabrictype.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblfabrictypemas  set   fabrictype='" + txtfabrictype.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblfabrictypemasid='" + txtfabrictypeid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtfabrictype.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("fabrictype " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void FabricTypeMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            txtfabrictypeid.Text = "";
            txtfabrictype.Text = "";
            checkactive.Checked = false;
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "   SELECT A.asptblfabrictypemasid, A.fabrictype , a.active  FROM  asptblfabrictypemas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblfabrictypemas");
                DataTable dt = ds.Tables["asptblfabrictypemas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblfabrictypemasid"].ToString());
                        list.SubItems.Add(myRow["fabrictype"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = System.Drawing.Color.White;

                        }
                        else
                        {
                            list.BackColor = System.Drawing.Color.WhiteSmoke;
                        }
                        listView1.Items.Add(list);
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

                    txtfabrictypeid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblfabrictypemasid, a.fabrictype , a.active    from  asptblfabrictypemas a    where a.asptblfabrictypemasid=" + txtfabrictypeid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblfabrictypemas");
                    DataTable dt = ds.Tables["asptblfabrictypemas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtfabrictypeid.Text = Convert.ToString(dt.Rows[0]["asptblfabrictypemasid"].ToString());
                        txtfabrictype.Text = Convert.ToString(dt.Rows[0]["fabrictype"].ToString());
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


                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length > 1)
                {

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
                                list.BackColor = System.Drawing.Color.White;

                            }
                            else
                            {
                                list.BackColor = System.Drawing.Color.WhiteSmoke;
                            }
                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {

                    ListView ll = new ListView();

                    listView1.Items.Clear(); item0 = listfilter.Items.Count;
                    foreach (ListViewItem item in listfilter.Items)
                    {

                        this.listView1.Items.Add((ListViewItem)item.Clone());

                        if (item0 % 2 == 0)
                        {
                            item.BackColor = System.Drawing.Color.White;

                        }
                        else
                        {
                            item.BackColor = System.Drawing.Color.WhiteSmoke;
                        }
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
            if (txtfabrictypeid.Text != "")
            {
                string sel1 = "select a.asptblfabrictypemasid from asptblfabrictypemas a join gtstatemast b on a.asptblfabrictypemasid=b.country where a.asptblfabrictypemasid='" + txtfabrictypeid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblfabrictypemas");
                DataTable dt = ds.Tables["asptblfabrictypemas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtfabrictype.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblfabrictypemas where asptblfabrictypemasid='" + Convert.ToInt64("0" + txtfabrictypeid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtfabrictype.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

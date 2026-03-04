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
    public partial class StateMaster : Form,ToolStripAccess
    {
       
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        private static StateMaster _instance;
        ListView listfilter = new ListView();
        public static StateMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new StateMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public StateMaster()
        {
            InitializeComponent();
         
 
        }
      

        private void StateMaster_Load(object sender, EventArgs e)
        {
           
        }

        public DataTable country()
        {
            string sel = " select a.GTCOUNTRYMASTID, a.countryname as country , a.active from  GTCOUNTRYMAST a    where a.active='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCOUNTRYMAST");
            DataTable dt = ds.Tables["GTCOUNTRYMAST"];
            combocountry.DataSource = dt;
            combocountry.DisplayMember = "country";
            combocountry.ValueMember = "GTCOUNTRYMASTID";
            return dt;
        }
        public void Saves()
        {
            try
            {
                if (txtstate.Text == "")
                {
                    MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtstate.Focus();
                    return;
                }
                if (combocountry.Text == "")
                {
                    MessageBox.Show("'Contry Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);this.combocountry.Focus(); return;
                }
                else
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsString(txtstate.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.GTSTATEMASTID    from  GTSTATEMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID   WHERE a.statename='" + txtstate.Text + "' and b.countryname='" + combocountry.Text + "' and a.active='" + chk + "' and a.GTSTATEMASTID='" + txtstateid.Text + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "GTSTATEMAST");
                        DataTable dt = ds.Tables["GTSTATEMAST"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtstate.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtstateid.Text) == 0 || Convert.ToInt32("0" + txtstateid.Text) == 0)
                        {
                            string ins = "insert into GTSTATEMAST(statename,country,active,createdby,modifiedby,ipaddress)  VALUES('" + txtstate.Text.ToUpper() + "','" + combocountry.SelectedValue + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtstate.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  GTSTATEMAST  set statename='" + txtstate.Text.ToUpper() + "', country='" + combocountry.SelectedValue + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where GTSTATEMASTID='" + txtstateid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtstate.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("country " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void StateMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            empty();
            country(); GridLoad();
        }
        private void empty()
        {
            txtstateid.Text = "";
            txtstate.Text = "";  combocountry.Text = "";
            combocountry.SelectedIndex = -1; listfilter.Items.Clear();
            //txtsearch.Text = ""; Class.Users.UserTime = 0;
            butheader.Text = Class.Users.ScreenName;
           // panel2.BackColor = Class.Users.BackColors;
            //panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            txtstate.Select();
        }
        public void GridLoad()
        {
            try
            {
                Class.Users.SearchQuery = "select a.GTSTATEMASTID AS ID,a.statename as STATE, b.countryname as COUNTRY , a.ACTIVE    from  GTSTATEMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID    order by 1";
                Class.Users.HideCols = new string[] { "ID" };
                listView1.Load_Details();

            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //private void ListView1_ItemActivate(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        if (listView1.Items.Count > 0)
        //        {

        //            txtstateid.Text = listView1.SelectedItems[0].SubItems[2].Text;
        //            string sel1 = " select a.GTSTATEMASTID,a.statename as state, b.countryname as country , a.active    from  GTSTATEMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID    where a.GTSTATEMASTID=" + txtstateid.Text;
        //            DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTSTATEMAST");
        //            DataTable dt = ds.Tables["GTSTATEMAST"];

        //            if (dt.Rows.Count > 0)
        //            {
        //                txtstateid.Text = Convert.ToString(dt.Rows[0]["GTSTATEMASTID"].ToString());
        //                txtstate.Text = Convert.ToString(dt.Rows[0]["state"].ToString());
        //                combocountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
        //                if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }


        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        //private void Txtsearch_TextChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        int item0 = 0; int i = 1;
        //        if (txtsearch.Text.Length > 0)
        //        {
        //            listView1.Items.Clear();
        //            foreach (ListViewItem item in listfilter.Items)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                if (item.SubItems[3].ToString().Contains(txtsearch.Text))
        //                {
        //                    list.Text = i.ToString();
        //                    list.SubItems.Add(item.SubItems[1].Text);
        //                    list.SubItems.Add(item.SubItems[2].Text);
        //                    list.SubItems.Add(item.SubItems[3].Text);
        //                    list.SubItems.Add(item.SubItems[4].Text);
        //                    list.SubItems.Add(item.SubItems[5].Text);

        //                    list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

        //                    listView1.Items.Add(list);
        //                }
        //                item0++;
        //            }
        //        }
        //        else
        //        {

        //            listView1.Items.Clear(); item0 = 0;
        //            foreach (ListViewItem item in listfilter.Items)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                list.Text = i.ToString();
        //                list.SubItems.Add(item.SubItems[1].Text);
        //                list.SubItems.Add(item.SubItems[2].Text);
        //                list.SubItems.Add(item.SubItems[3].Text);
        //                list.SubItems.Add(item.SubItems[4].Text);
        //                list.SubItems.Add(item.SubItems[5].Text);

        //                if (item0 % 2 == 0)
        //                {
        //                    list.BackColor = Color.White;
        //                }
        //                else
        //                {
        //                    list.BackColor = Color.WhiteSmoke;
        //                }
        //                listView1.Items.Add(list);

        //            }

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // MessageBox.Show("---" + ex.ToString());
        //    }
        //}

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            country();
        }

        public void Deletes()
        {
            if (txtstateid.Text != "")
            {
                string sel1 = "select a.GTSTATEMASTID from GTSTATEMAST a join GTCITYMAST b on a.GTSTATEMASTID=b.state where a.GTSTATEMASTID='" + txtstateid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTSTATEMAST");
                DataTable dt = ds.Tables["GTSTATEMAST"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtstate.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from GTSTATEMAST where GTSTATEMASTID='" + Convert.ToInt64("0" + txtstateid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtstate.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            try
            {
                if (EditID > 0)
                {

                    txtstateid.Text = EditID.ToString();
                    DataTable dt = Utility.SQLQuery("select a.GTSTATEMASTID,a.statename as state, b.countryname as country , a.active    from  GTSTATEMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID    where a.GTSTATEMASTID=" + txtstateid.Text + "");
                    if (dt.Rows.Count > 0)
                    {
                        txtstateid.Text = Convert.ToString(dt.Rows[0]["GTSTATEMASTID"].ToString());
                        txtstate.Text = Convert.ToString(dt.Rows[0]["state"].ToString());
                        combocountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void txtstate_TextChanged(object sender, EventArgs e)
        {
            if (txtstateid.Text == "" && txtstate.Text.Length >= 5)
            {
                string sel = "select gtstatemastid    from  gtstatemast    WHERE statename='" + txtstate.Text.ToUpper().Trim() + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtstatemast");
                DataTable dt = ds.Tables["gtstatemast"];
                if (dt.Rows.Count > 0)
                {
                    txtstate.Text = "";
                    MessageBox.Show("Child Recrod Found");
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class CountryMaster : Form,ToolStripAccess
    {
        private static CountryMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights(); ListView listfilter = new ListView();
        public static CountryMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new CountryMaster();             
                GlobalVariables.CurrentForm = _instance;               
                return _instance;
            }
            
        }
        public CountryMaster()
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
        private void CountryMaster_Load(object sender, EventArgs e)
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
                if (txtcountry.Text != "")
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsString(txtcountry.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select GTCOUNTRYMASTID    from  GTCOUNTRYMAST    WHERE countryname='" + txtcountry.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCOUNTRYMAST");
                        DataTable dt = ds.Tables["GTCOUNTRYMAST"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtcountry.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtcountryid.Text) == 0 || Convert.ToInt32("0" + txtcountryid.Text) == 0)
                        {
                            string ins = "insert into GTCOUNTRYMAST(countryname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtcountry.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtcountry.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  GTCOUNTRYMAST  set   countryname='" + txtcountry.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where GTCOUNTRYMASTID='" + txtcountryid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtcountry.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("'country  is Wrong'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("country " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void CountryMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            txtcountryid.Text = "";
            txtcountry.Text = "";           
            butheader.BackColor = Class.Users.BackColors;          
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            txtcountry.Select();
        }
        public void Searchs(int EditID)
        {
            try
            {
                if (EditID > 0)
                {

                    txtcountryid.Text = EditID.ToString();
                    DataTable dt = Utility.SQLQuery("select a.GTCOUNTRYMASTID, a.countryname as country , a.active    from  GTCOUNTRYMAST a    where a.GTCOUNTRYMASTID=" + txtcountryid.Text+"");
                    if (dt.Rows.Count > 0)
                    {
                        txtcountryid.Text = Convert.ToString(dt.Rows[0]["GTCITYMASTID"].ToString());
                        txtcountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
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
            Class.Users.SearchQuery = "select a.GTCOUNTRYMASTID AS ID, a.countryname as  COUNTRY , a.ACTIVE  from  GTCOUNTRYMAST a   order by 1";
            Class.Users.HideCols = new string[] { "ID" };
            listView1.Load_Details();
          
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            //try
            //{
            //    if (listView1.Items.Count > 0)
            //    {

            //        txtcountryid.Text = listView1.SelectedItems[0].SubItems[1].Text;
            //        string sel1 = " select a.GTCOUNTRYMASTID, a.countryname as country , a.active    from  GTCOUNTRYMAST a    where a.GTCOUNTRYMASTID=" + txtcountryid.Text;
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTCOUNTRYMAST");
            //        DataTable dt = ds.Tables["GTCOUNTRYMAST"];

            //        if (dt.Rows.Count > 0)
            //        {
            //            txtcountryid.Text = Convert.ToString(dt.Rows[0]["GTCOUNTRYMASTID"].ToString());
            //            txtcountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
            //            if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }
      
      
        public void Deletes()
        {
            if (txtcountryid.Text != "")
            {
                string sel1 = "select a.GTCOUNTRYMASTID from GTCOUNTRYMAST a join gtstatemast b on a.GTCOUNTRYMASTID=b.country where a.GTCOUNTRYMASTID='" + txtcountryid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTCOUNTRYMAST");
                DataTable dt = ds.Tables["GTCOUNTRYMAST"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcountry.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from GTCOUNTRYMAST where GTCOUNTRYMASTID='" + Convert.ToInt64("0" + txtcountryid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtcountry.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void txtcountry_TextChanged(object sender, EventArgs e)
        {
            if (txtcountryid.Text == "" && txtcountry.Text.Length >= 5)
            {
                string sel = "select gtcountrymastid    from  gtcountrymast    WHERE countryname='" + txtcountry.Text.ToUpper().Trim() + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcountrymast");
                DataTable dt = ds.Tables["gtcountrymast"];
                if (dt.Rows.Count > 0)
                {
                    txtcountry.Text = "";
                    MessageBox.Show("Child Recrod Found");
                }
            }
        }
    }
}

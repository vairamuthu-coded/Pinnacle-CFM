using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class CityMaster : Form,ToolStripAccess
    {
        private static CityMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
       
        public static CityMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new CityMaster();
                GlobalVariables.CurrentForm = _instance;
                _instance.Font = Class.Users.FontName;
                return _instance;
            }
        }
        public CityMaster()
        {
            InitializeComponent();
           


        }
        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    //for (int r = 0; r < dt1.Rows.Count; r++)
                    //{

                    //    if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                    //    if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                    //    if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                    //    if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                    //    if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                    //    if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                    //    if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                    //    if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                    //    if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                    //    if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                    //    if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                    //    if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }
                    //}
                }


            }
            else
            {

            }

        }

        private void CityMaster_Load(object sender, EventArgs e)
        {
          
            
        }
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        public DataTable state()
        {
            sb.Clear();
            sb.Append("SELECT  a.GTSTATEMASTID,a.statename from GTSTATEMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID  where a.active='T'");
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "GTSTATEMAST");
            DataTable dt = ds.Tables["GTSTATEMAST"];
            combostate.DataSource = dt;
            combostate.DisplayMember = "statename";
            combostate.ValueMember = "GTSTATEMASTID";
            return dt;
        }
        public DataTable country(Int64 s)
        {
            sb.Clear();
            sb.Append("SELECT  b.GTCOUNTRYMASTID,b.countryname as country from GTSTATEMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID where a.GTSTATEMASTID='" + s + "' ");
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "GTSTATEMAST");
            DataTable dt = ds.Tables["GTSTATEMAST"];
           
            return dt;
        }
        public void Saves()
        {
            try
            {
                sb.Clear();
                if (txtcity.Text == "")
                {
                    MessageBox.Show("'City Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtcity.Focus();
                    return;
                }
                if (combostate.SelectedValue == null)
                {
                    MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combostate.Focus();
                    return;
                }
                if (combocountry.SelectedValue == null)
                {
                    MessageBox.Show("'Contry Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combocountry.Focus(); return;
                }
                else
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsString(txtcity.Text) == true && va.IsString(combostate.Text) == true && va.IsString(combocountry.Text) == true)
                    {

                        string chk = ""; sb.Clear();
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        sb.Append("select a.GTCITYMASTID    from  GTCITYMAST a join GTCOUNTRYMAST b on a.country=b.GTCOUNTRYMASTID   WHERE a.cityname='" + txtcity.Text + "' and a.state='" + combostate.SelectedValue + "' and a.country='" + combocountry.SelectedValue + "' and a.active='" + chk + "' and a.GTCITYMASTID='" + txtcityid.Text + "'");
                        DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "GTCITYMAST");
                        DataTable dt = ds.Tables["GTCITYMAST"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtcity.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtcityid.Text) == 0 || Convert.ToInt32("0" + txtcityid.Text) == 0)
                        {
                            sb.Clear();
                            sb.Append("insert into GTCITYMAST(cityname, state,country,active,createdby,modifiedby,ipaddress)  VALUES('" + txtcity.Text.ToUpper() + "','" + combostate.SelectedValue + "','" + combocountry.SelectedValue + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )");
                            Utility.ExecuteNonQuery(sb.ToString());
                            MessageBox.Show("Record Saved Successfully " + txtcity.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            sb.Clear();
                            sb.Append("update  GTCITYMAST  set cityname='" + txtcity.Text.ToUpper() + "', state='" + combostate.SelectedValue + "',country='" + combocountry.SelectedValue + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where GTCITYMASTID='" + txtcityid.Text + "'");
                            Utility.ExecuteNonQuery(sb.ToString());
                            MessageBox.Show("Record Updated Successfully " + txtcity.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void CityMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        public void News()
        {
            GridLoad(); state(); empty();

        }
        private void empty()
        {
            txtcityid.Text = "";
            txtcity.Text = "";
            combostate.Text = ""; combostate.SelectedIndex = -1;  
            combocountry.Text = ""; combocountry.SelectedIndex = -1;            
          
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            txtcity.Select();
        }
        public void GridLoad()
        {
            try
            {
               

                Class.Users.SearchQuery = "select a.GTCITYMASTID  AS ID,a.CITYNAME, b.STATENAME, c.COUNTRYNAME , a.ACTIVE    from  GTCITYMAST a  join GTSTATEMAST b on a.state=b.GTSTATEMASTID   join GTCOUNTRYMAST c on b.country=c.GTCOUNTRYMASTID    order by 1";
                Class.Users.HideCols = new string[] { "ID" };
                listView1.Load_Details();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Searchs(int EditID)
        {
            try
            {
                if (EditID > 0)
                {

                    txtcityid.Text = EditID.ToString();
                    DataTable dt = Utility.SQLQuery("select a.GTCITYMASTID,a.cityname, b.statename , c.countryname  , a.active    from  GTCITYMAST a  join GTSTATEMAST b on a.state=b.GTSTATEMASTID   join GTCOUNTRYMAST c on b.country=c.GTCOUNTRYMASTID  where a.GTCITYMASTID=" + txtcityid.Text + "");
                    if (dt.Rows.Count > 0)
                    {
                        txtcityid.Text = Convert.ToString(dt.Rows[0]["GTCITYMASTID"].ToString());
                        txtcity.Text = Convert.ToString(dt.Rows[0]["cityname"].ToString());
                        combostate.Text = Convert.ToString(dt.Rows[0]["statename"].ToString());
                        combocountry.Text = Convert.ToString(dt.Rows[0]["countryname"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //private void ListView1_ItemActivate(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (listView1.Items.Count > 0)
        //        {

        //            txtcityid.Text = listView1.SelectedItems[0].SubItems[2].Text;
        //            string sel1 = "select a.GTCITYMASTID,a.cityname, b.statename , c.countryname  , a.active    from  GTCITYMAST a  join GTSTATEMAST b on a.state=b.GTSTATEMASTID   join GTCOUNTRYMAST c on b.country=c.GTCOUNTRYMASTID  where a.GTCITYMASTID=" + txtcityid.Text;
        //            DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTCITYMAST");
        //            DataTable dt = ds.Tables["GTCITYMAST"];

        //            if (dt.Rows.Count > 0)
        //            {
        //                txtcityid.Text = Convert.ToString(dt.Rows[0]["GTCITYMASTID"].ToString());
        //                txtcity.Text = Convert.ToString(dt.Rows[0]["cityname"].ToString());
        //                combostate.Text = Convert.ToString(dt.Rows[0]["statename"].ToString());
        //                combocountry.Text = Convert.ToString(dt.Rows[0]["countryname"].ToString());
        //                if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
        //                Combostate_SelectedIndexChanged(sender,e);


        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        //public void Txtsearch_TextChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        int item0 = 0;int i = 1;
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
        //                    list.SubItems.Add(item.SubItems[6].Text);
        //                    list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

        //                    listView1.Items.Add(list);
        //                }
        //                item0++;
        //            }
        //        }
        //        else
        //        {

        //                listView1.Items.Clear(); item0 = 0;
        //            foreach (ListViewItem item in listfilter.Items)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                list.Text = i.ToString();
        //                list.SubItems.Add(item.SubItems[1].Text);
        //                list.SubItems.Add(item.SubItems[2].Text);
        //                list.SubItems.Add(item.SubItems[3].Text);
        //                list.SubItems.Add(item.SubItems[4].Text);
        //                list.SubItems.Add(item.SubItems[5].Text);
        //                list.SubItems.Add(item.SubItems[6].Text);
        //                list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

        //                listView1.Items.Add(list);

        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // MessageBox.Show("---" + ex.ToString());
        //    }
        //}

        private void Combostate_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = country(Convert.ToInt64(combostate.SelectedValue));
                combocountry.DataSource = dt;
                combocountry.DisplayMember = "country";
                combocountry.ValueMember = "GTCOUNTRYMASTID";
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message);
            }



        }

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            state();
        }

        public void Deletes()
        {
            if (txtcityid.Text != "")
            {
                string sel1 = "select a.GTCITYMASTID from GTCITYMAST a join GTCOMPMAST b on a.GTCITYMASTID=b.city where a.GTCITYMASTID='" + txtcityid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTCITYMAST");
                DataTable dt = ds.Tables["GTCITYMAST"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcity.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from GTCITYMAST where GTCITYMASTID='" + Convert.ToInt64("0" + txtcityid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtcity.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
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

        private void txtcity_TextChanged(object sender, EventArgs e)
        {
            if (txtcityid.Text == "" && txtcity.Text.Length >= 5)
            {
                string sel = "select gtcitymastid    from  gtcitymast    WHERE state='" + combostate.SelectedValue.ToString() + "' and cityname='" + txtcity.Text.ToUpper().Trim() + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcitymast");
                DataTable dt = ds.Tables["gtcitymast"];
                if (dt.Rows.Count > 0)
                {
                    txtcity.Text = "";
                    MessageBox.Show("Child Recrod Found");
                }
            }
        }

    
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.TreeView
{
    public partial class NavigationMaster : Form,ToolStripAccess
    {
        private static NavigationMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.Navigation c = new Models.Navigation();
        Models.MenuName m = new Models.MenuName();
        public static NavigationMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new NavigationMaster();
                _instance.Font = Class.Users.FontName;
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public void ReadOnlys()
        {

        }
        public NavigationMaster()
        {
            InitializeComponent(); 
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;

        }
     
        public void GridLoad()
        {
            GlobalVariables.HideCols = new string[] { "MENUID1", "GTCOMPMASTID", "USERID","PARENTMENUID","MENUNAMEID1" };
            DataTable dt = c.select(combocompcode1.Text,combousername1.Text);
            GlobalVariables.WidthCols = new Int32[] {10,200,270};
            if (dt.Rows.Count > 0)
            {            
                mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
              
                Class.Users.TableNameGrid = "asptblnavigation";
                if (dt.Rows.Count > 0)
                {
                    int cnt = 0; dataGridView1.Rows.Clear();
                    int i = 0;

                    i = 0;

                    foreach (DataColumn item1 in dt.Columns)
                    {
                        mas.ColIndex.Add(item1.ColumnName);
                    }
                    i = 0;                   
                    if (mas.ColIndex.Count > 0)
                    {
                        CommonFunctions.AddColumnDynamic(dataGridView1, mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(), Class.Users.Query, GlobalVariables.HideCols,GlobalVariables.WidthCols);
                    }
                   
                    mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
                }
                lbltotal.Text = "Total Rows :  " + dataGridView1.Rows.Count.ToString();
            }
            
            Class.Users.Query = "";Class.Users.TableNameGrid = "";
        }
        private void combocode()
        {
            DataTable dt1 = mas.findcomcode(Class.Users.HCompcode);
            if (dt1.Rows.Count > 0)
            {
                combocompcode.DisplayMember = "COMPCODE";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt1;
                
            }

            DataTable dt2 = mas.findcomcode(Class.Users.HCompcode);
            if (dt2.Rows.Count > 0)
            {
              
                combocompcode1.DisplayMember = "COMPCODE";
                combocompcode1.ValueMember = "gtcompmastid";
                combocompcode1.DataSource = dt2;
            }
            //combocompcode.SelectedIndex = -1;
            //combocompcode1.SelectedIndex = -1;
        }
        //private void combocode2()
        //{
        //    DataTable dt1 = mas.comcode1();
        //    if (dt1.Rows.Count > 0)
        //    {


        //        combocompcode1.DisplayMember = "COMPCODE";
        //        combocompcode1.ValueMember = "gtcompmastid";
        //        combocompcode1.DataSource = dt1;


        //    }
        //    combocompcode1.SelectedIndex = -1;

        //}
        private void menuname()
        {

            DataTable dt1 = m.menuname();
            if (dt1.Rows.Count > 0)
            {
                combomenunameid.DisplayMember = "MENUNAME";
                combomenunameid.ValueMember = "MENUNAMEID";
                combomenunameid.DataSource = dt1;

            }
            combomenunameid.SelectedIndex = -1;
        }

        private void NavigationMaster_Load(object sender, EventArgs e)
        {
            combocode();
            menuname();

            GridLoad();
            
            //CommonFunctions.RemoveColumn(dataGridView1, GlobalVariables.HideCols);
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
        }

        private void NavigationMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        public void Saves()
        {


            try
            {
                if (combomenunameid.Text != null && txtmenuname.Text != null && txtnavurl.Text != null && txtparentmenuid.Text != null)
                {
                    c.MenuID = Convert.ToInt64("0" + txtmenuid.Text);
                    c.MenuNameID = Convert.ToInt64("0" + combomenunameid.SelectedValue);
                    c.MenuName = Convert.ToString(txtmenuname.Text);
                    c.NavURL = Convert.ToString(txtnavurl.Text);
                    c.ParentMenuID = Convert.ToInt64("0" + txtparentmenuid.Text);
                    c.CompCode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    c.UserName = Convert.ToInt64("0" + combousername.SelectedValue);

                    if (chk.Checked == true) c.Active = "T"; else c.Active = "F";
                    if (c.MenuNameID >= 0 && c.MenuName != null && c.NavURL != null && c.CompCode > 0 && c.UserName > 0)
                    {

                        DataTable dt = c.select(c.MenuName, c.NavURL, c.ParentMenuID, c.Active, c.MenuNameID, c.CompCode, c.UserName);

                        if (dt.Rows.Count != 0) { MessageBox.Show("Child Record Found.Can't Add Rows"); }
                        else if (dt.Rows.Count != 0 && c.MenuID == 0 || c.MenuID == 0)
                        {

                            c = new Models.Navigation(c.MenuName, c.NavURL, c.ParentMenuID, c.Active, c.MenuNameID, c.CompCode, c.UserName);
                            MessageBox.Show("Record Saved Successfully"); empty(); GridLoad();
                        }
                        else { c = new Models.Navigation(c.MenuName, c.NavURL, c.ParentMenuID, c.Active, c.MenuNameID, c.CompCode, c.UserName, c.MenuID); MessageBox.Show("Record Updated Sucessfully"); empty(); GridLoad(); }

                    }
                    else
                    {
                        MessageBox.Show("pls Enter the Mandatry Fields");
                    }
                }
                else
                {
                    MessageBox.Show("pls Enter the Mandatry Fields");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception" + ex.ToString());
            }
        }

       
        private void empty()
        {
            txtmenuid.Text = "";
            combomenunameid.SelectedIndex = -1;
            Class.Users.Query = "";
            Class.Users.TableNameGrid = "";
            txtmenuname.Text = "";
            txtnavurl.Text = "";
            txtparentmenuid.Text = "";
            combocompcode.SelectedIndex = -1;
            combousername.SelectedIndex = -1;lbltotal.Text = "";
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;           
            chk.Checked = false;
           
            CommonFunctions.SetRowNumber(dataGridView1);
            txtmenuname.Select();
        }

        private void Combomenunameid_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (combomenunameid.SelectedIndex >= 0)
                {
                    Int64 s = Convert.ToInt64(combomenunameid.SelectedValue);
                    DataTable dt1 = m.menuname(s);
                    if (dt1.Rows.Count > 0)
                    {
                        txtparentmenuid.Text = dt1.Rows[0]["parentmenuid"].ToString();
                        txtmenuname.Text = dt1.Rows[0]["MENUNAME"].ToString();
                        txtnavurl.Text = dt1.Rows[0]["MENUNAME"].ToString();
                        chk.Checked = true;
                    }
                }
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message.ToString());
            }
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {

                    DataGridViewRow dv = dataGridView1.Rows[e.RowIndex];
                    txtmenuid.Text = dv.Cells[0].Value.ToString();
                    combomenunameid.Text= dv.Cells[1].FormattedValue.ToString();
                    txtmenuname.Text = dv.Cells[1].Value.ToString();
                    txtnavurl.Text = dv.Cells[2].Value.ToString();
                    txtparentmenuid.Text = dv.Cells[3].Value.ToString();                  
                    if (dv.Cells[4].Value.ToString() == "T")
                        chk.Checked = true;
                    else
                        chk.Checked = false;
                    combocompcode.SelectedValue = dv.Cells[6].Value.ToString();
                    DataTable dt = mas.comcode2(Convert.ToInt64(dv.Cells[6].Value));

                    if (dt.Rows.Count > 0)
                    {
                      
                       // combocompcode.DisplayMember = "COMPCODE";
                      // combocompcode.ValueMember = "gtcompmastid";

                        combousername.DisplayMember = "USERNAME";
                        combousername.ValueMember = "USERID";


                        //combomenunameid.DisplayMember = "MENUNAME";
                        //combomenunameid.ValueMember = "MENUNAMEID";

                    }
                   // combocompcode.DataSource = dt;
                    combousername.DataSource = dt;
                   // combomenunameid.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (combocompcode.SelectedIndex >= 0)
                {
                    Int64 s = Convert.ToInt64(combocompcode.SelectedValue);
                    DataTable dt1 = mas.comcode1(s);
                    if (dt1.Rows.Count > 0)
                    {
                        combousername.DisplayMember = "USERNAME";
                        combousername.ValueMember = "USERID";
                        combousername.DataSource = dt1;
                      

                    }
                   // combousername.SelectedIndex = -1;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        public void News()
        {
            empty();
        }
      

        private void Combocompcode1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combocompcode1.SelectedValue) > 0)
                {
                    Int64 s = Convert.ToInt64(combocompcode1.SelectedValue);
                    DataTable dt1 = mas.comcode2(s);
                    if (dt1.Rows.Count > 0)
                    {
                        combousername1.DisplayMember = "USERNAME";
                        combousername1.ValueMember = "USERID";
                        combousername1.DataSource = dt1;


                    }
                //    combousername1.SelectedIndex = -1;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
           
        }

        private void Combousername1_SelectedIndexChanged(object sender, EventArgs e)
        {
        
        }

        private void MenuMasterRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            menuname();
            combocode(); GridLoad();
        }

        private void TabPage2_Click(object sender, EventArgs e)
        {

        }

      
        public void Prints()
        {
           
        }

        public void Searchs()
        {
           
        }

        public void Deletes()
        {
           if(txtmenuid.Text != "")
            {
                string del = "delete from asptblnavigation where menuid=" + txtmenuid.Text;
                Utility.ExecuteNonQuery(del);
                MessageBox.Show("Record Deleted Successfully.");
                GridLoad();
            }
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
            this.Hide();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

   

        public void Searchs(int EditID)
        {
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Class.Users.Query = "";
                if (txtsearch.Text != "")
                {

                    DataTable dt = c.select(combocompcode1.Text, combousername1.Text,txtsearch.Text);
                 
                    if (dt.Rows.Count > 0)
                    {
                       Class.Users.TableNameGrid = ""; Class.Users.TableNameGrid = "asptblnavigation";

                        int cnt = 0; dataGridView1.Rows.Clear();
                        int i = 0;

                        i = 0;

                        foreach (DataColumn item1 in dt.Columns)
                        {
                            mas.ColIndex.Add(item1.ColumnName);
                        }

                        i = 0;
                        //foreach (DataRow item2 in dt4.Rows)
                        //{
                        //    mas.SizeIndex.Add(item2[i].ToString());
                        //    i++;
                        //}

                        if (mas.ColIndex.Count > 0)
                        {
                            CommonFunctions.AddColumnDynamic(dataGridView1, mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(), Class.Users.Query, GlobalVariables.HideCols,GlobalVariables.WidthCols);
                        }
                        mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
                    }
                    //if (dt.Rows.Count > 0)
                    //{
                    //    dataGridView1.Rows.Clear();
                    //    for (int i = 0; i < dt.Rows.Count; i++)
                    //    {
                    //        dataGridView1.Rows.Add();
                    //        dataGridView1.Rows[i].Cells[0].Value = dt.Rows[i]["menuid1"].ToString();
                    //        dataGridView1.Rows[i].Cells[1].Value = dt.Rows[i]["menuname1"].ToString();
                    //        dataGridView1.Rows[i].Cells[2].Value = dt.Rows[i]["navurl"].ToString();
                    //        dataGridView1.Rows[i].Cells[3].Value = dt.Rows[i]["parentmenuid"].ToString();
                    //        if (dt.Rows[i]["ACTIVE"].ToString() == "T") { dataGridView1.Rows[i].Cells[4].Value = true; } else { dataGridView1.Rows[i].Cells[5].Value = false; }
                    //        dataGridView1.Rows[i].Cells[5].Value = dt.Rows[i]["menunameid1"].ToString();
                    //        dataGridView1.Rows[i].Cells[6].Value = dt.Rows[i]["gtcompmastid"].ToString();
                    //        dataGridView1.Rows[i].Cells[7].Value = dt.Rows[i]["USERID"].ToString();
                    //    }
                    //    CommonFunctions.SetRowNumber(dataGridView1);
                    //}



                }
                else
                {
                    dataGridView1.Rows.Clear();
                    GridLoad();
                }


            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
        }
    }
}


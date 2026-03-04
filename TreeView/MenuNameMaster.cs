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
    public partial class MenuNameMaster : Form,ToolStripAccess
    {
        private static MenuNameMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
        public static MenuNameMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new MenuNameMaster();
                _instance.Font = Class.Users.FontName;
                GlobalVariables.CurrentForm = _instance; 
                return _instance;
            }
        }
        public MenuNameMaster()
        {
            InitializeComponent();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            GlobalVariables.CurrentForm = this;
        }
   
        public void GridLoad()
        {
            GlobalVariables.HideCols = new string[] { "PARENTMENUID" };
            mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
            DataTable dt4 = c.select();
            Class.Users.TableNameGrid = "asptblmenuname";
            GlobalVariables.WidthCols = new Int32[] { 60,200};
            if (dt4.Rows.Count > 0)
            {
                int cnt = 0; dataGridView1.Rows.Clear();
                int i = 0;

                i = 0;

                foreach (DataColumn item1 in dt4.Columns)
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
                    CommonFunctions.AddColumnDynamic(dataGridView1, mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(), Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols);
                }
                lbltotal.Text = "Total : " + dataGridView1.Rows.Count;
                mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
            }
        }
        public void ReadOnlys()
        {

        }
        private void MenuNameMaster_Load(object sender, EventArgs e)
        {
            
                GridLoad();
            
            
        }

        private void MenuNameMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

  
        private void ToolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }


        public void News()
        {
            empty();
        }
        public void Saves()
        {
            try
            {
                if (txtmenuname.Text == "") { MessageBox.Show("pls Enter MenuName"); return; }
                if (txtdisplayname.Text == "") { MessageBox.Show("pls Enter Display Name"); return; }
                else
                {
                    c.MENUNAMEID = Convert.ToInt32("0" + txtmenunameid.Text);
                    c.MENUNAME = Convert.ToString(txtmenuname.Text);
                    c.PARENTMENUID = Convert.ToInt32("0" + comboparentmenuid.Text);
                    c.CREATEBY = Class.Users.HUserName;
                    c.IPADDRESS = GenFun.GetLocalIPAddress();
                    c.CREATEON = Convert.ToString(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
                    c.MODIFYON = Convert.ToString(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
                    c.ALIASNAME = txtdisplayname.Text;
                    if (chk.Checked == true) c.ACTIVE = "T"; else c.ACTIVE = "F";
                    DataTable dt = c.select(c.MENUNAME, c.ALIASNAME, c.ACTIVE, c.PARENTMENUID);
                    if (dt.Rows.Count != 0) { MessageBox.Show("Child Record Found.Can't Add Rows"); }
                    else if (dt.Rows.Count != 0 && c.MENUNAMEID == 0 || c.MENUNAMEID == 0) { c = new Models.MenuName(c.MENUNAME, c.ALIASNAME, c.ACTIVE, c.PARENTMENUID, c.CREATEON, c.CREATEBY, c.MODIFYON, c.IPADDRESS); MessageBox.Show("Record Saved Successfully"); GridLoad(); empty(); }
                    else { c = new Models.MenuName(c.MENUNAME, c.ALIASNAME, c.ACTIVE, c.PARENTMENUID, c.CREATEBY, c.MODIFYON, c.IPADDRESS, c.MENUNAMEID); MessageBox.Show("Record Updated Sucessfully"); GridLoad(); empty(); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception" + ex.ToString());
            }

        }
        private void empty()
        {
            txtmenunameid.Text = "";
            txtmenuname.Text = ""; Class.Users.Query = "";
            Class.Users.TableNameGrid = "";
            comboparentmenuid.SelectedIndex = -1;
            chk.Checked = false;
            _instance.Font = Class.Users.FontName;
          
            butheader.Text = Class.Users.ScreenName;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            CommonFunctions.SetRowNumber(dataGridView1);

            txtmenuname.Select();
        }
        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {

                    DataGridViewRow dv = dataGridView1.Rows[e.RowIndex];

                    txtmenunameid.Text = dv.Cells[0].Value.ToString();
                    txtmenuname.Text = dv.Cells[1].Value.ToString();
                    txtdisplayname.Text = dv.Cells[2].Value.ToString();
                    comboparentmenuid.Text = dv.Cells[3].Value.ToString();
                    if (dv.Cells[4].Value.ToString() == "T")
                        chk.Checked = true;
                    else
                        chk.Checked = false;


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void Comboparentmenuid_SelectedIndexChanged(object sender, EventArgs e)
        {
            chk.Checked = true;
        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void Txtmenusearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Class.Users.Query = "";
                Class.Users.TableNameGrid = "";
                if (txtmenusearch.Text != "")
                {
                    DataTable dt4 = c.menuname(txtmenusearch.Text);
                    if (dt4.Rows.Count > 0)
                    {
                        Class.Users.TableNameGrid = "asptblmenuname";

                        int cnt = 0; dataGridView1.Rows.Clear();
                        int i = 0;

                        i = 0;
                        
                            foreach (DataColumn item1 in dt4.Columns)
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
                }


            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
        }
        private void dataload()
        {
            //try
            //{
            //    DataTable dt = c.select();
            //    if (dt.Rows.Count > 0)
            //    {
            //        dataGridView1.DataSource = dt;
            //        for (int i = 0; i < dt.Rows.Count; i++)
            //        {

            //            if (dt.Rows[i]["ACTIVE"].ToString() == "T") { dataGridView1.Rows[i].Cells[4].Value = true; } else { dataGridView1.Rows[i].Cells[4].Value = false; }
            //            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

            //        }
            //        lbltotal.Text = "Total Rows  :  " + dataGridView1.Rows.Count.ToString();
            //    }
            //    else
            //    {
            //        dataGridView1.DataSource = dt;
            //        for (int i = 0; i < dt.Rows.Count; i++)
            //        {
            //            dataGridView1.Rows[i].Cells[i].Value = false;
            //        }

            //        MessageBox.Show("Screen Rights UnDefined", "Error");
            //    }
            //    CommonFunctions.SetRowNumber(dataGridView1);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

      

     

       
        public void Prints()
        {
          // 
        }

        public void Searchs()
        {
           //
        }

        public void Deletes()
        {
            if (txtmenunameid.Text != "")
            {
                string del = "delete from asptblmenuname where menunameid=" + txtmenunameid.Text;
                Utility.ExecuteNonQuery(del);
                MessageBox.Show("Record Deleted Successfully.");
                GridLoad();
            }
        }

        public void Imports()
        {
            //throw new NotImplementedException();
        }

        public void Pdfs()
        {
           //
        }

        public void ChangePasswords()
        {
           
        }

        public void DownLoads()
        {
           //throw new NotImplementedException();
        }

        public void ChangeSkins()
        {
            //throw new NotImplementedException();
        }

        public void Logins()
        {
            //throw new NotImplementedException();
        }

        public void GlobalSearchs()
        {
           //
        }

        public void TreeButtons()
        {
           //
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
            throw new NotImplementedException();
        }
    }
}

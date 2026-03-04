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
    public partial class DepartmentMaster : Form,ToolStripAccess
    {
        private static DepartmentMaster _instance;
       
        public static DepartmentMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DepartmentMaster();
                _instance.Font = Class.Users.FontName;
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights(); Models.Department dep = new Models.Department();
        public DepartmentMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS= GenFun.GetLocalIPAddress();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }
        public void ReadOnlys()
        {

        }
        public void usercheck(string s, string ss, string sss)
        {

            //DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            //if (dt1.Rows.Count > 0)
            //{
            //    if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
            //    {
            //        for (int r = 0; r < dt1.Rows.Count; r++)
            //        {


            //            if (dt1.Rows[r]["NEWS"].ToString() == "T") { GlobalVariables.News.Visible = true; } else { GlobalVariables.News.Visible = false; }
            //            if (dt1.Rows[r]["SAVES"].ToString() == "T") { GlobalVariables.Saves.Visible = true; } else { GlobalVariables.Saves.Visible = false; }
            //            if (dt1.Rows[r]["PRINTS"].ToString() == "T") { GlobalVariables.Prints.Visible = true; } else { GlobalVariables.Prints.Visible = false; }
            //            if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
            //            if (dt1.Rows[r]["SEARCH"].ToString() == "T") { GlobalVariables.Searchs.Visible = true; } else { GlobalVariables.Searchs.Visible = false; }
            //            if (dt1.Rows[r]["DELETES"].ToString() == "T") { GlobalVariables.Deletes.Visible = true; } else { GlobalVariables.Deletes.Visible = false; }
            //            if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { GlobalVariables.TreeButtons.Visible = true; } else { GlobalVariables.TreeButtons.Visible = false; }
            //            if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { GlobalVariables.GlobalSearchs.Visible = true; } else { GlobalVariables.GlobalSearchs.Visible = false; }
            //            if (dt1.Rows[r]["LOGIN"].ToString() == "T") { GlobalVariables.Logins.Visible = true; } else { GlobalVariables.Logins.Visible = false; }
            //            if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { GlobalVariables.ChangePasswords.Visible = true; } else { GlobalVariables.ChangePasswords.Visible = false; }
            //            if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { GlobalVariables.ChangeSkins.Visible = true; } else { GlobalVariables.ChangeSkins.Visible = false; }
            //            if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { GlobalVariables.DownLoads.Visible = true; } else { GlobalVariables.DownLoads.Visible = false; }
            //            if (dt1.Rows[r]["Pdf"].ToString() == "T") { GlobalVariables.Pdfs.Visible = true; } else { GlobalVariables.Pdfs.Visible = false; }
            //            if (dt1.Rows[r]["Imports"].ToString() == "T") { GlobalVariables.Imports.Visible = true; } else { GlobalVariables.Imports.Visible = false; }

            //        }
            //    }


            //}
            //else
            //{

            //}

        }
        private void DepartmentMaster_Load(object sender, EventArgs e)
        {
            GridLoad();
        }

        public void GridLoad()
        {
            try
            {
                DataTable dt = dep.select();
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        if (dt.Rows[i]["ACTIVE"].ToString() == "T") { dataGridView1.Rows[i].Cells[2].Value = true; } else { dataGridView1.Rows[i].Cells[2].Value = false; }

                    }
                    lbltotal.Text = "Total Rows  :  " + dataGridView1.Rows.Count.ToString();
                }
                else
                {
                    dataGridView1.DataSource = dt;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dataGridView1.Rows[i].Cells[i].Value = false;
                    }

                    MessageBox.Show("Screen Rights UnDefined", "Error");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void DepartmentMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            this.Hide();

            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
        }

        public void News()
        {
            empty(); GridLoad();
        }

        public void Saves()
        {
            try
            {
                Models.Department dep = new Models.Department();
                dep.DepartmentId = Convert.ToInt64("0" + txtdeptid.Text);
                dep.DepartmentName = Convert.ToString(txtdept.Text);
                dep.UserName = Convert.ToInt64(Class.Users.USERID);
                dep.IpAddress = Convert.ToString(Class.Users.IPADDRESS);
                dep.Createdby = Convert.ToString(Class.Users.CREATED);
                dep.Createdon = Convert.ToString(Class.Users.CREATED);
                dep.Modifiedon = Convert.ToString(Class.Users.CREATED);
                if (chkactive.Checked == true) { dep.Active = "T"; } else { dep.Active = "F"; }
                if (!string.IsNullOrEmpty(dep.DepartmentName))
                {


                    DataTable dt = dep.select(dep.DepartmentName, dep.Active);

                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("Child Record Found " + "        " + dep.DepartmentName, "Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        empty(); txtdept.Focus();
                    }
                    else if (dt.Rows.Count == 0 && dep.DepartmentId == 0 || dep.DepartmentId == 0)
                    {

                        dep = new Models.Department(dep.DepartmentName, dep.Active, dep.UserName, dep.IpAddress, dep.Createdon, dep.Createdby, dep.Modifiedon);

                        MessageBox.Show("Record Saved Successfully " + "        " + dep.DepartmentName, "Success Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        empty(); GridLoad();
                    }
                    else
                    {
                        dep = new Models.Department(dep.DepartmentName, dep.Active, dep.UserName, dep.IpAddress, dep.Createdby, dep.Modifiedon, dep.DepartmentId);
                        MessageBox.Show("Record Updated Successfully " + "        " + dep.DepartmentName, "Update Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        empty(); GridLoad();
                    }

                }
                else
                {
                    MessageBox.Show("'Qr Code Field'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Security Inventry " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

       
       
       

        private void empty()
        {
            txtdept.Text = "";
            txtdeptid.Text = "";
            chkactive.Checked = true;
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {

                    DataGridViewRow dv = dataGridView1.Rows[e.RowIndex];
                    
                    txtdeptid.Text = dv.Cells["gtdeptdesgmastid"].Value.ToString();
                    txtdept.Text = dv.Cells["department"].Value.ToString();                 
                    if (dv.Cells["active"].Value.ToString() == "True")

                        chkactive.Checked = true;
                    else
                        chkactive.Checked = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void Txtdeptsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {

                if (txtdeptsearch.Text.ToUpper() != "")
                {
                    
                    string sel1 = "select gtdeptdesgmastid,department,active from gtdeptdesgmast   where department like'%" + txtdeptsearch.Text.ToUpper() + "%'  order by 1 ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtdeptdesgmast");
                    DataTable dt = ds.Tables["gtdeptdesgmast"];
                    if (dt.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dt;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            if (dt.Rows[i]["active"].ToString() == "T") { dataGridView1.Rows[i].Cells[2].Value = true; } else { dataGridView1.Rows[i].Cells[2].Value = false; }

                        }
                        lbltotal.Text = "Total Rows  :  " + dataGridView1.Rows.Count.ToString();
                    }
                    else
                    {
                        MessageBox.Show("No Data Found");
                        txtdeptsearch.Focus();
                    }

                }
                else
                {
                    GridLoad();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void RefreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }

        
        public void Prints()
        {
           
        }

        public void Searchs()
        {
           
        }

        public void Deletes()
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

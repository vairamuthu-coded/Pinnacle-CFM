using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Pinnacle.Models;
namespace Pinnacle.TreeView
{
    public partial class UserMaster : Form,ToolStripAccess
    {
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        private static UserMaster _instance;
        private int compCode;
        private string empNo;
        private int dept;
        private string userName;
        private string gateName;
        private string strpass;
        private string active;
        private string ipaddress;
        private string createdon;

        public static UserMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new UserMaster();
                _instance.Font = Class.Users.FontName;
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public void ReadOnlys()
        {

        }

        public UserMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;

        }

    
       
        private void UserMaster_Load(object sender, EventArgs e)
        {
            try
            {
                
                DataTable dt1 = mas.comcode();
                DataTable dt2 = mas.dept();
               
               

                if (dt1.Rows.Count > 0)
                {
                    combo_compcode.DisplayMember = "COMPCODE";
                    combo_compcode.ValueMember = "gtcompmastid";
                    combo_compcode.DataSource = dt1;
                }
                if (dt2 != null)
                {
                    combo_dept.DisplayMember = "DEPARTMENT";
                    combo_dept.ValueMember = "gtdeptdesgmastid";
                    combo_dept.DataSource = dt2;
                }


                empty();
                combo_dept.SelectedIndex = -1;
                Comboempname_SelectedIndexChanged(sender,e); GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
           
            }

            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
        }
        private void empty()
        {
            txt_userid.Text = "";
           
            comboempname.Text = "";
            combo_dept.SelectedIndex = -1;
            txt_username.Text = "";
            txtgatename.Text = "";
            txt_password.Text = "";
            txtpassworddecript.Text = "";
            GlobalVariables.HideCols = new string[] { "USERID", "FINYEAR", "COMPCODE1" };          
            CommonFunctions.RemoveColumn(dataGridView1, GlobalVariables.HideCols);
            butheader.Text = Class.Users.ScreenName;
            this.Font = Class.Users.FontName;
            dataGridView1.Font = Class.Users.FontName;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = Class.Users.FontName;
            this.dataGridView1.RowTemplate.DefaultCellStyle.Font = Class.Users.FontName;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.DefaultCellStyle.BackColor = Class.Users.BackColors;
            dataGridView1.BackColor = Class.Users.BackColors;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            Activechk.Checked = true;
            txt_username.Select();
        }
        private void Btn_save_Click(object sender, EventArgs e)
        {

        }

        private void Btn_exit_Click(object sender, EventArgs e)
        {
            News();
            
        }



        private void UserMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        //private void Btn_Exit_Click_1(object sender, EventArgs e)
        //{
        //    this.Close();
        //}

        //private void Exit_Click(object sender, EventArgs e)
        //{
          
        //}

        private void Btn_Save_Click_1(object sender, EventArgs e)
        {

        }

        private void Combo_finyear_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Txt_userid_TextChanged(object sender, EventArgs e)
        {

        }

        private void Combo_compcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtgatename.Text = "";
            txtgatename.Text = combo_finyear.Text + "/" + combo_compcode.Text + "/";            
             GridLoad();
            CommonFunctions.SetRowNumber(dataGridView1);
        }
        public void GridLoad()
        {
            try
            {
                DataTable dt;
                if (Class.Users.HUserName == "VAIRAM")
                {
                    string sel1 = "SELECT A.userid,A.finyear,C.COMPCODE,A.empno,A.DEPT,A.username,A.PASWORD,A.ACTIVE,A.gatename,A.empname,A.newpassword,A.SESSIONTIME  FROM asptblusermas A join gtcompmast c on c.gtcompmastid = a.compcode    where c.compcode='" + combo_compcode.Text + "'  and  NOT  a.username='VAIRAMUTHU'  order by 4 ";// 
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblusermas");
                    dt = ds.Tables["asptblusermas"];
                }
                else
                {

                    string sel1 = "SELECT A.userid,A.finyear,C.COMPCODE,A.empno,A.DEPT,A.username,A.PASWORD,A.ACTIVE,A.gatename,A.empname,A.newpassword,A.SESSIONTIME  FROM asptblusermas A join gtcompmast c on c.gtcompmastid = a.compcode    where c.compcode='" + combo_compcode.Text + "'  and  NOT  a.username='VAIRAM' AND NOT  a.username='VAIRAMUTHU' and NOT  a.username='ADMIN' order by 4 ";// 
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblusermas");
                    dt = ds.Tables["asptblusermas"];
                }
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["ACTIVE"].ToString() == "T") { 
                            dataGridView1.Rows[i].Cells[9].Value = true; 
                        } else {
                            dataGridView1.Rows[i].Cells[9].Value = false; 
                        }

                    }                   
                }
                else
                {
                    do
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
                    }
                    while (dataGridView1.Rows.Count > 1);

                }
                lblcount.Text = "Total Count : " + dataGridView1.Rows.Count.ToString();
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
            CommonFunctions.SetRowNumber(dataGridView1);
        }
        public void Saves()
        {
            try
            {
                Models.UserMaster c = new Models.UserMaster();
                if (Class.Users.HCompcode != null && txt_username.Text !="")
                {
                    c.Userid = Convert.ToInt64("0" + txt_userid.Text);
                    c.FinYear = Convert.ToInt64("0" + combo_finyear.SelectedValue);
                    c.CompCode = Convert.ToInt64(combo_compcode.SelectedValue);
                    c.EMPNAME = Convert.ToInt64("0" + comboempname.SelectedValue);
                    c.Dept = Convert.ToInt64("1" + combo_dept.SelectedValue);
                    c.UserName = txt_username.Text;
                    string gate = "";
                    gate = txtgatename.Text;
                    c.GateName = gate;
                   
                    c.Ipaddress = Convert.ToString(Class.Users.IPADDRESS);
                    c.Createdon = Convert.ToString(Class.Users.CREATED);
                    c.SessionTime = Convert.ToString(txtsessiontime.Text);
                    c.Active = Activechk.Checked == true ? "T" : "F";

                    string strpass = sm.Encrypt(txt_password.Text);
                    c.Password = strpass;
                    DataTable dt1 = c.Select(c.CompCode, c.EMPNAME,c.Dept, c.UserName, c.GateName, c.Password, c.Active,c.SessionTime);
                    if (dt1.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found ------" + c.UserName); empty();


                    }
                    if (dt1.Rows.Count != 0 && Convert.ToInt64(c.Userid) == 0 || Convert.ToInt64(c.Userid) == 0)
                    {
                        c = new Models.UserMaster(c.FinYear, c.CompCode, c.EMPNAME, c.Dept, c.UserName, c.GateName, strpass, c.Active, c.Ipaddress, c.Createdon, c.SessionTime);
                        MessageBox.Show("Record Saved Successfully------" + c.UserName); empty(); GridLoad();
                    }
                    else
                    {

                        c = new Models.UserMaster(c.FinYear, c.CompCode, c.EMPNAME, c.Dept, c.UserName, c.GateName, strpass, c.Active, c.Ipaddress, c.Createdon, c.SessionTime,c.Userid);
                        MessageBox.Show("Record Updated Successfully-----" + c.UserName); empty(); GridLoad();
                    }

                }
                else
                {
                    MessageBox.Show("Invalid");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("error--------" + ex);
            }

        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {

                    DataGridViewRow dv = dataGridView1.Rows[e.RowIndex];
                    txt_userid.Text = dv.Cells["userid"].FormattedValue.ToString();
                    string sel1 = "select  a.userid,A.finyear,c.compcode,a.username ,a.gatename ,a.active,a.pasword,a.sessiontime from asptblusermas a join gtcompmast c on c.gtcompmastid = a.compcode where  a.userid='"+ txt_userid.Text + "'order by 1 ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblusermas");
                    DataTable dt = ds.Tables["asptblusermas"];
                    combo_compcode.Text = dt.Rows[0]["compcode"].ToString();               
                    txt_username.Text = dt.Rows[0]["username"].ToString();                   
                    if (dt.Rows[0]["active"].ToString() == "T")
                        Activechk.Checked = true;
                    else
                        Activechk.Checked = false;
                    txtsessiontime.Text= dt.Rows[0]["sessiontime"].ToString();
                    if (txt_username.Text == "VAIRAM" || txt_username.Text == "ADMIN") { txt_password.Enabled = false; Class.Users.Password = sm.Decrypt(dt.Rows[0]["pasword"].ToString()); }
                    else
                    {
                        txt_password.Text = sm.Decrypt(dt.Rows[0]["pasword"].ToString()); txt_password.Enabled = true;
                    }
                    txtpassworddecript.Text = dt.Rows[0]["pasword"].ToString();

                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message);
            }
            txtuserrightsusername.Text = "";

        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        private void Txtuserrightssearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void Txtuserrightsusername_TextChanged(object sender, EventArgs e)
        {
            try
            {

                if (txtuserrightsusername.Text.ToUpper() != "")
                {

                    string sel1 = "select  a.userid,A.finyear,c.compcode,a.username ,a.gatename ,a.active,a.pasword from asptblusermas a join gtcompmast c on c.gtcompmastid = a.compcode where  a.username like'%" + txtuserrightsusername.Text.ToUpper() + "%'  order by 1 ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblusermas");
                    DataTable dt = ds.Tables["asptblusermas"];
                    if (dt.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dt;


                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            if (dt.Rows[i]["ACTIVE"].ToString() == "T") { dataGridView1.Rows[i].Cells[9].Value = true; } else { dataGridView1.Rows[i].Cells[9].Value = false; }

                        }
                    }
                    else
                    {
                        MessageBox.Show("No Data Found");
                        txtuserrightsusername.Focus();
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

        private void Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ContextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void EmployeeNameRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt3 = mas.EmpName();

            if (dt3.Rows.Count > 0)
            {
                comboempname.DisplayMember = "empname";
                comboempname.ValueMember = "hremploymastid";
                comboempname.DataSource = dt3;
            }
        }

        private void DepartmentRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt2 = mas.dept(Convert.ToInt64("0" + comboempname.SelectedValue));
            if (dt2.Rows.Count > 0)
            {
                combo_dept.DisplayMember = "department";
                combo_dept.ValueMember = "gtdeptdesgmastid";
                combo_dept.DataSource = dt2;
            }
           // combo_dept.SelectedIndex = -1;
        }

        private void Combo_dept_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboempname.SelectedIndex >= 0  && combo_dept.SelectedIndex >= 0)
            //{
            //    string sel1 = "SELECT A.EMPIMAGE  FROM hremploymast A JOIN   GTCOMPMAST B ON B.gtcompmastid = A.COMPCODE JOIN   gtdeptdesgmast C ON C.gtdeptdesgmastid = A.DEPARTMENT   WHERE A.hremploymastID=" + comboempname.SelectedValue + " and c.gtdeptdesgmastid=" + combo_dept.SelectedValue;
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "hremploymast");
            //    DataTable dt = ds.Tables["hremploymast"];
            //    if (dt.Rows.Count > 0)
            //    {
            //        if (dt.Rows[0]["EMPIMAGE"].ToString() != "")
            //        {

            //            byte[] bytes = (byte[])dt.Rows[0]["EMPIMAGE"];
            //            Image img = Models.Device.ByteArrayToImage(bytes);
            //            pictureBox1.Image = img;


            //        }
            //    }
            //}
            txtgatename.Text = "";
            txtgatename.Text = combo_finyear.Text + "/" + combo_compcode.Text + "/";
        }

        private void Comboempname_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //DataTable dt2 = mas.dept(Convert.ToInt64("0" + comboempname.SelectedValue));
                //if (dt2.Rows.Count > 0)
                //{
                //    combo_dept.DisplayMember = "department";
                //    combo_dept.ValueMember = "gtdeptdesgmastid";
                //    combo_dept.DataSource = dt2;
                //}
                //combo_dept.SelectedIndex = -1;
                txtgatename.Text = "";
                txtgatename.Text = combo_finyear.Text + "/" + combo_compcode.Text + "/";
            }
            catch (Exception EX)
            { }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad(); 

        }

        public void News()
        {
            empty();
            GridLoad();
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

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void txt_password_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
           // CommonFunctions.SetRowNumber(dataGridView1, Class.Users.HideCols);

            //this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            txt_password.Text = Class.Users.Password;
            txt_password.Enabled = true; Class.Users.Password = "";
        }
    }

}
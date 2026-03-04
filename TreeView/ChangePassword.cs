using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.TreeView
{
    public partial class ChangePassword : Form
    {
        Models.UserRights sm = new Models.UserRights();
        public ChangePassword()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private void Btnsubmit_Click(object sender, EventArgs e)
        {
            Models.UserMaster c = new Models.UserMaster();

            if (txtpassword.Text == txtnewpassword.Text)
            {

                string sel = "SELECT A.USERID,A.FINYEAR,A.COMPCODE,A.EMPNAME ,A.DEPARTMENT,A.USERNAME ,A.GATENAME ,A.ACTIVE,A.PASWORD FROM asptblusermas A WHERE A.USERNAME='" + txtusername + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPINVENTRY");
                DataTable dt = ds.Tables["ASPINVENTRY"];
                if (dt.Rows.Count != 0)
                {
                    string frmpass = dt.Rows[0]["PASWORD"].ToString();
                    string topass = sm.Encrypt(txtnewpassword.Text);



                    if (frmpass == topass)
                    {
                        c.CompCode = Convert.ToInt64("0" + dt.Rows[0]["COMPCODE"].ToString());
                        c.Userid = Convert.ToInt64("0" + dt.Rows[0]["USERID"].ToString());
                        c.FinYear = Convert.ToInt64(dt.Rows[0]["FINYEAR"].ToString());
                        c.EMPNAME = Convert.ToInt64(dt.Rows[0]["EMPNAME"].ToString());
                        c.Dept = Convert.ToInt64(dt.Rows[0]["DEPARTMENT"].ToString());
                        c.UserName = dt.Rows[0]["USERNAME"].ToString();
                        c.GateName = dt.Rows[0]["GATENAME"].ToString();

                        c.Ipaddress = Convert.ToString(Class.Users.IPADDRESS);
                        c.Createdon = Convert.ToString(Class.Users.CREATED);
                        c.Active = "T";
                        c = new Models.UserMaster(c.FinYear, c.CompCode, c.EMPNAME, c.Dept, c.UserName, c.GateName, frmpass, topass, c.Active, c.Ipaddress, c.Createdon);
                        MessageBox.Show("Updated Successfully------" + c.UserName); empty();
                    }
                }
            }
        }

        private void empty()
        {
            txtusername.Text = "";
            txtnewpassword.Text = "";
            txtpassword.Text = "";
            txtconfirmPassword.Text = "";
        }
    }
}

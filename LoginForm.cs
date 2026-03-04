using FoxLearn.License;
using Pinnacle.Models;
using System;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Net;
using System.Windows.Forms;
namespace Pinnacle
{
    public partial class LoginForm : Form
    {
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        private static string s =ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        string q0 = "select a.host from mysql.user a  where a.host = '%'";
        string q1 = "CREATE USER";
        string q2 = "GRANT ALL  PRIVILEGES ON *.* TO ";
        string q3 = "FLUSH PRIVILEGES;";
        public LoginForm()
        {
            InitializeComponent();
            String[] data = s.Split(',');
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.ConString = s;
            Class.Users.HCompcode = data[1];
            String[] data1 = s.Split(';');
            if (Class.Users.HCompcode == "CFM")
            {
                Class.Users.PortNo = data[2];
                Class.Users.PortIP = data[3];
                Models.Serial.PortType = "TCP/IP";
            }            
            string[] words0 = data1[0].Split('=');
            words1 = data1[1].Split('=');
            pwords = data1[3].Split('=');
            string[] words = data1[2].Split('=');
            Class.Users.SqlProjectID = words1[1].ToString()+"@"+ words0[1].ToString();
            
            Class.Users.ProjectID = words[1].ToString();           
            Class.Users.DataBase = data1[0] + "  CompCode : " + Class.Users.HCompcode;
           
            combo_compcode.Text = Class.Users.HCompcode;
            logo(Class.Users.HCompcode);
        }
        string[] words1; string[] pwords; Image img;
        void companylogo(string c)
        {
            try
            {
                string sel = "select a.companylogo,a.compname from gtcompmast a where a.compcode='" + c + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dtCC = ds.Tables["gtcompmast"];
                
                if (dtCC != null)
                {
                    byte[] stdbytes; Int64 std;
                    if (dtCC.Rows.Count > 0)
                    {
                        foreach (DataRow myRow in dtCC.Rows)
                        {

                            if (myRow["companylogo"] != DBNull.Value)
                            {
                                byte[] bytes = (byte[])myRow["companylogo"];
                                button1.BackgroundImage = Models.Device.ByteArrayToImage(bytes);
                            }
                            else
                            {
                                button1.BackgroundImage = null;
                            }
                        }
                    }

                    else
                    {
                        img = null; stdbytes = null;
                    }
                    Class.Users.HCompName = dtCC.Rows[0]["compname"].ToString();
                }
            }
           
            catch(Exception ex) { }
        }
       
        void logo(string s)
        {
            companylogo(s);
           
                Class.Users.HCompcode = combo_compcode.Text;
                
        }
        private void Btn_sumbit_Click(object sender, EventArgs e)
        {
            try
            {

                string name = "";
                name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
                string MacAddress = "";
                MacAddress = GenFun.GetMacAddress();
                string create = "create table if not exists asptblsessionmas (asptblsessionmasid bigint auto_increment primary key,compcode varchar(100) default null,  username varchar(100) default null,  pasword varchar(100) default null,  systemdate varchar(50),  osuser varchar(50) default null,  sid int default null,  serial int default null,  ipaddress varchar(50) default null,remarks varchar(100) default null,MacAddress varchar(100) default null)";
                Utility.ExecuteNonQuery(create);
                if (ValidateChildren(ValidationConstraints.Enabled))
                {
                    string SS = "";
                    if (txtusername.Text == "VAIRAM")
                    {
                        Class.Users.HUserName = Convert.ToString(txtusername.Text);
                        SS = sm.Encrypt(txt_password.Text);
                        Class.Users.PWORD = SS; Class.Users.HCompcode = combo_compcode.Text;
                        string sel = "select  distinct a.compcode as cc ,b.userid, b.username ,a.compname ,b.gatename,a.gtcompmastid ,b.sessiontime  from   " + Class.Users.ProjectID + ".gtcompmast  a join asptblusermas b on a.gtcompmastid = b.compcode    where a.compcode='" + Class.Users.HCompcode + "'      and b.username='" + Class.Users.HUserName + "'  and b.pasword = '" + Class.Users.PWORD + "' and  b.active='T'  order by 1";//and b.pasword = '" + Class.Users.PWORD + "'
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                        DataTable dt = ds.Tables["asptblusermas"];
                        
                        if (dt.Rows.Count > 0)
                        {
                            //GlobalVariables.Dt = Utility.SQLQuery(q0);
                            //if (GlobalVariables.Dt.Rows.Count <= 0 || GlobalVariables.Dt == null)
                            //{
                            //    Class.Users.Query = q1 + " " + words1[1].ToString() + "'@'%' IDENTIFIED BY '" + pwords[1].ToString() + "'";
                            //    Utility.ExecuteNonQuery(Class.Users.Query);
                            //    Class.Users.Query = q2 + " " + words1[1].ToString() + "'@'%' WITH GRANT OPTION";
                            //    Utility.ExecuteNonQuery(Class.Users.Query);
                            //    Class.Users.Query = q3;
                            //    Utility.ExecuteNonQuery(Class.Users.Query);
                            //}
                            Class.Users.HUserName = dt.Rows[0]["username"].ToString();
                            Class.Users.USERID = Convert.ToInt64(dt.Rows[0]["userid"].ToString());
                            Class.Users.HGateName = System.DateTime.Now.Year + "/" + dt.Rows[0]["gatename"].ToString();
                            Class.Users.HCompName = dt.Rows[0]["compname"].ToString();
                            Class.Users.LoginTime = Convert.ToInt32("0" + dt.Rows[0]["sessiontime"].ToString());
                            Class.Users.COMPCODE = Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());
                            Class.Users.Log = System.DateTime.Now.AddDays(1);
                            Pinnacle.PinnacleMdi si = new Pinnacle.PinnacleMdi();
                            si.Show();
                            this.Hide();
                        }
                        else
                        {
                            MessageBox.Show("Invalid UserName and Password ", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        }

                    }
                    else
                    {
                        Class.Users.HUserName = Convert.ToString(txtusername.Text);
                        SS = sm.Encrypt(txt_password.Text);
                        Class.Users.PWORD = SS; Class.Users.HCompcode = combo_compcode.Text;
                        string sel = "select  distinct a.compcode as cc ,b.userid, b.username ,a.compname ,b.gatename,a.gtcompmastid ,b.sessiontime  from   " + Class.Users.ProjectID + ".gtcompmast  a join asptblusermas b on a.gtcompmastid = b.compcode   and a.compcode='" + Class.Users.HCompcode + "'    and b.username='" + Class.Users.HUserName + "' and b.pasword = '" + Class.Users.PWORD + "'   and  b.active='T'  order by 1";//and b.username='" + Class.Users.HUserName + "' and b.pasword = '" + Class.Users.PWORD + "'
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                        DataTable dt = ds.Tables["asptblusermas"];

                        if (dt.Rows.Count > 0)
                        {
                            string sell = "select max(a.asptblregistrationid) as asptblregistrationid  from asptblregistration a join gtcompmast b on a.compcode=b.gtcompmastid where b.gtcompmastid='" + combo_compcode.SelectedValue.ToString() + "'";
                            DataSet dss = Utility.ExecuteSelectQuery(sell, "asptblregistration");
                            DataTable dtt = dss.Tables["asptblregistration"];
                            if (dtt.Rows.Count > 0)
                            {
                                string sel0 = "select max(a.productkeys) productkeys,max(a.productid) productid from asptblregistration a join gtcompmast b on a.compcode=b.gtcompmastid where b.gtcompmastid='" + combo_compcode.SelectedValue.ToString() + "' and  a.asptblregistrationid='" + dtt.Rows[0]["asptblregistrationid"].ToString() + "'";
                                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "asptblregistration");
                                DataTable dt0 = ds0.Tables["asptblregistration"];
                                if (dt0.Rows.Count > 0)
                                {

                                    txtproductid.Text = dt0.Rows[0]["productid"].ToString();
                                    KeyManager km = new KeyManager(txtproductid.Text); bool val = false; Int64 days1 = 0;
                                    LicenseInfo lic = new LicenseInfo();
                                    string sel1 = ""; DataTable dt1; DataSet ds1;
                                    //lic.FullName = "Pinnacle.lic";
                                    // int value = km.LoadSuretyFile(string.Format(@"{0}\Pinnacle.lic", Application.StartupPath), ref lic);
                                    //string productKey = lic.ProductKey;
                                    string productKey = dt0.Rows[0]["productkeys"].ToString();
                                    if (km.ValidKey(ref productKey))
                                    {
                                        KeyValuesClass kv = new KeyValuesClass();
                                        if (km.DisassembleKey(productKey, ref kv))
                                        {
                                            txtproductkey.Text = productKey;
                                            Class.Users.Log = kv.Expiration;
                                            DateTime lastdate = Convert.ToDateTime(DateTime.Now.Date.AddDays(Convert.ToInt32("0" + 1)));
                                            if (kv.Type == LicenseType.FULL && lastdate >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                                            {
                                                Class.Users.LoginTime = Convert.ToInt32("0" + dt.Rows[0]["sessiontime"].ToString());
                                                Class.Users.HUserName = dt.Rows[0]["username"].ToString();
                                                Class.Users.USERID = Convert.ToInt64(dt.Rows[0]["userid"].ToString());
                                                Class.Users.HGateName = dt.Rows[0]["gatename"].ToString();
                                                Class.Users.HCompName = dt.Rows[0]["compname"].ToString();
                                                Class.Users.COMPCODE = Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());
                                                PinnacleMdi si = new PinnacleMdi();
                                                si.Show();
                                                this.Hide(); return;
                                            }

                                            if (kv.Type == LicenseType.TRIAL && kv.Expiration >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                                            {
                                             
                                                txtexperiencedays.Text = string.Format(@"{0}", (kv.Expiration - DateTime.Now.Date).Days);
                                                days1 = Convert.ToInt64(txtexperiencedays.Text);

                                                if (days1 == 1)
                                                {

                                                    MessageBox.Show("Your License  Expired Today. pls go to Your Administrator.");
                                                    val = true;
                                                    Class.Users.RemainingDays = 0;
                                                    Class.Users.RemainingDays = days1;
                                                }
                                                if (days1 > 1)
                                                {
                                                    val = true;
                                                    Class.Users.RemainingDays = 0;
                                                    Class.Users.RemainingDays = days1;
                                                }

                                            }
                                            else
                                            {
                                                val = false;

                                                MessageBox.Show("pls Contact your Administrator." + Class.Users.Log.ToString(), "Register Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                return;
                                            }

                                        }
                                        if (val == true)
                                        {
                                            Class.Users.LoginTime = Convert.ToInt32("0" + dt.Rows[0]["sessiontime"].ToString());
                                            Class.Users.HUserName = dt.Rows[0]["username"].ToString();
                                            Class.Users.USERID = Convert.ToInt64(dt.Rows[0]["userid"].ToString());
                                            Class.Users.HGateName = dt.Rows[0]["gatename"].ToString();
                                            Class.Users.HCompName = dt.Rows[0]["compname"].ToString();
                                            Class.Users.COMPCODE = Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());
                                            PinnacleMdi si = new PinnacleMdi();
                                            si.Show();
                                            this.Hide();
                                        }
                                        else
                                        {
                                            MessageBox.Show("pls Contact your Administrator. Last Date Expired ." + Class.Users.Log.ToString(), "Register Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("License Key doesn't have this solution", "Expired", MessageBoxButtons.OK, MessageBoxIcon.Information); return;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Please Create Licence Key ", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    txtusername.Select();
                                    return;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Invalid UserName and Password ", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            txtusername.Select();
                            return;
                        }



                    }
                    //else
                    //{
                    //    Class.Users.HUserName = Convert.ToString(txtusername.Text);
                    //    SS = sm.Encrypt(txt_password.Text);
                    //    Class.Users.PWORD = SS; Class.Users.HCompcode = combo_compcode.Text;

                    //    string sel = "select  distinct a.compcode as cc ,b.userid, b.username ,a.compname ,b.gatename,a.gtcompmastid ,b.sessiontime  from   " + Class.Users.ProjectID + ".gtcompmast  a join asptblusermas b on a.gtcompmastid = b.compcode   and a.compcode='" + Class.Users.HCompcode + "'      and b.username='" + Class.Users.HUserName + "' and b.pasword = '" + Class.Users.PWORD + "' and  b.active='T'  order by 1";
                    //    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                    //    DataTable dt = ds.Tables["asptblusermas"];
                    //    if (dt.Rows.Count > 0)
                    //    {

                    //        txtproductid.Text = ComputerInfo.GetComputerId();
                    //        KeyManager km = new KeyManager(txtproductid.Text); bool val = false; Int64 days1 = 0;
                    //        LicenseInfo lic = new LicenseInfo();
                    //        string sel1 = ""; DataTable dt1; DataSet ds1;
                    //       // int value = km.LoadSuretyFile(string.Format(@"{0}\Pinnacle" + name + ".lic", Application.StartupPath), ref lic);
                    //        string productKey = lic.ProductKey;
                    //        if (km.ValidKey(ref productKey))
                    //        {
                    //            KeyValuesClass kv = new KeyValuesClass();


                    //            if (km.DisassembleKey(productKey, ref kv))
                    //            {
                    //                txtproductkey.Text = productKey;
                    //                if (kv.Type == LicenseType.TRIAL || kv.Type == LicenseType.FULL)
                    //                {
                    //                    Class.Users.Log = kv.Expiration;
                    //                    if (string.IsNullOrEmpty(Class.Users.Log.ToString()))
                    //                    {
                    //                        Class.Users.LoginTime = Convert.ToInt32("0" + dt.Rows[0]["sessiontime"].ToString());
                    //                        Class.Users.HUserName = dt.Rows[0]["username"].ToString();
                    //                        Class.Users.USERID = Convert.ToInt64(dt.Rows[0]["userid"].ToString());
                    //                        Class.Users.HGateName = dt.Rows[0]["gatename"].ToString();
                    //                        Class.Users.HCompName = dt.Rows[0]["compname"].ToString();
                    //                        Class.Users.COMPCODE = Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());
                    //                        PinnacleMdi si = new PinnacleMdi();
                    //                        si.Show();
                    //                        this.Hide(); return;
                    //                    }
                    //                    if (kv.Expiration >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                    //                    {
                    //                        txtexperiencedays.Text = string.Format(@"{0}", (kv.Expiration - DateTime.Now.Date).Days);
                    //                        days1 = Convert.ToInt64(txtexperiencedays.Text);
                    //                        if (days1 == 1)
                    //                        {
                    //                            string up1 = "delete from asptblsessionmas  WHERE username='" + txtusername.Text + "'";
                    //                            Utility.ExecuteNonQuery(up1);
                    //                            MessageBox.Show("Your License  Expired Today. pls go to Your Administrator.");
                    //                            val = true;
                    //                            Class.Users.RemainingDays = 0;
                    //                            Class.Users.RemainingDays = days1;
                    //                        }
                    //                        if (days1 > 1)
                    //                        {
                    //                            val = true;
                    //                            Class.Users.RemainingDays = 0;
                    //                            Class.Users.RemainingDays = days1;
                    //                        }

                    //                    }
                    //                    else
                    //                    {
                    //                        val = false;

                    //                        MessageBox.Show("pls Contact your Administrator." + Class.Users.Log.ToString(), "Register Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //                        return;
                    //                    }
                    //                }
                    //            }
                    //            if (val == true)
                    //            {
                    //                Class.Users.LoginTime = Convert.ToInt32("0" + dt.Rows[0]["sessiontime"].ToString());
                    //                Class.Users.HUserName = dt.Rows[0]["username"].ToString();
                    //                Class.Users.USERID = Convert.ToInt64(dt.Rows[0]["userid"].ToString());
                    //                Class.Users.HGateName = dt.Rows[0]["gatename"].ToString();
                    //                Class.Users.HCompName = dt.Rows[0]["compname"].ToString();
                    //                Class.Users.COMPCODE = Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());

                    //                PinnacleMdi si = new PinnacleMdi();
                    //                si.Show();
                    //                this.Hide();


                    //            }
                    //            else
                    //            {
                    //                string up1 = "delete from asptblsessionmas WHERE username='" + txtusername.Text + "'";
                    //                Utility.ExecuteNonQuery(up1);
                    //                MessageBox.Show("pls Contact your Administrator. Last Date Expired ." + Class.Users.Log.ToString(), "Register Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //            }
                    //        }
                    //        else
                    //        {
                    //            string up1 = "delete from asptblsessionmas WHERE username='" + txtusername.Text + "'";
                    //            Utility.ExecuteNonQuery(up1);
                    //            MessageBox.Show("License Key doesn't have this solution", "Expired", MessageBoxButtons.OK, MessageBoxIcon.Information); return;
                    //        }

                    //    }
                    //    else
                    //    {
                    //        MessageBox.Show("Invalid UserName and Password ", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //        txtusername.Select();
                    //        return;
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Invalid UserName and Password ", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // txtusername.Select(); 
            }
        }

         
         
        private void LoginForm_Load(object sender, EventArgs e)
        {
           
            try
            {
              //  txtusername.Text = "admin";txt_password.Text = "cfm123";
                string up = "use  " + Class.Users.ProjectID + ";";
                Utility.ExecuteNonQuery(up);
                Class.Users.HUserName = txtusername.Text;     

                GlobalVariables.Dt = Utility.SQLQuery("SELECT SUBSTRING(now(),1,4) AS FINYEAR");
                if (GlobalVariables.Dt.Rows.Count > 0)
                {
                    Class.Users.Finyear = GlobalVariables.Dt.Rows[0]["FINYEAR"].ToString();
                    combofinyear.Text = Class.Users.Finyear;
                }               
                string ss = Class.Users.HCompcode;
                string sel = "select distinct b.gtcompmastid,  b.compcode from  gtcompmast b  join asptblusermas c on c.compcode = b.gtcompmastid  where c.active='T'  and b.compcode='" + ss + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                DataTable dt = ds.Tables["asptblusermas"];
               
                if (dt == null) { MessageBox.Show("Invalid Tables","Invalid",MessageBoxButtons.OK,MessageBoxIcon.Error);return; }
                else
                {
                    if (dt.Rows.Count == 1)
                    {                      
                        combo_compcode.DisplayMember = "compcode";
                        combo_compcode.ValueMember = "gtcompmastid";
                        combo_compcode.DataSource = dt;
                        combo_compcode.Text = dt.Rows[0]["compcode"].ToString();
                        Class.Users.COMPCODE = Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());
                    }
                    else
                    {
                        
                        combo_compcode.DisplayMember = "compcode";
                        combo_compcode.ValueMember = "gtcompmastid";
                        combo_compcode.DataSource = dt;
                        combo_compcode.Text = dt.Rows[0]["compcode"].ToString();
                        Class.Users.COMPCODE= Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }


        }
      
        private void Btn_Exit_Click(object sender, EventArgs e)
        {
            
            Application.Exit();
        }

       
      


        private void Txt_password_Enter(object sender, EventArgs e)
        {
            txt_password.BackColor = System.Drawing.Color.Wheat;
            System.Windows.Forms.Clipboard.Clear();
           
        }

        private void Txt_password_Leave(object sender, EventArgs e)
        {
            
            txt_password.ForeColor = System.Drawing.Color.RoyalBlue;
            txt_password.BackColor = System.Drawing.Color.White;
        }

        private void LoginForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void Buttblcreate_Click(object sender, EventArgs e)
        {
            Tables.butdroptable tt = new Tables.butdroptable();
            tt.Show();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            buttblcreate.Visible = true;
        }

   

        private void txt_password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_sumbit.Focus();
                //e.SuppressKeyPress = true;   // stop beep sound
                //SendKeys.Send("{TAB}");      // move to next control
            }
            
        }

    

        private void txtusername_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //if (string.IsNullOrEmpty(txtusername.Text))
            //{
            //    e.Cancel = true;
            //    txtusername.Focus();
            //    ErrProvider.SetError(txtusername, "Invalid Username");
            //}
            //else
            //{
            //    e.Cancel = false;
            //    ErrProvider.SetError(txtusername, null);
            //}
        }

        private void txt_password_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txt_password.Text))
            {
                e.Cancel = true;
                txt_password.Focus();
                ErrProvider.SetError(txt_password, "Invalid Password");
            }
            else
            {
                e.Cancel = false;
                ErrProvider.SetError(txt_password, null);
            }
        }

        private void combo_compcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(combo_compcode.Text != "")
            {
                logo(Class.Users.HCompcode);
            }
        }

        private void txtusername_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_password.Focus();
            }
        }
    }
}

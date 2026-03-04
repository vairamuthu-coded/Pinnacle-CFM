using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FoxLearn.License;
namespace Pinnacle.Registration
{
    public partial class RegisterMaster : Form,ToolStripAccess
    {
        private static RegisterMaster _instance;
      
        public static RegisterMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new RegisterMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        int totaldays = 0; string name = "";
        public RegisterMaster()
        {
            InitializeComponent();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.Text = Class.Users.ScreenName;           
            name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            textBox1.Text = name;
        }

        public void ReadOnlys()
        {

        }

        const int ProductCode = 1;

        private void butRegistration_Click(object sender, EventArgs e)
        {
            name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            KeyManager km = new KeyManager(txtproductid.Text);
            string productKey = txtproductkey.Text;
            if (km.ValidKey(ref productKey))
            {
                KeyValuesClass kv = new KeyValuesClass();
                if (km.DisassembleKey(productKey, ref kv))
                {
                    LicenseInfo lic = new LicenseInfo();
                    lic.ProductKey = productKey;

                    lic.FullName = "Pinnacle.lic";

                    if (kv.Type == LicenseType.TRIAL || kv.Type == LicenseType.FULL)
                    {
                        lic.Day = kv.Expiration.Day;
                        lic.Month = kv.Expiration.Month;
                        lic.Year = kv.Expiration.Year;
                        km.SaveSuretyFile(string.Format(lic.FullName, Application.StartupPath), lic);
                        txtexperiencedays.Text = string.Format(@"{0}", (kv.Expiration - DateTime.Now.Date).Days);

                        var chk = "";
                        if (checkactive.Checked == true) chk = "T"; else chk = "F";
                        if (name == textBox2.Text)
                        {
                            name = textBox1.Text;
                        }
                        else
                        {
                            if (textBox2.Text == "")
                            {

                            }
                            else
                            {
                                name = textBox2.Text;
                            }
                        }
                        string sel = "select a.productkeys from  asptblregistration a where a.productkeys='" + txtproductkey.Text + "' and days='" + txtexperiencedays.Text + "'  and a.active='" + chk + "' and a.fromdate='" + dateTimePicker1.Value.ToString().Substring(0, 10) + "' and a.todate='" + dateTimePicker2.Value.ToString().Substring(0, 10) + "' and a.productid='" + txtproductid.Text + "'  ;";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblregistration");
                        DataTable dt = ds.Tables["asptblregistration"];
                        if (dt.Rows.Count > 0)
                        {

                            MessageBox.Show("Child Record Found", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt64("0" + txtregistrationid.Text) == 0 || txtregistrationid.Text == "")
                        {
                            string ins = "insert into asptblregistration(finyear,fromdate,todate,productid,productkeys,active,compcode ,username,ipaddress,createdon,createdby,modifiedon,days,systemuser)values('" + Class.Users.Finyear + "','" + System.DateTime.Now.ToString() + "','" + kv.Expiration + "','" + txtproductid.Text + "','" + txtproductkey.Text + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.IPADDRESS + "','" + Class.Users.CREATED + "','" + Class.Users.HUserName + "','" + Class.Users.CREATED + "','" + txtexperiencedays.Text + "','" + name + "')";
                            Utility.ExecuteNonQuery(ins); empty();
                            MessageBox.Show("You have registered", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string ins = "update  asptblregistration set finyear='" + Class.Users.Finyear + "',fromdate='" + System.DateTime.Now.ToString() + "',todate='" + kv.Expiration + "',productid='" + txtproductid.Text + "',productkeys='" + txtproductkey.Text + "',active='" + chk + "',compcode='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',ipaddress='" + Class.Users.IPADDRESS + "',createdon='" + Class.Users.CREATED + "',createdby='" + Class.Users.HUserName + "',modifiedon='" + Class.Users.CREATED + "'  ,days='" + txtexperiencedays.Text + "',systemuser='" + name + "' where asptblregistrationid='" + txtregistrationid.Text + "';";
                            Utility.ExecuteNonQuery(ins);
                            empty();
                            MessageBox.Show("You have Updated", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Your Product Key is Invalid", "Message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }

            }
            else
            {
                MessageBox.Show("Invalid License Key.Pls Contact Your Administrator", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtproductkey.Text = ""; checkactive.Checked = false;
            }
        }
        ListView listfilter = new ListView();
        public  void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); 
                string sel1 = "select asptblregistrationid,productid,productkeys,active,fromdate,todate,systemuser from asptblregistration order by asptblregistrationid desc ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblregistration");
                DataTable dt = ds.Tables["asptblregistration"];
                if (dt !=null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblregistrationid"].ToString());
                        list.SubItems.Add(myRow["productid"].ToString());
                        list.SubItems.Add(myRow["productkeys"].ToString());
                        if (myRow["active"].ToString() == "T")
                            list.SubItems.Add("T");
                        else
                            list.SubItems.Add("F");
                        list.SubItems.Add(myRow["fromdate"].ToString().Substring(0,10));
                        list.SubItems.Add(myRow["todate"].ToString().Substring(0, 10));
                        list.SubItems.Add(myRow["systemuser"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        else
                        {
                            list.BackColor = Color.White;
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
        void empty()
        {
            txtregistrationid.Text = "";
            txtproductkey.Text = ""; textBox1.Text = ""; textBox2.Text = "";
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            txtexperiencedays.Text = ""; 
            this.BackColor = Class.Users.BackColors;
         
            this.Font = Class.Users.FontName;
            name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            textBox1.Text = name;
            txtproductid.Text = ComputerInfo.GetComputerId(); GridLoad();
        }
        private void RegistrationMaster_Load(object sender, EventArgs e)
        {
           // txtproductid.Text = ComputerInfo.GetComputerId(); GridLoad();
        }



        //private void butRegistration_Click(object sender, EventArgs e)
        //{
        //    KeyManager km = new KeyManager(txtproductid.Text);
        //    string productKey = txtproductkey.Text;
        //    if(km.ValidKey(ref productKey))
        //    {
        //        KeyValuesClass kv = new KeyValuesClass();
        //        if(km.DisassembleKey(productKey,ref kv))
        //        {
        //            LicenseInfo lic = new LicenseInfo();
        //            lic.ProductKey = productKey;
        //            lic.FullName = "FoxLearn";
        //            if(kv.Type==LicenseType.TRIAL)
        //            {
                     
        //                lic.Day = kv.Expiration.Day;
        //                lic.Month = kv.Expiration.Month;
        //                lic.Year = kv.Expiration.Year;

        //            }
        //            km.SaveSuretyFile(string.Format(@"{0}\VAI.lic",Application.StartupPath),lic);
        //            MessageBox.Show("You have registered","Message",MessageBoxButtons.OK,MessageBoxIcon.Information);
        //           // this.Close();


        //        }
        //        else
        //        {
        //            MessageBox.Show("Your Product Key is Invalid", "Message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        //        }
              
        //    }

        //}

        //private void RegistrationMaster_Load(object sender, EventArgs e)
        //{
        //    txtproductid.Text = ComputerInfo.GetComputerId();
        //}

      
        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

            totaldays = 0;
            DateTime StartingDate = Convert.ToDateTime(dateTimePicker1.Value.Date.ToString("dd-MM-yyyy"));
            DateTime EndingDate = Convert.ToDateTime(dateTimePicker2.Value.Date.ToString("dd-MM-yyyy"));
            TimeSpan countdays = EndingDate.Subtract(StartingDate);
            totaldays = Convert.ToInt32(countdays.Days);
            txtexperiencedays.Text = totaldays.ToString();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
         
            DateTime StartingDate = Convert.ToDateTime(dateTimePicker1.Value.Date.ToString("dd-MM-yyyy"));
            DateTime EndingDate = Convert.ToDateTime(dateTimePicker2.Value.Date.ToString("dd-MM-yyyy"));
            TimeSpan countdays = EndingDate.Subtract(StartingDate);
            totaldays = Convert.ToInt32(countdays.Days);
            txtexperiencedays.Text = totaldays.ToString();
        }
      
        public void News()
        {
            empty();
        }

        public void Saves()
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
            string del = "delete from asptblregistration a where a.asptblregistrationid='" + txtregistrationid.Text + "'";
            Utility.ExecuteNonQuery(del);
            MessageBox.Show("Record Deleted . "+ txtproductkey.Text);
            empty();
            GridLoad();
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
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
           

            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtregistrationid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select asptblregistrationid,productid,productkeys,active,days,fromdate,todate,systemuser from asptblregistration  where asptblregistrationid=" + txtregistrationid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblregistration");
                    DataTable dt = ds.Tables["asptblregistration"];
                    if (dt.Rows.Count > 0)
                    {
                        txtregistrationid.Text = Convert.ToString(dt.Rows[0]["asptblregistrationid"].ToString());

                        txtproductid.Text = Convert.ToString(dt.Rows[0]["productid"].ToString());
                        txtexperiencedays.Text = Convert.ToString(dt.Rows[0]["days"].ToString());
                        dateTimePicker1.Text = Convert.ToString(dt.Rows[0]["fromdate"].ToString());
                        dateTimePicker2.Text = Convert.ToString(dt.Rows[0]["todate"].ToString());
                        txtproductkey.Text = Convert.ToString(dt.Rows[0]["productkeys"].ToString());
                        textBox2.Text= Convert.ToString(dt.Rows[0]["systemuser"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

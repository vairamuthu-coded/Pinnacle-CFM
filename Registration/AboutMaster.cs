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
    public partial class AboutMaster : Form,ToolStripAccess
    {
        private static AboutMaster _instance;
      
        public static AboutMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new AboutMaster();
                GlobalVariables.CurrentForm = _instance; 
                return _instance;
            }
        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        public AboutMaster()
        {
            InitializeComponent();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            butheader.Text = Class.Users.ScreenName;
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
        }

  
        private void AboutMaster_Load(object sender, EventArgs e)
        {
            News();
        }

        private void butAbout_Click(object sender, EventArgs e)
        {
            string name = "";
            name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            txtproductid.Text = ComputerInfo.GetComputerId();
            KeyManager km = new KeyManager(txtproductid.Text);
            LicenseInfo lic = new LicenseInfo();
            int value = km.LoadSuretyFile(string.Format(@"{0}\Pinnacle.lic", Application.StartupPath), ref lic);
            string productKey = lic.ProductKey;
            if (km.ValidKey(ref productKey))
            {
                KeyValuesClass kv = new KeyValuesClass();
                if (km.DisassembleKey(productKey, ref kv))
                {
                    txtproductkeys.Text = productKey;
                    if (kv.Type == LicenseType.TRIAL)
                    {
                        txtexperiencedays.Text = string.Format(@"{0} days", (kv.Expiration - DateTime.Now.Date).Days);
                        txtlicensetype.Text = LicenseType.TRIAL.ToString();
                    }
                    else
                    {
                        txtexperiencedays.Text = string.Format(@"{0} days", (kv.Expiration));
                        txtlicensetype.Text = LicenseType.FULL.ToString();
                    }
                }
            }
        }

      

        public void News()
        {
            string name = "";
            name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            txtproductid.Text = ComputerInfo.GetComputerId();
            KeyManager km = new KeyManager(txtproductid.Text);
            LicenseInfo lic = new LicenseInfo();
            int value = km.LoadSuretyFile(string.Format(@"{0}\Pinnacle.lic", Application.StartupPath), ref lic);
            string productKey = lic.ProductKey;
            if (km.ValidKey(ref productKey))
            {
                KeyValuesClass kv = new KeyValuesClass();
                if (km.DisassembleKey(productKey, ref kv))
                {
                    txtproductkeys.Text = productKey;
                    if (kv.Type == LicenseType.TRIAL)
                    {
                        txtexperiencedays.Text = string.Format(@"{0} days", (kv.Expiration - DateTime.Now.Date).Days);
                        txtlicensetype.Text = LicenseType.TRIAL.ToString();
                    }
                    else
                    {
                        txtexperiencedays.Text = string.Format(@"{0} days", (kv.Expiration));
                        txtlicensetype.Text = LicenseType.FULL.ToString();
                    }
                }
            }
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
        public void GridLoad()
        {
           
        }
      
        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
           
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }
        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
        }
    }
}

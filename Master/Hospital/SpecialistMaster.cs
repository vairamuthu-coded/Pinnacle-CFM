using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master.Hospital
{
    public partial class SpecialistMaster : Form,ToolStripAccess
    {
        private static SpecialistMaster _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
        public static SpecialistMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new SpecialistMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public SpecialistMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            
        }
        string[] s; 
        private void pop()
        {

            //(select A.GARMENTIMAGE from ASPTBLBUYSAM a where A.ASPTBLBUYSAMID=1)

            string sel = "select b.asptbldocmasid, b.doctorname, a.department,b.experience,b.education,b.doctorphoto,b.medicalno from asptblhosdeptmas a join asptbldocmas b on a.asptblhosdeptmasid=b.department  order by 1;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblhosdeptmas");
            DataTable dt = ds.Tables["asptblhosdeptmas"];
            Pinnacle.UserControls.Hospital.FindDoctor[] items = new UserControls.Hospital.FindDoctor[dt.Rows.Count];
            flowLayoutPanel1.Controls.Clear();

            foreach (DataRow myRow in dt.Rows)
            {
                items[i] = new UserControls.Hospital.FindDoctor();

                if (Class.Users.HCompcode == "AMC")
                {
                    items[i].DoctorID.Text = Convert.ToString(myRow["medicalno"].ToString());
                    items[i].DoctorName.Text = Convert.ToString(myRow["doctorname"].ToString());
                    items[i].DoctorName.ForeColor = System.Drawing.Color.Teal;
                    items[i].DoctorEducation.Text = Convert.ToString(myRow["education"].ToString());
                    items[i].DoctorDept.Text = Convert.ToString(myRow["department"].ToString());
                    items[i].doctorPhoto.Text = Convert.ToString(myRow["asptbldocmasid"].ToString());
                    items[i].Paneldoctor.Text = Convert.ToString(myRow["asptbldocmasid"].ToString());
                    if (myRow["doctorphoto"].ToString() != "")
                    {
                        stdbytes = (byte[])myRow["doctorphoto"];
                        Image img = Models.Device.ByteArrayToImage(stdbytes);
                        items[i].doctorPhoto.BackgroundImage = img;
                    }
                    else
                    {
                        //items[i].doctorPhoto.BackgroundImage = Properties.Resources.AGF_IN;
                    }
                    flowLayoutPanel1.Controls.Add(items[i]);
                    items[i].Paneldoctor.Click += Paneldoctor_Click;
                    items[i].doctorPhoto.Click += DoctorPhoto_Click;
                    items[i].doctorPhoto.MouseHover += DoctorPhoto_MouseHover;
                    items[i].Panel1.BackColor= Class.Users.BackColors;
                    items[i].Panel2.BackColor = Class.Users.BackColors;
                }
            }
           
        }

        private void DoctorPhoto_MouseHover(object sender, EventArgs e)
        {
            
        }

        private void DoctorPhoto_MouseUp(object sender, MouseEventArgs e)
        {
           
        }

        private void DoctorPhoto_Click(object sender, EventArgs e)
        {
           
            s = sender.ToString().Split(',');
            sender = s[1].Substring(7).TrimEnd();
        }

        private void Paneldoctor_Click(object sender, EventArgs e)
        {
           
            s = sender.ToString().Split(',');
            sender = s[1].Substring(7).TrimEnd();

        }

       

        private void SpecialistMaster_Load(object sender, EventArgs e)
        {
            pop(); News();
        }

        public void News()
        {
           
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
        }

        public void Saves()
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

        public void Deletes()
        {
           
        }

        public void ReadOnlys()
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
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pop();
          
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
        }

        public void GridLoad()
        {
           
        }

     
        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

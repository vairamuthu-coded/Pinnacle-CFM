using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Threading.Tasks;

namespace Pinnacle.Fuel
{
    public partial class WeighBridge : Form
    {
        private static WeighBridge _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        public Thread ReadSerialDataThread;
        public string readserialvalue;
        public WeighBridge()
        {
            InitializeComponent();

            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            label1.Text = Class.Users.ScreenName;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());

        }

        public static WeighBridge Instance
        {
            get { if (_instance == null) _instance = new WeighBridge(); GlobalVariables.CurrentForm = _instance; return _instance; }
        }
        public void usercheck(string s, string ss, string sss)
        {
            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                        if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                        if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                        if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                        if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                        if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                        if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                        if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                        if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                        if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                        if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                        if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                        if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                        if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                        if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }

                    }
                }

            }
            else
            {
                MessageBox.Show("Invalid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        void empty()
        { }
        private void News_Click(object sender, EventArgs e)
        {
            empty();
        }

        private void WeighBridge_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();          
            comboport.Text = ports[0];
        }

    

        private void Butstop_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
                progressBar1.Value = 0;
            }
        }
        private void Butgetdata_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                ReadSerial();
                progressBar1.Value = 0;
            }
        }

      
        private void ReadSerial()
        {
            try
            {
                readserialvalue = serialPort1.ReadExisting();

                if (readserialvalue == "")
                {
                    txtname.Text = "0.000";
                    textBox1.Text = "0.000";
                    textBox2.Text = "0.000";
                }
                else
                {
                    string[] words = readserialvalue.Split(' ');
                    txtname.Text = Convert.ToDecimal(words[3]).ToString("0.000");
                    textBox1.Text = Convert.ToString(words[4]).ToString();
                    textBox2.Text = Convert.ToString(words[5]).ToString();
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(@"Read Serial Thread" + ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void Butclear_Click(object sender, EventArgs e)
        {
            txtname.Text = "";
            textBox1.Text = ""; textBox2.Text = "";

        }
            private void Butstart_Click(object sender, EventArgs e)
        {
            try
            {
                serialPort1.Close(); progressBar1.Value = 0;
                serialPort1.PortName = comboport.Text;
                serialPort1.BaudRate = Convert.ToInt32("0" + combobaudrate.Text);
                serialPort1.DataBits = Convert.ToInt32("0" + combodatabits.Text);
                serialPort1.StopBits = (StopBits)Enum.Parse(typeof(StopBits), combostopbits.Text);
                serialPort1.Parity = (Parity)Enum.Parse(typeof(Parity), comboparity.Text);

                serialPort1.Open();

                progressBar1.Value = 100;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

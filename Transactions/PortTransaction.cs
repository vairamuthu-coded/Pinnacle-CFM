using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
namespace Pinnacle.Transactions
{
    public partial class PortTransaction : Form,ToolStripAccess
    {
        public PortTransaction()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
           
        }
    
        public static PortTransaction Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PortTransaction();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        private static PortTransaction _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        public void usercheck(string s, string ss, string sss)
        {

          

        }
        private void News_Click(object sender, EventArgs e)
        {

        }

        private void butgetdata1_Click(object sender, EventArgs e)
        {

        }

        private void butclear_Click(object sender, EventArgs e)
        {

        }

        private void butstop_Click(object sender, EventArgs e)
        {

        }

        private void butstart_Click(object sender, EventArgs e)
        {

        }

        private void butgetdata_Click(object sender, EventArgs e)
        {
            try
            {
                this.serialPort1.Close();
                progressBar1.Value = 0;
                this.serialPort1.PortName = comboport.Text;
                this.serialPort1.BaudRate = Convert.ToInt32("0" + combobaudrate.Text);
                this.serialPort1.DataBits = Convert.ToInt32("0" + combodatabits.Text);
                this.serialPort1.StopBits = (StopBits)Enum.Parse(typeof(StopBits), combostopbits.Text);
                this.serialPort1.Parity = (Parity)Enum.Parse(typeof(Parity), comboparity.Text);

                Models.Serial.PortName = comboport.Text;
                Models.Serial.BaudRate = Convert.ToInt32("0" + combobaudrate.Text);
                Models.Serial.DataBits = Convert.ToInt32("0" + combodatabits.Text);
                Models.Serial.StopBits = combostopbits.Text;
                Models.Serial.Parity = comboparity.Text;
                serialPort1.Open();
                serialPort1.DiscardInBuffer();
                serialPort1.DiscardOutBuffer();
                string readserialvalue1 = "";
                lblcon.Text = "";
                readserialvalue1 = serialPort1.ReadExisting();
                textBox1.Text = readserialvalue1.ToString();        //     50 kg    G 000000 A;    -20 kg    G 000000 A;,     00 kg    G 000000 A;  35770 kg    G 000000 A,   4830 kg    G 000000 A;


                if (serialPort1.IsOpen)
                {

                    lblcon.Text = "SerialPort Connected";
                    lblcon.ForeColor = Color.DarkGreen;
                    string[] words1 = readserialvalue1.Split(' ');
                    if (words1.Length >= 1)
                    {
                        Int32 ss = 0; Field0.Text = ""; textBox1.Text = "";

                        if (Convert.ToInt64("0" + words1[2].ToString()) >= 1)
                        {
                            ss = 0; Field0.Text = "";
                            ss = Convert.ToInt32("0" + words1[2].ToString());
                            Field0.Text = ss.ToString();

                            //return;
                        }
                        if (Convert.ToInt64("0" + words1[3].ToString()) >= 1)
                        {
                            ss = 0; Field0.Text = "";
                            ss = Convert.ToInt32("0" + words1[3].ToString());
                            Field0.Text = ss.ToString();
                            // return;
                        }
                        if (Convert.ToInt64("0" + words1[4].ToString()) >= 1)
                        {
                            ss = 0; Field0.Text = "";
                            ss = Convert.ToInt32("0" + words1[4].ToString());
                            Field0.Text = ss.ToString();
                            // return;
                        }
                        if (Convert.ToInt64("0" + words1[5].ToString()) >= 1 || Convert.ToInt64("0" + words1[5].ToString()) == 00)
                        {
                            ss = 0; Field0.Text = "";
                            ss = Convert.ToInt32("0" + words1[5].ToString());
                            Field0.Text = ss.ToString();
                            //return;
                        }
                        if (Convert.ToInt64("0" + words1[6].ToString()) >= 1 || Convert.ToInt64("0" + words1[6].ToString()) == 00)
                        {
                            ss = 0; Field0.Text = "";
                            ss = Convert.ToInt32("0" + words1[6].ToString());
                            Field0.Text = ss.ToString();
                            //return;
                        }


                        if (Convert.ToInt64("0" + words1[7].ToString()) >= 1 || Convert.ToInt64("0" + words1[7].ToString()) == 00)
                        {
                            ss = 0; Field0.Text = "";
                            ss = Convert.ToInt32("0" + words1[7].ToString());
                            Field0.Text = ss.ToString();

                            //MessageBox.Show("7-----" + words1[7].ToString());

                        }
                        if (Convert.ToInt64("0" + words1[8].ToString()) >= 1 || Convert.ToInt64("0" + words1[8].ToString()) == 00)
                        {
                            ss = 0; Field0.Text = "";
                            ss = Convert.ToInt32("0" + words1[8].ToString());
                            Field0.Text = ss.ToString();

                            //MessageBox.Show("8-----" + words1[8].ToString());

                        }
                        //if (Convert.ToInt64("0" + words1[9].ToString()) >= 1 || Convert.ToInt64("0" + words1[9].ToString()) == 00)
                        //{
                        //    ss = 0; Field0.Text = "";
                        //    ss = Convert.ToInt32("0" + words1[9].ToString());
                        //    Field0.Text = ss.ToString();

                        //}
                        //if (Convert.ToInt64("0" + words1[10].ToString()) >= 1 || Convert.ToInt64("0" + words1[10].ToString()) == 00)
                        //{
                        //    ss = 0; Field0.Text = "";
                        //    ss = Convert.ToInt32("0" + words1[10].ToString().Trim());
                        //    Field0.Text = ss.ToString();


                        //}
                        //if (Convert.ToInt64("0" + words1[11].ToString()) >= 1 || Convert.ToInt64("0" + words1[11].ToString()) == 00)
                        //{
                        //    ss = 0; Field0.Text = "";
                        //    ss = Convert.ToInt32("0" + words1[11].ToString().Trim());
                        //    Field0.Text = ss.ToString();

                        //}
                        //if (Convert.ToInt64("0" + words1[12].ToString()) >= 1 || Convert.ToInt64("0" + words1[12].ToString()) == 00)
                        //{
                        //    ss = 0; Field0.Text = "";
                        //    ss = Convert.ToInt32("0" + words1[12].ToString().Trim());
                        //    Field0.Text = ss.ToString();
                        //}
                        //if (Convert.ToInt64("0" + words1[13].ToString()) >= 1 || Convert.ToInt64("0" + words1[13].ToString()) == 00)
                        //{
                        //    ss = 0; Field0.Text = "";
                        //    ss = Convert.ToInt32("0" + words1[13].ToString().Trim());
                        //    Field0.Text = ss.ToString();
                        //}
                        //if (Convert.ToInt64("0" + words1[14].ToString()) >= 1 || Convert.ToInt64("0" + words1[14].ToString()) == 00)
                        //{
                        //    ss = 0; Field0.Text = "";
                        //    ss = Convert.ToInt32("0" + words1[14].ToString().Trim());
                        //    Field0.Text = ss.ToString();
                        //}
                        //if (Convert.ToInt64("0" + words1[15].ToString()) >= 1 || Convert.ToInt64("0" + words1[15].ToString()) == 00)
                        //{
                        //    ss = 0; Field0.Text = "";
                        //    ss = Convert.ToInt32("0" + words1[15].ToString().Trim());
                        //    Field0.Text = ss.ToString();
                        //}
                    }
                    else
                    {
                        MessageBox.Show("Disconnted" + readserialvalue1.ToString());
                    }
                }
                else
                {
                    lblcon.Text = "SerialPort Closed";
                    progressBar1.Value = 0;
                    lblcon.ForeColor = Color.DarkRed;
                    serialPort1.Close();
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Error: " + ex.ToString(), " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            serialPort1.Close();
        }

        private void PortTransaction_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            comboport.Items.AddRange(ports);
            comboport.Text = "COM1";
        }

       void empty()
        {
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
        }

        public void News()
        {
            
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

        public void Searchs(int EditID)
        {
           
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        public void GridLoad()
        {
           
        }
    }
}

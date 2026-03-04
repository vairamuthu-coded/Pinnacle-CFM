using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO.Ports;
using System.Net;
using System.Net.Sockets;
using System.Text;
using SuperSimpleTcp;
using System.Text.RegularExpressions;

namespace Pinnacle.Master.CFM
{
    public partial class PortMaster : Form,ToolStripAccess
    {
        private static PortMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        SimpleTcpServer server; SimpleTcpClient client;
 

        public PortMaster()
        {
            InitializeComponent();
           
            Class.Users.bisconnected = false;
        }
        public static PortMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PortMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;

            }
        }
       
       public void News()
        {

            empty(); 
        }
        private void empty()
        {
            txtportid.Text = "";
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            Class.Users.UserTime = 0;
            textBox1.Text = "";s = "";words1 = null;
            comboport.Select(); readserialvalue1 = "";
            string[] ports = SerialPort.GetPortNames();
            comboport.Items.Add("TCP/IP");
            comboport.Text = "TCP/IP";
            comboport.Items.AddRange(ports);
            try
            {
                if (client.IsConnected == false)
                {
                    client.Connect();
                }

                client.Send("0");
            }
            catch(Exception ex) { }

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

        string readserialvalue1 = "";
        string[] words1; string[] words2;
        private void butgetdata_Click(object sender, EventArgs e)
        {
            try
            {
                Field0.Clear();
                lblcon.Text = "";
                progressBar1.Value = 0;

                if (comboport.Text == "TCP/IP")
                {
                    lblcon.ForeColor = Color.Red;
                    lblcon.Text = "TCP/IP Disconnected";
                    Class.Users.bisconnected = false;
                    return;
                }

                // ---------- Serial Settings ----------
                Models.Serial.PortType = comboport.Text;
                Models.Serial.PortName = comboport.Text;

                serialPort3.PortName = comboport.Text;
                serialPort3.BaudRate = int.TryParse(combobaudrate.Text, out var br) ? br : 0;
                serialPort3.DataBits = int.TryParse(combodatabits.Text, out var db) ? db : 0;
                serialPort3.StopBits = (StopBits)Enum.Parse(typeof(StopBits), combostopbits.Text);
                serialPort3.Parity = (Parity)Enum.Parse(typeof(Parity), comboparity.Text);

                Models.Serial.BaudRate = serialPort3.BaudRate;
                Models.Serial.DataBits = serialPort3.DataBits;
                Models.Serial.StopBits = combostopbits.Text;
                Models.Serial.Parity = comboparity.Text;

                // ---------- Open Port ----------
                if (!serialPort3.IsOpen)
                    serialPort3.Open();

                serialPort3.DiscardInBuffer();
                serialPort3.DiscardOutBuffer();

                readserialvalue1 = serialPort3.ReadLine().Trim();
                textBox1.Text = "COM : " + readserialvalue1;

                lblcon.ForeColor = Color.DarkGreen;
                lblcon.Text = "SerialPort Connected";
                Class.Users.bisconnected = true;

                words1 = readserialvalue1.Split(' ');

                // ---------- Extract Value Safely ----------
                Field0.Clear();

                for (int i = 2; i < words1.Length && i <= 8; i++)
                {
                    if (long.TryParse(words1[i], out long value))
                    {
                        Field0.Text = value.ToString();
                        break; // stop at first valid value
                    }
                }
            }
            catch (Exception ex)
            {
                // log if needed
                // MessageBox.Show(ex.Message);
            }
            finally
            {
                if (serialPort3.IsOpen)
                    serialPort3.Close();
            }

          
        }
        private void PortMaster_Load(object sender, EventArgs e)
        {
            
            string[] ports = SerialPort.GetPortNames();
            comboport.Items.Add("TCP/IP");
            comboport.Text = "TCP/IP";
            comboport.Items.AddRange(ports);
            

            txtip.Text = Class.Users.PortIP+":"+ Class.Users.PortNo;
            txtportno.Text = Class.Users.PortNo;
            server = new SimpleTcpServer(txtip.Text);
            server.Events.ClientConnected += Events_ClientConnected;
            server.Events.ClientDisconnected += Events_ClientDisconnected;
            server.Events.DataReceived += Events_DataReceived;
            client = new SimpleTcpClient(txtip.Text);
            client.Events.Connected += Events_Connected;
            client.Events.Disconnected += Events_Disconnected;
            client.Events.DataReceived += Events_DataReceived;
          
            txtMessage.Text = "0";
            txtMessageclient.Text = "0";
            if (server.IsListening)
            {
                server.Start();
                MessageBox.Show("server start");
            }

            btnsendclient_Click(sender, e);
        }
        string s = "";
        private void Events_DataReceived(object sender, DataReceivedEventArgs e)
        {
            s = "";
            this.Invoke((MethodInvoker)delegate
           {
               s += Encoding.UTF8.GetString(e.Data.Array).Trim();
           });


            // client.Disconnect();            
            // MessageBox.Show("Events_DataReceived1     Cleint Stoped" + words1.ToString());
        }

        private void Events_Disconnected(object sender, ConnectionEventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            {
                lblcon.Refresh();
                lblcon.ForeColor = System.Drawing.Color.Red;
                lblcon.Text = "Sever DisConnected ";

            });
        }

        private void Events_Connected(object sender, ConnectionEventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            {
                lblcon.Refresh(); lblcon.ForeColor = System.Drawing.Color.Green;
                lblcon.Text = "Sever Connected ";
            });
        }


        //private void Events_DataReceived(object sender, DataReceivedEventArgs e)
        //{
        //    s = "";
        //    this.Invoke((MethodInvoker)delegate
        //    {

        //        s += Encoding.UTF8.GetString(e.Data.Array).Trim();
        //    });

        //    // client.Disconnect();

        //}

        private void Events_ClientDisconnected(object sender, ConnectionEventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
           {
               lblcon.Refresh();
               lblcon.ForeColor = System.Drawing.Color.Red;
               lblcon.Text = " TCP DisConnected ";

           });
        }

        private void Events_ClientConnected(object sender, ConnectionEventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            {
                lblcon.Refresh();
                lblcon.ForeColor = System.Drawing.Color.Green;
                lblcon.Text = "TCP Connected";
            });

        }


        private void btnsend_Click(object sender, EventArgs e)
        {
            if (server.IsListening)
            {
                server.Start();
                server.Send(txtip.Text, txtMessage.Text);
            }
            if (!string.IsNullOrEmpty(txtMessage.Text))
            {
                client.Connect();
                client.Send(txtMessage.Text);
            }


        }
      
        private void btnsendclient_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (!client.IsConnected)
                    client.Connect();

                client.Send("0");

                textBox1.Text = s;

                string cleaned = Regex.Replace(s, @"Gross", "", RegexOptions.IgnoreCase).Trim();
                string[] words = cleaned.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                lblcon.Text = "TCP/IP Connected";
                lblcon.ForeColor = Color.DarkGreen;
                Field0.Text = string.Empty;
                int value;
                // Try first valid numeric value from index 0 to 8
                for (int i = 0; i < words.Length && i <= 8; i++)
                {
                    if (int.TryParse(words[i], out value) && value >= 0)
                    {
                        Field0.Text = value.ToString();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                lblcon.ForeColor = Color.Red;
                lblcon.Text = "TCP/IP Disconnected";

                // Log or show error (choose one)
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Class.Users.bisconnected = false;
            }

        }
        private void Exit_Click(object sender, EventArgs e)
        {
            
        }

  

        //private void Saves_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (comboport.Text == "")
        //        {
        //            MessageBox.Show("Product Name is empty " + " Alert " + comboport.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //            return;
        //        }
        //        if (comboport.Text != "")
        //        {
                   


        //            string chk = "";
        //            if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
        //            string sel = "select asptblportmasid    from  asptblportmas    WHERE comport='" + comboport.Text + "' and active='" + chk + "' ";
        //            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblportmas");
        //            DataTable dt = ds.Tables["asptblportmas"];
        //            if (dt.Rows.Count != 0)
        //            {
        //                MessageBox.Show("Child Record Found " + " Alert " + comboport.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
        //            }
        //            else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtportid.Text) == 0 || Convert.ToInt32("0" + txtportid.Text) == 0)
        //            {
        //                string ins = "insert into asptblportmas(comport,active,createdby,modifiedby,ipaddress)  VALUES('" + comboport.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
        //                Utility.ExecuteNonQuery(ins);
        //                MessageBox.Show("Record Saved Successfully " + comboport.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //               empty();
        //            }
        //            else
        //            {
        //                string up = "update  asptblportmas  set   comport='" + comboport.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblportmasid='" + txtportid.Text + "';";
        //                Utility.ExecuteNonQuery(up);
        //                MessageBox.Show("Record Updated Successfully " + comboport.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
        //                empty();
        //            }


        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        MessageBox.Show("comport " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //    }
        //}
        void disable()
        {
            combodatabits.Enabled = true;
            comboparity.Enabled = true;
            combobaudrate.Enabled = true;
            combostopbits.Enabled = true; txtip.Enabled = false;txtportno.Enabled = false;
            txtip.Text = "";txtportno.Text = "";
        }
        void enable()
        {
            combodatabits.Enabled = false;
            comboparity.Enabled = false;
            combobaudrate.Enabled = false;
            combostopbits.Enabled = false; txtip.Enabled = true; txtportno.Enabled = true;
            txtip.Text = Class.Users.PortIP+":"+Class.Users.PortNo;
           



        }
        private void comboport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboport.Text == "TCP/IP")
            {
             
              
                enable();
               
            }
            else
            {
                disable(); 
            }
        }

        private void checkactive_CheckedChanged(object sender, EventArgs e)
        {
            if (checkactive.Checked == true)
            {

                Form1 fm = new Form1();
                fm.Show();
            }
        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
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
            throw new NotImplementedException();
        }

        public void Pdfs()
        {
            throw new NotImplementedException();
        }

        public void ChangePasswords()
        {
            throw new NotImplementedException();
        }

        public void DownLoads()
        {
            throw new NotImplementedException();
        }

        public void ChangeSkins()
        {
            throw new NotImplementedException();
        }

        public void Logins()
        {
            throw new NotImplementedException();
        }

        public void GlobalSearchs()
        {
            throw new NotImplementedException();
        }

        public void TreeButtons()
        {
            throw new NotImplementedException();
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        public void GridLoad()
        {
           
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboport.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            comboport.Items.Add("TCP/IP");
            comboport.Text = "TCP/IP";
            comboport.Items.AddRange(ports);
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Net;
using Pinnacle.Models;
using System.Runtime;
using System.Runtime.InteropServices;
namespace Pinnacle
{
    public partial class IPDisconnnect : Form
    {
        public IPDisconnnect()
        {
            InitializeComponent();
        }

        private void IPDisconnnect_Load(object sender, EventArgs e)
        {
            IPHostEntry iph;
            string myip = "";
            iph = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in iph.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    myip = ip.ToString();
                }
            }
            button1.Text = myip;
            //MessageBox.Show(myip.ToString());
        }
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);

        private void button1_Click(object sender, EventArgs e)
        {
            int Desc;
                      MessageBox.Show(InternetGetConnectedState(out Desc, 0).ToString());
        }
    }

}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class DashBoard : Form
    {
        private static DashBoard _instance;
        public static DashBoard Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DashBoard();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public DashBoard()
        {
            InitializeComponent();
        }

        private void DashBoard_Load(object sender, EventArgs e)
        {

        }
    }
}

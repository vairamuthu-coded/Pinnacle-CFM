using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.UserControls
{
    public partial class MonthControl : UserControl
    {
        public MonthControl()
        {
            InitializeComponent();
        }
        public Label finyear
        {
            get { return lbluserfinyear; }
            set { lbluserfinyear = value; }
        }
        public Label compcode
        {
            get { return lblusercompcode; }
            set { lblusercompcode = value; }
        }
        public Button month
        {
            get { return lblusermonth; }
            set { lblusermonth = value; }
        }
       
    }
}

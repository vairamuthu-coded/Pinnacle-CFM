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
    public partial class PaySlipPrint : UserControl
    {
        public PaySlipPrint()
        {
            InitializeComponent();
        }
        public UserControl PaySlipPrint1
        {          
            set {  }
        }
        public Button MidCardButton
        {
            get { return button1; }
            set { button1 = value; }
        }
        public Panel PanelPaySlip
        {
            get { return userpayslippanel; }
            set { userpayslippanel = value; }
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
        public Label month
        {
            get { return lblusermonth; }
            set { lblusermonth = value; }
        }
        public Label EmpName
        {
            get { return lbluserempname; }
            set { lbluserempname = value; }
        }
        public Label MidCard
        {
            get { return lblusermidcard; }
            set { lblusermidcard = value; }
        }
    }
}

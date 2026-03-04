using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.UserControls.Hospital
{
    public partial class FindDoctor : UserControl
    {
        public FindDoctor()
        {
            InitializeComponent();
        }
        public Button doctorPhoto
        {
            get { return butdoctor; }
            set { butdoctor=value; }
        }

        public Panel Panel1
        {
            get { return panel1; }
            set { panel1 = value; }
        }
        public Panel Panel2
        {
            get { return panel2; }
            set { panel2 = value; }
        }
        public Button Paneldoctor
        {
            get { return panelimage; }
            set { panelimage = value; }
        }
        public TextBox DoctorName
        {
            get { return txtdoctor; }
            set { txtdoctor = value; }
        }
        
        public TextBox DoctorEducation
        {
            get { return txteducation; }
            set { txteducation = value; }
        }

        public TextBox DoctorDept
        {
            get { return txtdept; }
            set { txtdept = value; }
        }
        public Label DoctorID
        {
            get { return txtdoctorid; }
            set { txtdoctorid = value; }
        }

        private void txtdoctorid_KeyUp(object sender, KeyEventArgs e)
        {
           
        }

        private void txtdoctorid_MouseHover(object sender, EventArgs e)
        {
            this.txtdoctorid.BackColor = System.Drawing.Color.WhiteSmoke;
        }
    }
}

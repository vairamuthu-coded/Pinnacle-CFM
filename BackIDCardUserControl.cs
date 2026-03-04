using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle
{
    public partial class BackIDCardUserControl : UserControl
    {
        public BackIDCardUserControl()
        {
            InitializeComponent();
        }
        public TextBox fathername
        {
            get { return userfather; }
            set { userfather = value; }
        }
        public TextBox dob
        {
            get { return userdob; }
            set { userdob = value; }
        }
       

        public TextBox permanentaddress
        {
            get { return userpermanentaddress; }
            set { userpermanentaddress = value; }
        }
        public TextBox bloodgroup
        {
            get { return userbloodgroup; }
            set { userbloodgroup = value; }
        }
        public TextBox natureofemployeement
        {
            get { return usernatureofemployeeemnt; }
            set { usernatureofemployeeemnt = value; }
        }
     

      
        public TextBox dateofissue
        {
            get { return userdateofissue; }
            set { userdateofissue = value; }
        }
      
        public Button CloseButton
        {
            get { return userbutclose; }
            set { userbutclose = value; }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}

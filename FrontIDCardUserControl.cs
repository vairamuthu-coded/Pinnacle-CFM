using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle
{
    public partial class FrontIDCardUserControl : UserControl
    {
        public FrontIDCardUserControl()
        {
            InitializeComponent();
        }
        public Label heading
        {
            get { return userlblheading; }
            set { userlblheading = value; }
        }
        public Label underline1
        {
            get { return userlblunderline1; }
            set { userlblunderline1 = value; }
        }
        public TextBox compname
        {
            get { return userlblcompcode; }
            set { userlblcompcode = value; }
        }

        public TextBox address
        {
            get { return userlbladdress; }
            set { userlbladdress = value; }
        }
        public Label underline2
        {
            get { return userlblunderline2; }
            set { userlblunderline2 = value; }
        }
        public Label registration
        {
            get { return usrlblregisteration; }
            set { usrlblregisteration = value; }
        }

        public Image empimage
        {
            get { return pictureBox1.BackgroundImage; }
            set { pictureBox1.BackgroundImage = value; }
        }
        public TextBox empname
        {
            get { return userlblname; }
            set { userlblname = value; }
        }
        public TextBox midcard
        {
            get { return userlblmidcard; }
            set { userlblmidcard = value; }
        }
        public Button CloseButton
        {
            get { return userbutclose; }
            set { userbutclose = value; }
        }
        public Button CloseButton1
        {
            get { return userbutclose1; }
            set { userbutclose1 = value; }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle
{
    public partial class Form1 : Form
    {


        public static bool Data_Validation(Control RootCtrl, ErrorProvider ErrProvider)
        {
            foreach (Control Ctrl in RootCtrl.Controls)
            {
                if (ErrProvider.GetError(Ctrl) != "")
                {
                    return false;
                }
            }
            return true;
        }
       
        public Form1()
        {
            InitializeComponent();
            DataTable dt = new DataTable("Table1");         
            dt.Columns.Add("FirstName");
            dt.Columns.Add("LastName");
            dt.Columns.Add("occupation");
            dt.Columns.Add("Place");
            DataSet ds = new DataSet();
            ds.Tables.Add(dt); comboBox1.Items.Clear(); comboBox2.Items.Clear(); comboBox3.Items.Clear();
            dt.Rows.Add(new string[] { "John", "Tina", "Doctor", "Italy" });
            dt.Rows.Add(new string[] { "Mary", "anu", "Teacher", "America" });
            dt.Rows.Add(new string[] { "asha", "roy", "Staff", "London" });
            dt.Rows.Add(new string[] { "George", "Gaskin", "Nurse", "germany" });
            dt.Rows.Add(new string[] { "sam", "jens", "Engineer", "Russia" });
            dt.Rows.Add(new string[] { "Ben", "Geo", "Developer", "India" });
            DataView view = new DataView(dt);
            this.buttonEditExt1.DataSource = view;
            this.buttonEditExt1.DisplayMember = "FirstName";           
            this.comboBox1.DisplayMember = "LastName";
            this.comboBox1.DataSource = view;
            this.comboBox2.DisplayMember = "occupation";
            this.comboBox2.DataSource = view;
            this.comboBox3.DisplayMember = "place";
            this.comboBox3.DataSource = view;
            List<States> STATE = new List<States>();
            STATE.Add(new States() { ID = "", Name = "" });
            STATE.Add(new States() {ID="1",Name="1111"});

            this.comboBox4.ValueMember = "ID";
            this.comboBox4.DisplayMember = "Name";
            this.comboBox4.DataSource = STATE;
        }
        public class States
        {
            public string ID { get; set; }
            public string Name { get; set; }
        }
        public bool ValidEmailAddress(string emailAddress, out string errorMessage)
        {
            // Confirm that the email address string is not empty.
            if (emailAddress.Length == 0)
            {
                errorMessage = "email address is required.";
                return false;
            }

            // Confirm that there is an "@" and a "." in the email address, and in the correct order.
            if (emailAddress.IndexOf("@") > -1)
            {
                if (emailAddress.IndexOf(".", emailAddress.IndexOf("@")) > emailAddress.IndexOf("@"))
                {
                    errorMessage = "";
                    return true;
                }
            }

            errorMessage = "email address must be valid email address format.\n" +
               "For example 'someone@example.com' ";
            return false;
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            if (ValidateChildren(ValidationConstraints.Enabled))
            {
                MessageBox.Show("success");
            }
        }

        private void textBox2_Validating(object sender, CancelEventArgs e)
        {
             
            if (string.IsNullOrEmpty(textBox2.Text))
            {
                e.Cancel = true;
                textBox2.Focus();
                ErrProvider.SetError(textBox2, "error");

            }
            else
            {
                e.Cancel = false;
                ErrProvider.SetError(textBox2, null);
            }
        }
    }
    
}

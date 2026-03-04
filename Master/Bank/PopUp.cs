using Microsoft.Office.Interop.Word;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Pinnacle.Master.Bank
{
    public partial class PopUp : Form
    {
        public PopUp()
        {
            InitializeComponent();
            this.Font = Class.Users.FontName;
            button1.ForeColor = Class.Users.Color1;
            button1.BackColor = Class.Users.BackColors;
            button1.Text = Class.Users.TableName;

            string filepath = "";


            filepath = System.IO.Directory.GetCurrentDirectory();
            if (Class.Users.Extension == "xls" || Class.Users.Extension == "xlsx")
            {
                filepath += "\\temp1." + Class.Users.Extension;
                FileStream FS = new FileStream(filepath, System.IO.FileMode.Create);
                FS.Write(Class.Users.StaticByts, 0, Class.Users.StaticByts.Length);
                FS.Close();
                DataTable dt = ReadExcel(filepath, Class.Users.Extension);
                dataGridView1.DataSource = null;
                dataGridView1.DataSource = dt;
                CommonFunctions.SetRowNumber(dataGridView1);
                tabControl1.SelectTab(tabPage3);
            }
            if (Class.Users.Extension == "pdf")
            {
                filepath += "\\temp1." + Class.Users.Extension;
                FileStream FS = new FileStream(filepath, System.IO.FileMode.Create);
                FS.Write(Class.Users.StaticByts, 0, Class.Users.StaticByts.Length);
                FS.Close();
                axAcroPDF1.src = null;
                axAcroPDF1.src = filepath;
                tabControl1.SelectTab(tabPage1);
            }
            if (Class.Users.Extension == "jpeg" || Class.Users.Extension == "jpg" || Class.Users.Extension == "png" || Class.Users.Extension == "JPG")
            {
                pictureBox1.Image = null;
                Image img = Models.Device.ByteArrayToImage(Class.Users.StaticByts);
                pictureBox1.Image = img;
                tabControl1.SelectTab(tabPage2);
            }
            
        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo("xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES';"; //for above excel 2007  
            using (System.Data.OleDb.OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    System.Data.OleDb.OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }

        private void PopUp_FormClosing(object sender, FormClosingEventArgs e)
        {
            dataGridView1.DataSource = null;
             pictureBox1.Image = null;
           // axAcroPDF1.src = "D:\\temp2.pdf";
            this.Hide();

        }

        private void PopUp_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }
    }
}

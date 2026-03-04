using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master.Bank
{
    public partial class ReportPopUp : Form
    {
        public ReportPopUp()
        {
            InitializeComponent();
            this.Font = Class.Users.FontName;
            button1.ForeColor = Class.Users.Color1;
            button1.BackColor = Class.Users.BackColors;          
            //button1.Text = Class.Users.TableName;

            DataTable dt1 = Utility.SQLQuery("SELECT  a.asptbladvpaydetid,a.asptbladvpaymasid,a.compcode,b.department,c.partyname,a.invoicetype,a.invoice,a.INVBLOB,a.INVPROBLOB,a.QUABLOB,a.powoblob,a.OTHBLOB from asptbladvpaydet a join asptbldeptmas b on b.asptbldeptmasid=a.department join asptblpartymas c on c.asptblpartymasid=a.partyname   where a.asptbladvpaymasid='" + Class.Users.Paramid + "'");
            dataGridView1.Rows.Clear();
            int i = 0; Class.Users.Paramid = 0;
            foreach (DataRow myRow in dt1.Rows)
            {
               
                
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = myRow["asptbladvpaydetid"].ToString();
                dataGridView1.Rows[i].Cells[1].Value = myRow["invoicetype"].ToString();
                dataGridView1.Rows[i].Cells[2].Value = myRow["invoice"].ToString();
                string[] ext = myRow["invoice"].ToString().Split('.');
                if (i == 0)
                {
                    stdbytes = (byte[])myRow["INVBLOB"];

                    //filepath1 = "D:\\temp1."+ext[1].ToString();
                    // FS = new FileStream(filepath1, System.IO.FileMode.Create);
                    // FS.Write(stdbytes, 0, stdbytes.Length);
                    // FS.Close();
                }
                if (i == 1)
                {
                    stdbytes2 = (byte[])myRow["INVPROBLOB"];
                    //filepath1 = "";
                    //filepath1 = "D:\\temp1." + ext[1].ToString();
                    //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                    //FS.Write(stdbytes2, 0, stdbytes2.Length);
                    //FS.Close();
                }
                if (i == 2)
                {
                    stdbytes3 = (byte[])myRow["QUABLOB"];

                    //filepath1 = "D:\\temp1." + ext[1].ToString();
                    //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                    //FS.Write(stdbytes3, 0, stdbytes3.Length);
                    //FS.Close();
                }
                if (i == 3)
                {
                    stdbytes4 = (byte[])myRow["powoblob"];
                    //filepath1 = "";
                    //filepath1 = "D:\\temp1." + ext[1].ToString();
                    //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                    //FS.Write(stdbytes4, 0, stdbytes4.Length);
                    //FS.Close();
                }
                if (i == 4)
                {
                    stdbytes5 = (byte[])myRow["OTHBLOB"];
                    //filepath1 = "";
                    //filepath1 = "D:\\temp1." + ext[1].ToString();
                    //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                    //FS.Write(stdbytes5, 0, stdbytes5.Length);
                    //FS.Close();
                }
                ////Class.Users.Extension = extension;
                //if (i == 0)
                //{
                //    stdbytes = (byte[])myRow["INVBLOB"];
                //    //string filepath1 = "";
                //    //filepath1 = "D:\\temp1.pdf";
                //    //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                //    //FS.Write(stdbytes, 0, stdbytes.Length);
                //    //FS.Close();
                //}
                //if (i == 1)
                //{
                //    stdbytes2 = (byte[])myRow["INVPROBLOB"];
                //    //string filepath2 = "";
                //    //filepath2 = "D:\\temp2.pdf";
                //    //FS = new FileStream(filepath2, System.IO.FileMode.Create);
                //    //FS.Write(stdbytes2, 0, stdbytes2.Length);
                //    //FS.Close();
                //}
                //if (i == 2)
                //{
                //    stdbytes3 = (byte[])myRow["QUABLOB"];
                //    //string filepath3 = "";
                //    //filepath3 = "D:\\temp3.pdf";
                //    //FS = new FileStream(filepath3, System.IO.FileMode.Create);
                //    //FS.Write(stdbytes3, 0, stdbytes3.Length);
                //    //FS.Close();
                //}
                //if (i == 3)
                //{
                //    stdbytes4 = (byte[])myRow["powoblob"];
                //    //string filepath4 = "";
                //    //filepath4 = "D:\\temp4.pdf";
                //    //FS = new FileStream(filepath4, System.IO.FileMode.Create);
                //    //FS.Write(stdbytes4, 0, stdbytes4.Length);
                //    //FS.Close();
                //}
                //if (i == 4)
                //{
                //    stdbytes5 = (byte[])myRow["OTHBLOB"];
                //    //string filepath5 = "";
                //    //filepath5 = "D:\\temp4.pdf";
                //    //FS = new FileStream(filepath5, System.IO.FileMode.Create);
                //    //FS.Write(stdbytes5, 0, stdbytes5.Length);
                //    //FS.Close();
                //}
                i++;

            }
            CommonFunctions.SetRowNumber(dataGridView1);
        }

        private void PopUp_FormClosing(object sender, FormClosingEventArgs e)
        {
         
            
            this.Dispose();
        }
        string filepath = "D:\\temp.pdf"; FileStream FS = null;
        byte[] stdbytes; byte[] stdbytes2; byte[] stdbytes3; byte[] stdbytes4; byte[] stdbytes5;
        public void checkcellvalue(int index, DataGridView dgrid)
        {
           
            if (dgrid.Rows[index].Cells[2].FormattedValue != "" && dgrid.CurrentCell.ColumnIndex.Equals(3) && index != -1)
            {
                if (dgrid.CurrentCell != null)
                {
                    string[] ext = dgrid.CurrentRow.Cells[2].FormattedValue.ToString().Split('.');
                    Class.Users.Extension = ext[1].ToString();
                    if (index == 0)
                        {
                        Class.Users.StaticByts = null;
                        //Class.Users.StaticByts = null;
                        //Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                        //Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                        ////FS.Write(stdbytes, 0, stdbytes.Length);
                        ////FS.Close();
                        //Class.Users.StaticByts = stdbytes;
                        //Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                        //Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                        //Master.Bank.PopUp pop = new Master.Bank.PopUp();
                        //pop.Show();

                        //filepath = "D:\\temp1.pdf";
                        //    FS = new FileStream(filepath, System.IO.FileMode.Create);
                        //    FS.Write(stdbytes, 0, stdbytes.Length);
                        //    FS.Close();
                        Class.Users.StaticByts = stdbytes;
                        Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();

                        }
                        if (index == 1)
                        {
                        Class.Users.StaticByts = null;
                        //filepath = "";
                        //filepath = "D:\\temp1.pdf";
                        //FS = new FileStream(filepath, System.IO.FileMode.Create);
                        //FS.Write(stdbytes2, 0, stdbytes2.Length);
                        //FS.Close();
                        Class.Users.StaticByts = stdbytes2;                          
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();
                        }
                    if (index == 2)
                    {
                        Class.Users.StaticByts = null;
                        //    filepath = "";
                        //    filepath = "D:\\temp1.pdf";
                        //    Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                        //    FS = new FileStream(filepath, System.IO.FileMode.Create);
                        //    FS.Write(stdbytes3, 0, stdbytes3.Length);
                        //    FS.Close();
                        Class.Users.StaticByts = stdbytes3;                      
                        Master.Bank.PopUp pop = new Master.Bank.PopUp();
                        pop.Show();

                    }
                    if (index == 3)
                    {
                        Class.Users.StaticByts = null;
                        //filepath = "";
                        //filepath = "D:\\temp1.pdf";
                        //Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                        //FS = new FileStream(filepath, System.IO.FileMode.Create);
                        //FS.Write(stdbytes4, 0, stdbytes4.Length);
                        //FS.Close();
                        Class.Users.StaticByts = stdbytes4;                        
                        Master.Bank.PopUp pop = new Master.Bank.PopUp();
                        pop.Show();
                    }
                    if (index == 4)
                    {
                        Class.Users.StaticByts = null;
                        //filepath = "";
                        //filepath = "D:\\temp1.pdf";
                        //Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                        //FS = new FileStream(filepath, System.IO.FileMode.Create);
                        //FS.Write(stdbytes5, 0, stdbytes5.Length);
                        //FS.Close();
                        Class.Users.StaticByts = stdbytes5;
                        Master.Bank.PopUp pop = new Master.Bank.PopUp();
                        pop.Show();
                    }
                }
            }
            else
            {
                //MessageBox.Show("No Attachment  : Index of "+ index.ToString());
                //this.Dispose();
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            checkcellvalue(e.RowIndex, dataGridView1);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Class.Users.UserTime = 0;
        }

        private void ReportPopUp_Load(object sender, System.EventArgs e)
        {

        }

        private void ReportPopUp_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
        }
    }
}

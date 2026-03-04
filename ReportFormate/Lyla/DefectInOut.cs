using CrystalDecisions.Shared;
using Pinnacle.Master;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static QRCoder.PayloadGenerator;

namespace Pinnacle.ReportFormate.Lyla
{
    public partial class DefectInOut : Form,ToolStripAccess
    {
        private static DefectInOut _instance;
        public static DefectInOut Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DefectInOut();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public DefectInOut()
        {
            InitializeComponent();
        }

        private void ProductionStatusReport_Load(object sender, EventArgs e)
        {


            
        }
        private void empty()
        {
pictureBox1.Image = null; comboctype.SelectedIndex = 0;
            this.BackColor = Class.Users.BackColors; comboctype.SelectedIndex =0;
            butheader.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            //string sel1 = "SELECT distinct B.GTCOMPMASTID,b.compcode,b.compname,a.finyear FROM  asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + Class.Users.HCompcode + "'";
            //DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
            //DataTable dt = ds.Tables["asptblpur"];
           
            string sel2 = "SELECT  0 asptblpurID,  ''pono FROM  dual union all SELECT  0 asptblpurID,  'ALL'pono FROM  dual union all SELECT distinct X.asptblpurID,  X.pono FROM ( SELECT distinct A.asptblpurID,  a.pono FROM  asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  where b.compcode='" + Class.Users.HCompcode+"' ORDER BY A.asptblpurid DESC) X ";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
            DataTable dt2 = ds2.Tables["asptblpur"];
            combopono.DisplayMember = "pono";
            combopono.ValueMember = "pono";
            combopono.DataSource = dt2;Class.Users.UserTime = 0;
     
        }
        public void GridLoad()
        {

        }
        public void News()
        {
            empty();

        }

        public void Saves()
        {
           
        }

        public void Prints()
        {
            
        }

        public void Searchs()
        {
            
        }

        public void Searchs(int EditID)
        {
            
        }

        public void Deletes()
        {
            
        }

        public void ReadOnlys()
        {
            
        }

        public void Imports()
        {
            
        }

        public void Pdfs()
        {
            
        }

        public void ChangePasswords()
        {
            
        }

        public void DownLoads()
        {
           

                DialogResult result = MessageBox.Show("Do you want to Formate ??", "PRINT OUT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result.Equals(DialogResult.OK))
            {

                //if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage2"])//your specific tabname
                //{                   
                   rd2.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\Nyla-Production.xls");
                //}
                //if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])//your specific tabname
                //{
                //    rd2.ExportToDisk(ExportFormatType.ExcelWorkbook, "d:\\Nyla-ProductionDateWiseReport.xls");
                //}

            }

            else
            {
                MessageBox.Show("Pls Select Combo Box Value");
            }
        }

        public void ChangeSkins()
        {
            
        }

        public void Logins()
        {
            
        }

        public void GlobalSearchs()
        {
            
        }

        public void TreeButtons()
        {
           
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }
        Models.CommonClass CC = new Models.CommonClass();
        private void button1_Click(object sender, EventArgs e)
        {
            if (combopono.Text != "")
            {

               

                Class.Users.Query = "SELECT A.asptblpurid,B.compcode,B.compname,B.address, a.barcode,A.PONO,A.colorname,A.sizename,'PRODUCTION PENDING' AS INWARDDATA,b.companylogo,A.ISSUETYPE,A.REMARKS,A.restitching FROM asptblpurdet1 A JOIN gtcompmast B ON A.compcode = B.gtcompmastid JOIN asptblpur C ON C.asptblpurid = A.asptblpurid AND C.compcode = B.GTCOMPMASTID  WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND A.pono='" + combopono.Text + "'  UNION ALL SELECT A.asptblpurid,B.compcode,B.compname,B.address, a.barcode,A.PONO,A.colorname,A.sizename,'CHECKING PENDING' AS INWARDDATA,b.companylogo,A.ISSUETYPE,A.REMARKS,A.restitching FROM asptblpurdet1 A JOIN gtcompmast B ON A.compcode = B.gtcompmastid JOIN asptblpur C ON C.asptblpurid = A.asptblpurid AND C.compcode = B.GTCOMPMASTID WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND A.pono='" + combopono.Text + "' ";
                Class.Users.dt = CC.select(Class.Users.Query, "asptblpurdet1");
                if (Class.Users.dt.Rows.Count > 0)
                {
                    Class.Users.dt1 = Class.Users.dt.Clone();
                }

                int trow = 0; Int64 inqty = 0, inbal = 0, inrew = 0, indel = 0, inchin = 0, inchch = 0, inchst = 0, inchdel = 0;
                if (Class.Users.dt.Rows.Count > 0)
                {
                    foreach (DataRow row in Class.Users.dt.Rows)
                    {
                        Class.Users.dt1.Rows.Add();
                        Class.Users.dt1.Rows[trow]["asptblpurid"] = Convert.ToInt64("0" + row[0].ToString());
                        Class.Users.dt1.Rows[trow]["compcode"] = row["compcode"].ToString();
                        Class.Users.dt1.Rows[trow]["compname"] = row["compname"].ToString();
                        Class.Users.dt1.Rows[trow]["address"] = row["address"].ToString();
                        Class.Users.dt1.Rows[trow]["pono"] = row["pono"].ToString();
                        Class.Users.dt1.Rows[trow]["barcode"] = row["barcode"].ToString();
                        Class.Users.dt1.Rows[trow]["colorname"] = row["colorname"].ToString();
                        Class.Users.dt1.Rows[trow]["sizename"] = row["sizename"].ToString();
                        Class.Users.dt1.Rows[trow]["INWARDDATA"] = row["INWARDDATA"].ToString();
                        if (row["companylogo"].ToString() != "")
                        {
                            Class.Users.dt1.Rows[trow]["companylogo"] = row["companylogo"];
                        }
                        trow++;
                    }


                    if (Class.Users.dt1.Rows.Count == Class.Users.dt.Rows.Count)
                    {
                        rd2.SetDataSource(Class.Users.dt1);
                        crystalReportViewer1.ReportSource = null;
                        crystalReportViewer1.ReportSource = rd2;
                        crystalReportViewer1.Refresh();
                    }
                    else
                    {

                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in Production Entry.");
                }

            }
            else
            {
                MessageBox.Show("Po-Number empty");

            }
            Cursor.Current = Cursors.Default;
            
        }
        Report.Lyla.Production rd2 = new Report.Lyla.Production();

             byte[] stdbytes; Int64 std;
        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(combopono.Text != "ALL")
            {

            }
            string sel="select a.garmentimage from asptblpur a where a.pono='"+combopono.Text+"'";
            DataTable dt = CC.select(sel, "asptblpur");
            int c = dt.Rows.Count;
            if (c >= 1)
            {
                if (c >= 1 && dt.Rows[0]["garmentimage"].ToString() !="")
                {
                    pictureBox1.Image = null; stdbytes = null;
                    stdbytes = (byte[])dt.Rows[0]["garmentimage"];
                    Image img = Models.Device.ByteArrayToImage(stdbytes);
                    pictureBox1.Image = img;
                }
                else
                {
                    pictureBox1.Image = null; stdbytes = null;
                }
            }
        }

        private void downLoadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //if (dataGridView1.Rows.Count > 0)
            //{
            //    Microsoft.Office.Interop.Excel.Application xcelapp = new Microsoft.Office.Interop.Excel.Application();
            //    xcelapp.Application.Workbooks.Add(Type.Missing);
               
            //    for (int i = 1; i < dataGridView1.Columns.Count+1; i++)
            //    {
            //        xcelapp.Cells[1,i] = dataGridView1.Columns[i - 1].HeaderText;
                    
            //    }
            //    for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < dataGridView1.Columns.Count; j++)
            //        {
            //            xcelapp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;

            //         }
            //    }
         
            //    xcelapp.Columns.AutoFit();
            //    xcelapp.Visible = true;
            //}
        }

        private void butdatewise_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;combopono.Text = "";
            
                Class.Users.Query = "SELECT A.asptblpurid,B.compcode,B.compname,B.address, a.barcode,A.PONO,A.colorname,A.sizename,'INWARD PENDING' AS INWARDDATA,b.companylogo FROM asptblpurdet1 A JOIN gtcompmast B ON A.compcode=B.gtcompmastid JOIN asptblpur C ON C.asptblpurid=A.asptblpurid AND C.compcode=B.GTCOMPMASTID  WHERE A.restitching='F' AND A.rechecking='F' AND A.ISSUETYPE='INWARD' AND A.delivery='F' AND    B.COMPCODE='" + Class.Users.HCompcode + "' and C.modified between '" + frmdate.Value.ToString("dd-MM-yyyy") + "' and  '" + todate.Value.ToString("dd-MM-yyyy") + "'  ORDER BY 1";
                Class.Users.dt = CC.select(Class.Users.Query, "asptblpurdet1");
               
            
            Class.Users.dt1 = Class.Users.dt.Clone();
            int trow = 0; Int64 inqty = 0, inbal = 0, inrew = 0, indel = 0, inchin = 0, inchch = 0, inchst = 0, inchdel = 0;
            if (Class.Users.dt.Rows.Count > 0)
            {
                foreach (DataRow row in Class.Users.dt.Rows)
                {
                 
                    Class.Users.dt1.Rows.Add();
                    Class.Users.dt1.Rows[trow]["asptblpurid"] = Convert.ToInt64("0" + row[0].ToString());
                    Class.Users.dt1.Rows[trow]["compcode"] = row["compcode"].ToString();
                    Class.Users.dt1.Rows[trow]["compname"] = row["compname"].ToString();
                    Class.Users.dt1.Rows[trow]["address"] = row["address"].ToString();
                    Class.Users.dt1.Rows[trow]["pono"] = row["pono"].ToString();
                    Class.Users.dt1.Rows[trow]["barcode"] = row["barcode"].ToString();
                    Class.Users.dt1.Rows[trow]["colorname"] = row["colorname"].ToString();
                    Class.Users.dt1.Rows[trow]["sizename"] = row["sizename"].ToString();
                    Class.Users.dt1.Rows[trow]["INWARDDATA"] = row["INWARDDATA"].ToString();
                    if (row["companylogo"].ToString() != "")
                    {
                        Class.Users.dt1.Rows[trow]["companylogo"] = row["companylogo"];
                    }
                    trow++;
                }

                if (Class.Users.dt1.Rows.Count == Class.Users.dt.Rows.Count)
                {

                  
                    rd2.SetDataSource(Class.Users.dt1);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd2;
                    crystalReportViewer1.Refresh();

                }
                else
                {
                    
                }
            }
            else
            {
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.Refresh();

                MessageBox.Show("No Data Found in Production Entry.");
            }
            Cursor.Current = Cursors.Default;
            
        }

        private void combotype_SelectedIndexChanged(object sender, EventArgs e)
        {

            reportwise();
        }
        void reportwise()
        {
            if (comboctype.Text == "PO WISE")
            {
                frmdate.Enabled = false;
                todate.Enabled = false;
                butdatewise.Enabled = false;
                combopono.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                combopono.Enabled = false;
                button1.Enabled = false;
                frmdate.Enabled = true;
                todate.Enabled = true;
                butdatewise.Enabled = true;
            }
        }
    }
}

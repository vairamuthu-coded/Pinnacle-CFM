using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using System.IO;
using System.Drawing.Imaging;
using QRCoder;
namespace Pinnacle.Report
{
    public partial class QrCode : Form
    {
        Report.QrCodeCrystalReport crypt = new Report.QrCodeCrystalReport();
       // ReportDocument rd = new ReportDocument();
        private static QrCode _instance;
      //  Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        public QrCode()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            lbl_Header.Text = Class.Users.ScreenName;
        }
        public static QrCode Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new QrCode();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

        public ImageFormat ImageFormate { get; private set; }

        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                        if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                        if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                        if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                        if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                        if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                        if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                        if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                        if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                        if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                        if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                        if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                        if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true;  } else { this.DownLoads.Visible = false; }
                        if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                        if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }
                    }
                }


            }
            else
            {

            }

        }
        private void CrystalReportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }
        QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
        private void QrCode_Load(object sender, EventArgs e)
        {

            string sel = "SELECT ASPTBLBUYSAMID,BUYERCODE,FABRIC,CONTENT,GARMENTIMAGE,GIMAGE,QRCODE  FROM  ASPTBLBUYSAM A ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLBUYSAM");
            DataTable dt = ds.Tables["ASPTBLBUYSAM"];
            string str = "";
               Byte[] qrbytes;
            qrbytes = (byte[])dt.Rows[0]["QRCODE"];
            str = Encoding.Default.GetString(qrbytes);
            PictureBox P = new PictureBox();
            var mydata = qc.CreateQrCode(str, QRCoder.QRCodeGenerator.ECCLevel.L);
            var code = new QRCoder.QRCode(mydata);
            P.Image = code.GetGraphic(50);

            // QRCodecrystalReportViewer1.ReportSource = null;
            crypt.SetDataSource(dt);
            crypt.Load("Report.QrCodeCrystalReport.rpt");
            QRCodecrystalReportViewer1.ReportSource = crypt;
            QRCodecrystalReportViewer1.Refresh();

        }

        private void Button1_Click(object sender, EventArgs e)
        {

        }

        private void CrystalReportViewer1_Load_1(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void QRCodecrystalReportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void Saves_Click(object sender, EventArgs e)
        {

        }
    }
}

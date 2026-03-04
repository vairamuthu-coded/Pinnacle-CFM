using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Report
{
    public partial class HostelGatePassPrint : Form
    {
        public HostelGatePassPrint()
        {
            InitializeComponent();
        }
        public zkemkeeper.CZKEM axCZKEM1 = new zkemkeeper.CZKEM();
        private bool bIsConnected = false;//the boolean value identifies whether the device is connected
        private int iMachineNumber;//the serial number of the device.After connecting the device ,this value will be changed.      
        string sdwEnrollNumber = "";
        int idwVerifyMode = 0;
        int idwInOutMode = 0;
        int idwYear = 0;
        int idwMonth = 0;
        int idwDay = 0;
        int idwHour = 0;
        int idwMinute = 0;
        int idwSecond = 0;
        int idwWorkcode = 0;
        int idwErrorCode = 0;
        int iGLCount = 0;
        int iIndex = 0;
           Label lblheader2 = new Label();
               
                Label lblqrcode = new Label();
               
        private void HostelGatePassPrint_Load(object sender, EventArgs e)
        {
            string sel1 = " SELECT  MAX(A.ASPTBLHOSTELGATEPASSID) ID  FROM ASPTBLHOSTELGATEPASS A";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
            DataTable dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
            if (dt.Rows.Count > 0)
            {
                string IDD = Convert.ToString(dt.Rows[0]["ID"].ToString());
                string sel2 = " SELECT  A.ASPTBLHOSTELGATEPASSID,B.COMPCODE,C.IDCARDNO,C.EMPNAME,D.DEPARTMENT,C.CONTACT,E.HOSTELNAME,E.HOSTELBLOCK,  E.HOSTELROOM,F.REASON,A.PERMISSIONHRS,A.SYSTEMDATE,A.SYSTEMTIME FROM ASPTBLHOSTELGATEPASS A JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE JOIN   ASPTBLEMP C ON C.COMPCODE = A.COMPCODE AND C.COMNAME=B.GTCOMPMASTID  JOIN ASPTBLDEP D ON D.ASPTBLDEPID = A.DEPARTMENT JOIN ASPTBLHOSTELMAS E ON E.EMPNAME=A.EMPNAME AND E.EMPNAME=C.ASPTBLEMPID JOIN ASPTBLREASONMAS F ON F.ASPTBLREASONMASID=A.REASON WHERE A.ASPTBLHOSTELGATEPASSID=" + IDD;
                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                DataTable dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                lblHOSTELGATEPASSID.Text = dt2.Rows[0]["ASPTBLHOSTELGATEPASSID"].ToString();
                lblCOMPCODE.Text = dt2.Rows[0]["COMPCODE"].ToString();
                lblIDCARDNO.Text = dt2.Rows[0]["IDCARDNO"].ToString();
                lblEMPNAME.Text = dt2.Rows[0]["EMPNAME"].ToString();
                lblDEPARTMENT.Text = dt2.Rows[0]["DEPARTMENT"].ToString();
                lblCONTACT.Text = dt2.Rows[0]["CONTACT"].ToString();
                lblHOSTELNAME.Text = dt2.Rows[0]["HOSTELNAME"].ToString();
                lblHOSTELBLOCK.Text = dt2.Rows[0]["HOSTELBLOCK"].ToString();
                lblHOSTELROOM.Text = dt2.Rows[0]["HOSTELROOM"].ToString();
                lblREASON.Text = dt2.Rows[0]["REASON"].ToString();
                lblPERMISSIONHRS.Text = dt2.Rows[0]["PERMISSIONHRS"].ToString();
                lblSYSTEMDATE.Text = dt2.Rows[0]["SYSTEMDATE"].ToString();
                lblSYSTEMTIME.Text = dt2.Rows[0]["SYSTEMTIME"].ToString();

             
                lblqrcode.Text = "-------------------------------------------------------\n" + "TOKENNO           : " + dt2.Rows[0]["ASPTBLHOSTELGATEPASSID"].ToString() + "\n" + "COMPCODE        : " + dt2.Rows[0]["COMPCODE"].ToString() + "\n" + "IDCARDNO          : " + dt2.Rows[0]["IDCARDNO"].ToString() + "\n" + "EMP NAME          : " + dt2.Rows[0]["EMPNAME"].ToString() + "\n" + "DEPARTMENT     : " + dt2.Rows[0]["DEPARTMENT"].ToString() + "\n" + "CONTACT NO      : " + dt2.Rows[0]["CONTACT"].ToString() + "\n" + "HOSTEL NAME    : " + dt2.Rows[0]["HOSTELNAME"].ToString() + "\n" + "HOSTEL BLOCK  : " + dt2.Rows[0]["HOSTELBLOCK"].ToString() + "\n" + "HOSTEL ROOM   : " + dt2.Rows[0]["HOSTELROOM"].ToString() + "\n" + "REASON               : " + dt2.Rows[0]["REASON"].ToString() + "\n" + "PER HRS              : " + dt2.Rows[0]["PERMISSIONHRS"].ToString() + "\n" + "DATE TIME           : " + System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + "\n-------------------------------------------------------\n";
                lblheader2.Text = lblgatepass.Text + "\n" + lblqrcode.Text;
                QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                var mydata1 = qc.CreateQrCode(IDD, QRCoder.QRCodeGenerator.ECCLevel.L);
                var code1 = new QRCoder.QRCode(mydata1);
                pictureBox2.Image = code1.GetGraphic(5, Color.Black, Color.White, true);
                rlblcompcode.Text = "For   " + Class.Users.HCompName.ToLower();


            }
        }

        private void CrystalReportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString(lblgatepass.Text, new Font("Calibri", 8, FontStyle.Regular), Brushes.DarkBlue, 111, 29);
            //e.Graphics.DrawString(lblqrcode.Text, new Font("Calibri", 8, FontStyle.Regular), Brushes.DarkBlue, 14, 42);
            e.Graphics.DrawImage(pictureBox2.Image, 26, 267, pictureBox2.Width, pictureBox2.Height);
          //  e.Graphics.DrawImage(pictureBox1.Image, 160, 256, pictureBox1.Width, pictureBox1.Height);
         ///   e.Graphics.DrawString(rlblcompcode.Text.Substring(0, 3).ToLower(), new Font("Calibri", 8, FontStyle.Regular), Brushes.DarkBlue, 27, 364); ;

         

        }

        private void Butprint1_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {


                printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                printDocument1.Print();

                //if (bIsConnected == true)
                //{
                //    int idwErrorCode = 0;

                //    int iDataFlag = 1;
                //    if (axCZKEM1.ClearData(iMachineNumber, iDataFlag))
                //    {
                //        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                //    }
                //    else
                //    {
                //        axCZKEM1.GetLastError(ref idwErrorCode);
                //    }
                //    axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
                //}

            }
            else
            {
                //printpanel.Hide();
                //printpanel.Refresh();
            }

            this.Hide();
        }
    }
}

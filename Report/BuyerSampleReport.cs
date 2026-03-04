using System;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using QRCoder;
namespace Pinnacle.Report
{
    public partial class BuyerSampleReport : Form
    {
        Int64 qridd = 0;
        Byte[] qrbytes;
        Byte[] bytes;
        string myString = "";
        // Models.BuyerSample buy = new Models.BuyerSample();
        public BuyerSampleReport(int qrid)
        {
            InitializeComponent();
         
            qridd = qrid;
           // rlblcompcode.Text = Class.Users.HCompName;
            //rlbldate.Text = Class.Users.SysDate;
        }
        public BuyerSampleReport()
        {

        }
        QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
        public static string CapitalizeFirstLetters(string sValue)
        {



            char[] array = sValue.ToCharArray();

            if (array.Length >= 1)
            {
                if (char.IsLower(array[0]))
                {
                    array[0] = char.ToUpper(array[0]);
                }
            }

            for (int i = 1; i < array.Length; i++)
            {
                if (array[i] == ' ')
                {
                    if (char.IsLower(array[i + 1]))
                    {
                        array[i + 1] = char.ToUpper(array[i + 1]); // space is identified at Index i so next i+1 has to converted to Upper!!
                    }
                }
            }


            return new string(array);

        }
        Report.Sample.SampleCollectionPrintFormate rd2 = new Sample.SampleCollectionPrintFormate();
        //private BuyerSampleReport GetData(string query)
        //{
        //    string conString = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
        //    SqlCommand cmd = new SqlCommand(query);
        //    using (SqlConnection con = new SqlConnection(conString))
        //    {
        //        using (SqlDataAdapter sda = new SqlDataAdapter())
        //        {
        //            cmd.Connection = con;
        //            sda.SelectCommand = cmd;
        //            using (BuyerSampleReport dsCustomers = new BuyerSampleReport())
        //            {
        //                sda.Fill(dsCustomers);
        //                dsCustomers.Tables["Table"].Columns.Add(new DataColumn("QRCode", typeof(byte[])));
        //                foreach (DataRow dr in dsCustomers.Tables["Table"].Rows)
        //                {
        //                    dr["QRCode"] = GenerateQrCode(dr["CustomerID"].ToString());
        //                }
        //                return dsCustomers;
        //            }
        //        }
        //    }
        //}
        DataTable dt = new DataTable();


        private void BuyerSampleReport_Load(object sender, EventArgs e)
        {
              string sel = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,C.BRAND,A.AGFSAMPLE,E.SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.GARMENTIMAGE,A.QRCODE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE  where A.ACTIVE='T' AND  A.ASPTBLBUYSAMID=" + Convert.ToInt64("0" + qridd);
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLBUYSAM");
            dt = null;
            dt = ds.Tables["ASPTBLBUYSAM"];
            if (dt.Rows.Count > 0)
            {
                rd2.SetDataSource(dt);
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.ReportSource = rd2;
                crystalReportViewer1.Refresh();
            }
            else
            {
                MessageBox.Show("This Sample Active'F' "); this.Close(); return;
            }
            //string sel0 = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,C.BRAND,A.AGFSAMPLE AS REFCODE,E.SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.GARMENTIMAGE,A.QRCODE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR AS BUYERCODE   FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.RATE JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE");
            //DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLBUYSAM");
            //DataTable dt0 = ds0.Tables["ASPTBLBUYSAM"];

            //string folderLocation = "C:\\SampleCollections-Download\\";
            //if (!Directory.Exists(folderLocation)) { Directory.CreateDirectory(folderLocation); }
            //for (int i = 0; i < dt0.Rows.Count; i++)
            //{
            //    rd2.ExportToDisk(ExportFormatType.Excel, folderLocation + "-" + dt.Rows[0]["DATE1"].ToString() + " SampleCollections.xls");
            //}
            //string sel = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,C.BRAND,A.AGFSAMPLE AS BUYERCODE,E.SEASON,F.DEPARTMENT,G.STYLENAME,H.SUBSTYLE,I.FABRIC,J.COUNTS,K.GG AS GAUGE,L.GSM,M.COLORNAME,N.SIZENAME,A.RATE,A.REMARKS,A.ACTIVE,A.GARMENTIMAGE,A.QRCODE  FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT JOIN ASPTBLSTYMAS G ON G.ASPTBLSTYMASID=A.STYLENAME  JOIN ASPTBLSUBSTYMAS H ON H.ASPTBLSUBSTYMASID=A.SUBSTYLE AND G.ASPTBLSTYMASID=H.STYLENAME JOIN  ASPTBLFABMAS I ON I.ASPTBLFABMASID=A.FABRIC JOIN  GTCOUNTRYMAST J ON J.GTCOUNTRYMASTID=A.COUNTS JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE JOIN ASPTBLGSMMAS L ON L.ASPTBLGSMMASID=A.GSM  JOIN ASPTBLCOLMAS M ON M.ASPTBLCOLMASID=A.COLORNAME JOIN ASPTBLSIZMAS N ON N.ASPTBLSIZMASID=A.SIZENAME where A.ASPTBLBUYSAMID=" + Convert.ToInt64("0" + qridd);
            //DataSet ds2 = Utility.ExecuteSelectQuery(sel, "ASPTBLBUYSAM");
            //DataTable dt2 = ds2.Tables["ASPTBLBUYSAM"];
            //CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //reportdocument.Load(Application.StartupPath + "\\Report\\SampleCollectionCrystalReport.rpt");
            //reportdocument.SetDataSource(dt2);
            //reportdocument.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
            //reportdocument.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);
            //this.Dispose();
        }
        private void button1_Click(object sender, EventArgs e)
        {

            string sel = "SELECT A.ASPTBLBUYSAMID,  A.DATE1,B.COMPCODE,C.BRAND,A.AGFSAMPLE,E.SEASON,F.DEPARTMENT || '(' ||  O.CATEGORY || ')' AS DEPARTMENT ,O.CATEGORY,A.STYLENAME,A.SUBSTYLE,A.FABRIC,A.FABRICCONTENT,A.COUNTS, K.GG AS GAUGE,A.GSM, A.COLORNAME,P.ORDERPACKTYPE,A.SIZENAME,N.CURRENCYNAME,A.REMARKS,A.ACTIVE,A.GARMENTIMAGE,A.QRCODE,A.RISK1,A.RISK2,A.RISK3,A.RISK4,A.RISK5,A.FABRICCOMPLIANT,A.REMARKS,A.MFYEAR    FROM ASPTBLBUYSAM A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID join asptblBRANDMAS C ON C.ASPTBLBRANDMASID=A.BRAND  JOIN ASPTBLSEASONMAS E ON E.ASPTBLSEASONMASID=A.SEASON JOIN ASPTBLSAMDEPTMAS F ON F.ASPTBLSAMDEPTMASID=A.DEPARTMENT   JOIN ASPTBLGGMAS K ON K.ASPTBLGGMASID=A.GAUGE      JOIN ASPTBLCURMAS N ON N.ASPTBLCURMASID=A.CURRENCYNAME JOIN ASPTBLSAMCATMAS O ON O.ASPTBLSAMCATMASID=A.CATEGORY  JOIN ASPTBLORDPACKMAS  P ON P.ASPTBLORDPACKMASID=A.ORDERPACKTYPE  where A.ACTIVE='T' AND  A.ASPTBLBUYSAMID=" + Convert.ToInt64("0" + qridd);
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLBUYSAM");
            dt = null;
            dt = ds.Tables["ASPTBLBUYSAM"];
            rd2.SetDataSource(dt);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd2;
            crystalReportViewer1.Refresh();
            rd2.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
            rd2.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);
            this.Dispose();
        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
        }



        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Close();


        }

    }
}

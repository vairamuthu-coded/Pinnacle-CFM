using System;

using System.Data;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace Pinnacle.Report
{
    public partial class FuelPrint : Form
    {
       
       

        public FuelPrint()
        {
            InitializeComponent();

           
        }

       

        public object MatrixBarcode { get; private set; }
        Report.FuelCrystalReport rd = new Report.FuelCrystalReport();
        private void FuelPrint_Load(object sender, EventArgs e)
        {


            string sel2 = "SELECT A.ASPTBLVEHTOKENID1,C.COMPNAME ,L.CITYNAME ,I.BUNKADDRESS  AS QRCODE ,TO_CHAR(C.PINCODE) AS PINCODE, TO_CHAR(A.ASPTBLVEHTOKENID1)AS TOKENNO,TO_CHAR(SUBSTR(A.TOKENDATE,0,10)) AS TOKENDATE , (  SELECT  AA.VEHICLENO    FROM HRVEHMAST AA  JOIN HRVEHTYPEMAST Ab on AA.VEHICLETYPE=AB.HRVEHTYPEMASTID     JOIN GTCOMPMAST AC ON AC.GTCOMPMASTID=AA.COMPCODE     JOIN ASPTBLVEHTOKEN AD ON AD.VEHICLENO=AA.HRVEHMASTID   AND AA.ACTIVE='T'  WHERE  AD.ASPTBLVEHTOKENID='" + Class.Users.Paramid + "'      UNION ALL  SELECT BA.VEHICLENO  FROM ASPTBLVEHMAS BA  JOIN HRVEHTYPEMAST BB ON BA.VEHICLETYPE=BB.HRVEHTYPEMASTID   JOIN GTCOMPMAST BC ON BC.GTCOMPMASTID=BA.COMPCODE  AND BA.ACTIVE='T'   JOIN ASPTBLVEHTOKEN BD ON BD.VEHICLENO=BA.ASPTBLVEHMASID    WHERE  BD.ASPTBLVEHTOKENID='" + Class.Users.Paramid + "'  ) AS VEHICLENO,CONCAT(E.FNAME, concat('-', F.MIDCARD)) as EMPNAME , K.ITEMNAME,I.BUNKNAME,J.LITRES,J.KM,''QRCODE  FROM ASPTBLVEHTOKEN A   JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  JOIN HREMPLOYMAST  E ON E.HREMPLOYMASTID = A.EMPNAME JOIN HREMPLOYDETAILS F ON  F.HREMPLOYMASTID = E.HREMPLOYMASTID  AND E.IDCARDNO = F.IDCARD  JOIN GTFINANCIALYEAR G ON G.GTFINANCIALYEARID = A.FINYEAR JOIN ASPTBLPETMAS I ON I.COMPCODE = A.COMPCODE AND I.COMPCODE = C.GTCOMPMASTID AND I.ASPTBLPETMASID=A.BUNKNAME JOIN ASPTBLVEHTOKENDET J ON J.ASPTBLVEHTOKENID = A.ASPTBLVEHTOKENID JOIN GTGENITEMMAST K ON K.GTGENITEMMASTID = J.ITEMNAME JOIN GTCITYMAST L ON L.GTCITYMASTID = C.CITY  WHERE A.ASPTBLVEHTOKENID= '" + Class.Users.Paramid + "' ";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLVEHTOKEN");
            DataTable dt2 = ds2.Tables["ASPTBLVEHTOKEN"];
            rd.SetDataSource(dt2);
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = rd;


        }
        private void butprint_Click(object sender, EventArgs e)
        {
            //rd.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
            //rd.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);


          //  FuelPrint_Load(sender, e);
            this.Close();
        }
        private void CrystalReport11_InitReport(object sender, EventArgs e)
        {

        }

        private void CrystalReportViewer1_Load(object sender, EventArgs e)
        {
            
           
        }

        private void CrystalReport41_InitReport(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
           
        }

        private void butcancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

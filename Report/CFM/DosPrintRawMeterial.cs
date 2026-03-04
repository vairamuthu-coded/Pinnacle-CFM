using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Report.CFM
{
    public partial class DosPrintRawMeterial : Form
    {
        private long v;string vechileno="";
        public DosPrintRawMeterial()
        {
            InitializeComponent();
            this.v = Class.Users.Paramid;
            this.vechileno = Class.Users.PayPeriod;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString(labelFromCompany.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 317, 0);
            e.Graphics.DrawString(labelfromaddress.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, 295, 24);
            e.Graphics.DrawString(labelheading.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 342, 50); 
            e.Graphics.DrawString(labelcertificate1.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 469, 50);

            e.Graphics.DrawString(labelDate.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 39, 83);
            e.Graphics.DrawString(label1.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 169, 83);
            e.Graphics.DrawString(labelDate1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 177, 83);

            e.Graphics.DrawString(labelreceivedfrom.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 39, 109);
            e.Graphics.DrawString(label2.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 169, 109);
            e.Graphics.DrawString(labelToCompany.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 177, 109);

            e.Graphics.DrawString(labelVechileNo.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 39, 135);
            e.Graphics.DrawString(label3.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 169, 135);
            e.Graphics.DrawString(labelVechileNo1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 177, 135);

            RectangleF rect = new RectangleF(10.0F, 10.0F, 200.0F, 30.0F);
            StringFormat format = new StringFormat(StringFormatFlags.DirectionRightToLeft);
            e.Graphics.DrawString(labelFirstWeight.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 39, 160);
            e.Graphics.DrawString(label5.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 169, 160);
            e.Graphics.DrawString(labelFirstWeight1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 256, 160, format);

            e.Graphics.DrawString(labelsecondweight.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 39, 186);
            e.Graphics.DrawString(label6.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 169, 186);
            e.Graphics.DrawString(labelsecondweight1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 256, 186, format);

            e.Graphics.DrawString(labelnetwegiht.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 39, 212);
            e.Graphics.DrawString(label7.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 169, 212);
            e.Graphics.DrawString(labelnetwegiht1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 256, 212, format);

            e.Graphics.DrawString(labelitemvariety.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 464, 83);
            e.Graphics.DrawString(label8.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 615, 83);
            e.Graphics.DrawString(labelitemvariety1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 625, 83);
       
            e.Graphics.DrawString(labelnoofbags.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 464, 109);
            e.Graphics.DrawString(label9.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 615, 109);
            e.Graphics.DrawString(labelnoofbags1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 625, 109);

            e.Graphics.DrawString(labelwargonReceiptno.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 464, 135);
            e.Graphics.DrawString(label10.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 615, 135);
            e.Graphics.DrawString(labelwargonReceiptno1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 625, 135);


            e.Graphics.DrawString(labelsignature.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 32, 289);
            e.Graphics.DrawString(labelforauthority.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 469, 289);



        }

        private void DosPrintRawMeterial_Load(object sender, EventArgs e)
        {
            try
            {
                if (v > 0 && vechileno != "")
                {
                    Cursor = Cursors.WaitCursor;
                    string sel2 = "select a.asptblrawmaterialid,a.asptblrawmaterialid1,a.certificateno,b.compcode, b.compname,a.finyear as fromdate, c.partyname as RecevedFrom, g.cityname, b.address,b.email,b.gstno,b.gstdate,b.panno,b.pincode, a.vechileno,date_format(a.datetime1, '%d-%m-%Y') as datetime1,a.createdon,d.itemname,a.tripwagonno,a.lotno,e.godownname, a.sampledby,a.certifiedby,a.grossweight,a.tareweight,a.netweight,a.noofbag from asptblrawmaterial a join gtcompmast b on a.compname = b.gtcompmastid  join asptblpartymas c on a.receivedFrom = c.asptblpartymasid join asptblitemmast d on a.itemnamevarity = d.asptblitemmastid join asptblgodwonmas e on a.godownname = e.asptblgodwonmasid join gtcitymast g on g.gtcitymastid = b.city  where a.asptblrawmaterialid='" + v + "' and a.vechileNo='" + vechileno + "';";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblrawmaterial");
                    DataTable dt2 = ds2.Tables["asptblrawmaterial"];
                    // labelFromCompany.Text = dt2.Rows[0]["compname"].ToString();
                    labelfromaddress.Text = dt2.Rows[0]["address"].ToString();
                    labelcertificate1.Text = " " + dt2.Rows[0]["compcode"].ToString() + "/" + dt2.Rows[0]["fromdate"].ToString() + "/" + dt2.Rows[0]["asptblrawmaterialid1"].ToString();
                    labelDate1.Text = dt2.Rows[0]["createdon"].ToString();
                    // labelsendto.Text = dt2.Rows[0]["RecevedFrom"].ToString(); 
                    labelToCompany.Text = dt2.Rows[0]["RecevedFrom"].ToString();
                    //   labelVechileNo.Text = dt2.Rows[0]["RecevedFrom"].ToString();
                    labelVechileNo1.Text = dt2.Rows[0]["vechileno"].ToString();
                    //  labelcertificate.Text = dt2.Rows[0]["RecevedFrom"].ToString();

                    //  labelFirstWeight.Text = dt2.Rows[0]["RecevedFrom"].ToString();
                    labelFirstWeight1.Text = dt2.Rows[0]["grossweight"].ToString();
                    //  labelsecondweight.Text = dt2.Rows[0]["RecevedFrom"].ToString();
                    labelsecondweight1.Text = dt2.Rows[0]["tareweight"].ToString();
                    //    labelnetwegiht.Text = dt2.Rows[0]["RecevedFrom"].ToString();
                    labelnetwegiht1.Text = dt2.Rows[0]["netweight"].ToString();

                    labelitemvariety1.Text = dt2.Rows[0]["itemname"].ToString();

                    labelnoofbags1.Text = dt2.Rows[0]["noofbag"].ToString();

                    labelwargonReceiptno1.Text = dt2.Rows[0]["tripwagonno"].ToString();
                    //  labelsignature.Text = dt2.Rows[0]["RecevedFrom"].ToString();
                    // labelforauthority.Text = "For   " + dt2.Rows[0]["compname"].ToString();
                    printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                    printDocument1.Print(); this.Close();
                }
                else
                {
                    MessageBox.Show("invalid .  ID :- " + v + " Vechile No" + vechileno);
                    this.Close();
                }
            }
            catch(Exception ex) { Cursor = Cursors.Default; }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

      
    }
}

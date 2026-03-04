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
    public partial class DosPrint_DeliveryCertificatecs : Form
    {
        private long v;string vechile = "";
        public DosPrint_DeliveryCertificatecs()
        {
            InitializeComponent();
            this.v =  Class.Users.Paramid;
            this.vechile = Class.Users.PayPeriod;
         //  usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
           // lbl_Header.Text = Class.Users.ScreenName;
        }
        private static DosPrint_DeliveryCertificatecs _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        public static DosPrint_DeliveryCertificatecs Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DosPrint_DeliveryCertificatecs();
                return _instance;
            }
        }

        //public void usercheck(string s, string ss, string sss)
        //{
        //    try
        //    {
        //        DataTable dt1 = sm.headerdropdowns(s, ss, sss);
        //        if (dt1.Rows.Count > 0)
        //        {
        //            if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
        //            {
        //                for (int r = 0; r < dt1.Rows.Count; r++)
        //                {

        //                    if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
        //                    if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
        //                    if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
        //                    if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
        //                    if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
        //                    if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
        //                    if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = false;} else { this.TreeButtons.Visible = false;  }
        //                    if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
        //                    if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
        //                    if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
        //                    if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
        //                    if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
        //                    if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
        //                    if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }

        //                }
        //            }


        //        }
        //        else
        //        {
        //            MessageBox.Show("Invalid");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("usercheck: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        int i = 1; decimal totweight = 0;
        private void DosPrint_DeliveryCertificatecs_Load(object sender, EventArgs e)
        {
            try
            {

                if (v > 0 && vechile != "")
                {
                    totweight = 0; Cursor = Cursors.WaitCursor;
                    string sel2 = " SELECT a.asptbldeliveryid, a.asptbldeliveryid1,    a.finyear,  c.compcode,a.finyear as  fromdate, c.compname as receivedfrom, d.cityname,c.address,c.gstno, c.gstdate,c.panno,c.pincode,c.devision, a.certificateno,a.vechileNo,  date_format(a.deliverydate, '%d-%m-%Y') as deliverydate,  a.deliverytime,    a.firstweight,    a.secondweight, a.netweight,    a.productweight , a.sendto, a.certifiedby,h.productname1 as productname, format(g.productweight,0,4) as productweight1,g.productkgs  FROM asptbldelivery a  join gtcompmast c on c.gtcompmastid=a.compcode  join gtcitymast d on d.gtcitymastid=c.city  join asptbldeliverydet g on g.asptbldeliveryid=a.asptbldeliveryid join asptblproductweightmas h on h.asptblproductweightmasid=g.productname where a.asptbldeliveryid='" + v + "'  and a.vechileNo='" + vechile + "'  ;";

                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldelivery");
                    DataTable dt2 = ds2.Tables["asptbldelivery"];
                    if (dt2.Rows.Count == 0)
                    {
                        string sel1 = " SELECT a.asptbldeliveryid, a.asptbldeliveryid1,    a.finyear,  c.compcode,a.finyear as fromdate, c.compname as receivedfrom, d.cityname,c.address,c.gstno, c.gstdate,c.panno,c.pincode,c.devision, a.certificateno,a.vechileNo,  date_format(a.deliverydate, '%d-%m-%Y') as deliverydate,  a.deliverytime,    a.firstweight,    a.secondweight, abs(a.netweight) as netweight,    a.productweight , a.sendto, a.certifiedby,'' as productname FROM asptbldelivery a  join gtcompmast c on c.gtcompmastid = a.compcode  join gtcitymast d on d.gtcitymastid = c.city    where a.asptbldeliveryid='" + v + "'   and a.vechileNo='" + vechile + "' ;";
                        ds2 = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                        dt2 = ds2.Tables["asptbldelivery"];
                    }

                    labelfromaddress.Text = dt2.Rows[0]["address"].ToString();
                    labelToCompany.Text = dt2.Rows[0]["sendto"].ToString();
                    labelDate1.Text = dt2.Rows[0]["deliverydate"].ToString();
                    labelcertificate1.Text = " " + dt2.Rows[0]["compcode"].ToString() + "/" + dt2.Rows[0]["fromdate"].ToString() + "/" + dt2.Rows[0]["asptbldeliveryid1"].ToString();
                    labletime1.Text = dt2.Rows[0]["deliverytime"].ToString();
                    labelVechileNo1.Text = dt2.Rows[0]["vechileNo"].ToString();

                    labelFirstWeight1.Text = dt2.Rows[0]["firstweight"].ToString();
                    labelsecondweight1.Text = dt2.Rows[0]["secondweight"].ToString();
                    labelnetwegiht1.Text = dt2.Rows[0]["netweight"].ToString();
                    if (dt2.Rows[0]["productname"].ToString() != "")
                    {
                        foreach (DataRow myRow in dt2.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            i += listView1.Items.Count;
                            list.Text = i.ToString();
                            list.SubItems.Add(myRow["productname"].ToString());
                            list.SubItems.Add(myRow["productweight1"].ToString());
                            list.SubItems.Add(myRow["productkgs"].ToString());
                            listView1.Items.Add(list);
                            totweight += Convert.ToDecimal(myRow["productkgs"].ToString());
                        }
                    }
                    else
                    {
                        totweight += Convert.ToDecimal(dt2.Rows[0]["netweight"].ToString());
                    }
                    labeltotalproductweight1.Text = Convert.ToDecimal(string.Format("{0:00.000}", totweight)).ToString() + " Kg";



                    printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                    printDocument1.Print();
                    Cursor = Cursors.Default;
                    this.Close();


                }
                else
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("invalid .  ID :- " + v + " Vechile No" + vechile);
                    this.Close();
                }
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void Prints_Click(object sender, EventArgs e)
        {
            printDocument1.PrinterSettings = printDialog1.PrinterSettings;
            printDocument1.Print();

        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            RectangleF rect = new RectangleF(10.0F, 10.0F, 200.0F, 30.0F);
            StringFormat format = new StringFormat(StringFormatFlags.DirectionRightToLeft);

            e.Graphics.DrawString(labelFromCompany.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 329, 0);

            e.Graphics.DrawString(labelfromaddress.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.Black, 310, 18);
            e.Graphics.DrawString(labelheading.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 359, 38);
            e.Graphics.DrawString(labelcertificate1.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 498, 38);

            e.Graphics.DrawString(labelDate.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 25, 81);
            e.Graphics.DrawString(label1.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 159, 83);
            e.Graphics.DrawString(labelDate1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 172, 81);


            e.Graphics.DrawString(labletime.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 25, 105);
            e.Graphics.DrawString(label4.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 159, 107);
            e.Graphics.DrawString(labletime1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 172, 105);


            e.Graphics.DrawString(labelVechileNo.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 25, 130);
            e.Graphics.DrawString(label3.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 159, 130);
            e.Graphics.DrawString(labelVechileNo1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 172, 130);
            StringFormat drawFormat = new StringFormat();
            drawFormat.Alignment = StringAlignment.Near;
            e.Graphics.DrawString(labelFirstWeight.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 25, 155);
            e.Graphics.DrawString(label5.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 159, 152);
            e.Graphics.DrawString(labelFirstWeight1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 245, 155, format);

            e.Graphics.DrawString(labelsecondweight.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 25, 179);
            e.Graphics.DrawString(label6.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 159, 179);
            e.Graphics.DrawString(labelsecondweight1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 245, 181, format);

            e.Graphics.DrawString(labelnetwegiht.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 25, 207);
            e.Graphics.DrawString(label7.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 159, 207);
            e.Graphics.DrawString(labelnetwegiht1.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 245, 207, format);

            e.Graphics.DrawString(labelheading1.Text, new Font("Arial", 9, FontStyle.Bold), Brushes.Black, 270, 84);
            e.Graphics.DrawString(labelheading2.Text, new Font("Arial", 9, FontStyle.Bold), Brushes.Black, 322, 84);
            e.Graphics.DrawString(labelheading3.Text, new Font("Arial", 9, FontStyle.Bold), Brushes.Black, 675, 84);
            e.Graphics.DrawString(labelheading4.Text, new Font("Arial", 9, FontStyle.Bold), Brushes.Black, 713, 84);
            if (listView1.Items.Count >= 1)
            {
                int yPos = 107;
                int i = 1;

                foreach (ListViewItem xItem in listView1.Items)
                {
                    e.Graphics.DrawString(i.ToString(), new Font("Arial", 8), Brushes.Black, 278, yPos);
                    e.Graphics.DrawString(xItem.SubItems[1].Text, new Font("Arial", 8), Brushes.Black, 321, yPos);
                    e.Graphics.DrawString(xItem.SubItems[2].Text, new Font("Arial", 8), Brushes.Black, 709, yPos, format);
                    e.Graphics.DrawString(xItem.SubItems[3].Text.ToString(), new Font("Arial", 8, FontStyle.Regular), Brushes.Black, 795, yPos, format);
                    yPos += 17; i++;

                }
                e.HasMorePages = false;

            }
            e.Graphics.DrawString(labelsendto.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 25, 379);
            e.Graphics.DrawString(label2.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 159, 379);
            e.Graphics.DrawString(labelToCompany.Text, new Font("Arial", 10, FontStyle.Regular), Brushes.Black, 172, 379);

            e.Graphics.DrawString(labeltotalproductweight.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 322, 342);
            e.Graphics.DrawString(labeltotalproductweight1.Text, new Font("Arial", 10, FontStyle.Bold), Brushes.Black, 695, 342);
            e.Graphics.DrawString(labelsignature.Text, new Font("Arial", 9, FontStyle.Regular), Brushes.Black, 25, 458);
            e.Graphics.DrawString(labelforauthority.Text, new Font("Arial", 9, FontStyle.Regular), Brushes.Black, 520, 458);

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

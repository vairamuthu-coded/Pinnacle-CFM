using QRCoder;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using QRCoder;

namespace Pinnacle.ReportFormate.Lyla
{
    public partial class LylaReport : Form
    {
        public LylaReport()
        {
            InitializeComponent();
        }
        Byte[] qrbytes; byte[] byteImage;
        QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
        
        //void QrcodeGenerate(string col1)
        //{
        //    try
        //    {
        //        PictureBox picturegrcode = new PictureBox();
        //        qrbytes = null;


        //            var mydata = qc.CreateQrCode(col1, QRCoder.QRCodeGenerator.ECCLevel.L);
        //            var code = new QRCoder.QRCode(mydata);

        //            picturegrcode.Image = code.GetGraphic(5, Color.Black, Color.White, true);
        //            MemoryStream stream = new MemoryStream();
        //            picturegrcode.Image.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
        //            qrbytes = stream.ToArray();
        //            //string str = Encoding.Default.GetString(qrbytes);
        //            //myString = str;
                

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("error" + ex.ToString());
        //    }
        //}
        QRCoder.QRCodeGenerator qrGenerator = new QRCoder.QRCodeGenerator();
        //private byte[] GenerateQrCode(string qrmsg)
        //{

        //    var mydata = qc.CreateQrCode(qrmsg, QRCoder.QRCodeGenerator.ECCLevel.Q);
        //    var qrCode = new QRCoder.QRCode(mydata);           
        //    System.Web.UI.WebControls.Image imgBarCode = new System.Web.UI.WebControls.Image();
        //    imgBarCode.Height = 150;
        //    imgBarCode.Width = 150;
        //    using (Bitmap bitMap = qrCode.GetGraphic(20))
        //    {
        //        using (MemoryStream ms = new MemoryStream())
        //        {
        //            bitMap.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
        //            byteImage = ms.ToArray();
        //            return byteImage;
        //        }
        //    }
        //}
        //CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        //Report.Lyla.BarcodeGeneral rd1 = new Report.Lyla.BarcodeGeneral();
        DataTable dtgeneral = new DataTable();
        DataTable reversedDt; DataTable reversedDt1;
        DataTable reversedDt2;int cnt = 0;
        private void LylaReport_Load(object sender, EventArgs e)
        {

            try
            {

                Class.Users.UserTime = 0;
                if (Class.Users.QrCode.Rows.Count >= 1)
                {
                    var orderedRows = from row in Class.Users.QrCode.AsEnumerable() select row;
                    reversedDt = orderedRows.CopyToDataTable();
                    if (reversedDt.Rows.Count > 0)
                    {
                        //reversedDt2.Rows.Clear();
                        //reversedDt1.Rows.Clear();
                        reversedDt1 = reversedDt.Clone();
                        reversedDt2 = reversedDt.Clone();
                    }

                    cnt = 0;

                    foreach (DataRow dr in Class.Users.QrCode.Rows)
                    {


                        if (cnt % 2 == 0)
                        {
                            reversedDt1.Rows.Add(dr.ItemArray);
                            reversedDt1.AcceptChanges();
                        }
                        else
                        {
                            reversedDt2.Rows.Add(dr.ItemArray);
                            reversedDt2.AcceptChanges();
                        }


                        cnt++;
                    }

                    if (reversedDt1.Rows.Count >= 2)
                    {
                        //Report\Lyla\BarcodeGeneral.rpt
                        //Report.Lyla.BarcodeGeneral
                        crystalReportViewer1.ReportSource = null;
                        CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                        reportdocument.Load(Application.StartupPath + "\\Report\\Lyla\\BarcodeGeneral.rpt");
                        reportdocument.Database.Tables["DataTable1"].SetDataSource(reversedDt1);
                        reportdocument.Database.Tables["DataTable11"].SetDataSource(reversedDt2);
                        crystalReportViewer1.ReportSource = reportdocument;
                        crystalReportViewer1.Refresh();
                    }
                    if (reversedDt1.Rows.Count <= 2)
                    {

                        crystalReportViewer1.ReportSource = null;
                          Report.Sample.SampleCollectionPrintFormate reportdocument = new Report.Sample.SampleCollectionPrintFormate();
                        reportdocument.Database.Tables["DataTable1"].SetDataSource(reversedDt1);
                        reportdocument.Database.Tables["DataTable11"].SetDataSource(reversedDt2);
                        crystalReportViewer1.ReportSource = reportdocument;
                        crystalReportViewer1.Refresh();
                      
                    }

                }
                else
                {
                    MessageBox.Show("Minimum TWO Sample should be Choose in Checkbox !.");
                  
                }
            }
            catch (Exception EX)
            {
                MessageBox.Show("--------------------Install CrystalReport---------------------------" + EX.ToString(), "Install Crystal Report", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //Report.Lyla.BarcodeGeneral rd1 = new Report.Lyla.BarcodeGeneral();
            //Class.Users.QrCode.TableName = "DataTable1";


            //Models.CommonClass com = new Models.CommonClass();

            //string[] sarray = Class.Users.Query.Split(':');
            //dt = com.select(sarray[0], sarray[4]);

            //int n = 2;          
            //int cnt = 0, k = 0, cnt1 = 0, cnt2 = 1, row = 0, col = 0, rowcount = 0, tot = 0;
            //col = Class.Users.QrCodeArray.Columns.Count - 2;
            //int totcount = 0;
            //rowcount = Class.Users.QrCodeArray.Rows.Count;
            //int cc = 0;
            //for (int i1 = 0; i1 < rowcount; i1++)
            //{
            //    tot = 0; row = 0; cc++;
            //    cnt1 = 0; row = 2;

            //    for (int i = 0; i < col; i++)
            //    {
            //        if (col != tot)
            //        {
            //            cnt2 = cc - 1; cnt1 = i + row; cnt = 0;
            //            cnt = Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2]["P-COUNT"].ToString()) * Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString());
            //            if (cnt >= 1)
            //            {
            //                k++;
            //                string sel2 = "select a.asptblpurdetid ,a.asptblpurid,b.asptblpurid1,b.pono   from  asptblpurdet a join asptblpur b on a.asptblpurid=b.asptblpurid join asptblsizmas  c on c.asptblsizmasid=a.sizename  where  b.compcode=" + sarray[1] + "  and b.finyear=" + sarray[2] + " and b.pono=" + sarray[3] + " and c.sizename='"+ Class.Users.QrCodeArray.Columns[cnt1].ToString() + "' ";
            //                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpurdet");
            //                DataTable dt2 = ds2.Tables["asptblpurdet"];
            //                int id1 = Convert.ToInt32(dt2.Rows[i1]["asptblpurdetid"].ToString());
            //                int id2 = Convert.ToInt32(dt2.Rows[i1]["asptblpurid"].ToString());
            //                int id3 = Convert.ToInt32(dt2.Rows[i1]["asptblpurid1"].ToString());
            //                string id4 = Convert.ToString(dt2.Rows[i1]["pono"].ToString());
            //                for (int a = 0; a < cnt; a++)
            //                {
            //                    totcount++;
            //                    dt.Rows.Add(0);
            //                    dt.Rows[dt.Rows.Count - 1][1] = id1;
            //                    dt.Rows[dt.Rows.Count - 1][2] = id2;
            //                    dt.Rows[dt.Rows.Count - 1][3] = id3;
            //                    dt.Rows[dt.Rows.Count - 1][4] = Class.Users.COMPCODE; 
            //                    dt.Rows[dt.Rows.Count - 1][5] = id4;
            //                    dt.Rows[dt.Rows.Count - 1][6] = Class.Users.QrCodeArray.Rows[i1]["COLORNAME"].ToString(); //GenerateQrCode(Class.Users.QrCodeArray.Rows[i]["COLORNAME"].ToString());
            //                    dt.Rows[dt.Rows.Count - 1][7] = Class.Users.QrCodeArray.Columns[cnt1].ToString();
            //                    dt.Rows[dt.Rows.Count - 1][8] = Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString();
            //                    dt.Rows[dt.Rows.Count - 1][9] = Class.Users.QrCodeArray.Rows[cnt2]["P-COUNT"].ToString();
            //                    dt.Rows[dt.Rows.Count - 1][10] = Class.Users.QrCodeArray.Rows[i1]["COLORNAME"].ToString()+" " + i1 + " "+ k.ToString() + " " + Class.Users.QrCodeArray.Columns[cnt1].ToString() + " "+ Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString() + " "+ totcount;
            //                    dt.Rows[dt.Rows.Count - 1][11] = "";
            //                }

            //                tot++;
            //            }
            //        }

            //    }
            //}
            //if (Class.Users.QrCode.Rows.Count > 0)
            //{
            //    rd1.SetDataSource(Class.Users.QrCode);
            //    crystalReportViewer1.ReportSource = null;
            //    crystalReportViewer1.ReportSource = rd1;
            //    crystalReportViewer1.Refresh();
            //}
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
namespace Pinnacle.Hostel
{
    public partial class HostelGatePass : Form, ToolStripAccess
    {
        private static HostelGatePass _instance;

        public HostelGatePass()
        {
            InitializeComponent();
            panelprint.Hide();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            btnhostelsave.Focus();
            DateTime dateForButton = DateTime.Now;
            frmdate.Value = dateForButton;
            btnhostelsave.Visible = false;
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }

        public void ReadOnlys()
        {

        }
        ListView listfilter = new ListView();
        public static HostelGatePass Instance
        {
            get { if (_instance == null) _instance = new HostelGatePass(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
        Models.Device dev = new Models.Device(); byte[] qrbytes; byte[] bytes;



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


        public void News()
        {
            GridLoad(); empty();
        }

        private void reason()
        {
            try
            {
                string sel3 = " SELECT  C.ASPTBLREASONMASID,C.REASON  FROM  GTCOMPMAST A JOIN  asptblusermas B ON A.GTCOMPMASTID= B.COMPCODE   JOIN ASPTBLREASONMAS C ON C.COMPCODE=A.GTCOMPMASTID     WHERE C.ACTIVE='T'   AND A.COMPCODE='" + Class.Users.HCompcode + "'   AND B.USERNAME='" + Class.Users.HUserName + "'";
                DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLREASONMAS");
                DataTable dt3 = ds3.Tables["ASPTBLREASONMAS"];
                if (dt3.Rows.Count > 0)
                {
                    comboreason.DisplayMember = "REASON";
                    comboreason.ValueMember = "ASPTBLREASONMASID";
                    comboreason.DataSource = dt3;


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
        }

        private void HostelGatePass_Load(object sender, EventArgs e)
        {

            reason();
            frmdate.Text = DateTime.Now.ToShortDateString(); todate.Text = DateTime.Now.ToShortDateString();
            btnhostelsave.Focus(); GridLoad();
        }


        private void LvLogs_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                empty();
                if (lvLogs.Items.Count > 0)
                {

                    pictureempimage.Image = null; bytes = null;
                    txtempid.Text = Convert.ToString(lvLogs.SelectedItems[0].SubItems[1].Text);
                    DataTable dt = new DataTable();
                    
                        dt.Rows.Clear();
                        string sel0 = "SELECT A.ASPTBLHOSTELGATEPASSID,B.COMPCODE, D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,   B.PHONENO AS  CONTACTNO, A.MANUALTIME,G.REASON,A.PERMISSIONHRS,A.SYSTEMDATE,A.SYSTEMTIME,A.OUTTIME,A.INTIME,I.QRCODE, I.EMPIMAGE,A.REMARKS, A.IPADDRESS, A.IPADDRESS1,    A.NATIVE  FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE AND C.HREMPLOYMASTID=A.EMPNAME     JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO    AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE    AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON       JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID    WHERE     A.ASPTBLHOSTELGATEPASSID=" + txtempid.Text + "  ORDER BY A.ASPTBLHOSTELGATEPASSID desc";

                        DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLHOSTELGATEPASS");
                        dt = ds0.Tables["ASPTBLHOSTELGATEPASS"];

                    if (dt == null)
                    {
                        string sel1 = "SELECT  A.ASPTBLHOSTELGATEPASSID,B.COMPCODE,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.DISPNAME AS DEPARTMENT,B.PHONENO AS  CONTACTNO,''AS HOSTELNAME,'' AS HOSTELBLOCK, '' AS HOSTELROOM,A.MANUALTIME,G.REASON,A.PERMISSIONHRS,A.SYSTEMDATE,A.SYSTEMTIME,A.OUTTIME,A.INTIME,I.QRCODE, I.EMPIMAGE,A.REMARKS, A.IPADDRESS, A.IPADDRESS1,    A.NATIVE FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE           JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE     JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID         AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO       JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT         JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON       LEFT OUTER JOIN ASPTBLEMP I ON I.COMPCODE=B.GTCOMPMASTID AND I.IDCARDNO=A.IDCARDNO WHERE A.ASPTBLHOSTELGATEPASSID=" + txtempid.Text;
                        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
                        dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
                    }



                    txtempid.Text = Convert.ToString(dt.Rows[0]["ASPTBLHOSTELGATEPASSID"].ToString());
                    combo_compcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                    txtidcardno.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                    comboempname.Text = Convert.ToString(dt.Rows[0]["EMPNAME"].ToString());
                    combo_dept.Text = Convert.ToString(dt.Rows[0]["DEPARTMENT"].ToString());
                    combohostel.Text = Convert.ToString(dt.Rows[0]["HOSTELNAME"].ToString());
                    combohostelblock.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                    combohostelroom.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                    txtmanualTime.Text = Convert.ToDateTime(dt.Rows[0]["MANUALTIME"].ToString()).ToString("HH:mm:ss");
                    comboreason.Text = Convert.ToString(dt.Rows[0]["REASON"].ToString());
                    txtpermissionhrs.Text = Convert.ToDateTime(dt.Rows[0]["PERMISSIONHRS"].ToString()).ToString("HH:mm:ss");
                    txtsysdate.Text = Convert.ToString(dt.Rows[0]["SYSTEMDATE"].ToString());
                    txtsystime.Text = Convert.ToString(dt.Rows[0]["SYSTEMTIME"].ToString());
                    txtoutime.Text = Convert.ToString(dt.Rows[0]["OUTTIME"].ToString());
                    txtintime.Text = Convert.ToString(dt.Rows[0]["INTIME"].ToString());
                    txtRemarks.Text = Convert.ToString(dt.Rows[0]["REMARKS"].ToString());
                    txtcontactno.Text= Convert.ToString(dt.Rows[0]["CONTACTNO"].ToString());
                    QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                    var mydata = qc.CreateQrCode(txtempid.Text, QRCoder.QRCodeGenerator.ECCLevel.L);
                    var code = new QRCoder.QRCode(mydata);
                    qrbytes = Encoding.ASCII.GetBytes(txtempid.Text);
                    pictureBox1.Image = code.GetGraphic(50, Color.Black, Color.White, true);
                    if (dt.Rows[0]["EMPIMAGE"].ToString() != "")
                    {

                        bytes = (byte[])dt.Rows[0]["EMPIMAGE"];
                        Image img = Models.Device.ByteArrayToImage(bytes);
                        pictureempimage.Image = img;


                    }
                    else
                    {
                        pictureempimage.Image = Pinnacle.Properties.Resources.close_image1;
                    }
                    if (dt.Rows[0]["NATIVE"].ToString() == "T")
                    {
                        checknative.Checked = true;
                    }
                    else
                    {
                        checknative.Checked = false;
                    }
                    panelprint.Hide();
                    btnhostelsave.Enabled = false;

                    comboreason.Enabled = false;
                    txtmanualTime.Enabled = false;
                    txtpermissionhrs.Enabled = false;

                    // txtintime.Enabled = false;txtoutime.Enabled = false;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Btnsaves_Click(object sender, EventArgs e)
        {

        }

        private void empty()
        {
            txtempid.Text = ""; pictureempimage.Image = null; pictureBox1.Image = null;
            combo_compcode.Text = "";
            txtidcardno.Text = "";txtcontactno.Text = "";
            combohostel.Text = "";
            combohostelblock.Text = ""; btnhostelsave.Visible = false;
            combohostelroom.Text = "";
            comboempname.Text = "";
            combo_dept.Text = ""; txtdept.Text = ""; txthostelblock.Text = ""; txthostelroom.Text = "";

            txtsysdate.Text = "";
            txtsystime.Text = "";
            pictureBox1.Image = null;
            txtcompcode.Text = "";
            txtempname.Text = "";
            txtdept.Text = "";
            txthostelblock.Text = "";
            txthostelroom.Text = "";
            comboreason.Text = ""; comboreason.SelectedIndex = -1;
            txtmanualTime.Text = "";
            btnhostelsave.Enabled = true;
            txtintime.Text = ""; txtoutime.Text = ""; txtRemarks.Text = "";
            comboreason.Enabled = true;
            txtmanualTime.Enabled = true; checknative.Checked = false;
            //lbloutime.Visible = false;
            //lblintime.Visible = false;
            //txtintime.Visible = false;
            //txtoutime.Visible = false;pictureempimage
            pictureempimage.Image = null; txtRemarks.Text = "";

            txtpermissionhrs.Enabled = true; panelprint.Hide();
            panelprint.Refresh(); butGetData.Visible = true;
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
        }

        private void Butcancel_Click(object sender, EventArgs e)
        {
            panelprint.Hide();
        }

        public void Prints()
        {
            string sel1 = " SELECT  MAX(A.ASPTBLHOSTELGATEPASSID) ID  FROM ASPTBLHOSTELGATEPASS A";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
            DataTable dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
            if (dt.Rows.Count > 0)
            {
                DataTable dt2 = new DataTable();
                string sel2 = "SELECT '" + System.DateTime.Now.Year + "' AS FINYEAR,H.DESIGNATION AS DESIGN,B.PHONENO  || ',' || B.FAXNO AS CONTACTAGF,A.ASPTBLHOSTELGATEPASSID AS TOKENNO ,B.COMPCODE,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT,A.CONTACTNO,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,G.REASON,A.PERMISSIONHRS AS PERHRS,A.SYSTEMDATE ,A.QRCODE,I.EMPIMAGE  FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON     JOIN GTDESIGNATIONMAST H ON H.GTDESIGNATIONMASTID=D.DESIGNATION LEFT OUTER JOIN ASPTBLEMP I ON I.COMPCODE=B.GTCOMPMASTID AND I.IDCARDNO=A.IDCARDNO  WHERE A.ASPTBLHOSTELGATEPASSID=" + Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                if (dt2.Rows.Count <= 0)
                {                  
                    string sel3 = "SELECT '"+System.DateTime.Now.Year+ "' AS FINYEAR,H.DESIGNATION AS DESIGN,B.PHONENO  || ',' || B.FAXNO AS CONTACTAGF, A.ASPTBLHOSTELGATEPASSID  AS TOKENNO,B.COMPCODE,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT,A.CONTACTNO ,G.REASON,A.PERMISSIONHRS,A.SYSTEMDATE,A.QRCODE,I.EMPIMAGE   FROM ASPTBLHOSTELGATEPASS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE    JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO      JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN GTDESIGNATIONMAST H ON H.GTDESIGNATIONMASTID=D.DESIGNATION  LEFT OUTER JOIN ASPTBLEMP I ON I.COMPCODE=B.GTCOMPMASTID AND I.IDCARDNO=A.IDCARDNO     WHERE   A.ASPTBLHOSTELGATEPASSID=" + Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLHOSTELGATEPASS");
                    dt2 = ds3.Tables["ASPTBLHOSTELGATEPASS"];
                }

                string IDD = "TOKENNO: " + Convert.ToString(dt2.Rows[0]["TOKENNO"].ToString()) + ",\nIDCARD : " + Convert.ToString(dt2.Rows[0]["IDCARDNO"].ToString()) + ",\nNAME   : " + Convert.ToString(dt2.Rows[0]["EMPNAME"].ToString());
                PictureBox picturegrcode = new PictureBox();               
                MemoryStream stream = new MemoryStream();
                QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                var mydata1 = qc.CreateQrCode(IDD, QRCoder.QRCodeGenerator.ECCLevel.L);
                var code1 = new QRCoder.QRCode(mydata1);
                picturegrcode.Image = code1.GetGraphic(50, Color.Black, Color.White, true);
                picturegrcode.Image.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
                qrbytes = stream.ToArray();
                dt2.Rows[0]["QRCODE"] = qrbytes;
                crystalReportViewer1.ReportSource = null;
                CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();

               
                if (combohostel.Text == "MENS HOTEL" || combohostel.Text == "WORKING GENTS HOSTEL" || combohostel.Text == "WOMENS HOSTEL" || combohostel.Text == "GENTS STAFF HOSTEL" || combohostel.Text == "BOYS HOSTEL")
                {

                    reportdocument.Load(Application.StartupPath + "\\Report\\AGF\\HostelReport.rpt");
                    reportdocument.Database.Tables["DataTable1"].SetDataSource(dt2);
                }
                if (combohostel.Text != "MENS HOTEL" || combohostel.Text != "WORKING GENTS HOSTEL" || combohostel.Text != "WOMENS HOSTEL" || combohostel.Text != "GENTS STAFF HOSTEL" || combohostel.Text != "BOYS HOSTEL")
                {

                    reportdocument.Load(Application.StartupPath + "\\Report\\AGF\\OutPassReport.rpt");
                    reportdocument.Database.Tables["DataTable1"].SetDataSource(dt2);
                }
                crystalReportViewer1.ReportSource = reportdocument;
                crystalReportViewer1.Refresh();
                reportdocument.PrintOptions.PrinterName = printDialog1.PrinterSettings.PrinterName;
                reportdocument.PrintToPrinter(printDialog1.PrinterSettings.Copies, printDialog1.PrinterSettings.Collate, printDialog1.PrinterSettings.FromPage, printDialog1.PrinterSettings.ToPage);

               
                panelprint.Hide();
                panelprint.Refresh();
                if (bIsConnected == true)
                {
                    int idwErrorCode = 0;
                    int iDataFlag = 1;
                    if (axCZKEM1.ClearData(iMachineNumber, iDataFlag))
                    {
                        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                    }
                    else
                    {
                        axCZKEM1.GetLastError(ref idwErrorCode);
                    }
                    axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
                }
                btnhostelsave.Focus();
            }

            
            panelprint.Visible = false;
            butprint1.Focus();
            btnhostelsave.Visible = false;
            butGetData.Visible = true;
        }



        private void Butprint_Click(object sender, EventArgs e)
        {

            //if (printDialog1.ShowDialog() == DialogResult.OK)

            //else
            //{
            //    panelprint.Hide();
            //    panelprint.Refresh();
            //}

        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString(lblgatepass.Text, new Font("Calibri", 12, FontStyle.Bold), Brushes.Black, 82, 16);
            e.Graphics.DrawString(lblqrcode.Text, new Font("Calibri", 7, FontStyle.Bold), Brushes.Black, 0, 37);
            e.Graphics.DrawImage(pictureBox2.Image, 168, 70, pictureBox2.Width, pictureBox2.Height);
            // e.Graphics.DrawImage(panelprint.BorderStyle=BorderStyle.FixedSingle, Brushes.DarkBlue, 530, 3);
            // e.Graphics.DrawImage(pictureBox1.Image, 160, 256, pictureBox1.Width, pictureBox1.Height);

            //  e.Graphics.DrawString(lblempsign.Text,  new Font("Arial", 8, FontStyle.Regular), Brushes.DarkBlue, 14, 374);
            //  e.Graphics.DrawString(lblwardensign.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.DarkBlue, 92, 374);
            //e.Graphics.DrawString(lblsecuritysing.Text, new Font("Arial", 8, FontStyle.Regular), Brushes.DarkBlue, 183, 374);

            // e.Graphics.DrawString(rlblcompcode.Text.ToLower(), new Font("Arial", 8, FontStyle.Regular), Brushes.DarkBlue, 14, 400);
        }






        private void Btnhostelsave_Click(object sender, EventArgs e)
        {
            try
            {


                if (txtidcardno.Text != "")
                {
                    panelprint.Hide();
                    panelprint.Refresh();

                    if (Convert.ToInt32("0" + comboreason.SelectedValue) >= 1 && txtpermissionhrs.Text != "" && txtmanualTime.Text != "")
                    {
                        string native = "";
                        if (checknative.Checked == true) { native = "T"; } else { native = "F"; }

                        DateTime modified = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy"));
                        DateTime CreatedOn = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
                        string ins = "INSERT INTO ASPTBLHOSTELGATEPASS(COMPCODE,  IDCARDNO,  EMPNAME ,  DEPARTMENT,  HOSTELNAME, HOSTELBLOCK,HOSTELROOM,CONTACTNO, SYSTEMDATE,  SYSTEMTIME ,REASON,MANUALTIME,USERNAME,  MODIFIED,  CREATEDON,  IPADDRESS1,IPADDRESS,PERMISSIONHRS,Remarks,NATIVE)VALUES(" + txtcompcode.Text + ",'" + txtidcardno.Text + "','" + txtempname.Text + "','" + txtdept.Text + "' ,'" + Class.Users.HostelName + "' ,'" + txthostelblock.Text + "' ,'" + txthostelroom.Text + "' ,'" + txtcontactno.Text + "','" + txtsysdate.Text + "','" + txtsystime.Text + "','" + comboreason.SelectedValue + "' ,'" + Convert.ToDateTime(txtmanualTime.Text) + "'," + Class.Users.USERID + ",to_date('" + Convert.ToDateTime(modified).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),to_date('" + Convert.ToDateTime(CreatedOn.ToString()) + "', 'dd/MM/yyyy hh24:MI:SS'),'" + txtipaddress.Text + "','" + Class.Users.IPADDRESS + "', '" + Convert.ToDateTime(txtpermissionhrs.Text) + "','" + txtRemarks.Text + "','" + native + "')";
                        Utility.ExecuteNonQuery(ins);
                        empty(); GridLoad();
                        Prints();
                    }
                    else
                    {
                        MessageBox.Show("pls Select Mandatary Fields", " Gate Pass ", MessageBoxButtons.OK, MessageBoxIcon.Error); comboreason.Focus(); comboreason.BackColor = System.Drawing.Color.Red;
                        return;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Gate Pass Cancelled    " + ex.Message.ToString() + "", " Gate Pass ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return;
            }
        }

        private void Txthostelgatesearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txthostelgatesearch.Text.Length > 1)
                {
                    lvLogs.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txthostelgatesearch.Text) || listfilter.Items[item0].SubItems[3].ToString().Contains(txthostelgatesearch.Text))
                        {


                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
                            if (item0 % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            lvLogs.Items.Add(list);


                        }
                        item0++;
                    }

                }
                else
                {
                    ListView ll = new ListView(); item0 = 1;
                    lvLogs.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {

                        if (item0 % 2 == 0)
                        {
                            item.BackColor = Color.White;
                        }
                        else
                        {
                            item.BackColor = Color.WhiteSmoke;
                        }
                        this.lvLogs.Items.Add((ListViewItem)item.Clone());
                        item0++;
                    }
                    lblattcount.Text = "Total Count: " + lvLogs.Items.Count;
                }


            }
            catch (Exception ex)
            {

            }

        }

        private void MenuRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ListViewRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Butview_Click(sender, e);
        }

        private void Pictureempimage_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    bytes = null;
            //    PictureBox p = sender as PictureBox;
            //    if (p != null)
            //    {


            //            p.Image = new Bitmap(pictureempimage.Image);
            //            bytes = Models.Device.ImageToByteArray(p);


            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }
        public void GridLoad()
        {
            iGLCount = 1; listfilter.Items.Clear(); lvLogs.Items.Clear();
            DataTable dt = new DataTable();
            if (Class.Users.HostelName == "WOMENS HOSTEL" || Class.Users.HostelName == "WORKING GENTS HOSTEL")
            {
                string sel1 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM, B.PHONENO || B.FAXNO AS CONTACTNO, A.SYSTEMDATE,A.MODIFIED   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID   JOIN ASPTBLMACHINEMAS I ON I.COMPCODE=B.GTCOMPMASTID AND I.COMPCODE=A.COMPCODE AND I.HOSTELNAME=A.HOSTELNAME  WHERE    A.HOSTELNAME='" + Class.Users.HostelName + "' AND A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')  ORDER BY A.ASPTBLHOSTELGATEPASSID desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
                dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
            }
            //if (Class.Users.HostelName == "WORKING GENTS HOSTEL")
            //{
            //    dt.Rows.Clear();
            //    string sel1 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,  B.PHONENO || B.FAXNO AS CONTACTNO, A.SYSTEMDATE,A.MODIFIED   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  JOIN ASPTBLMACHINEMAS I ON I.COMPCODE=B.GTCOMPMASTID AND I.COMPCODE=A.COMPCODE AND I.HOSTELNAME=A.HOSTELNAME  WHERE  A.HOSTELNAME='" + Class.Users.HostelName + "'  AND A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')  ORDER BY A.ASPTBLHOSTELGATEPASSID desc";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
            //    dt = ds.Tables["ASPTBLHOSTELGATEPASS"];

            //}
            //if (Class.Users.HostelName == "AGF" || Class.Users.HostelName == "FLF" || Class.Users.HostelName == "FLFD" || Class.Users.HostelName == "AGFM" || Class.Users.HostelName == "AGFC" || Class.Users.HostelName == "AGFMGII")
            //{
            if (Class.Users.HostelName != "WORKING GENTS HOSTEL" && Class.Users.HostelName != "WOMENS HOSTEL")
            {
                dt.Rows.Clear();
                string sel1 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT,'' AS HOSTELNAME,''AS HOSTELBLOCK, '' AS HOSTELROOM,  H.CONTACTNO, A.SYSTEMDATE,A.MODIFIED   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE     JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO   AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT          JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE  A.HOSTELNAME='" + Class.Users.HostelName + "'    AND A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')  ORDER BY A.ASPTBLHOSTELGATEPASSID desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
                dt = ds.Tables["ASPTBLHOSTELGATEPASS"];

            }
            if (dt.Rows.Count > 0)
            {

                foreach (DataRow myRow in dt.Rows)
                {
                    ListViewItem list = new ListViewItem();
                    list.Text = iGLCount.ToString();
                    list.SubItems.Add(myRow["ASPTBLHOSTELGATEPASSID"].ToString());
                    list.SubItems.Add(myRow["IDCARDNO"].ToString());
                    list.SubItems.Add(myRow["EMPNAME"].ToString());
                    list.SubItems.Add(myRow["DEPARTMENT"].ToString());
                    list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                    list.SubItems.Add(myRow["HOSTELROOM"].ToString());
                    list.SubItems.Add(myRow["HOSTELROOM"].ToString());
                    list.SubItems.Add(myRow["CONTACTNO"].ToString());
                    list.SubItems.Add(myRow["SYSTEMDATE"].ToString());
                    list.SubItems.Add(myRow["MODIFIED"].ToString());
                    this.listfilter.Items.Add((ListViewItem)list.Clone());
                    if (iGLCount % 2 == 0)
                    {
                        list.BackColor = Color.White;
                    }
                    else
                    {
                        list.BackColor = Color.WhiteSmoke;
                    }
                    lvLogs.Items.Add(list);
                    iGLCount++;
                }
                lblattcount.Text = "Total Count    :" + lvLogs.Items.Count;
            }
        }
        private void Butview_Click(object sender, EventArgs e)
        {
            try
            {
                listfilter.Items.Clear(); lvLogs.Items.Clear();
                iGLCount = 1;
                DataTable dt = new DataTable();
               

                     string sel1 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,   B.PHONENO AS  CONTACTNO, A.SYSTEMDATE,A.MODIFIED  FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE AND C.HREMPLOYMASTID=A.EMPNAME     JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO    AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE    AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON       JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID        WHERE  A.HOSTELNAME='" + Class.Users.HostelName + "'  AND A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy') ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
                    dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
                if (dt == null || dt.Rows.Count <= 0)
                {


                    string sel0 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.MNNAME1 AS DEPARTMENT, H.CONTACTNO, A.SYSTEMDATE,A.MODIFIED   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE     JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO    AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT       JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID    WHERE  A.HOSTELNAME='" + Class.Users.HostelName + "'  AND A.MODIFIED= TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')  ORDER BY A.ASPTBLHOSTELGATEPASSID";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel0, "ASPTBLHOSTELGATEPASS");
                    dt = ds1.Tables["ASPTBLHOSTELGATEPASS"];
                }
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();

                        list.Text = iGLCount.ToString();
                        list.SubItems.Add(myRow["ASPTBLHOSTELGATEPASSID"].ToString());
                        list.SubItems.Add(myRow["IDCARDNO"].ToString());
                        list.SubItems.Add(myRow["EMPNAME"].ToString());
                        list.SubItems.Add(myRow["DEPARTMENT"].ToString());
                        if (Class.Users.HostelName == "WORKING GENTS HOSTEL" || Class.Users.HostelName == "WOMENS HOSTEL")
                        {
                            list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                            list.SubItems.Add(myRow["HOSTELROOM"].ToString());
                            list.SubItems.Add(myRow["HOSTELROOM"].ToString());
                        } 
                        list.SubItems.Add(myRow["CONTACTNO"].ToString());
                        list.SubItems.Add(myRow["SYSTEMDATE"].ToString());
                        list.SubItems.Add(myRow["MODIFIED"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        if (iGLCount % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        lvLogs.Items.Add(list);
                        iGLCount++;
                    }
                    lblattcount.Text = "Total Count    :" + lvLogs.Items.Count;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            empty();
        }

        private void ViewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string ccode = "";


            try
            {
                ccode = Class.Users.HCompcode;
                lvLogs1.Items.Clear();

                int k = 0;
                iIndex = 0;


                iGLCount = 0;

                string ip = "";


                txtipaddress.Text = "";

                // DataTable dt = Utility.SQLQuery("SELECT DISTINCT  B.HRMACIPENTRYDETID,B.MACIP  FROM ASPTBLMACHINEMAS A JOIN HRMACIPENTRYDET B ON B.HRMACIPENTRYDETID=A.IPADDRESS JOIN asptblusermas C ON C.userid=A.WARDENNAME JOIN GTCOMPMAST E ON   E.GTCOMPMASTID = A.COMPCODE AND E.GTCOMPMASTID=C.COMPCODE WHERE  A.ACTIVE='T' AND E.COMPCODE = '" + Class.Users.HCompcode + "' AND C.USERNAME='" + Class.Users.HUserName + "'");
                DataTable dt = Utility.SQLQuery("SELECT A.ASPTBLMACIPID AS HRMACIPENTRYDETID, A.MACIP     FROM  ASPTBLMACIP   A JOIN GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE   JOIN ASPTBLUSERMAS  D ON D.COMPCODE = B.GTCOMPMASTID AND D.COMPCODE = A.COMPCODE AND D.USERNAME = A.USERNAME   WHERE  A.ACTIVE = 'T' AND B.COMPCODE = '" + Class.Users.HCompcode + "' AND D.USERNAME = '" + Class.Users.HUserName + "' UNION ALL SELECT DISTINCT B.HRMACIPENTRYDETID, B.MACIP  FROM ASPTBLMACHINEMAS A JOIN HRMACIPENTRYDET B ON B.HRMACIPENTRYDETID = A.IPADDRESS JOIN asptblusermas C ON C.userid = A.WARDENNAME JOIN GTCOMPMAST E ON   E.GTCOMPMASTID = A.COMPCODE AND E.GTCOMPMASTID = C.COMPCODE   AND B.CURMAC='YES' WHERE  A.ACTIVE = 'T' AND E.COMPCODE = '" + Class.Users.HCompcode + "' AND C.USERNAME = '" + Class.Users.HUserName + "'");
                int maxip = dt.Rows.Count;
                if (maxip == 0)
                {
                    MessageBox.Show("IP Address not assign this User.   : " + Class.Users.HUserName);
                }
                if (maxip == 1)
                {
                    int i = 0;
                    for (i = 0; i < maxip; i++)
                    {
                        bIsConnected = axCZKEM1.Connect_Net(dt.Rows[i]["MACIP"].ToString(), Convert.ToInt32(4370));
                        ip = dt.Rows[i]["MACIP"].ToString();
                        txtipaddress.Text = dt.Rows[i]["HRMACIPENTRYDETID"].ToString();
                        if (bIsConnected == true)
                        {
                            if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                            {
                                while (axCZKEM1.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode, out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                                {
                                    DateTime inputDate = new DateTime(idwYear, idwMonth, idwDay);
                                    if (Convert.ToDateTime(inputDate) >= frmdate.Value.Date)//&& Convert.ToDateTime(inputDate) <= todate.Value.Date.AddDays(1).AddTicks(-1)
                                    {

                                        iGLCount++;
                                        lvLogs1.Items.Add(iGLCount.ToString());
                                        lvLogs1.Items[iIndex].SubItems.Add(sdwEnrollNumber);
                                        iIndex++;
                                    }
                                }
                            }
                            else
                            {
                                Cursor = Cursors.Default; butGetData.Visible = true; btnhostelsave.Visible = false;
                                MessageBox.Show("No Data Found this Machine...." + ip.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device    

                        }
                        else
                        {
                            axCZKEM1.GetLastError(ref idwErrorCode);
                            Cursor = Cursors.Default; butGetData.Visible = true; btnhostelsave.Visible = false;
                            MessageBox.Show("Unable to connect the device , ErrorCode=" + idwErrorCode.ToString() + "---IP-----" + dt.Rows[i]["MACIP"].ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Only one IPAddress assign this User.pls check Machine Master   : " + Class.Users.HUserName);
                    return;
                }
            }
            catch (Exception ex)
            {
                butGetData.Visible = true; btnhostelsave.Visible = false;
                MessageBox.Show(ex.Message.ToString());

            }

            try
            {
                empty();

                if (lvLogs1.Items.Count >= 1)
                {

                    var idd = lvLogs1.Items[lvLogs1.Items.Count - 1].SubItems[1].Text;
                    bytes = null;
                    pictureempimage.Image = null;
                    DataTable dt = new DataTable();
                    string sel1 = "SELECT '" + System.DateTime.Now.Year + "' AS FINYEAR,C.HREMPLOYMASTID AS ASPTBLEMPID, B.COMPCODE ,B.GTCOMPMASTID,D.MIDCARD,C.FNAME,E.MNNAME1 as DISPNAME,E.GTDEPTDESGMASTID ,A.HOSTELNAME,A.BLOCKFLOOR,A.ROOMNO,A.IDCARDNO,B.PHONENO || B.FAXNO AS CONTACTNO FROM HOSTELLIVEDATA A JOIN GTCOMPMAST B ON B.COMPCODE= A.COMPCODE   JOIN HREMPLOYMAST C ON C.COMPCODE = B.GTCOMPMASTID AND C.IDCARDNO = A.IDCARDNO   JOIN HREMPLOYDETAILS D ON C.HREMPLOYMASTID = D.HREMPLOYMASTID AND D.HOSTEL='YES' JOIN GTDEPTDESGMAST E ON E.GTDEPTDESGMASTID = D.DEPTNAME   WHERE D.MIDCARD=" + idd.ToString();
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HOSTELLIVEDATA");
                    dt = ds.Tables["HOSTELLIVEDATA"];
                    if (dt.Rows.Count <= 0)
                    {
                        string sel2 = "SELECT '" + System.DateTime.Now.Year + "' AS FINYEAR,C.HREMPLOYMASTID AS ASPTBLEMPID, B.COMPCODE ,B.GTCOMPMASTID,D.MIDCARD,C.FNAME,E.MNNAME1 as DISPNAME,E.GTDEPTDESGMASTID,'" + Class.Users.HCompcode + "' as HOSTELNAME,'0' as BLOCKFLOOR,'0' AS ROOMNO,C.IDCARDNO,B.PHONENO || B.FAXNO AS CONTACTNO  FROM  GTCOMPMAST B JOIN  HREMPLOYMAST C ON C.COMPCODE = B.GTCOMPMASTID    JOIN HREMPLOYDETAILS D ON C.HREMPLOYMASTID = D.HREMPLOYMASTID   JOIN GTDEPTDESGMAST E ON E.GTDEPTDESGMASTID = D.DEPTNAME     WHERE D.MIDCARD='" + idd.ToString()+"' " ;
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "HREMPLOYDETAILS");
                        dt = ds2.Tables["HREMPLOYDETAILS"];
                    }

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("This IDCardno  '" + idd + "' is empty in HostelMaster", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {

                        txtempid.Text = Convert.ToString(dt.Rows[0]["ASPTBLEMPID"].ToString());
                        txtcompcode.Text = Convert.ToString(dt.Rows[0]["GTCOMPMASTID"].ToString());
                        combo_compcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                        txtidcardno.Text = Convert.ToString(dt.Rows[0]["MIDCARD"].ToString());
                        comboempname.Text = Convert.ToString(dt.Rows[0]["FNAME"].ToString());
                        txtempname.Text = Convert.ToString(dt.Rows[0]["ASPTBLEMPID"].ToString());
                        combo_dept.Text = Convert.ToString(dt.Rows[0]["DISPNAME"].ToString());
                        txtdept.Text = Convert.ToString(dt.Rows[0]["GTDEPTDESGMASTID"].ToString());
                        txthostelroom.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                        combohostel.Text = Convert.ToString(dt.Rows[0]["HOSTELNAME"].ToString());
                        combohostelroom.Text = Convert.ToString(dt.Rows[0]["ROOMNO"].ToString());
                        txthostelblock.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                        combohostelblock.Text = Convert.ToString(dt.Rows[0]["BLOCKFLOOR"].ToString());
                        txtcontactno.Text = Convert.ToString(dt.Rows[0]["CONTACTNO"].ToString());
                        txtsysdate.Text = Convert.ToString(Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss tt")));
                        txtsystime.Text = Convert.ToString(System.DateTime.Now.ToString("HH:mm:ss tt"));
                        txtmanualTime.Text = Convert.ToString(DateTime.Now.ToString("HH:mm:ss"));
                        string IDD = "TOKENNO: " + Convert.ToString(dt.Rows[0]["ASPTBLEMPID"].ToString()) + ",\nIDCARD : " + Convert.ToString(dt.Rows[0]["MIDCARD"].ToString()) + ",\nNAME   : " + Convert.ToString(dt.Rows[0]["FNAME"].ToString());

                        QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                        var mydata1 = qc.CreateQrCode(IDD, QRCoder.QRCodeGenerator.ECCLevel.L);
                        var code1 = new QRCoder.QRCode(mydata1);
                        pictureBox1.Image = code1.GetGraphic(50, Color.Black, Color.White, true);
                        string sel2 = ""; DataTable dt2 = new DataTable();
                        sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,D.MIDCARD ,C.FNAME ,E.MNNAME1 AS DISPNAME,substr(A.SYSTEMDATE,1,10) AS CONTACTNO,F.HOSTELNAME,F.BLOCKFLOOR , F.ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME, A.REMARKS   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID   WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND A.INTIME IS NULL and D.MIDCARD='" + dt.Rows[0]["MIDCARD"].ToString() + "'  AND  A.MODIFIED >= TO_DATE('" + System.DateTime.Now.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')   ORDER BY 1";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                        dt2 = ds2.Tables["ASPTBLHOSTELGATEPASS"];
                        if (dt2.Rows.Count <= 0)
                        {
                           
                             sel2 = "SELECT to_char(A.ASPTBLHOSTELGATEPASSID) as ASPTBLHOSTELGATEPASSID,B.COMPNAME AS COMPCODE,D.MIDCARD ,C.FNAME ,E.MNNAME1 AS DISPNAME,  substr(A.SYSTEMDATE,1,10) AS CONTACTNO,'" + Class.Users.HCompcode + "' as HOSTELNAME,''as BLOCKFLOOR , ''as ROOMNO,G.REASON,A.PERMISSIONHRS,A.INTIME,A.OUTTIME, A.REMARKS     FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE       JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT  JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON  WHERE  A.INTIME IS NULL and D.MIDCARD='" + dt.Rows[0]["MIDCARD"].ToString() + "'  AND  A.MODIFIED >= TO_DATE('" + System.DateTime.Now.ToString("dd-MM-yyyy") + "','dd-MM-yyyy')   ORDER BY 1";
                            DataSet ds3 = Utility.ExecuteSelectQuery(sel2, "ASPTBLHOSTELGATEPASS");
                            dt2 = ds3.Tables["ASPTBLHOSTELGATEPASS"];
                        }
                        if (dt2.Rows.Count == 0)
                        {
                            butGetData.Visible = false; btnhostelsave.Visible = true;
                        }
                        else
                        {
                            MessageBox.Show("This Employee IDCard: -   '" + dt2.Rows[0]["MIDCARD"].ToString() + "====" + dt2.Rows[0]["FNAME"].ToString() + "'  not Closed Privious Pass.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lvLogs1.Items.Clear(); btnhostelsave.Visible = false;
                            empty();
                            return;
                        }
                    }
                }
                else
                {
                    butGetData.Visible = true; btnhostelsave.Visible = false;
                    MessageBox.Show("No Data Found in Finger Print Machine");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Cursor = Cursors.Default;
        }

        private void Comboreason_MouseHover(object sender, EventArgs e)
        {
            comboreason.BackColor = Color.White;
        }



        private void ReasonMasterRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            reason();
        }

        private void MenuRefreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {

        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }

        public void Saves()
        {
            if (txtempid.Text != "" && txtRemarks.Text != "")
            {
                string native = "";
                if (checknative.Checked == true) { native = "T"; } else { native = "F"; }
                if (txtintime.Text == "" || txtoutime.Text == "")
                {
                    DialogResult result = MessageBox.Show("This is Administrator Issue.Do You want to save this Record  'IDCardNo :-'" + txtempid.Text + "'' ,Name:-  '" + comboempname.Text + "'' this Record ??", "Gate Pass", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

                    if (result.Equals(DialogResult.OK))
                    {
                        DateTime endtime = Convert.ToDateTime(System.DateTime.Now.ToString());
                        DateTime statetime = Convert.ToDateTime(System.DateTime.Now.ToString());
                        var statetime1 = Convert.ToDateTime(System.DateTime.Now.ToString());
                        TimeSpan differ = endtime.Subtract(statetime);
                        TimeSpan differ1 = endtime - statetime1;

                        string ins = "update  ASPTBLHOSTELGATEPASS set OUTTIME='" + System.DateTime.Now.ToString() + "', INTIME='" + System.DateTime.Now.ToString() + "',NATIVE='" + native + "', TIMEDIFF='" + differ1.ToString() + "',REMARKS='UserName" + Class.Users.HUserName + "CompCode"+ Class.Users.HCompcode+ " "+ txtRemarks.Text + "',IPADDRESS='" + Class.Users.IPADDRESS + "' WHERE ASPTBLHOSTELGATEPASSID=" + txtempid.Text;
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("This is Record Saved Successfully.  'IDCardNo :-'" + txtempid.Text + "'' ,Name:-  '" + comboempname.Text + "'' this Record ??", "Gate Pass", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        empty();
                    }

                    else
                    {

                        MessageBox.Show("Invalid.  'IDCardNo :-'" + txtempid.Text + "'' ,Name:-  '" + comboempname.Text + "'' this Record ??", "Gate Pass", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }
                else
                {
                    if (txtempid.Text != "" && txtintime.Text != "" && txtoutime.Text != "")
                    {
                        MessageBox.Show("Invalid.", "Gate Pass", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid.  pls go to Your Administrator.??", "Gate Pass", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        empty();
                    }
                }
            }
            else
            {
                MessageBox.Show("Pls Enter Remarks Field    : ", "Employee Name :-" + comboempname.Text + " Remarks Field Empty", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }



        private void butprintcancel_Click(object sender, EventArgs e)
        {
            this.panelprint.Visible = false;
        }

        private void butGetData_Click(object sender, EventArgs e)
        {
            empty(); Cursor = Cursors.WaitCursor;
            if (txtidcardno.Text == "")
            {

                ViewToolStripMenuItem_Click(sender, e);

            }
            Cursor = Cursors.Default;
        }

        private void comboreason_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboreason.Text == "NATIVE")
            {
                checknative.Checked = true;
            }
            else
            {
                checknative.Checked = false;
            }
        }

        private void pictureempimage_MouseHover(object sender, EventArgs e)
        {
            pictureempimage.Height = 170;
            pictureempimage.Width = 180;
        }

        private void pictureempimage_MouseLeave(object sender, EventArgs e)
        {
            pictureempimage.Height = 128;
            pictureempimage.Width = 148;
        }

        public void Searchs()
        {

        }

        public void Deletes()
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
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void checkAGF_CheckedChanged(object sender, EventArgs e)
        {

        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

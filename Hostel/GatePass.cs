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
    public partial class GatePass : Form
    {
        private static GatePass _instance; 

        public GatePass()
        {
            InitializeComponent();
            label1.Text = Class.Users.ScreenName; panelprint.Hide();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            btnhostelsave.Focus();
            DateTime dateForButton = DateTime.Now;
            frmdate.Value = dateForButton;
           
        }
        ListView listfilter = new ListView();
        public static GatePass Instance
        {
            get { if (_instance == null) _instance = new GatePass(); GlobalVariables.CurrentForm = _instance; return _instance; }

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


        private void News_Click(object sender, EventArgs e)
        {
            empty();
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

            reason();// gridload();
            frmdate.Text = DateTime.Now.ToShortDateString(); todate.Text = DateTime.Now.ToShortDateString();
            btnhostelsave.Focus();
           // Btnhostelsave_Click(sender,e);
        }

      
        private void LvLogs_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (lvLogs.Items.Count > 0)
                {
                    Cursor = Cursors.WaitCursor;
                    pictureempimage.Image = null; bytes = null;
                    txtempid.Text = Convert.ToString(lvLogs.SelectedItems[0].SubItems[1].Text);

                    string sel1 = "SELECT  A.ASPTBLHOSTELGATEPASSID,B.COMPCODE,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.DISPNAME AS DEPARTMENT,H.CONTACTNO,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,A.MANUALTIME,G.REASON,A.PERMISSIONHRS,A.SYSTEMDATE,A.SYSTEMTIME,A.OUTTIME,A.INTIME,'' AS EMPIMAGE,A.REMARKS FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT AND E.COMPCODE=B.GTCOMPMASTID   JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE A.ASPTBLHOSTELGATEPASSID=" + txtempid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
                    DataTable dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
                    txtempid.Text = Convert.ToString(dt.Rows[0]["ASPTBLHOSTELGATEPASSID"].ToString());
                    combo_compcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                    txtidcardno.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                    comboempname.Text = Convert.ToString(dt.Rows[0]["EMPNAME"].ToString());
                    combo_dept.Text = Convert.ToString(dt.Rows[0]["DEPARTMENT"].ToString());
                    txtcontactno.Text = Convert.ToString(dt.Rows[0]["CONTACTNO"].ToString());
                    combohostel.Text = Convert.ToString(dt.Rows[0]["HOSTELNAME"].ToString());
                    combohostelblock.Text = Convert.ToString(dt.Rows[0]["HOSTELBLOCK"].ToString());
                    combohostelroom.Text = Convert.ToString(dt.Rows[0]["HOSTELROOM"].ToString());
                    txtmanualTime.Text = Convert.ToDateTime(dt.Rows[0]["MANUALTIME"].ToString()).ToString("hh:mm");
                    comboreason.Text = Convert.ToString(dt.Rows[0]["REASON"].ToString());
                    txtpermissionhrs.Text = Convert.ToDateTime(dt.Rows[0]["PERMISSIONHRS"].ToString()).ToString("hh:mm");
                    txtsysdate.Text = Convert.ToString(dt.Rows[0]["SYSTEMDATE"].ToString());
                    txtsystime.Text = Convert.ToString(dt.Rows[0]["SYSTEMTIME"].ToString());
                    txtoutime.Text = Convert.ToString(dt.Rows[0]["OUTTIME"].ToString());
                    txtintime.Text = Convert.ToString(dt.Rows[0]["INTIME"].ToString());
                    txtRemarks.Text = Convert.ToString(dt.Rows[0]["REMARKS"].ToString());
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

                    panelprint.Hide();
                 
                    comboreason.Enabled = false;
                    txtmanualTime.Enabled = false;
                    lbloutime.Visible = true;
                    lblintime.Visible = true;
                    txtintime.Visible = true;
                    txtoutime.Visible = true; txtpermissionhrs.Enabled = false;
                    Cursor = Cursors.Default;
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
            txtempid.Text = "";
            combo_compcode.Text = "";
            txtidcardno.Text = "";
            combohostel.Text = "";
            combohostelblock.Text = "";
            combohostelroom.Text = "";
            comboempname.Text = "";
            combo_dept.Text = "";
            txtcontactno.Text = "";
            txtsysdate.Text = "";
            txtsystime.Text = "";
            pictureBox1.Image = null;
            txtcompcode.Text = "";
            txtempname.Text = "";
            txtdept.Text = "";
            txthostelblock.Text = "";
            txthostelroom.Text = "";
            // comboreason.Text = "";
            txtmanualTime.Text = "";
            btnhostelsave.Enabled = true;

            comboreason.Enabled = true;
            txtmanualTime.Enabled = true;
            //lbloutime.Visible = false;
            //lblintime.Visible = false;
            //txtintime.Visible = false;
            //txtoutime.Visible = false;
            pictureempimage.Image = null; txtRemarks.Text = "";
            txtpermissionhrs.Enabled = true; panelprint.Hide(); panelprint.Refresh(); 
        }

        private void Butcancel_Click(object sender, EventArgs e)
        {
            panelprint.Hide();
        }

        private void Prints_Click(object sender, EventArgs e)
        {

            string sel1 = " SELECT  MAX(A.ASPTBLGATEPASSID) ID  FROM ASPTBLGATEPASS A";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLGATEPASS");
            DataTable dt = ds.Tables["ASPTBLGATEPASS"];
            if (dt.Rows.Count > 0)
            {
                string sel2 = "SELECT A.ASPTBLGATEPASSID,B.COMPCODE,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.DISPNAME AS DEPARTMENT,H.CONTACTNO AS CONTACT,G.REASON,A.PERMISSIONHRS,A.SYSTEMDATE   FROM ASPTBLGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO    JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT AND E.COMPCODE=B.GTCOMPMASTID JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID WHERE A.ASPTBLGATEPASSID=" + Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLGATEPASS");
                DataTable dt2 = ds2.Tables["ASPTBLGATEPASS"];
                string IDD = "TOKENNO: " + Convert.ToString(dt2.Rows[0]["ASPTBLGATEPASSID"].ToString()) + ",\nIDCARD : "+ Convert.ToString(dt2.Rows[0]["IDCARDNO"].ToString()) + ",\nNAME   : " + Convert.ToString(dt2.Rows[0]["EMPNAME"].ToString());
                Label lblheader2 = new Label();
                lblheader2.Name = "lblheader2";
                lblqrcode.Text = "-------------------------------------------------------------------------------------------\n" + "TOKENNO           : " + dt2.Rows[0]["ASPTBLGATEPASSID"].ToString() + "\n" + "COMPCODE        : " + dt2.Rows[0]["COMPCODE"].ToString() + "\n" + "IDCARDNO          : " + dt2.Rows[0]["IDCARDNO"].ToString() + "\n" + "EMP NAME         : " + dt2.Rows[0]["EMPNAME"].ToString() + "\n" + "DEPARTMENT     : " + dt2.Rows[0]["DEPARTMENT"].ToString() + "\n" + "CONTACT NO      : " + dt2.Rows[0]["CONTACT"].ToString() + "\n" + "REASON               : " + "" + "\n" + "PER HRS               : " + dt2.Rows[0]["PERMISSIONHRS"].ToString() + "\n" + "OUT TIME           : " + System.DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt") + "\n-------------------------------------------------------------------------------------------\n";
                lblheader2.Text = lblgatepass.Text + " - " + Class.Users.HCompcode + "\n" + lblqrcode.Text;
                QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                var mydata1 = qc.CreateQrCode(IDD, QRCoder.QRCodeGenerator.ECCLevel.L);
                var code1 = new QRCoder.QRCode(mydata1);
                pictureBox2.Image = code1.GetGraphic(5, Color.Black, Color.White, true);
                rlblcompcode.Text = "For   " + Class.Users.HCompName;


            }
          Butprint_Click(sender, e);
            //panelprint.Width = 320;
            //panelprint.Height = 431;
          
            panelprint.Visible=false;
            butprint1.Focus();
            Cursor = Cursors.Default;
        }



        private void Butprint_Click(object sender, EventArgs e)
        {

            //if (printDialog1.ShowDialog() == DialogResult.OK)
            try
            {

                //printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                //printDocument1.Print();
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
            catch(Exception ex)
            {

            }
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
                Cursor = Cursors.WaitCursor;
                ViewToolStripMenuItem_Click(sender, e);
                panelprint.Hide();
                    panelprint.Refresh();

                Cursor = Cursors.Default;




            }
            catch (Exception ex)
            {
               // MessageBox.Show("Gate Pass Cancelled    " + ex.Message.ToString() + "", " Gate Pass ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void Txthostelgatesearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txthostelgatesearch.Text.Length > 0)
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
                            lvLogs.Items.Add(list);


                        }
                        item0++;
                    }

                }
                else
                {
                    ListView ll = new ListView();
                    lvLogs.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {


                        this.lvLogs.Items.Add((ListViewItem)item.Clone());



                        item0++;
                    }

                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
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
        private void gridload()
        {
            iGLCount = 1; listfilter.Items.Clear(); lvLogs.Items.Clear();
            string sel1 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.DISPNAME AS DEPARTMENT,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,  H.CONTACTNO, A.SYSTEMDATE,A.MODIFIED   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT AND E.COMPCODE=B.GTCOMPMASTID  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE  B.COMPCODE='" + Class.Users.HCompcode + "'  AND  A.INTIME IS NULL OR  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')  ORDER BY 1";
            // string sel1 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.DISPNAME AS DEPARTMENT,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,  H.CONTACTNO, A.SYSTEMDATE,A.MODIFIED FROM ASPTBLHOSTELGATEPASS A  JOIN GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN HREMPLOYMAST    C ON C.COMPCODE = B.GTCOMPMASTID AND C.COMPCODE = A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID = C.HREMPLOYMASTID AND D.IDCARD = C.IDCARDNO AND D.DEPTNAME = A.DEPARTMENT AND D.MIDCARD = A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT AND E.COMPCODE = B.GTCOMPMASTID  JOIN HOSTELLIVEDATA F ON F.COMPCODE = B.COMPCODE AND F.IDCARDNO = C.IDCARDNO AND F.IDCARDNO = D.IDCARD AND F.HOSTELNAME = A.HOSTELNAME  AND D.IDCARD = F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID = A.REASON    JOIN HRECONTACTDETAILS H ON H.HREMPLOYMASTID = C.HREMPLOYMASTID ORDER BY A.ASPTBLHOSTELGATEPASSID DESC";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
            DataTable dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
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
                string sel1 = "SELECT A.ASPTBLHOSTELGATEPASSID,D.MIDCARD AS IDCARDNO,C.FNAME AS EMPNAME,E.DISPNAME AS DEPARTMENT,F.HOSTELNAME,F.BLOCKFLOOR AS HOSTELBLOCK, F.ROOMNO AS HOSTELROOM,  H.CONTACTNO, A.SYSTEMDATE,A.MODIFIED   FROM ASPTBLHOSTELGATEPASS A  JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN   HREMPLOYMAST    C ON C.COMPCODE=B.GTCOMPMASTID AND C.COMPCODE=A.COMPCODE   JOIN HREMPLOYDETAILS D ON D.HREMPLOYMASTID=C.HREMPLOYMASTID AND D.IDCARD=C.IDCARDNO AND  D.DEPTNAME=A.DEPARTMENT AND D.MIDCARD=A.IDCARDNO   JOIN GTDEPTDESGMAST  E ON E.GTDEPTDESGMASTID = A.DEPARTMENT AND E.COMPCODE=B.GTCOMPMASTID  JOIN HOSTELLIVEDATA F ON F.COMPCODE=B.COMPCODE AND F.IDCARDNO=C.IDCARDNO AND F.IDCARDNO=D.IDCARD AND F.HOSTELNAME=A.HOSTELNAME  AND D.IDCARD=F.IDCARDNO   JOIN ASPTBLREASONMAS G ON G.ASPTBLREASONMASID=A.REASON    JOIN  HRECONTACTDETAILS H ON H.HREMPLOYMASTID=C.HREMPLOYMASTID  WHERE  B.COMPCODE='" + Class.Users.HCompcode + "'  AND  A.INTIME IS NULL OR  A.MODIFIED between TO_DATE('" + frmdate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy') and TO_DATE('" + todate.Value.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')  ORDER BY 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELGATEPASS");
                DataTable dt = ds.Tables["ASPTBLHOSTELGATEPASS"];
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
                DataTable dt = Utility.SQLQuery("SELECT DISTINCT  B.HRMACIPENTRYDETID,B.MACIP  FROM ASPTBLMACHINEMAS A JOIN HRMACIPENTRYDET B ON B.HRMACIPENTRYDETID=A.IPADDRESS JOIN asptblusermas C ON C.userid=A.WARDENNAME JOIN GTCOMPMAST E ON   E.GTCOMPMASTID = A.COMPCODE AND E.GTCOMPMASTID=C.COMPCODE WHERE  A.ACTIVE='T' AND E.COMPCODE = '" + Class.Users.HCompcode + "' AND C.USERNAME='" + Class.Users.HUserName + "'");

                //DataTable dt = Utility.SQLQuery("SELECT DISTINCT  a.HRMACIPENTRYDETID,a.MACIP  FROM HRMACIPENTRYDET A WHERE A.MACIP='192.168.101.19'");
                int maxip = dt.Rows.Count;
                if (maxip == 0)
                {
                    MessageBox.Show("IP Address not assign this User.   : " + Class.Users.HUserName);
                }
                if (maxip >= 1)
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
                                Cursor = Cursors.Default;  axCZKEM1.Disconnect(); Cursor = Cursors.Default;
                                MessageBox.Show("No Data Found this Machine...." + ip.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device    
                            
                        }
                        else
                        {
                            axCZKEM1.GetLastError(ref idwErrorCode);
                            Cursor = Cursors.Default; axCZKEM1.Disconnect(); Cursor = Cursors.Default;
                            MessageBox.Show("Unable to connect the device , ErrorCode=" + idwErrorCode.ToString() + "---IP-----" + dt.Rows[i]["MACIP"].ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
               
                //MessageBox.Show(ex.Message.ToString()); 

            }

            try
            {
                //empty();

                if (lvLogs1.Items.Count >= 1)
                {

                    var idd = lvLogs1.Items[lvLogs1.Items.Count - 1].SubItems[1].Text;
                    bytes = null;
                    pictureempimage.Image = null;

                    string sel1 = "SELECT C.HREMPLOYMASTID AS ASPTBLEMPID, B.COMPCODE ,B.GTCOMPMASTID,D.MIDCARD,C.FNAME,E.DISPNAME,E.GTDEPTDESGMASTID, F.CONTACTNO ,C.IDCARDNO FROM GTCOMPMAST B JOIN HREMPLOYMAST C ON C.COMPCODE = B.GTCOMPMASTID JOIN HREMPLOYDETAILS D ON C.HREMPLOYMASTID = D.HREMPLOYMASTID  JOIN GTDEPTDESGMAST E ON E.GTDEPTDESGMASTID = D.DEPTNAME  JOIN HRECONTACTDETAILS F ON F.HREMPLOYMASTID = C.HREMPLOYMASTID AND F.HREMPLOYMASTID = D.HREMPLOYMASTID WHERE D.MIDCARD=" + idd.ToString();
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HREMPLOYDETAILS");
                    DataTable dt = ds.Tables["HREMPLOYDETAILS"];
                    if (dt.Rows.Count == 0)
                    {
                       // MessageBox.Show("This IDCardno  '" + idd + "' is empty in HostelMaster", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                        txtcontactno.Text = Convert.ToString(dt.Rows[0]["CONTACTNO"].ToString());
                        //combohostel.Text = Convert.ToString(dt.Rows[0]["HOSTELNAME"].ToString());
                        //combohostelblock.Text = Convert.ToString(dt.Rows[0]["BLOCKFLOOR"].ToString());
                        //combohostelroom.Text = Convert.ToString(dt.Rows[0]["ROOMNO"].ToString());
                        txthostelblock.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                        txthostelroom.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                        txtsysdate.Text = Convert.ToString(Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt")));
                        txtsystime.Text = Convert.ToString(System.DateTime.Now.ToString("hh:mm:ss tt"));
                        txtmanualTime.Text = Convert.ToString(System.DateTime.Now.ToString("hh:mm"));
                        QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                        var mydata1 = qc.CreateQrCode(txtidcardno.Text, QRCoder.QRCodeGenerator.ECCLevel.L);
                        var code1 = new QRCoder.QRCode(mydata1);
                        pictureBox1.Image = code1.GetGraphic(50, Color.Black, Color.White, true);
                        //string sel2 = "select d.imagefieldvalu AS EMPIMAGE from hremploymast a  join hremploydetails b on a.hremploymastid=b.hremploymastid join gtcompmast c on c.gtcompmastid=a.compcode left join hremploymastimage d on d.hremploymastid=a.hremploymastid  where b.hostel='YES' and c.compcode='" + Class.Users.HCompcode + "'  and  A.FNAME='" + dt.Rows[0]["FNAME"].ToString() + "'";
                        //DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hremploymast");
                        //DataTable dt2 = ds2.Tables["hremploymast"];
                        //if (dt2 != null){ 
                        //    bytes = (byte[])dt2.Rows[0]["EMPIMAGE"];
                        //    Image img = Models.Device.ByteArrayToImage(bytes);
                        //    pictureempimage.Image = img;
                        //}

                        string ins = "INSERT INTO ASPTBLGATEPASS(COMPCODE,  IDCARDNO,  EMPNAME ,  DEPARTMENT,CONTACTNO, SYSTEMDATE,  SYSTEMTIME ,REASON,MANUALTIME,USERNAME,  MODIFIED,  CREATEDON,  IPADDRESS1,IPADDRESS,PERMISSIONHRS,Remarks)VALUES(" + txtcompcode.Text + ",'" + txtidcardno.Text + "'," + txtempname.Text + ",'" + txtdept.Text + "','" + txtcontactno.Text + "','" + txtsysdate.Text + "','" + txtsystime.Text + "','" + comboreason.SelectedValue + "' ,'" + txtmanualTime.Text + "' ," + Class.Users.USERID + ",to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),'" + txtipaddress.Text + "','" + Class.Users.IPADDRESS + "', '" + txtpermissionhrs.Text + "','" + txtRemarks.Text + "')";
                        Utility.ExecuteNonQuery(ins);//empty();
                        Prints_Click(sender, e);
                    }
                }
                else
                {
                    
                    MessageBox.Show("No Data Found in Finger Print Machine","Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.ToString());
            }

            Cursor = Cursors.Default;
        }

        private void Comboreason_MouseHover(object sender, EventArgs e)
        {
            comboreason.BackColor = Color.White;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Hide();
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

        private void Saves_Click(object sender, EventArgs e)
        {
            //if(txtempid.Text != "")
            //{//COMPCODE,  IDCARDNO,  EMPNAME ,  DEPARTMENT,  HOSTELNAME, HOSTELBLOCK,HOSTELROOM,CONTACTNO, SYSTEMDATE,  SYSTEMTIME ,REASON,MANUALTIME,USERNAME,  MODIFIED,  CREATEDON,  IPADDRESS1,IPADDRESS,PERMISSIONHRS,Remarks
            //    string up = " update ASPTBLHOSTELGATEPASS set COMPCODE='" + txtcompcode.Text + "',IDCARDNO='" + txtidcardno.Text + "',EMPNAME='" + txtempname.Text+ "'," +
            //        "DEPARTMENT='" + combo_dept.Text + "' ,HOSTELNAME='" + combohostel.Text + "' ,HOSTELBLOCK='" + combohostelblock.Text + "' ,HOSTELROOM='" + combohostelroom.Text + "' ," +
            //        "CONTACTNO='" + txtcontactno.Text + "',SYSTEMDATE='" + txtsysdate.Text + "',SYSTEMTIME='" + txtsystime.Text + "',REASON='" + comboreason.Text + "' ," +
            //        "MANUALTIME='" + txtmanualTime.Text + "' ,USERNAME=" + Class.Users.USERID + ",MODIFIED=to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),CREATEDON=to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),IPADDRESS1='" + txtipaddress.Text + "',IPADDRESS='" + Class.Users.IPADDRESS + "', PERMISSIONHRS='" + txtpermissionhrs.Text + "',Remarks='" + txtRemarks.Text + "' where ASPTBLHOSTELGATEPASSID='" + txtempid.Text + "'";
            //    Utility.ExecuteNonQuery(up);
            //}
        }

        private void Rlblcompcode_Click(object sender, EventArgs e)
        {

        }

        private void Lblqrcode_Click(object sender, EventArgs e)
        {

        }

        private void Lblgatepass_Click(object sender, EventArgs e)
        {

        }

        private void PictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void butprintcancel_Click(object sender, EventArgs e)
        {
            this.panelprint.Visible = false;
        }

        private void butGetData_Click(object sender, EventArgs e)
        {
           
        }

       
    }
}

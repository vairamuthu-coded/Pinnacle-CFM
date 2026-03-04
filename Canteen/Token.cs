using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Canteen
{
    public partial class Token : Form, ToolStripAccess
    {
        public Token()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString()); this.panelprint.Hide();
            Class.Users.ScreenName = "Token";
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }
        public void Searchs(int EditID)
        {

        }
        private static Token _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
        Pinnacle.Models.MailModel obj = new Models.MailModel();
        public static Token Instance
        {
            get { if (_instance == null) _instance = new Token(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }      
        public void usercheck(string s, string ss, string sss)
        {
          
            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                       // if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.go = true;combooptions.SelectedIndex = 0;combooptions.Enabled = false; } else { this.News.Visible = false; combooptions.SelectedIndex = 1; combooptions.Enabled = true; }
                        //if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                        //if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                        //if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                        //if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                        //if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                        //if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                        //if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                        //if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                        //if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                        //if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                        //if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                        //if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                        //if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }
                    }
                }


            }
            else
            {

                GlobalVariables.Toolstrip1.Enabled = false;
            }

        }
        int amt = 0;
        public zkemkeeper.CZKEM axCZKEM1 = new zkemkeeper.CZKEM();
        private bool bIsConnected = false;//the boolean value identifies whether the device is connected
        private int iMachineNumber = 1;//the serial number of the device.After connecting the device ,this value will be changed.      
        //private static Int32 MyCount;
        ////  private static Int32 ToIPCount;
        //private bool bIsConnectedToIP = false;
        //bool bAddControl = true;
        ////  private static Int32 MyCountFinger;
        ////  private static Int32 MyCountFace;
        //string sdwEnrollNumber = "";
        //string sName = "";
        //string sPassword = "";
        //int iPrivilege = 0;
        //bool bEnabled = false;
        //int idwFingerIndex = 0;
        //string sTmpData = "";
        //string sTmpData1 = "";
        //int iTmpLength = 0;
        //int iFlag = 0;
        //string sEnabled = "";
        ////  int idwTMachineNumber = 0;      
        //int idwVerifyMode = 0;
        //int idwInOutMode = 0;
        //int idwYear = 0;
        //int idwMonth = 0;
        //int idwDay = 0;
        //int idwHour = 0;
        //int idwMinute = 0;
        //int idwSecond = 0;
        //int idwWorkcode = 0;
        //int idwErrorCode = 0;
        //int iGLCount = 0;
        //int iIndex = 0;
        //// int iUpdateFlag = 0;
        //string suserid = "";
        //int iFaceIndex = 50;//the only possible parameter value     
        //int iLength = 0;
        //string sLastEnrollNumber = "";//the former enrollnumber you have upload(define original value as 0)
        //string sCardnumber = "";
        //string MacIP = "";
       // int i = 0;
      
        private void Token_Load(object sender, EventArgs e)
        {
            if (Convert.ToInt64(Class.Users.TOKENEMPID) > 0)
            {

                string sel1 = "  SELECT  A.ASPTBLEMPID,A.EMPNAME,A.IDCARDNO,A.EMPLOYEETYPE FROM ASPTBLEMP A   WHERE A.IDCARDNO='" + Class.Users.TOKENEMPID + "' AND A.ACTIVE='T'";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLEMP");
                DataTable dt1 = ds1.Tables["ASPTBLEMP"];

                txtempid.Text = dt1.Rows[0]["ASPTBLEMPID"].ToString();

                comboempname.DisplayMember = "EMPNAME";
                comboempname.ValueMember = "ASPTBLEMPID";

                comboidcardno.DisplayMember = "IDCARDNO";
                comboidcardno.ValueMember = "ASPTBLEMPID";
                comboempname.DataSource = dt1;
                comboidcardno.DataSource = dt1;
                comboemptype.Text = dt1.Rows[0]["EMPLOYEETYPE"].ToString();
                //comboemptype.DisplayMember = "EMPLOYEETYPE";
                //comboemptype.ValueMember = "EMPLOYEETYPE";
                //comboemptype.DataSource = dt1;


                string sel = " SELECT  A.ASPTBLCANITEMMASID,A.ITEMCODE, A.ITEMNAME1,A.EMPLOYEECOST,A.CONTRACTORCOST  FROM  ASPTBLCANITEMMAS A  JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY=B.ASPTBLCANCATEGORYMASID  WHERE A.ITEMNAME1='" + Class.Users.CANTEENMENUNAME + "' AND A.ACTIVE='T'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLCANITEMMAS");
                DataTable dt = ds.Tables["ASPTBLCANITEMMAS"];

                comboitemcode.DisplayMember = "ITEMCODE";
                comboitemcode.ValueMember = "ASPTBLCANITEMMASID";
                comboitemname.DisplayMember = "ITEMNAME1";
                comboitemname.ValueMember = "ASPTBLCANITEMMASID";
                txtitemcost.Text = "";

                comboitemcode.DataSource = dt;
                comboitemname.DataSource = dt;
              
                if (comboemptype.Text == "EMPLOYEE")
                {
                    txtitemcost.Text = dt.Rows[0]["EMPLOYEECOST"].ToString();
                }
                else
                {

                    txtitemcost.Text = dt.Rows[0]["CONTRACTORCOST"].ToString();
                }
                amt = Sum(Convert.ToInt32("0" + txtitemcost.Text), Convert.ToInt32("0" + txtQuantity.Value), Convert.ToInt32("0" + txtDays.Value));
                txtTotalAmount.Text = amt.ToString();
            }
            else
            {
               //  MessageBox.Show("INVALID");
               
            }
            txtQuantity.Select(); this.panelprint.Show();
        }

        private void Token_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public static int Sum(int num1, int num2)
        {
            int total;
            total = num1 * num2;
            return total;
        }
        public static int add(int num1, int num2)
        {
            int total;
            total = num1 + num2;
            return total;
        }
        public static int Sum(int num1, int num2,int num3)
        {
            int total;
            total = num1 * num2 * num3;
            return total;
        }
        private void TxtQuantity_TextChanged(object sender, EventArgs e)
        {
         
        }

        private void TxtDays_TextChanged(object sender, EventArgs e)
        {
          
        }
        private void TxtQuantity_ValueChanged(object sender, EventArgs e)
        {
             amt = Sum(Convert.ToInt32("0" + txtitemcost.Text), Convert.ToInt32("0" + txtQuantity.Value), Convert.ToInt32("0" + txtDays.Value));
            txtTotalAmount.Text = amt.ToString();
        }
        private void TxtDays_ValueChanged(object sender, EventArgs e)
        {
             amt = Sum(Convert.ToInt32("0" + txtitemcost.Text), Convert.ToInt32("0" + txtQuantity.Value), Convert.ToInt32("0" + txtDays.Value));
            txtTotalAmount.Text = amt.ToString();
        }

        //public bool OGSendEMail(string sendmail, string sendpass, string toEmail, string subject, string emailBody)
        //{
        
        //    subject = "'"+Class.Users.CANTEENMENUNAME+"'"; 
        

        //    MailMessage mm = new MailMessage(sendmail, toEmail);
        //    mm.Subject = subject;
        //    mm.Body = emailBody;
        //    mm.IsBodyHtml = true;
        //    SmtpClient smtp = new SmtpClient();
        //    smtp.Host = "smtp.gmail.com";
        //    smtp.Port = 587;
        //    smtp.EnableSsl = true;
        //    NetworkCredential nc = new NetworkCredential(sendmail, sendpass);
        //    smtp.UseDefaultCredentials = false;
        //    smtp.Credentials = nc;
        //    smtp.Send(mm);
        //    return true;
        //}

        private void OGSendMailToUser()
        {

            //bool result = false;
            //string sendpass = System.Configuration.ConfigurationManager.AppSettings["OGPassword"].ToString();
            //string fromEmail = System.Configuration.ConfigurationManager.AppSettings["OGEmail"].ToString();
            //var sub = "";
            //string emailMsg = "<html><body style='box-shadow: 0 4px 10px 0 rgba(0,0,0,0.2),0 4px 20px 0 rgba(0,0,0,0.19);width:auto;height:auto;'> <div style='height:auto;width:auto; background-color:teal;'>     <div style='height:30px;width:auto;'><div style='float:left;color:white; font-weight:bold;'>	 " + "Your Card Used in    :" + Class.Users.HCompcode + "   Canteen " + " </div> <div style='float:right;color:white; font-weight:bold;margin:0;width:auto;'><input type=button style='color:red;font-weight:large;background-color:red border:0;margin:0;' value='❌'></input> </div></div> <div style='width:auto;background-color:white;padding:1%;'><table style='width:100%;box-shadow: 0 4px 10px 0 rgba(0,0,0,0.2),0 4px 20px 0 rgba(0,0,0,0.19);border-bottom: 1px solid teal'><tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Name </td><td> " + obj.VisitorName + " </td></tr> <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Company </td><td>" + obj.Company + "</td></tr> <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'> Mobile</td><td>" + obj.MobileNo + "</td></tr>  <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Amount</td><td>" + obj.Amount + "</td></tr> <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Purpose</td><td>" + obj.Purpose + "</td></tr> </table></div><div> <hr></hr></body></html>";
            //obj.Body = emailMsg;
            //obj.To = System.Configuration.ConfigurationManager.AppSettings["OGEmail"].ToString();
            //result = OGSendEMail(fromEmail, sendpass, obj.To, obj.Subject, obj.Body);

        }

        private void Btntokenprint_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    string chk = "";

            //    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }                
            //        string ins = "INSERT INTO ASPTBLCANTOKEN(EMPID,EMPNAME,IDCARDNO,ITEMCODE,ITEMNAME1,ITEMCOST,ITEMQTY,NOOFDAYS ,TOTALAMOUNT,ACTIVE  ,USERNAME,MODIFIED,CREATEDON,IPADDRESS )VALUES(" + txtempid.Text + "," + comboempname.SelectedValue + "," + comboidcardno.SelectedValue + "," + comboitemcode.SelectedValue + "," + comboitemname.SelectedValue + "," + txtitemcost.Text + "," + txtQuantity.Text + "," + txtDays.Text + "," + txtTotalAmount.Text + ",'" + chk + "'," + Class.Users.USERID + ",'" + Convert.ToString(Class.Users.CREATED) + "','" + Convert.ToString(Class.Users.CREATED) + "','" + Class.Users.IPADDRESS + "' )";
            //        Utility.ExecuteNonQuery(ins);
            //        MessageBox.Show("Record Saved Successfully     " + comboidcardno.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message.ToString());
            //}
            if (txtQuantity.Value >= 1 && txtDays.Value >= 1)
            {
                TxtDays_ValueChanged(sender, e);
                Butok_Click(sender, e);
            }
            else
            {
                MessageBox.Show("Invalid", "Info", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }

            //string sel0 = "select max(A.ASPTBLCANTOKENID) id FROM ASPTBLCANTOKEN A ";
            //DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLCANTOKEN");
            //DataTable dt0 = ds0.Tables["ASPTBLCANTOKEN"];

            //string sel1 = "   SELECT  A.ASPTBLCANTOKENID, B.ASPTBLEMPID AS EMPID,B.EMPNAME,B.IDCARDNO,C.ITEMCODE,C.ITEMNAME1,C.ITEMCOST,A.ITEMQTY ,A.NOOFDAYS,A.TOTALAMOUNT,B.EMPLOYEETYPE FROM ASPTBLCANTOKEN A  JOIN ASPTBLEMP B ON  A.EMPID = B.ASPTBLEMPID  JOIN ASPTBLCANITEMMAS C ON C.ASPTBLCANITEMMASID = A.ITEMNAME1   WHERE A.ASPTBLCANTOKENID=" + dt0.Rows[0]["id"].ToString();
            //DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANTOKEN");
            //DataTable dt1 = ds1.Tables["ASPTBLCANTOKEN"];
            //lbltoken2.Text = dt1.Rows[0]["ASPTBLCANTOKENID"].ToString();
            //lblempid2.Text = dt1.Rows[0]["EMPID"].ToString();
            //lblempname2.Text = dt1.Rows[0]["EMPNAME"].ToString();
            //lblitemname2.Text = dt1.Rows[0]["ITEMNAME1"].ToString();
            //lblqty2.Text = dt1.Rows[0]["ITEMQTY"].ToString();
            //lblnoofdays2.Text = dt1.Rows[0]["NOOFDAYS"].ToString();
            //comboemptype.Text = dt1.Rows[0]["EMPLOYEETYPE"].ToString();
            //if (comboemptype.Text == "EMPLOYEE")
            //{
            //    lblitemcost.Text = "Rate : " + dt1.Rows[0]["ITEMCOST"].ToString();
            //}
            //else
            //{
            //    int addtot = add(Convert.ToInt32("0" + dt1.Rows[0]["ITEMCOST"].ToString()), Contractorrate);
            //    lblitemcost.Text = "Rate : " + addtot;
            //}
            //lblcompcode.Text = Class.Users.HCompName;
            //lbldatetime.Text = Convert.ToString(Class.Users.CREATED);
            //this.panelprint.Show();
           
        }

        private void Btnexit_Click(object sender, EventArgs e)
        {
           
            this.Dispose();
            
        }

       

        private void Butok_Click(object sender, EventArgs e)
        {
            Saves();

        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                e.Graphics.DrawString(lblheading.Text.ToUpper(), new Font("Arial", 12, FontStyle.Bold), Brushes.DarkBlue, 46, 35);
                e.Graphics.DrawString(lbltoken1.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 16, 63);
                e.Graphics.DrawString(lbltoken2.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 105, 63);

                e.Graphics.DrawString(lblempid1.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 16, 80);
                e.Graphics.DrawString(lblempid2.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 105, 80);

                e.Graphics.DrawString(lblempname1.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 16, 95);
                e.Graphics.DrawString(lblempname2.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 105, 96);

                e.Graphics.DrawString(lblitemname1.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 16, 112);
                e.Graphics.DrawString(lblitemname2.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 105, 112);

                e.Graphics.DrawString(lblqty1.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 16, 128);
                e.Graphics.DrawString(lblqty2.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 105, 127);


                e.Graphics.DrawString(lblnoofdays1.Text.ToUpper(), new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 16, 144);
                e.Graphics.DrawString(lblnoofdays2.Text, new Font("Arial", 9, FontStyle.Bold), Brushes.DarkBlue, 105, 143);

                e.Graphics.DrawString(lblitemcost.Text.ToUpper(), new Font("Arial", 12, FontStyle.Bold), Brushes.DarkBlue, 72, 164);
                e.Graphics.DrawImage(pictureBox1.Image, 62, 190, pictureBox1.Width, pictureBox1.Height);


                e.Graphics.DrawString(lblcompcode.Text.ToUpper(), new Font("Calibri", 8, FontStyle.Regular), Brushes.DarkBlue, 16, 271);
                e.Graphics.DrawString(lbldatetime.Text, new Font("Calibri", 8, FontStyle.Regular), Brushes.DarkBlue, 16, 287);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

     
        private void Btncancel_Click(object sender, EventArgs e)
        {
            this.panelprint.Hide();
        }

      
        private void Combooptions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combooptions.Text == "MULTIPLE")
            {
                txtDays.Enabled = false;
                txtDays.Value = 1;
            }
            else
            {
                txtDays.Enabled = true;
            }
        }

        public void News()
        {
           
        }

        public void Saves()
        {
            try
            {
              
                QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                int count = Convert.ToInt32(txtQuantity.Value) * Convert.ToInt32(txtDays.Value);
                string token = Class.Users.Finyear + "/" + Class.Users.HCompcode + "CAN";

                int totamt = 0;
                string chk = ""; if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }
                //if (printDialog1.ShowDialog() == DialogResult.OK)
                //{

                try
                {
                    if (combooptions.Text == "SINGLE")
                    {

                        for (int i = 0; i < count; i++)
                        {
                            amt = 0;
                            totamt = 0;


                            string sel2 = "select max(A.ASPTBLCANTOKENID)+1 id FROM ASPTBLCANTOKEN A ";
                            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLCANTOKEN");
                            DataTable dt2 = ds2.Tables["ASPTBLCANTOKEN"];
                            string ins = "INSERT INTO ASPTBLCANTOKEN(TOKENNO,EMPID,EMPNAME,IDCARDNO,ITEMCODE,ITEMNAME1,ITEMCOST,ITEMQTY,NOOFDAYS ,TOTALAMOUNT,ACTIVE  ,USERNAME,MODIFIED,CREATEDON,IPADDRESS,TOKENNOCANCEL,EMPLOYEETYPE)VALUES('" + token + "/" + dt2.Rows[0]["id"].ToString() + "'," + txtempid.Text + "," + comboempname.SelectedValue + "," + comboidcardno.SelectedValue + "," + comboitemcode.SelectedValue + "," + comboitemname.SelectedValue + "," + txtitemcost.Text + "," + txtQuantity.Text + "," + txtDays.Text + "," + txtTotalAmount.Text + ",'" + chk + "'," + Class.Users.USERID + ",'" + Convert.ToString(Class.Users.CREATED) + "','" + Convert.ToString(Class.Users.CREATED) + "','" + Class.Users.IPADDRESS + "','','" + comboemptype.Text + "' )";
                            Utility.ExecuteNonQuery(ins);
                            string sel3 = "select max(A.ASPTBLCANTOKENID) id3 FROM ASPTBLCANTOKEN A ";
                            DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLCANTOKEN");
                            DataTable dt3 = ds3.Tables["ASPTBLCANTOKEN"];

                            string sel4 = "   SELECT  A.TOKENNO, B.ASPTBLEMPID AS EMPID,B.EMPNAME,B.IDCARDNO,C.ITEMCODE,C.ITEMNAME1, C.EMPLOYEECOST,C.CONTRACTORCOST,A.ITEMQTY ,A.NOOFDAYS,A.TOTALAMOUNT FROM ASPTBLCANTOKEN A  JOIN ASPTBLEMP B ON  A.EMPID = B.ASPTBLEMPID  JOIN ASPTBLCANITEMMAS C ON C.ASPTBLCANITEMMASID = A.ITEMNAME1   WHERE A.ASPTBLCANTOKENID=" + dt3.Rows[0]["id3"].ToString();
                            DataSet ds4 = Utility.ExecuteSelectQuery(sel4, "ASPTBLCANTOKEN");
                            DataTable dt4 = ds4.Tables["ASPTBLCANTOKEN"];
                            amt = Sum(Convert.ToInt32("0" + txtQuantity.Value), Convert.ToInt32("0" + txtDays.Value));
                            lbltoken2.Text = dt4.Rows[0]["TOKENNO"].ToString();
                            lblempid2.Text = dt4.Rows[0]["EMPID"].ToString();
                            lblempname2.Text = dt4.Rows[0]["EMPNAME"].ToString();
                            lblitemname2.Text = dt4.Rows[0]["ITEMNAME1"].ToString();
                            obj.ID = dt4.Rows[0]["TOKENNO"].ToString();
                            obj.VisitorName= dt4.Rows[0]["EMPNAME"].ToString();
                            obj.Company = Class.Users.HCompName;
                            obj.MobileNo = dt4.Rows[0]["IDCARDNO"].ToString();
                            obj.Purpose = "Lunch Purpose";
                            if (comboemptype.Text == "EMPLOYEE")
                            {
                                totamt = Sum(Convert.ToInt32(dt4.Rows[0]["ITEMQTY"].ToString()), Convert.ToInt32(dt4.Rows[0]["EMPLOYEECOST"].ToString()));
                                lblitemcost.Text = "Rate : " + totamt;
                                lblqty2.Text = dt4.Rows[0]["ITEMQTY"].ToString() + " / Rate: " + dt4.Rows[0]["EMPLOYEECOST"].ToString();
                                obj.Amount = lblqty2.Text;
                            }
                            else
                            {
                                totamt = Sum(Convert.ToInt32("0" + txtQuantity.Value), Convert.ToInt32("0" + dt4.Rows[0]["CONTRACTORCOST"].ToString()));

                                lblitemcost.Text = "Rate : " + totamt;
                                lblqty2.Text = dt4.Rows[0]["ITEMQTY"].ToString() + " / Rate: " + dt4.Rows[0]["CONTRACTORCOST"].ToString();
                                obj.Amount = lblqty2.Text;
                            }

                            lblnoofdays2.Text = txtDays.Value + "/" + count + "  " + "No's";
                            lblcompcode.Text = Class.Users.HCompName;
                            lbldatetime.Text = Convert.ToString(Class.Users.CREATED);
                            var mydata = qc.CreateQrCode(lbltoken2.Text, QRCoder.QRCodeGenerator.ECCLevel.L);
                            var code = new QRCoder.QRCode(mydata);
                            pictureBox1.Image = code.GetGraphic(50, Color.Black, Color.White, true);
                            this.panelprint.Refresh();
                            this.panelprint.Show();
                          //  OGSendMailToUser();
                            printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                            printDocument1.Print();
                              MessageBox.Show("TOKEN NO     " + lbltoken2.Text, "" + Class.Users.HCompcode + "Canteen", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        }
                    }
                    else
                    {
                        amt = 0;
                        amt = Sum(Convert.ToInt32("0" + txtQuantity.Value), Convert.ToInt32("0" + txtDays.Value));
                        string sel2 = "select max(A.ASPTBLCANTOKENID)+1 id FROM ASPTBLCANTOKEN A ";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLCANTOKEN");
                        DataTable dt2 = ds2.Tables["ASPTBLCANTOKEN"];
                        string ins = "INSERT INTO ASPTBLCANTOKEN(TOKENNO,EMPID,EMPNAME,IDCARDNO,ITEMCODE,ITEMNAME1,ITEMCOST,ITEMQTY,NOOFDAYS ,TOTALAMOUNT,ACTIVE  ,USERNAME,MODIFIED,CREATEDON,IPADDRESS,TOKENNOCANCEL,EMPLOYEETYPE )VALUES('" + token + "/" + dt2.Rows[0]["id"].ToString() + "'," + txtempid.Text + "," + comboempname.SelectedValue + "," + comboidcardno.SelectedValue + "," + comboitemcode.SelectedValue + "," + comboitemname.SelectedValue + "," + txtitemcost.Text + "," + txtQuantity.Text + "," + txtDays.Text + "," + txtTotalAmount.Text + ",'" + chk + "'," + Class.Users.USERID + ",'" + Convert.ToString(Class.Users.CREATED) + "','" + Convert.ToString(Class.Users.CREATED) + "','" + Class.Users.IPADDRESS + "','','" + comboemptype.Text + "')";
                        Utility.ExecuteNonQuery(ins);
                        string sel3 = "select max(A.ASPTBLCANTOKENID) id3 FROM ASPTBLCANTOKEN A ";
                        DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLCANTOKEN");
                        DataTable dt3 = ds3.Tables["ASPTBLCANTOKEN"];

                        string sel4 = "   SELECT  A.TOKENNO, B.ASPTBLEMPID AS EMPID,B.EMPNAME,B.IDCARDNO,C.ITEMCODE,C.ITEMNAME1,C.EMPLOYEECOST,C.CONTRACTORCOST,A.ITEMQTY ,A.NOOFDAYS,A.TOTALAMOUNT FROM ASPTBLCANTOKEN A  JOIN ASPTBLEMP B ON  A.EMPID = B.ASPTBLEMPID  JOIN ASPTBLCANITEMMAS C ON C.ASPTBLCANITEMMASID = A.ITEMNAME1   WHERE A.ASPTBLCANTOKENID=" + dt3.Rows[0]["id3"].ToString();
                        DataSet ds4 = Utility.ExecuteSelectQuery(sel4, "ASPTBLCANTOKEN");
                        DataTable dt4 = ds4.Tables["ASPTBLCANTOKEN"];

                        lbltoken2.Text = dt4.Rows[0]["TOKENNO"].ToString();
                        lblempid2.Text = dt4.Rows[0]["EMPID"].ToString();
                        lblempname2.Text = dt4.Rows[0]["EMPNAME"].ToString();
                        lblitemname2.Text = dt4.Rows[0]["ITEMNAME1"].ToString();
                        lblnoofdays2.Text = dt4.Rows[0]["NOOFDAYS"].ToString();
                        obj.ID = dt4.Rows[0]["TOKENNO"].ToString();
                        obj.VisitorName = dt4.Rows[0]["EMPNAME"].ToString();
                        obj.Company = Class.Users.HCompName;
                        obj.MobileNo = dt4.Rows[0]["IDCARDNO"].ToString();
                        obj.Purpose = "Lunch Purpose";
                        if (comboemptype.Text == "EMPLOYEE")
                        {
                            int emprate = Sum(Convert.ToInt32(dt4.Rows[0]["ITEMQTY"].ToString()), Convert.ToInt32(dt4.Rows[0]["EMPLOYEECOST"].ToString()));
                            lblitemcost.Text = "Rate : " + emprate;
                            lblqty2.Text = dt4.Rows[0]["ITEMQTY"].ToString() + " / Rate: " + dt4.Rows[0]["EMPLOYEECOST"].ToString();
                            obj.Amount = lblqty2.Text;
                        }
                        else
                        {
                            int emprate = Sum(Convert.ToInt32(dt4.Rows[0]["ITEMQTY"].ToString()), Convert.ToInt32(dt4.Rows[0]["CONTRACTORCOST"].ToString()));
                            lblitemcost.Text = "Rate : " + emprate;
                            lblqty2.Text = dt4.Rows[0]["ITEMQTY"].ToString() + " / Rate: " + dt4.Rows[0]["CONTRACTORCOST"].ToString();
                            obj.Amount = lblqty2.Text;
                        }
                        lblnoofdays2.Text = txtDays.Value + "/" + count + "  " + "No's";
                        lblcompcode.Text = Class.Users.HCompName;
                        lbldatetime.Text = Convert.ToString(Class.Users.CREATED);
                      //  OGSendMailToUser();

                      //  MessageBox.Show("TOKEN NO     " + lbltoken2.Text, "MULTIPLE", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        var mydata = qc.CreateQrCode(lbltoken2.Text, QRCoder.QRCodeGenerator.ECCLevel.L);
                        var code = new QRCoder.QRCode(mydata);
                        pictureBox1.Image = code.GetGraphic(50, Color.Black, Color.White, true);
                        this.panelprint.Show();
                        printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                        printDocument1.Print();
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    //  DialogResult result1 = MessageBox.Show("Do You want to continue another items ??  " + comboempname.Text, ""+Class.Users.HCompcode+ " Canteen", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                    DataTable dt = Utility.SQLQuery("SELECT DISTINCT B.MACIP,B.MACNO FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE  B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES'   AND C.COMPCODE='" + Class.Users.HCompcode + "' AND B.MACIP='192.168.101.19'");
                    bIsConnected = axCZKEM1.Connect_Net(dt.Rows[0]["MACIP"].ToString(), Convert.ToInt32(4370));
                    if (bIsConnected == true)
                    {
                        axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device


                        axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
                        if (axCZKEM1.ClearGLog(iMachineNumber))
                        {
                            axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                        }

                        axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
                        MessageBox.Show("Print Completed Successfully.  " + comboempname.Text, "" + Class.Users.HCompcode + " Canteen", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }
                //}
                //else
                //{
                //    panelprint.Hide();
                //    panelprint.Refresh();
                //}

            }
            catch (Exception ex)
            {
                
            }
            
        }

        public void Prints()
        {
          
        }

        public void Searchs()
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
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
            this.Hide();
        }

        public void GridLoad()
        {
           
        }
    }
}

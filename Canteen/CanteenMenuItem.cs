using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Web;
using System.Collections.Specialized;

namespace Pinnacle.Canteen
{
    public partial class CanteenMenuItem : Form,ToolStripAccess
    {
        public CanteenMenuItem()
        {
            InitializeComponent();
            label1.Text = Class.Users.ScreenName;
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }
        private static CanteenMenuItem _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        int i = 0; string canitems = "";
     
        public static CanteenMenuItem Instance
        {
            get { if (_instance == null) _instance = new CanteenMenuItem(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }


      
        private void News_Click(object sender, EventArgs e)
        {

        }
        private void TEA()
        {

            //flowLayoutPanel1.Controls.Clear();
            //string sel = " SELECT  A.ASPTBLCANITEMMASID, A.ITEMNAME1,A.ITEMCOST, A.ITEMIMAGE  FROM  ASPTBLCANITEMMAS A  JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY=B.ASPTBLCANCATEGORYMASID   WHERE B.CATEGORY='" + Class.Users.CANTEENMENUNAME + "' AND A.ACTIVE='T'";
            //DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLCANITEMMAS");
            //DataTable dt = ds.Tables["ASPTBLCANITEMMAS"];
            //CustomControl[] items = new CustomControl[dt.Rows.Count];

            //foreach (DataRow myRow in dt.Rows)
            //{

            //    items[i] = new CustomControl();
            //    if (myRow["ITEMIMAGE"].ToString() != "")
            //    {

            //        byte[] bytes = (byte[])myRow["ITEMIMAGE"];
            //        Image img = Models.Device.ByteArrayToImage(bytes);
            //        items[i].userimage = img;
            //    }
            //    items[i].username.Text = Convert.ToString(myRow["ITEMNAME1"].ToString());
            //    items[i].subtitle.Text = "Rate   :"+Convert.ToString(myRow["ITEMCOST"].ToString());
            //    flowLayoutPanel1.Controls.Add(items[i]);
            //    items[i].username.Click += Username_Click1;
            //}
        }

        private void Username_Click1(object sender, EventArgs e)
        {
             canitems = sender.ToString();
            string[] data = canitems.Split(',');
            Class.Users.CANTEENMENUNAME = data[1].Substring(7);
            Token TT = new Token();
            TT.Show();
        }

        private void Username_Click(object sender, EventArgs e)
        {

        }

       

        private void Username_Click2(object sender, EventArgs e)
        {
         
          

            canitems = sender.ToString();
            string[] data = canitems.Split(',');
            Class.Users.CANTEENMENUNAME = data[1].Substring(7);
            Token TT = new Token();
            TT.Show();
        }

      

        private void Username_Click4(object sender, EventArgs e)
        {
            canitems = sender.ToString();
            string[] data = canitems.Split(',');
            Class.Users.CANTEENMENUNAME = data[1].Substring(7);
            Token TT = new Token();
            TT.Show();
        }

     

        private void Username_Click5(object sender, EventArgs e)
        {
            canitems = sender.ToString();
            string[] data = canitems.Split(',');
            Class.Users.CANTEENMENUNAME = data[1].Substring(7);
            Token TT = new Token();
            TT.Show();
        }

        

        private void Username_Click6(object sender, EventArgs e)
        {
            canitems = sender.ToString();
            string[] data = canitems.Split(',');
            Class.Users.CANTEENMENUNAME = data[1].Substring(7);
            Token TT = new Token();
            TT.Show();
        }

        private void Username_Click3(object sender, EventArgs e)
        {
           
           
            canitems = sender.ToString();
            string[] data = canitems.Split(',');
            Class.Users.CANTEENMENUNAME = data[1].Substring(7);
            Token TT = new Token();
            TT.Show();
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

                        //if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
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

        private void CanteenMenuItem_Load(object sender, EventArgs e)
        {
           
            //if (Convert.ToInt32(Class.Users.TOKENEMPID) >= 1)
            //{
            //    string sel = "  SELECT  A.ASPTBLEMPID,A.EMPNAME,A.IDCARDNO,A.EMPIMAGE FROM ASPTBLEMP A   WHERE A.ASPTBLEMPID='" + Class.Users.TOKENEMPID + "' AND A.ACTIVE='T'";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLEMP");
            //    DataTable dt = ds.Tables["ASPTBLEMP"];
            //    lblempid.Text ="EMPID             :  "+ Class.Users.TOKENEMPID;
            //    lblempname.Text = "EMP NAME     :  " + dt.Rows[0]["EMPNAME"].ToString();
            //    lblIdcardno.Text = "IDCARDNO    :  " + dt.Rows[0]["IDCARDNO"].ToString();
            //    if (dt.Rows[0]["EMPIMAGE"].ToString() != "")
            //    {
            //        foreach (DataRow row in dt.Rows)
            //        {
            //            byte[] bytes = (byte[])row["GARMENTIMAGE"];
            //            Image img = Models.Device.ByteArrayToImage(bytes);
            //            pictureemp.Image = img;
            //        }
            //    }   
            //}
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
       
        private void Button1_Click(object sender, EventArgs e)
        {

           

        }

   
        private void FlowLayoutPanel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Lblidcardno_Click(object sender, EventArgs e)
        {

        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void Lblempname_Click(object sender, EventArgs e)
        {

        }

        private void LblIdcardno_Click_1(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public void News()
        {
            
        }

        public void Saves()
        {
            
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

        private void butsend_Click(object sender, EventArgs e)
        {
            try
            {

                //WebClient client = new WebClient();
                //Stream s = client.OpenRead(string.Format("https://platform.clickatell.com/messages/http/send?apiKey=XF7eR4hzQS-1_Eis75GBFA==&to={0}&content={1}", txtphone.Text, txtmesage.Text));

                //StreamReader reader = new StreamReader(s);
                //string result = reader.ReadToEnd();
                //MessageBox.Show("Message Send Successfully");



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        Pinnacle.Models.MailModel obj = new Models.MailModel();
        public bool OGSendEMail(string sendmail, string sendpass, string toEmail, string subject, string emailBody)
        {

            subject = "'" + Class.Users.CANTEENMENUNAME + "'";

            SmtpClient smtp = new SmtpClient();
            MailMessage mm = new MailMessage(sendmail, toEmail);
            mm.Subject = subject;
            mm.Body = emailBody;
            mm.IsBodyHtml = true;
          
            smtp.Host = "ipipi.com";
            smtp.Port = 587;
            smtp.EnableSsl = true;
            NetworkCredential nc = new NetworkCredential(sendmail, sendpass);
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = nc;
            smtp.Send(mm);
            return true;
        }

        private void OGSendMailToUser()
        {

            bool result = false;
            string sendpass = System.Configuration.ConfigurationManager.AppSettings["OGPassword"].ToString();
            string fromEmail = System.Configuration.ConfigurationManager.AppSettings["OGEmail"].ToString();
            var sub = "";
            string emailMsg = "<html><body style='box-shadow: 0 4px 10px 0 rgba(0,0,0,0.2),0 4px 20px 0 rgba(0,0,0,0.19);width:auto;height:auto;'> <div style='height:auto;width:auto; background-color:teal;'>     <div style='height:30px;width:auto;'><div style='float:left;color:white; font-weight:bold;'>	 " + "Your Card Used in    :" + Class.Users.HCompcode + "   Canteen " + " </div> <div style='float:right;color:white; font-weight:bold;margin:0;width:auto;'><input type=button style='color:red;font-weight:large;background-color:red border:0;margin:0;' value='❌'></input> </div></div> <div style='width:auto;background-color:white;padding:1%;'><table style='width:100%;box-shadow: 0 4px 10px 0 rgba(0,0,0,0.2),0 4px 20px 0 rgba(0,0,0,0.19);border-bottom: 1px solid teal'><tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Name </td><td> " + obj.VisitorName + " </td></tr> <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Company </td><td>" + obj.Company + "</td></tr> <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'> Mobile</td><td>" + obj.MobileNo + "</td></tr>  <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Amount</td><td>" + obj.Amount + "</td></tr> <tr style='width:100%;border:1px solid #ccc'><td style='color:teal; font-weight:bold'>Purpose</td><td>" + obj.Purpose + "</td></tr> </table></div><div> <hr></hr></body></html>";
            obj.Body = emailMsg;
            obj.To = System.Configuration.ConfigurationManager.AppSettings["OGEmail"].ToString();
            result = OGSendEMail(fromEmail, sendpass, obj.To, obj.Subject, obj.Body);

        }
        public void sendSMS()
        {
            //String message = HttpUtility.UrlEncode("This is your message");
            //using (var wb = new WebClient())
            //{
            //    byte[] response = wb.UploadValues("https://api.textlocal.in/send/", new NameValueCollection()
            //    {
            //    {"apikey" , "yourapiKey"},
            //    {"numbers" , "9751828323"},
            //    {"message" , message},
            //    {"sender" , "9751828323"}
            //    });
            //    string result = System.Text.Encoding.UTF8.GetString(response);
            //    return result;
            //}
        }

        public void Searchs(int EditID)
        {
            
        }
        //private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    TEA();
        //    COFFEE();
        //    TIFFEN();
        //    DINNER();
        //    SNACKS();
        //    LUNCH();
        //}
    }
}

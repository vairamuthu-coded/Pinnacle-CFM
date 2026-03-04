using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using zkemkeeper;
namespace Pinnacle.Hostel
{
    public partial class Form1 : Form
    {
        private static Form1 _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
        Models.Device dev = new Models.Device();
        public Form1()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            label2.Text = Class.Users.ScreenName; 
          
        }
        public static Form1 Instance
        {
            get { if (_instance == null) _instance = new Form1(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }
        public zkemkeeper.CZKEM axCZKEM1 = new zkemkeeper.CZKEM();
        private bool bIsConnected = false;//the boolean value identifies whether the device is connected
        private int iMachineNumber;//the serial number of the device.After connecting the device ,this value will be changed.      
        private static Int32 MyCount;
        //  private static Int32 ToIPCount;
        private bool bIsConnectedToIP = false;
        bool bAddControl = true;
        //  private static Int32 MyCountFinger;
        //  private static Int32 MyCountFace;
        string sdwEnrollNumber = "";
        string sdwEnrollNumber1 = "";
        string sName = "";
        string sPassword = "";
        int iPrivilege = 0;
        bool bEnabled = false;
        int idwFingerIndex = 0;
        string sTmpData = "";
        string sTmpData1 = "";
        int iTmpLength = 0;
        int iFlag = 0;
        string sEnabled = "";
        //  int idwTMachineNumber = 0;      
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
        // int iUpdateFlag = 0;
        string suserid = "";
        int iFaceIndex = 50;//the only possible parameter value     
        int iLength = 0;
        string sLastEnrollNumber = "";//the former enrollnumber you have upload(define original value as 0)
        string sCardnumber = "";
        string MacIP = "";


        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                        if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                        if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                        if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                        if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                        if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                        if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                        if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                        if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                        if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                        if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                        if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                        if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                        if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                        if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }
                    }
                }


            }
            else
            {

            }

        }
        private void Btnsaves_Click(object sender, EventArgs e)
        {
            string ccode = "";
            ccode = Class.Users.HCompcode;
            try
            {

                lvLogs.Items.Clear();

                int k = 0;
                iIndex = 0;

               
                    iGLCount = 0; Cursor = Cursors.WaitCursor;

                    string ip = "";
                    string macno = "";
                    string mactype = "";
                    string mactype2 = "";


                    Cursor = Cursors.WaitCursor;
                  
                    //try
                    //{

                    int i = 0;

                ip = "192.168.101.20";

                        bIsConnected = axCZKEM1.Connect_Net(ip, Convert.ToInt32(4370));
                       

                        if (bIsConnected == true)
                        {


                    if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                    {

                        while (axCZKEM1.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode, out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                        {


                            //   DateTime inputDate = new DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, idwMinute);

                            string id = sdwEnrollNumber.ToString().Length.ToString();
                            int card = Convert.ToInt32(id.ToString());
                            string idcard = sdwEnrollNumber;
                            string sdate = idwDay.ToString().Length.ToString();
                            int date = Convert.ToInt32(sdate.ToString());
                            string ss;
                            if (date < 2)
                            {
                                ss = "0" + idwDay.ToString();
                            }
                            else
                            {
                                ss = idwDay.ToString();
                            }

                            string smonth = idwMonth.ToString().Length.ToString();
                            int month = Convert.ToInt32(smonth.ToString());
                            string sss;

                            if (month < 2)
                            {
                                sss = "0" + idwMonth.ToString();
                            }
                            else
                            {
                                sss = idwMonth.ToString();
                            }

                            string shour = idwHour.ToString().Length.ToString();
                            int hour = Convert.ToInt32(shour.ToString());
                            string h;
                            if (hour < 2)
                            {
                                h = "0" + idwHour.ToString();
                            }
                            else
                            {
                                h = idwHour.ToString();
                            }

                            string sminits = idwMinute.ToString().Length.ToString();
                            int minits = Convert.ToInt32(sminits.ToString());
                            string m;
                            if (minits < 2)
                            {
                                m = "0" + idwMinute.ToString();
                            }
                            else
                            {
                                m = idwMinute.ToString();
                            }

                            string sseconds = idwSecond.ToString().Length.ToString();
                            int seconds = Convert.ToInt32(sseconds.ToString());
                            string s;
                            if (seconds < 2)
                            {
                                s = "0" + idwSecond.ToString();
                            }
                            else
                            {
                                s = idwSecond.ToString();
                            }
                            var dat = ss.ToString() + "/" + sss.ToString() + "/" + idwYear.ToString();
                            string time = h.ToString() + ":" + m.ToString() + ":" + s.ToString();
                            iGLCount++;


                            lvLogs.Items.Add(iGLCount.ToString());
                            lvLogs.Items[iIndex].SubItems.Add(idcard.ToString());//modify by Darcy on Nov.26 2009
                            lvLogs.Items[iIndex].SubItems.Add(ip.ToString());//ipaddresss.ToString());
                            lvLogs.Items[iIndex].SubItems.Add(idwInOutMode.ToString());
                            lvLogs.Items[iIndex].SubItems.Add(dat + " " + time);
                            lvLogs.Items[iIndex].SubItems.Add(dat);
                            lvLogs.Items[iIndex].SubItems.Add(time);
                            lvLogs.Items[iIndex].SubItems.Add(macno.ToString());
                            lvLogs.Items[iIndex].SubItems.Add(mactype);
                            lvLogs.Items[iIndex].SubItems.Add(mactype2);
                            lvLogs.Items[iIndex].SubItems.Add(mactype);
                            lvLogs.Items[iIndex].SubItems.Add(mactype2);
                            iIndex++;




                        }

                        if (lvLogs.Items.Count > 0)
                        {

                            var result = lvLogs.Items.Count > 0 ? lvLogs.Items[lvLogs.Items.Count - 1] : null;

                            MessageBox.Show(lvLogs.Items[lvLogs.Items.Count - 1].SubItems[1].Text);
                        }
                    }
                    else
                    {
                        Cursor = Cursors.Default;
                        MessageBox.Show("No Data Found this Machine...." + ip.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                            axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device    

                        }
                        else
                        {
                            axCZKEM1.GetLastError(ref idwErrorCode);


                            Cursor = Cursors.Default;
                            MessageBox.Show("Unable to connect the device , ErrorCode=" + idwErrorCode.ToString() + "---IP-----","Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);


                        }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());

            }


            Cursor = Cursors.Default;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        private void News_Click(object sender, EventArgs e)
        {

        }
    }
}

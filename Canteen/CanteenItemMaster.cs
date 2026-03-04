using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using zkemkeeper;

namespace Pinnacle.Canteen
{
    public partial class CanteenItemMaster : Form,ToolStripAccess
    {
        private static CanteenItemMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
        public static CanteenItemMaster Instance
        {
            get { if (_instance == null) _instance = new CanteenItemMaster(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }


        public CanteenItemMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
           
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
         zkemkeeper.CZKEM axCZKEM1 = new zkemkeeper.CZKEM();
        private bool bIsConnected = false;//the boolean value identifies whether the device is connected
        private int iMachineNumber;//the serial number of the device.After connecting the device ,this value will be changed.      
        private static Int32 MyCount;
        //  private static Int32 ToIPCount;
        private bool bIsConnectedToIP = false;
        bool bAddControl = true;
        //  private static Int32 MyCountFinger;
        //  private static Int32 MyCountFace;
        string sdwEnrollNumber = "";
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
        int i = 0;
        private void CanteenItemMaster_Load(object sender, EventArgs e)
        {
            comboitemload();   
        }

        void comboitemload()
        {
            string sel1 = "  SELECT A.ASPTBLCANCATEGORYMASID,  A.Category  FROM  ASPTBLCANCATEGORYMAS A   ORDER BY 1";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANCATEGORYMAS");
            DataTable dt1 = ds1.Tables["ASPTBLCANCATEGORYMAS"];
            comboitem.DataSource = dt1;
            comboitem.ValueMember = "ASPTBLCANCATEGORYMASID";
            comboitem.DisplayMember = "Category";
        }
        private void comboitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            pop();
        }
      
        private void pop()
        {
           

            flowLayoutPanel1.Controls.Clear();
            string sel = " SELECT  A.ASPTBLCANITEMMASID, A.ITEMNAME1,A.EMPLOYEECOST, A.ITEMIMAGE,B.CATEGORY  FROM  ASPTBLCANITEMMAS A  JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY=B.ASPTBLCANCATEGORYMASID   WHERE A.ACTIVE='T' and b.category='" + comboitem.Text + "' ORDER BY 5";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLCANITEMMAS");
            DataTable dt = ds.Tables["ASPTBLCANITEMMAS"];
            UserControls.CanteenCustom[] items = new UserControls.CanteenCustom[dt.Rows.Count];     
            foreach (DataRow myRow in dt.Rows)
            {

                items[i] = new UserControls.CanteenCustom();
                if (myRow["ITEMIMAGE"].ToString() != "")
                {
                    byte[] bytes = (byte[])myRow["ITEMIMAGE"];
                    Image img = Models.Device.ByteArrayToImage(bytes);
                    items[i].userimage = img;
                }
                items[i].menuname.Text = Convert.ToString(myRow["ITEMNAME1"].ToString());
                items[i].subtitle.Text = "Rate   :" + Convert.ToString(myRow["EMPLOYEECOST"].ToString());
                flowLayoutPanel1.Controls.Add(items[i]);

               

                items[i].menuname.Click += Menuname_Click;

            }



        }

        private void Menuname_Click(object sender, EventArgs e)
        {
            try
            {

                Timercanteen_Tick(sender, e);


                if (lvLogs.Items.Count > 0)
                {
                    Class.Users.TOKENEMPID = 0; pictureemp.Image = null;
                    Class.Users.TOKENEMPID = Convert.ToInt64(lvLogs.Items[lvLogs.Items.Count - 1].SubItems[1].Text);

                    string sel = "  SELECT  A.ASPTBLEMPID,A.EMPNAME,A.IDCARDNO,A.EMPIMAGE FROM ASPTBLEMP A   WHERE A.IDCARDNO='" + Class.Users.TOKENEMPID + "' AND A.ACTIVE='T'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLEMP");
                    DataTable dt = ds.Tables["ASPTBLEMP"];
                    lblempid.Text = "EMPID             :  " + dt.Rows[0]["IDCARDNO"].ToString();
                    lblempname.Text = "EMP NAME     :  " + dt.Rows[0]["EMPNAME"].ToString();
                    lblIdcardno.Text = "IDCARDNO    :  " + dt.Rows[0]["IDCARDNO"].ToString();
                    Class.Users.IDCARDNO = Convert.ToInt64(Class.Users.TOKENEMPID);
                    if (dt.Rows[0]["EMPIMAGE"].ToString() != "")
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            byte[] bytes = (byte[])row["EMPIMAGE"];
                            Image img = Models.Device.ByteArrayToImage(bytes);
                            pictureemp.Image = img;
                        }
                    }
                    if (lblempid.Text == "") { MessageBox.Show("Pls Select Employee Name"); }
                    else
                    {

                        timercanteen.Enabled = false;
                        Class.Users.CANTEENMENUNAME = "";
                        string s = sender.ToString();
                        string[] data = s.Split(',');
                        Class.Users.CANTEENMENUNAME = data[1].Substring(7);

                        foreach (Form a in Application.OpenForms) { if (a.Name == "Token") { a.Close(); break; } }
                        Token TT = new Token();
                        TT.Show();
                    }
                }
                else
                {
                    Class.Users.TOKENEMPID = 0; pictureemp.Image = null;
                    Cursor = Cursors.Default; lblempid.Text = ""; lblempname.Text = ""; lblIdcardno.Text = "";

                    MessageBox.Show("Finger Print is empty   ", Class.Users.HCompcode + " Canteen", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception EX)
            {
                MessageBox.Show("  Username_Click  ", " Canteen", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                timercanteen.Enabled = false;

            }
        }

        private void Username_Click1(object sender, EventArgs e)
        {
           
        }

   

        private void CanteenItemMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

      
    

        private void ItemRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pop();
        }

      

        private  void Timercanteen_Tick(object sender, EventArgs e)
        {
            string ccode = "";
            try
            {
                ccode = Class.Users.HCompcode;
                lvLogs.Items.Clear();
               Cursor = Cursors.WaitCursor;
                int k = 0;

              //  timercanteen.Enabled = true;

                iGLCount = 0;

                string ip = "";
                //string macno = "";
                //string mactype = "";
                //string mactype2 = "";

             
                    

                DataTable dt = Utility.SQLQuery("SELECT DISTINCT  B.ASPTBLMACIPID,B.MACIP  FROM ASPTBLUSERMAS A JOIN ASPTBLMACIP B ON  B.USERNAME =A.USERNAME AND B.ACTIVE='T' JOIN GTCOMPMAST E ON   E.GTCOMPMASTID = A.COMPCODE    WHERE  A.ACTIVE='T'   AND     E.COMPCODE = '" + Class.Users.HCompcode + "' AND A.USERNAME='" + Class.Users.HUserName + "' ");
                int maxip = dt.Rows.Count;
                if (maxip == 0)
                {
                    MessageBox.Show("IP Address not assign this User.   : " + Class.Users.HUserName);
                }
                if (maxip >= 1)
                {
                    int i = 0; iIndex = 0;
                    for (i = 0; i < maxip; i++)
                    {
                      
                        bIsConnected = axCZKEM1.Connect_Net(dt.Rows[i]["MACIP"].ToString(), Convert.ToInt32(4370));
                        ip = dt.Rows[i]["MACIP"].ToString();
                       
                        if (bIsConnected == true)
                        {


                            if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                            {

                                while (axCZKEM1.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode, out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                                {
                                    // DateTime inputDate = new DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, idwMinute);
                                    //if (Convert.ToDateTime(inputDate) >= frmdate.Value.Date && Convert.ToDateTime(inputDate) <= todate.Value.Date.AddDays(1).AddTicks(-1))
                                    //{

                                    string idcard = sdwEnrollNumber;
                                    string sel1 = "  SELECT  A.ASPTBLEMPID,A.EMPNAME,A.IDCARDNO FROM ASPTBLEMP A   WHERE A.IDCARDNO='" + idcard + "' AND A.ACTIVE='T'";
                                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLEMP");
                                    DataTable dt2 = ds1.Tables["ASPTBLEMP"];
                                    if (dt2.Rows.Count > 0)
                                    {
                                        iGLCount++;

                                        lvLogs.Items.Add(iGLCount.ToString());
                                        lvLogs.Items[iIndex].SubItems.Add(idcard.ToString());//modify by Darcy on Nov.26 2009
                                        lvLogs.Items[iIndex].SubItems.Add(dt2.Rows[0]["EMPNAME"].ToString());//modify by Darcy on Nov.26 2009
                                        lvLogs.Items[iIndex].SubItems.Add(dt2.Rows[0]["IDCARDNO"].ToString());//modify by Darcy on Nov.26 2009

                                        iIndex++;

                                    }
                                    //else
                                    //{
                                    //    MessageBox.Show("This ID Card/Employee null value  in Employee Master : " +  idcard.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    //}


                                }

                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                //   MessageBox.Show("No Data Found this Machine...." + ip.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device    
                        }
                        else
                        {
                            Cursor = Cursors.Default;
                            axCZKEM1.GetLastError(ref idwErrorCode);
                            
                            MessageBox.Show("Unable to connect the device","Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        }
                        Cursor = Cursors.Default;
                    }
                }
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                Cursor = Cursors.Default;
            }
        }

        private void Saves_Click(object sender, EventArgs e)
        {

        }

        private void News_Click(object sender, EventArgs e)
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

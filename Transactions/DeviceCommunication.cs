using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Collections.Generic;
using System.Linq;

namespace Pinnacle.Transactions
{


    public partial class DeviceCommunication : Form, ToolStripAccess
    {
        private static DeviceCommunication _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.Device dev = new Models.Device();string systemuser = "";


        ListView allip = new ListView();
        ListView allip1 = new ListView();
        ListView allip2 = new ListView();
        ListView removeuserid = new ListView(); UICO uti = new UICO();
        ListView listfilter = new ListView();
        
        public static DeviceCommunication Instance
        {

            get
            {
                if (_instance == null)
                    _instance = new DeviceCommunication();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }

        }


        public DeviceCommunication()
        {
            InitializeComponent();

            DateTime dateForButton = DateTime.Now;
            frmdate.Value = dateForButton.AddDays(-1);
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; 
           //SecondtabControl2.TabPages.Remove(tab2RemovefrmMachine);
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            SecondtabControl2.SelectedTab= tab6Attlots;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            splitter1.BackColor = Class.Users.BackColors;
            splitter2.BackColor = Class.Users.BackColors;
            splitter4.BackColor = Class.Users.BackColors;
            splitter3.BackColor = Class.Users.BackColors;
            splitter5.BackColor = Class.Users.BackColors;
            splitercardreader.BackColor = Class.Users.BackColors;
            splitterFaceTemp.BackColor = Class.Users.BackColors;
            splitterFingerTemp.BackColor = Class.Users.BackColors;
         
            label10.BackColor = Class.Users.BackColors;
            label12.BackColor = Class.Users.BackColors;
            lblattsearch.BackColor = Class.Users.BackColors;
            checkremoveall.BackColor = Class.Users.BackColors;
            lblremovesearch.BackColor = Class.Users.BackColors;
            label15.BackColor = Class.Users.BackColors;
            label16.BackColor = Class.Users.BackColors;
            label5.BackColor = Class.Users.BackColors;
            label6.BackColor = Class.Users.BackColors;
            lblfingerconstate.BackColor = Class.Users.BackColors;
            lblfingersearch.BackColor = Class.Users.BackColors;
            label8.BackColor = Class.Users.BackColors;

            label27.BackColor = Class.Users.BackColors;
            label35.BackColor = Class.Users.BackColors;

            checkCard.BackColor = Class.Users.BackColors;
            label13.BackColor = Class.Users.BackColors;

            label14.BackColor = Class.Users.BackColors;
            checkface.BackColor = Class.Users.BackColors;

            checkallrows.BackColor = Class.Users.BackColors;
            lblsearch.BackColor = Class.Users.BackColors;
        }

        public void usercheck(string s, string ss, string sss)
        {
            s = Class.Users.HCompcode;
            ss = Class.Users.HUserName;
            sss = Class.Users.ScreenName;
            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                       // if (dt1.Rows[r]["NEWS"].ToString() == "T") { GlobalVariables.News.Visible = false; SecondtabControl2.TabPages.Add(tab2RemovefrmMachine); SecondtabControl2.SelectTab(tab2RemovefrmMachine); } else { GlobalVariables.News.Visible = false; SecondtabControl2.TabPages.Remove(tab2RemovefrmMachine); }


                    }
                }


            }
            else
            {

            }
        }

     
         zkemkeeper.CZKEM axCZKEM1 = new zkemkeeper.CZKEM();
        private bool bIsConnected = false;//the boolean value identifies whether the device is connected
        private int iMachineNumber = 10;//the serial number of the device.After connecting the device ,this value will be changed.
        private static Int32 MyCount;
     //  private static Int32 ToIPCount;     
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
        string sUserID = "";
        int iFaceIndex = 50;//the only possible parameter value     
        int iLength = 0;
        string sCardnumber = "";
        string MacIP = "";
        int iBackupNumber = 1;


        private void DeviceCommunication_Load(object sender, EventArgs e)
        {

            try
            {
                DataTable dt = dev.FromIp(Class.Users.HCompcode);
             
                if (dt !=null)
                {
                   

                        comboMasterIp.DisplayMember = "MACIP";
                        comboMasterIp.ValueMember = "MACIP";
                        comboMasterIp.DataSource = dt;
                    

                }
                else
                {
                    comboMasterIp.DataSource = null;
                }
                if (Class.Users.HUserName == "VAIRAM") {   } else { SecondtabControl2.TabPages.Remove(tab2RemovefrmMachine); }
               
                listfilter.Items.Clear();                 progressBar1.Value = 0; lblattcount.Text = ""; lblprogress1.Text = ""; listviewchecklistip.Items.Clear(); listviewchecklistip1.Items.Clear(); listviewchecklistip2.Items.Clear();
                Lvdownremove.Items.Clear(); listviewattdown.Items.Clear(); listremovechecklistip.Items.Clear(); allip2.Items.Clear();
                Lvdownall.Items.Clear(); allip.Items.Clear(); allip1.Items.Clear();
                comboMasterIp.SelectedIndex = -1; AttIPLoad(); faceip();
            }
            catch(Exception EX)
            { }
        }



        //private void Exit_Click(object sender, EventArgs e)
        //{

            
          
        //    this.Hide();
           
        //}

        private void DeviceCommunication_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        private void BtnConnect_Click(object sender, EventArgs e)
        {
            if (comboMasterIp.Text.Trim() == "" || txtPort.Text.Trim() == "")
            {
                MessageBox.Show("IP and Port cannot be null", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
          
            SecondtabControl2.SelectTab(tab1finger); listremovechecklistip.Items.Clear();
            lblprogress1.Text = ""; listViewupload.Items.Clear(); Class.Users.UserTime = 0;
            int idwErrorCode = 0; lblattcount.Text = ""; checkremoveall.Checked = false; checkallrows.Checked = false;
               Cursor = Cursors.WaitCursor;
            string macip = "";
            if (btnConnect.Text == "Finger DisConnect")
            {
                axCZKEM1.Disconnect();
                bIsConnected = false;
                btnConnect.Text = "Finger Download ??";
                lblState.Text = "Current State:DisConnected";
             
                Cursor = Cursors.Default;
                return;
            }
            axCZKEM1.PullMode = 1;
            bIsConnected = axCZKEM1.Connect_Net(comboMasterIp.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == true)
            {
                LvDownload.Items.Clear(); Lvdownremove.Items.Clear();
                Lvdownall.Items.Clear(); lvCard.Items.Clear();
                LvDownload.BeginUpdate();
                lvCard.BeginUpdate(); Class.Users.UserTime = 0;
                listfilter.Items.Clear();
                axCZKEM1.EnableDevice(iMachineNumber, false);
                Cursor = Cursors.WaitCursor;
                btnConnect.Text = "Finger DisConnect";
                lblState.Text = "Current State:Connected";
                axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory
                axCZKEM1.IsTFTMachine(iMachineNumber);   // to distingush machines
                axCZKEM1.ReadAllTemplate(iMachineNumber);//read all the users' fingerprint templates to the memory
                int r = 1;
                
               
                while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sdwEnrollNumber, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
                {
                    string sss = "";
                    string ssss = "";
                    string card = "";
                    
                      macip = comboMasterIp.Text;
                    for (idwFingerIndex = 0; idwFingerIndex < 10; idwFingerIndex++)
                    {
                        Class.Users.UserTime = 0;
                        if (axCZKEM1.GetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, out iFlag, out sTmpData, out iTmpLength))//get the corresponding templates string and length from the memory
                        {
                            ssss = sTmpData;
                     
                       
                            if (axCZKEM1.GetStrCardNumber(out sCardnumber))//get the card number from the memory  
                            {


                            }
                            ListViewItem list = new ListViewItem();
                        ssss = sTmpData;
                            card = sCardnumber;
                            list.SubItems.Add(r.ToString());
                            list.SubItems.Add(sdwEnrollNumber.ToString());
                            list.SubItems.Add(sName.ToUpper());
                            list.SubItems.Add(idwFingerIndex.ToString());
                            list.SubItems.Add(ssss);
                            list.SubItems.Add(sss);
                            list.SubItems.Add(card);
                            list.SubItems.Add(iPrivilege.ToString());
                            list.SubItems.Add(sPassword);
                            if (bEnabled == true) { list.SubItems.Add("True"); } else { list.SubItems.Add("False"); }
                            list.SubItems.Add(iFlag.ToString());
                            list.SubItems.Add(macip.ToString());
                            this.listfilter.Items.Add((ListViewItem)list.Clone());
                            LvDownload.Items.Add(list);

                            ListViewItem list1 = new ListViewItem();
                            list1.SubItems.Add(r.ToString());
                            list1.SubItems.Add(sdwEnrollNumber.ToString());
                            list1.SubItems.Add(sName.ToUpper());
                            list1.SubItems.Add(idwFingerIndex.ToString());
                            list1.SubItems.Add(ssss);
                            list1.SubItems.Add(sss);
                            list1.SubItems.Add(card);
                            list1.SubItems.Add(iPrivilege.ToString());
                            list1.SubItems.Add(sPassword);
                            if (bEnabled == true) { list1.SubItems.Add("True"); } else { list1.SubItems.Add("False"); }
                            list1.SubItems.Add(iFlag.ToString());
                            list1.SubItems.Add(macip.ToString());
                            Lvdownall.Items.Add(list1);

                            ListViewItem list2 = new ListViewItem();

                            list2.SubItems.Add(r.ToString());
                            list2.SubItems.Add(sdwEnrollNumber.ToString());
                            list2.SubItems.Add(sName.ToUpper());
                            list2.SubItems.Add(idwFingerIndex.ToString());
                            list2.SubItems.Add(ssss);
                            list2.SubItems.Add(sss);
                            list2.SubItems.Add(card);
                            list2.SubItems.Add(iPrivilege.ToString());
                            list2.SubItems.Add(sPassword);
                            if (bEnabled == true) { list2.SubItems.Add("True"); } else { list2.SubItems.Add("False"); }
                            list2.SubItems.Add(iFlag.ToString());
                            lvCard.Items.Add(list2);
                            if (r % 2 == 0)
                            {
                                list.BackColor = Color.White;
                                list1.BackColor = Color.White;
                                list2.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                                list1.BackColor = Color.WhiteSmoke;
                                list2.BackColor = Color.WhiteSmoke;

                            }

                            r++;


                            sss = "";
                            ssss = "";
                            card = "";
                            iIndex++;
                      
                           
                            lblprogress1.Text = "Data DownLoading From Machine : " + card + " %" + "User ID:-" + sdwEnrollNumber.ToString();
                        lblprogress1.Refresh();
                        
                    }
                    }

                }
                LvDownload.EndUpdate();
                Lvdownall.EndUpdate();
                lvCard.EndUpdate();
                axCZKEM1.EnableDevice(iMachineNumber, true);
                Cursor = Cursors.Default;
                lblattcount.Text = "Total Employee Finger Rows Count  :" + LvDownload.Items.Count.ToString() + " and IP Addres   :" + comboMasterIp.Text;
                allip.Items.Clear(); allip1.Items.Clear();


            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Unable to connect the device,ErrorCode=" + idwErrorCode.ToString(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            listviewchecklistip.Items.Clear();
            axCZKEM1.RefreshData(iMachineNumber);
            combo_ToIPload();
            Cursor = Cursors.Default; bIsConnected = false;
        }
        private void AttIPLoad()
        {
            string ss = Class.Users.HCompcode;
            DataTable dt1 = dev.ToIp(ss);iIndex = 1;
            if (dt1.Rows.Count >= 0)
            {

                foreach (DataRow row in dt1.Rows)
                {

                  
                    ListViewItem list3 = new ListViewItem();                 

                    list3.SubItems.Add(row["MACIP"].ToString());
                   list3.SubItems.Add("------");
                    if (iIndex % 2 == 0)
                    {
                        list3.BackColor = Color.White;
                        
                    }
                    else
                    {
                        list3.BackColor = Color.WhiteSmoke;
                        
                    }
                    listviewattdown.Items.Add(list3);
                    iIndex++;
                }
               
            }
            else
            {
                MessageBox.Show("No data found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            btnConnect.Text = "Finger Download ??";
            lblState.Text = "Current State:DisConnected";
        }
        private void combo_ToIPload()
        {
            listviewchecklistip1.Items.Clear();
            listviewchecklistip2.Items.Clear();
            string ss = Class.Users.HCompcode;
             DataTable dt1 = dev.AllIp(ss, comboMasterIp.Text);
         
            if (dt1.Rows.Count > 0)
            {
                iIndex = 1;
                foreach (DataRow row in dt1.Rows)
                {
                    
                    ListViewItem list = new ListViewItem();                  
                    ListViewItem list2 = new ListViewItem();
                    list.SubItems.Add(row["MACIP"].ToString());
                    list.SubItems.Add("------");
                    
                    
                    ListViewItem listremove = new ListViewItem();


                    listremove.SubItems.Add(row["MACIP"].ToString());
                    listremove.SubItems.Add("------");
                   

                    ListViewItem listcard = new ListViewItem();
                    listcard.SubItems.Add(row["MACIP"].ToString());
                    listcard.SubItems.Add("------");
                 

                    ListViewItem listface = new ListViewItem();
                    listface.SubItems.Add(row["MACIP"].ToString());
                    listface.SubItems.Add("------");
                   
                    if (iIndex % 2 == 0)
                    {
                        list.BackColor = Color.White;
                        listremove.BackColor = Color.White;
                        listcard.BackColor = Color.White;
                        listface.BackColor = Color.White;
                    }
                    else
                    {
                        list.BackColor = Color.WhiteSmoke;
                        listremove.BackColor = Color.WhiteSmoke;
                        listcard.BackColor = Color.WhiteSmoke;
                        listface.BackColor = Color.WhiteSmoke;

                    }
                    listviewchecklistip.Items.Add(list);
                    listremovechecklistip.Items.Add(listremove);
                    listviewchecklistip1.Items.Add(listcard);
                    listviewchecklistip2.Items.Add(listface);
                    iIndex++;
                }
                ListViewItem listremove1 = new ListViewItem();
                if (comboMasterIp.Text != "")
                {
                    listremove1.SubItems.Add(comboMasterIp.Text);
                    listremove1.SubItems.Add("Connected");
                    listremove1.SubItems.Add("True");
                    allip2.Items.Add(listremove1);
                }
            }
            else
            {
                MessageBox.Show("No data found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void combo_RemoveIPload()
        {

            string ss = Class.Users.HCompcode;
            DataTable dt1 = dev.AllIp(ss, comboMasterIp.Text);
            if (dt1.Rows.Count >= 0)
            {
                ListViewItem listremove1 = new ListViewItem(); iIndex = 1;
                foreach (DataRow row in dt1.Rows)
                {


                    ListViewItem listremove = new ListViewItem();


                    listremove.SubItems.Add(row["MACIP"].ToString());
                    listremove.SubItems.Add("------");
                    listremovechecklistip.Items.Add(listremove);
                    if (iIndex % 2 == 0)
                    {
                        listremove.BackColor = Color.White;

                    }
                    else
                    {
                        listremove.BackColor = Color.WhiteSmoke;

                    }
                   
                    iIndex++;
                }
               
            }
            else
            {
                MessageBox.Show("No data found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void faceip()
        {
           
                string ss = Class.Users.HCompcode;
                DataTable dt2 = dev.ToIp(ss);
            if (dt2.Rows.Count >= 0)
            {

                combofingerboxip.DisplayMember = "MACIP";
                combofingerboxip.ValueMember = "MACIP";
                combofingerboxip.DataSource = dt2;

                combofaceboxip.DisplayMember = "MACIP";
                combofaceboxip.ValueMember = "MACIP";
                combofaceboxip.DataSource = dt2;

                combocardreaderipbox.DisplayMember = "MACIP";
                combocardreaderipbox.ValueMember = "MACIP";
                combocardreaderipbox.DataSource = dt2;

                DataTable dt3 = new DataTable();
                dt3 = dev.AllIp();
                comboremoveip.DisplayMember = "MACIP";
                comboremoveip.ValueMember = "MACIP";
                comboremoveip.DataSource = dt3;
            }
            else
            {
                MessageBox.Show("No data found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


                combofingerboxip.SelectedIndex = -1;
                comboremoveip.SelectedIndex = -1;
                combofaceboxip.SelectedIndex = -1;
                combocardreaderipbox.SelectedIndex = -1;
           
         //   comboremoveip.DisplayMember = "---";
           // comboremoveip.Items.Add("-----");
            //comboremoveip.ValueMember = "-1";
        }
       

        private void Contextallitemcheck_Click(object sender, EventArgs e)
        {
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void LvDownload_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListViewItem itt = new ListViewItem(); Class.Users.UserTime = 0;
            iIndex = listViewupload.Items.Count; progressBar1.Value = 0;
            if (e.Item.Checked == true)
            {

                itt.SubItems.Add(e.Item.SubItems[1].Text);
                itt.SubItems.Add(e.Item.SubItems[2].Text);
                itt.SubItems.Add(e.Item.SubItems[3].Text);
                itt.SubItems.Add(e.Item.SubItems[4].Text);
                itt.SubItems.Add(e.Item.SubItems[5].Text);
                itt.SubItems.Add(e.Item.SubItems[6].Text);
                itt.SubItems.Add(e.Item.SubItems[7].Text);
                itt.SubItems.Add(e.Item.SubItems[8].Text);
                itt.SubItems.Add(e.Item.SubItems[9].Text);
                itt.SubItems.Add(e.Item.SubItems[10].Text);
                itt.SubItems.Add(e.Item.SubItems[11].Text);
                itt.SubItems.Add(e.Item.SubItems[12].Text);

                if (iIndex % 2 == 0)
                {
                    itt.BackColor = Color.White;
                }
                else
                {
                    itt.BackColor = Color.WhiteSmoke;
                }
                listViewupload.Items.Add(itt);
                iIndex++;
            }
        }
        private void LvDownload_ItemActivate(object sender, EventArgs e)
        {
            //try
            //{
            //    if (LvDownload.Items.Count > 0)
            //    {
            //        ListViewItem item1 = new ListViewItem(); iIndex = listViewupload.Items.Count;
            //        for (int c = 0; c < LvDownload.SelectedItems[0].SubItems.Count; c++)
            //        {
            //            item1.SubItems.Add(LvDownload.SelectedItems[0].SubItems[c].Text);
            //        }

            //        if (iIndex % 2 == 0)
            //        {
            //            item1.BackColor = Color.White;
            //        }
            //        else
            //        {
            //            item1.BackColor = Color.WhiteSmoke;
            //        }
            //        iIndex++;
            //        listViewupload.Items.Add(item1);
            //    }


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }
        private void Contextallitemcheck_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void ListViewupload_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listViewupload.Items.Count > 0)
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        int i = 0;
                        for (i = 0; i < listViewupload.Items.Count; i++)
                        {

                            if (listViewupload.Items[i].Selected)
                            {
                                MessageBox.Show("UserID:   " + listViewupload.Items[i].SubItems[1].Text + "      Name:  " + listViewupload.Items[i].SubItems[2].Text, "Delete this Record");

                                listViewupload.Items[i].Remove();
                                i--;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("pls upload data from Master IP", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Btntransfer_Click(object sender, EventArgs e)
        {
            try
            {
                txtsearch.Text = ""; Class.Users.UserTime = 0;
                if (listviewchecklistip.CheckedItems.Count >= 0)
                {
                   
                    DialogResult result = MessageBox.Show("Do You want to Export  Finger Index??", "Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result.Equals(DialogResult.OK))
                    {                       
                       
                        for (int j = 0; j < allip.Items.Count; j++)
                        {

                            if (allip.Items.Count >= 0)
                            {

                                Cursor = Cursors.WaitCursor;
                                if (allip.Items[j].SubItems[1].Text.Length > 10)
                                {
                                    lblattcount.Text = ""; bIsConnected = false;
                                    bIsConnected = axCZKEM1.Connect_Net(allip.Items[j].SubItems[1].Text, Convert.ToInt32(txtPort.Text));
                                    if (bIsConnected == true)
                                    {

                                        lblState.Text = "Current State:Connected";
                                        int idwErrorCode = 0;
                                        int iFlag = 1;

                                        axCZKEM1.EnableDevice(iMachineNumber, false);

                                        int i = 0;                                     
                                        progressBar1.Minimum = 0;
                                        progressBar1.Maximum = listViewupload.Items.Count;
                                        for (i = 0; i < listViewupload.Items.Count; i++)
                                        {

                                            string uid = "0";
                                            int index1 = 0;
                                            string c = "";
                                            sdwEnrollNumber = listViewupload.Items[i].SubItems[2].Text;
                                            sName = listViewupload.Items[i].SubItems[3].Text;
                                            idwFingerIndex = Convert.ToInt32("0" + listViewupload.Items[i].SubItems[4].Text);
                                            sTmpData = listViewupload.Items[i].SubItems[5].Text;
                                            sTmpData1 = listViewupload.Items[i].SubItems[6].Text;
                                            sCardnumber = listViewupload.Items[i].SubItems[7].Text;
                                            iPrivilege = Convert.ToInt32("0" + listViewupload.Items[i].SubItems[8].Text);
                                            sPassword = listViewupload.Items[i].SubItems[9].Text;
                                            sEnabled = listViewupload.Items[i].SubItems[10].Text;
                                            iFlag = Convert.ToInt32("0" + listViewupload.Items[i].SubItems[11].Text);
                                            MacIP = allip.Items[j].SubItems[1].Text;
                                            if (sEnabled == "True")
                                            {
                                                bEnabled = true;

                                            }
                                            else
                                            {
                                                bEnabled = false;
                                            }
                                            axCZKEM1.SetStrCardNumber(sCardnumber);//Before you using function SetUserInfo,set the card number to make sure you can upload it to the device

                                            if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//face templates are part of users' information
                                            {
                                                axCZKEM1.SetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, iFlag, sTmpData);//upload templates information to the memory


                                                uid = sdwEnrollNumber;
                                                index1 = idwFingerIndex;

                                                lblattcount.Text = "Total Employee Finger Rows Count  : " + listViewupload.Items.Count.ToString() + " and IP Addres   :" + allip.Items[j].SubItems[1].Text;

                                                decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(listViewupload.Items.Count)) * (i + 1);
                                                lblprogress1.Text = " Data Transfer Machine to Machine : " + (per).ToString("N0") + " % '" + i.ToString() + "' ID Card No: -" + sdwEnrollNumber;

                                                lblprogress1.Refresh();

                                                progressBar1.Value = i + 1;

                                            }
                                            else
                                            {
                                                axCZKEM1.GetLastError(ref idwErrorCode);

                                                // axCZKEM1.EnableDevice(iMachineNumber, true); Cursor = Cursors.Default;
                                                //  return;
                                            }
                                        }

                                        axCZKEM1.BatchUpdate(iMachineNumber);//upload all the information in the memory
                                        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed

                                        axCZKEM1.EnableDevice(iMachineNumber, true);


                                        MessageBox.Show("Successfully upload fingerprint, " + "total:" + listViewupload.Items.Count.ToString() + "IP      :" + allip.Items[j].SubItems[1].Text, "Success");
                                        Cursor = Cursors.Default;
                                        progressBar1.Value = 0;
                                    }
                                    else
                                    {
                                        Cursor = Cursors.Default;
                                        MessageBox.Show("Machine DisConnected" + allip.Items[j].SubItems[1].Text);
                                    }

                                }                                
                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                MessageBox.Show("Invalid");
                            }


                            Cursor = Cursors.Default;
                        }
                    }
                   
                }
                else
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("Machine not connected.pls send IP Address","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                Cursor = Cursors.Default;
                comboMasterIp.Enabled = true;
                btnConnect.Enabled = true;
                lblState.Text = "Current State:DisConnected";
                
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("pls Connect Device", "error"); Cursor = Cursors.Default;
            }
            listViewupload.Items.Clear();
        }
     
     
        private void BtnDeleteEnrollData_Click(object sender, EventArgs e)
        {
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cbUserIDDE.Text.Trim() == "" || cbBackupDE.Text.Trim() == "")
            {
                MessageBox.Show("Please input the UserID and BackupNumber first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            idwErrorCode = 0;

            sUserID = cbUserIDDE.Text.Trim();
            int iBackupNumber = Convert.ToInt32(cbBackupDE.Text.Trim());

            Cursor = Cursors.WaitCursor;
            axCZKEM1.EnableDevice(iMachineNumber, true);
            if (axCZKEM1.SSR_DeleteEnrollData(iMachineNumber, sUserID, iBackupNumber))
            {
                axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
               // UICO ui = new UICO();
                //ui.DeleteAllEmpTmTFT(sUserID, iBackupNumber);
                cbUserIDDE.SelectedIndex = -1; cbUserIDTmp.SelectedIndex = -1;
                MessageBox.Show("DeleteEnrollData,UserID=" + sUserID + " BackupNumber=" + iBackupNumber.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor = Cursors.Default;
        }

        //Delete a certain user's fingerprint template of specified index
        //You shuold input the the user id and the fingerprint index you will delete
        //The difference between the two functions "SSR_DelUserTmpExt" and "SSR_DelUserTmp" is that the former supports 24 bits' user id.

        private void BtnSSR_DelUserTmpExt_Click(object sender, EventArgs e)
        {
            try
            {
                if (bIsConnected == false)
                {
                    MessageBox.Show("Please connect the device first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (bIsConnected == true)
                {
                    if (cbUserIDTmp.Text.Trim() == "" || cbFingerIndex.Text.Trim() == "")
                    {
                        MessageBox.Show("Please input the UserID and FingerIndex first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    idwErrorCode = 0;
                    sUserID = cbUserIDTmp.Text.Trim();
                    idwFingerIndex = Convert.ToInt32(cbFingerIndex.Text.Trim());

                    Cursor = Cursors.WaitCursor;
                    if (axCZKEM1.SSR_DelUserTmpExt(iMachineNumber, sUserID, idwFingerIndex))
                    {
                        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed


                        cbUserIDDE.SelectedIndex = -1; cbUserIDTmp.SelectedIndex = -1;
                        MessageBox.Show("SSR_DelUserTmpExt,UserID:" + sUserID + " FingerIndex:" + idwFingerIndex.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        axCZKEM1.GetLastError(ref idwErrorCode);
                        MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("--" + ex.ToString());
            }
        }


        //Clear all the administrator privilege(not clear the administrators themselves)
        private void BtnClearAdministrators_Click(object sender, EventArgs e)
        {
          
            if (comboremoveip.Text != "" )
            {
                MyCount = 0; bIsConnected = false ; Cursor = Cursors.WaitCursor;
                bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
                if (bIsConnected == true)
                {
                    idwErrorCode = 0;

                    Cursor = Cursors.WaitCursor;
                    if (axCZKEM1.ClearAdministrators(iMachineNumber))
                    {
                        axCZKEM1.RefreshData(iMachineNumber);
                        Cursor = Cursors.Default;
                        //the data in the device should be refreshed
                        MessageBox.Show("Successfully clear administrator privilege from teiminal!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                    }
                    else
                    {
                        axCZKEM1.GetLastError(ref idwErrorCode); Cursor = Cursors.Default;
                        MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            Cursor = Cursors.Default;
        }

        //Delete all the user information in the device,while the related fingerprint templates will be deleted either. 
        //(While the parameter DataFlag  of the Function "ClearData" is 5 )
        private void BtnClearDataUserInfo_Click(object sender, EventArgs e)
        {

            if (comboremoveip.SelectedIndex >= 0)
            {
                MyCount = 0;
                bIsConnected = false;
                bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
                if (bIsConnected == true)
                {
                    //------------------------Clear Data Start----------------------------//
                    DialogResult result1 = MessageBox.Show("Do You want to Clear All Data from Machine??\n'   IP ADDRESS ---" + comboremoveip.Text + "'", "Attendance Logs", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result1.Equals(DialogResult.OK))
                    {
                        int idwErrorCode = 0;

                        int iDataFlag = 5;

                        Cursor = Cursors.WaitCursor;
                        if (axCZKEM1.ClearData(iMachineNumber, iDataFlag))
                        {
                            axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                            MessageBox.Show("Clear all the UserInfo data!", "Success");

                        }
                        else
                        {
                            axCZKEM1.GetLastError(ref idwErrorCode);
                            MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    Cursor = Cursors.Default;
                    lblattcount.Text = "Employee All Details are Deleted from Machine  :" + lvremove.Items.Count.ToString() + " and Connected IP Addres   :" + comboremoveip.Text;

                }
                //    ////---------------------Clear Data End--------------------------------//


            }
            Cursor = Cursors.Default;
        }

        //Clear all the fingerprint templates in the device(While the parameter DataFlag  of the Function "ClearData" is 2 )
        private void BtnClearDataTmps_Click(object sender, EventArgs e)
        {
            if (comboremoveip.SelectedIndex > 0)
            {
                MyCount = 0;
                bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
                int idwErrorCode = 0;

                int iDataFlag = 1;

                Cursor = Cursors.WaitCursor;
                DialogResult result1 = MessageBox.Show("Do You want to Clear from Machine??\n'    IP ADDRESS ---" + comboremoveip.Text + "'", "Attendance Logs", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result1.Equals(DialogResult.OK))
                {
                    axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
                    if (axCZKEM1.ClearData(iMachineNumber, iDataFlag))
                    {
                        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                        MessageBox.Show("Clear all the fingerprint templates!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        axCZKEM1.GetLastError(ref idwErrorCode);
                        MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
                    lblattcount.Text = "Attendance Logs are Cleared From Machine  :" + lvremove.Items.Count.ToString() + " and Connected IP Addres   :" + comboremoveip.Text;

                }



            }
            Cursor = Cursors.Default;
        }




       public void News() { 
            comboMasterIp.SelectedIndex = -1;
            comboMasterIp.SelectedIndex = -1;
            combocardreaderipbox.SelectedIndex = -1;
            cbUserIDTmp.SelectedIndex = -1;
            cbUserIDDE.SelectedIndex = -1; cbUserIDTmp.SelectedIndex = -1;
            allip.Items.Clear();
            allip1.Items.Clear();
            allip2.Items.Clear();
           // ComboMasterIp_SelectedIndexChanged(sender, e);
            listViewupload.Items.Clear();
            LvDownload.Items.Clear();
            listviewchecklistip.Refresh();
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            splitter1.BackColor = Class.Users.BackColors;
            splitter2.BackColor = Class.Users.BackColors;
            splitter4.BackColor = Class.Users.BackColors;
            splitter3.BackColor = Class.Users.BackColors;
            splitter5.BackColor = Class.Users.BackColors;
            splitercardreader.BackColor = Class.Users.BackColors;
            splitterFaceTemp.BackColor = Class.Users.BackColors;
            splitterFingerTemp.BackColor = Class.Users.BackColors;
          
            label10.BackColor = Class.Users.BackColors;
            label12.BackColor = Class.Users.BackColors;
            lblattsearch.BackColor = Class.Users.BackColors;
            checkremoveall.BackColor = Class.Users.BackColors;
            lblremovesearch.BackColor = Class.Users.BackColors;
            label15.BackColor = Class.Users.BackColors;
            label16.BackColor = Class.Users.BackColors;
            label5.BackColor = Class.Users.BackColors;
            label6.BackColor = Class.Users.BackColors;
            lblfingerconstate.BackColor = Class.Users.BackColors;
            lblfingersearch.BackColor = Class.Users.BackColors;
            label8.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            label27.BackColor = Class.Users.BackColors;
            label35.BackColor = Class.Users.BackColors;

            checkCard.BackColor = Class.Users.BackColors;
            label13.BackColor = Class.Users.BackColors;

            label14.BackColor = Class.Users.BackColors;
            checkface.BackColor = Class.Users.BackColors;

            checkallrows.BackColor = Class.Users.BackColors;
            lblsearch.BackColor = Class.Users.BackColors;
        }

        public void Saves()
        {

           // Btntransfer_Click(sender, e);
        }

        class Data
        {
            public string Id { get; set; }
            public string Name { get; set; }
        }
        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txtsearch.Text.Length >= 1)
                {
                    LvDownload.Items.Clear();

                    foreach (ListViewItem item in Lvdownall.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (Lvdownall.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text) || Lvdownall.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text))
                        {

                            list.Text = Lvdownall.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[9].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[10].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[11].Text);
                            list.SubItems.Add(Lvdownall.Items[item0].SubItems[12].Text);
                          
                            if (item0 % 2 == 0)
                            {
                                list.BackColor = Color.White;
                               
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                               
                            }

                            

                            LvDownload.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {
                    try
                    {
                        LvDownload.Items.Clear();
                        foreach (ListViewItem item in Lvdownall.Items)
                        {
                            this.LvDownload.Items.Add((ListViewItem)item.Clone());
                            if (item0 % 2 == 0)
                            {
                                item.BackColor = Color.White;

                            }
                            else
                            {
                                item.BackColor = Color.WhiteSmoke;

                            }
                            item0++;
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("---" + ex.ToString());
            }
            
        }

        private void BtnSetStrCardNumber_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                bIsConnected = axCZKEM1.Connect_Net(combocardreaderipbox.Text, Convert.ToInt32(txtPort.Text));
                if (combocardreaderipbox.SelectedIndex >= 0)
                {
                    if (bIsConnected == false)
                    {
                        Cursor = Cursors.Default;
                        MessageBox.Show("Please connect the device first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (bIsConnected == true)
                    {
                        if (cbUserId_Card.Text.Trim() == "" || txtCardnumber.Text.Trim() == "" || txtName.Text.Trim() == "")
                        {
                            MessageBox.Show("UserID,Privilege,Cardnumber,Name must be inputted first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            cbUserId_Card.Focus(); Cursor = Cursors.Default;
                            return;
                        }
                        idwErrorCode = 0;

                        sdwEnrollNumber = cbUserId_Card.Text;
                        sName = txtName.Text;
                        idwFingerIndex = Convert.ToInt32(txtcardfindex.Text);
                        sTmpData = txtcardfingerimage.Text;
                        sTmpData1 = txtcardfaceimage.Text;
                        sCardnumber = txtCardnumber.Text.Trim();
                        iPrivilege = Convert.ToInt32("0" + cbPrivilegecard.Text);
                        sPassword = txtPassword.Text;
                        idwFingerIndex = Convert.ToInt32(txtcardfindex.Text);
                        if (chbEnabled.Checked) { bEnabled = true; } else { bEnabled = false; }
                        Cursor = Cursors.WaitCursor;
                        axCZKEM1.EnableDevice(iMachineNumber, false);
                        // UICO uti = new UICO();
                        if (checkcard1.Checked == true)
                        {
                            sCardnumber = null; checkcard1.Checked = false;
                            axCZKEM1.SetStrCardNumber(sCardnumber);//Before you using function SetUserInfo,set the card number to make sure you can upload it to the device

                            if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//face templates are part of users' information
                            {
                                axCZKEM1.SetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, iFlag, sTmpData);//upload templates information to the memory

                                axCZKEM1.SetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, sTmpData1, iLength);//upload face templates information to the device

                                MessageBox.Show("SetUserInfo,UserID:" + sdwEnrollNumber + " Privilege:" + iPrivilege.ToString() + " Enabled:" + bEnabled.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);


                            }
                            else
                            {
                                axCZKEM1.GetLastError(ref idwErrorCode);
                                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                axCZKEM1.EnableDevice(iMachineNumber, true);
                                return;
                            }
                        }
                        if (checkfinger1.Checked == true)
                        {
                            sTmpData = ""; checkfinger1.Checked = false;
                            axCZKEM1.SetStrCardNumber(sCardnumber);//Before you using function SetUserInfo,set the card number to make sure you can upload it to the device

                            if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//face templates are part of users' information
                            {
                                axCZKEM1.SetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, iFlag, sTmpData);//upload templates information to the memory

                                axCZKEM1.SetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, sTmpData1, iLength);//upload face templates information to the device

                                MessageBox.Show("SetUserInfo,UserID:" + sdwEnrollNumber + " Privilege:" + iPrivilege.ToString() + " Enabled:" + bEnabled.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);


                            }
                            else
                            {
                                axCZKEM1.GetLastError(ref idwErrorCode);
                                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                axCZKEM1.EnableDevice(iMachineNumber, true);
                                return;
                            }
                        }
                        if (checkface1.Checked == true)
                        {
                            sTmpData1 = null; checkface1.Checked = false;
                            axCZKEM1.SetStrCardNumber(sCardnumber);//Before you using function SetUserInfo,set the card number to make sure you can upload it to the device

                            if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//face templates are part of users' information
                            {
                                axCZKEM1.SetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, iFlag, sTmpData);//upload templates information to the memory

                                axCZKEM1.SetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, sTmpData1, iLength);//upload face templates information to the device

                                MessageBox.Show("SetUserInfo,UserID:" + sdwEnrollNumber + " Privilege:" + iPrivilege.ToString() + " Enabled:" + bEnabled.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);


                            }
                            else
                            {
                                axCZKEM1.GetLastError(ref idwErrorCode);
                                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                axCZKEM1.EnableDevice(iMachineNumber, true);
                                return;
                            }
                        }

                        if (checkcard1.Checked == false && checkfinger1.Checked == false && checkface1.Checked == false)
                        {

                            axCZKEM1.SetStrCardNumber(sCardnumber);//Before you using function SetUserInfo,set the card number to make sure you can upload it to the device

                            if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//face templates are part of users' information
                            {
                                axCZKEM1.SetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, iFlag, sTmpData);//upload templates information to the memory

                                axCZKEM1.SetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, sTmpData1, iLength);//upload face templates information to the device

                                MessageBox.Show("SetUserInfo,UserID:" + sdwEnrollNumber + " Privilege:" + iPrivilege.ToString() + " Enabled:" + bEnabled.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);


                            }
                            else
                            {
                                axCZKEM1.GetLastError(ref idwErrorCode);
                                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                axCZKEM1.EnableDevice(iMachineNumber, true);
                                return;
                            }
                           
                        }
                        axCZKEM1.BatchUpdate(iMachineNumber);//upload all the information in the memory
                        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed

                        axCZKEM1.EnableDevice(iMachineNumber, true);

                        Cursor = Cursors.Default;
                        txtName.Text = ""; txtPassword.Text = ""; txtCardnumber.Text = ""; chbEnabled.Checked = false;
                    }
                }
                else
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("Pls Select Ip List");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("-------" + sCardnumber + "--" + ex.ToString());
            }
           // BtnGetStrCardNumber_Click(sender, e);
        }



        private void BtnConnects_Click(object sender, EventArgs e)
        {
            string ccode = "";
            ccode = Class.Users.HCompcode; Class.Users.UserTime = 0;
            try
            {

                lvLogs.Items.Clear(); lblprogress1.Text = "";
                progressBar1.Value = 0;
                int k = 0;
                iIndex = 0;
                DialogResult result = MessageBox.Show("Download Attendance Logs??", "Download Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {

                    iGLCount = 0; Cursor = Cursors.WaitCursor;
                    if (listviewattdown.CheckedItems.Count > 0)
                    {
                        string ip = "";
                        string macno = "";
                        string mactype = "";
                        string mactype2 = "";
                        for (int j = 0; j < allip1.Items.Count; j++)
                        {
                            if (allip1.Items[j].SubItems[3].Text == "True")
                            {

                                Cursor = Cursors.WaitCursor;

                                DataTable dt = dev.IPLOAD(Class.Users.HCompcode, allip1.Items[j].SubItems[1].Text);
                                int maxip = dt.Rows.Count;
                                listfilter.Items.Clear();
                                int i = 0;
                                for (i = 0; i < maxip; i++)
                                {

                                    lblattcount.Text = ""; ip = ""; macno = ""; mactype = ""; mactype2 = "";
                                    bIsConnected = axCZKEM1.Connect_Net(dt.Rows[i]["MACIP"].ToString(), Convert.ToInt32(txtPort.Text));
                                    ip = dt.Rows[i]["MACIP"].ToString();
                                    macno = dt.Rows[i]["MACNO"].ToString();
                                    mactype = dt.Rows[i]["MTYPE"].ToString();
                                    mactype2 = dt.Rows[i]["MTYPE2"].ToString();

                                    if (bIsConnected == true)
                                    {
                                        btnConnects.Refresh();
                                        btnConnects.Text = "DisConnect"; lblState.Refresh();
                                        lblState.Text = "Current State:Connected";
                                        axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
                                        if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                                        {
                                            lvLogs.BeginUpdate();
                                            while (axCZKEM1.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode, out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                                            {
                                                Class.Users.UserTime = 0;
                                                DateTime inputDate = new DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, idwSecond);
                                                if (Convert.ToDateTime(inputDate) >= frmdate.Value.Date && Convert.ToDateTime(inputDate) <= todate.Value.Date.AddDays(1).AddTicks(-1))
                                                {
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
                                                    var dat = ss.ToString() + "-" + sss.ToString() + "-" + idwYear.ToString();
                                                    string time = h.ToString() + ":" + m.ToString() + ":" + s.ToString();
                                                    iGLCount++;


                                                    // {
                                                    ListViewItem list = new ListViewItem();
                                                    list.SubItems.Add(iGLCount.ToString());
                                                    list.SubItems.Add(idcard.ToString());//modify by Darcy on Nov.26 2009
                                                    list.SubItems.Add(ip.ToString());//ipaddresss.ToString());
                                                    list.SubItems.Add(idwInOutMode.ToString());
                                                    list.SubItems.Add(dat + " " + time);
                                                    list.SubItems.Add(dat);
                                                    list.SubItems.Add(time);
                                                    list.SubItems.Add(macno.ToString());
                                                    list.SubItems.Add(mactype);
                                                    list.SubItems.Add(mactype2);
                                                    list.SubItems.Add(mactype);
                                                    list.SubItems.Add(mactype2);
                                                    this.listfilter.Items.Add((ListViewItem)list.Clone());
                                                    lvLogs.Items.Add(list);
                                                    if (iIndex % 2 == 0)
                                                    {
                                                        list.BackColor = Color.White;

                                                    }
                                                    else
                                                    {
                                                        list.BackColor = Color.WhiteSmoke;


                                                    }

                                                    iIndex++;
                                                    lvLogs.EndUpdate();
                                                }
                                            }
                                        }
                                        else
                                        {
                                            Cursor = Cursors.Default;
                                            MessageBox.Show("No Data Found this Machine...." + ip.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                        axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device    
                                        lblattcount.Text = "Downloading from Machine.   " + "    Rows Count  :" + lvLogs.Items.Count.ToString() + " and IP Addres   :" + ip.ToString();
                                       
                                    }

                                    if (lvLogs.Items.Count > 0)
                                    {

                                        progressBar1.Minimum = 0;
                                        progressBar1.Maximum = lvLogs.Items.Count;
                                        lblattcount.Text = "";
                                        string del1 = "delete from " + Class.Users.HCompcode + "TRS_TEMPATTLOG";
                                        Utility.ExecuteNonQuery(del1);
                                        foreach (ListViewItem item in lvLogs.Items)
                                        {
                                            decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(lvLogs.Items.Count)) * (item.Index + 1);
                                            lblprogress1.Text = "Downloading : " + (per).ToString("N0") + " %";
                                            lblprogress1.Refresh();
                                            string del = "DELETE FROM " + Class.Users.HCompcode + "TRS_ATTLOG   WHERE ENROLLNO ='" + Convert.ToString(item.SubItems[2].Text) + "' AND TO_CHAR(TO_TIMESTAMP(DATETIMERECORD,'DD-MM-YYYY HH24:MI:SS'),'DD-MM-YYYY HH24:MI:SS') = TO_CHAR(TO_TIMESTAMP('" + Convert.ToDateTime(item.SubItems[5].Text) + "','DD-MM-YYYY HH24:MI:SS'),'DD-MM-YYYY HH24:MI:SS') AND IPAddress ='" + Convert.ToString(item.SubItems[3].Text) + "'";
                                            Utility.ExecuteNonQuery(del);

                                            string ins = "INSERT INTO " + Class.Users.HCompcode + "TRS_TEMPATTLOG(MACHINENUMBER, IPADDRESS, ENROLLNO, VERIFYMODE, INOUTMODE, WORKCODE, DATETIMERECORD)VALUES(" + item.SubItems[8].Text + ",'" + item.SubItems[3].Text + "'," + item.SubItems[2].Text + "," + 0 + "," + 0 + "," + 0 + ",'" + item.SubItems[5].Text + "')";
                                            Utility.ExecuteNonQuery(ins);

                                            string ins1 = "INSERT INTO " + Class.Users.HCompcode + "TRS_ATTLOG(MACHINENUMBER, IPADDRESS, ENROLLNO, VERIFYMODE, INOUTMODE, WORKCODE, DATETIMERECORD)VALUES(" + item.SubItems[8].Text + ",'" + item.SubItems[3].Text + "'," + item.SubItems[2].Text + "," + 0 + "," + 0 + "," + 0 + ",'" + item.SubItems[5].Text + "')";
                                            Utility.ExecuteNonQuery(ins1);
                                            progressBar1.Value = Convert.ToInt32(item.Index + 1);
                                            this.listfilter.Items.Add((ListViewItem)item.Clone());
                                        }
                                        string exec = "BEGIN  ATTINSERT('" + Class.Users.HCompcode + "','" + Class.Users.HUserName + "'); END;";
                                        Utility.ExecuteNonQuery(exec);
                                        lblattcount.Text = "Total Count  :" + lvLogs.Items.Count;
                                        progressBar1.Value = 0;
                                        //DialogResult result1 = MessageBox.Show("Do You want to Clear from Machine??\n' THIS MACHINE NO  :" + dt.Rows[i]["MACNO"].ToString() + "   IP ADDRESS ---" + dt.Rows[i]["MACIP"].ToString() + "'", "Attendance Logs", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                        //if (result1.Equals(DialogResult.OK))
                                        //{
                                        //    axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
                                        //    if (axCZKEM1.ClearGLog(iMachineNumber))
                                        //    {
                                        //        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                                        //    }
                                        //    else
                                        //    {
                                        //        axCZKEM1.GetLastError(ref idwErrorCode);
                                        //        MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error");
                                        //    }
                                        //    axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
                                        //    lblattcount.Text = "Attendance Logs are Cleared From Machine  :" + lvremove.Items.Count.ToString() + " and Connected IP Addres   :" + comboremoveip.Text;
                                        //}
                                    }
                                    else
                                    {
                                        axCZKEM1.GetLastError(ref idwErrorCode);
                                        Cursor = Cursors.Default;
                                        MessageBox.Show("No Data Found , ErrorCode=" + idwErrorCode.ToString() + "---IP-----" + dt.Rows[i]["MACIP"].ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                    }
                                    Class.Users.UserTime = 0;
                                    axCZKEM1.RefreshData(iMachineNumber);
                                    MessageBox.Show("Attendance Downloaded. :\t" + dt.Rows[i]["MACIP"].ToString() + "", "Confirmation Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }


                            }
                        }

                    }
                    else
                    {
                        Cursor = Cursors.Default;
                        MessageBox.Show("Bio-Metric Device DisConnected \t", "Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("Download Canceled \t", "Information", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                   
                }
              
              
            }
            catch (Exception ex)
            {
                sessiondelete();
                MessageBox.Show(ex.Message.ToString());
               
            }
           
            btnConnects.Refresh(); axCZKEM1.RefreshData(iMachineNumber);
            btnConnects.Text = "Connect / Import";
            lblState.Text = "Current State:DisConnected";
            Cursor = Cursors.Default;
        }

        void sessiondelete()
        {
            systemuser = Environment.UserName;
            string selchek1 = "select a.ASPTBLSESSIONMASID   from  ASPTBLSESSIONMAS a join gtcompmast  b on a.compcode=b.gtcompmastid join asptblusermas c on  c.compcode = a.compcode AND C.COMPCODE=B.GTCOMPMASTID  and A.USERNAME=C.USERID   and B.compcode='" + Class.Users.HCompcode + "'      and C.username='" + Class.Users.HUserName + "' and C.PASWORD='" + Class.Users.PWORD + "' AND A.OSUSER='" + systemuser + "'  ORDER BY a.ASPTBLSESSIONMASID DESC ";//and A.SYSTEMDATE = to_date('" + System.DateTime.Now.ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy') and  C.active='T'  order by 1
            DataSet dschk1 = Utility.ExecuteSelectQuery(selchek1, "ASPTBLSESSIONMAS");
            DataTable dtchk1 = dschk1.Tables["ASPTBLSESSIONMAS"];
            if (dtchk1.Rows.Count > 0)
            {
                Class.Users.SessionID = Convert.ToInt32(dtchk1.Rows[0]["ASPTBLSESSIONMASID"].ToString());
                string del = "delete from  ASPTBLSESSIONMAS a where a.ASPTBLSESSIONMASID=" + Convert.ToInt32(dtchk1.Rows[0]["ASPTBLSESSIONMASID"].ToString());
                Utility.ExecuteNonQuery(del);
            }
        }
        private void BtnGetDeviceStatus_Click(object sender, EventArgs e)
        {

        }
        // combocardreaderipbox
        private void Txtattlogssearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txtattlogssearch.Text.Length >= 1)
                {
                    lvLogs.Items.Clear(); 
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtattlogssearch.Text))
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
                    ListView ll = new ListView();
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

                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }

            //try
            //{
            //    int i;
            //    if (txtattlogssearch.Text.Length > 1)
            //    {

            //        for (i = 0; i < lvLogs.Items.Count; i++)
            //        {

            //            if (lvLogs.Items[i].SubItems[1].ToString().Contains(txtattlogssearch.Text) || lvLogs.Items[i].SubItems[2].ToString().Contains(txtattlogssearch.Text))
            //            {
            //                lvLogs.Items[i].BackColor = Color.Navy;
            //                lvLogs.Items[i].ForeColor = Color.White;
            //            }
            //            else
            //            {
            //                lvLogs.Items[i].ForeColor = Color.Blue;
            //                lvLogs.Items[i].BackColor = Color.Gainsboro;
            //            }

            //        }
            //    }
            //    else
            //    {
            //        try
            //        {
            //            i = 0;
            //            for (i = 0; i < lvLogs.Items.Count; i++)
            //            {
            //                lvLogs.Items[i].ForeColor = Color.Blue;
            //                lvLogs.Items[i].BackColor = Color.Gainsboro;
            //            }
            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("---" + ex.ToString());
            //}

        }



        private void BtnDownloadTmp_Click(object sender, EventArgs e)
        {


            if (combofingerboxip.SelectedIndex >= 0)
            {
                Cursor = Cursors.WaitCursor;
                bIsConnected = axCZKEM1.Connect_Net(combofingerboxip.Text, Convert.ToInt32(txtPort.Text));
                if (bIsConnected == true)
                {
                    listfilter.Items.Clear();
                    comboMasterIp.SelectedIndex = -1;
                    btnConnect.Enabled = true;
                    comboMasterIp.Enabled = true;
                    btnConnect.Refresh();
                    btnConnect.Text = "Connect && Download ??";
                    bIsConnected = false;
                    lblState.Text = "Current State:DisConnected";
                    listViewFingerTemp.Items.Clear();
                    listViewFingerTemp.BeginUpdate();
                    MacIP = combofingerboxip.Text;
                    axCZKEM1.EnableDevice(iMachineNumber, false);

                    axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory

                    axCZKEM1.IsTFTMachine(iMachineNumber);   // to distingush machines
                    iIndex = 1;
                    axCZKEM1.ReadAllTemplate(iMachineNumber);//read all the users' fingerprint templates to the memory
                    while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sdwEnrollNumber, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
                    {
                        if (axCZKEM1.GetStrCardNumber(out sCardnumber))//get the card number from the memory  
                        {

                        }

                        for (idwFingerIndex = 0; idwFingerIndex < 10; idwFingerIndex++)
                        {
                            if (axCZKEM1.GetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, out iFlag, out sTmpData, out iTmpLength))//get the corresponding templates string and length from the memory
                            {

                                ListViewItem list = new ListViewItem();
                                list.Text = sdwEnrollNumber.ToString();
                                list.SubItems.Add(sName.ToUpper());
                                list.SubItems.Add(idwFingerIndex.ToString());
                                list.SubItems.Add(sTmpData);
                                list.SubItems.Add(sCardnumber);
                                list.SubItems.Add(iPrivilege.ToString());
                                list.SubItems.Add(sPassword);
                                if (bEnabled == true) { list.SubItems.Add("True"); } else { list.SubItems.Add("False"); }
                                list.SubItems.Add(iFlag.ToString());
                                list.SubItems.Add(iTmpLength.ToString());
                                if (iIndex % 2 == 0)
                                {
                                    list.BackColor = Color.White;
                                }
                                else
                                {
                                    list.BackColor = Color.WhiteSmoke;
                                }
                                iIndex++;
                                listViewFingerTemp.Items.Add(list);
                            }


                        }

                    }

                    listViewFingerTemp.EndUpdate();
                    if (listViewFingerTemp.Items.Count == 0)
                    {
                        MessageBox.Show("No Data Found in Machine.");
                    }
                    axCZKEM1.EnableDevice(iMachineNumber, true);
                    Cursor = Cursors.Default;
                    lblattcount.Text = "Total Employee Finger Rows Count  :" + listViewFingerTemp.Items.Count.ToString() + " and IP Addres   :" + combofingerboxip.Text;

                }
                else
                {
                    MessageBox.Show("Please select Ip in combo boxlist", "Error");
                }
            }
            else
            {
                MessageBox.Show("Pls select IP Address from Combobox");
            }
            Cursor = Cursors.Default;
            lblState.Text = "Current State:DisConnected";
        }
        private void BtnDownLoadFace_Click(object sender, EventArgs e)
        {

            //if (combofaceboxip.SelectedIndex >= 0)
            //{
            //    Cursor = Cursors.WaitCursor; lblattcount.Text = "";
            //    bIsConnected = axCZKEM1.Connect_Net(combofaceboxip.Text, Convert.ToInt32(txtPort.Text));
            //    if (bIsConnected == true)
            //    {
            //        listfilter.Items.Clear();
            //        lblState.Text = "Current State:Connected";
            //        lvFace.Items.Clear();
            //        lvFace.BeginUpdate();
            //        Cursor = Cursors.WaitCursor;
            //        axCZKEM1.EnableDevice(iMachineNumber, false);
            //        axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory

            //        while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sUserID, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
            //        {

            //            if (axCZKEM1.GetUserFaceStr(iMachineNumber, sUserID, iFaceIndex, ref sTmpData, ref iLength))//get the face templates from the memory
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = sUserID;
            //                list.SubItems.Add(sName.ToUpper());
            //                list.SubItems.Add(iFaceIndex.ToString());
            //                list.SubItems.Add(sTmpData.ToString());
            //                list.SubItems.Add(iPrivilege.ToString());
            //                list.SubItems.Add(sPassword);

            //                if (bEnabled == true)
            //                {
            //                    list.SubItems.Add("True");
            //                }
            //                else
            //                {
            //                    list.SubItems.Add("False");
            //                }
            //                list.SubItems.Add(iLength.ToString());
            //                lvFace.Items.Add(list);
            //                this.listfilter.Items.Add((ListViewItem)list.Clone());
            //            }
            //        }
            //        axCZKEM1.EnableDevice(iMachineNumber, true);
            //        lvFace.EndUpdate();
            //        if (lvFace.Items.Count == 0)
            //        {
            //            MessageBox.Show("No Data Found");
            //        }
            //        Cursor = Cursors.Default;
            //        lblattcount.Text = "Total Employee Face Rows Count  :" + lvFace.Items.Count.ToString() + " and IP Addres   :" + comboMasterIp.Text;
            //    }
            //    else
            //    {

            //        axCZKEM1.GetLastError(ref idwErrorCode);
            //        MessageBox.Show("Unable to connect the device,ErrorCode=" + idwErrorCode.ToString(), "Error");

            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Pls select IP Address from Combobox");
            //}
            //Cursor = Cursors.Default;
            //lblState.Text = "Current State:DisConnected";
        }

   
        private void btnUploadTmp_Click(object sender, EventArgs e)
        {
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error");
                return;
            }

            if (listViewFingerTemp.Items.Count == 0)
            {
                MessageBox.Show("There is no data to upload!", "Error");
                return;
            }
           

            Cursor = Cursors.WaitCursor;
            axCZKEM1.EnableDevice(iMachineNumber, false);
            for (int i = 0; i < listViewFingerTemp.Items.Count; i++)
            {
                sdwEnrollNumber = listViewFingerTemp.Items[i].SubItems[0].Text.Trim();
                sName = listViewFingerTemp.Items[i].SubItems[1].Text.Trim();
                idwFingerIndex = Convert.ToInt32(listViewFingerTemp.Items[i].SubItems[2].Text.Trim());
                sTmpData = listViewFingerTemp.Items[i].SubItems[3].Text.Trim();
                iPrivilege = Convert.ToInt32(listViewFingerTemp.Items[i].SubItems[4].Text.Trim());
                sPassword = listViewFingerTemp.Items[i].SubItems[5].Text.Trim();

                sEnabled = listViewFingerTemp.Items[i].SubItems[6].Text.Trim();
                iFlag = Convert.ToInt32(listViewFingerTemp.Items[i].SubItems[7].Text);
                if (sEnabled == "True")
                {
                    bEnabled = true;
                }
                else
                {
                    bEnabled = false;
                }


                if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//upload user information to the device
                {
                    axCZKEM1.SetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, iFlag, sTmpData);//upload templates information to the device
                }
                else
                {
                    axCZKEM1.GetLastError(ref idwErrorCode);
                    MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error");
                    Cursor = Cursors.Default;
                    axCZKEM1.EnableDevice(iMachineNumber, true);
                    return;
                }
            }
            axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
            Cursor = Cursors.Default;
            axCZKEM1.EnableDevice(iMachineNumber, true);
            MessageBox.Show("Successfully Upload fingerprint templates, " + "total:" + LvDownload.Items.Count.ToString(), "Success");
        }

    
        private void BtnDatabase_Click(object sender, EventArgs e)
        {
            //if (Class.Users.HCompcode == "AGF")
            //{
            //    try
            //    {


            //        listViewFingerTemp.Items.Clear(); lblattcount.Text = "";

            //        //axCZKEM1.EnableDevice(iMachineNumber, false);

            //        // select data from database
            //        UICO uti = new UICO();
            //        DataTable dt = uti.UploadDataTFT(combofingerboxip.Text);
            //        if (dt.Rows.Count > 0)
            //        {
            //            // start select data from database to upload in listview
            //            for (int i = 0; i < dt.Rows.Count; i++)
            //            {
            //                sdwEnrollNumber = string.IsNullOrEmpty(dt.Rows[i]["User_Id"].ToString()) ? " " : dt.Rows[i]["User_Id"].ToString();
            //                sName = string.IsNullOrEmpty(dt.Rows[i]["Name"].ToString()) ? " " : dt.Rows[i]["Name"].ToString();
            //                idwFingerIndex = string.IsNullOrEmpty(dt.Rows[i]["Finger_Index"].ToString()) ? 0 : Convert.ToInt32(dt.Rows[i]["Finger_Index"].ToString());
            //                sTmpData = string.IsNullOrEmpty(dt.Rows[i]["Finger_Image"].ToString()) ? " " : dt.Rows[i]["Finger_Image"].ToString();
            //                iPrivilege = string.IsNullOrEmpty(dt.Rows[i]["Privilege"].ToString()) ? 0 : Convert.ToInt32(dt.Rows[i]["Privilege"].ToString());
            //                sPassword = string.IsNullOrEmpty(dt.Rows[i]["Passwords"].ToString()) ? null : dt.Rows[i]["Passwords"].ToString();
            //                if (dt.Rows[i]["Enabled"].ToString() == "True")
            //                {
            //                    sEnabled = "True";
            //                }
            //                else
            //                {
            //                    sEnabled = "False";
            //                }
            //                int iFlag = Convert.ToInt32(dt.Rows[i]["Flag"].ToString());

            //                ListViewItem list = new ListViewItem();
            //                list.Text = sdwEnrollNumber.ToString();
            //                list.SubItems.Add(sName.ToUpper());
            //                list.SubItems.Add(idwFingerIndex.ToString());
            //                list.SubItems.Add(sTmpData);
            //                list.SubItems.Add(iPrivilege.ToString());
            //                list.SubItems.Add(sPassword);
            //                list.SubItems.Add(sEnabled.ToString());
            //                list.SubItems.Add(iFlag.ToString());
            //                listViewFingerTemp.Items.Add(list);
            //                lblattcount.Text = "Total Rows Count:  " + listViewFingerTemp.Items.Count.ToString();

            //            }
            //        }
            //        else
            //        {
            //            MessageBox.Show("No Data Found");
            //            lblattcount.Text = "Total Rows Count:  " + listViewFingerTemp.Items.Count.ToString();
            //        }


            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("-----" + ex.ToString());
            //    }
            //}
        }

        private void DeleteFpTm_Click(object sender, EventArgs e)
        {

            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error");
                return;
            }

            //UICO obj = new UICO();
            //obj.DeleteAllEmpTmTFT();
            MessageBox.Show("Successfully Data deleted from database");
        }

      

        private void Button2_Click(object sender, EventArgs e)
        {

        }

        private void BtnDelUserFace_Click(object sender, EventArgs e)
        {
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error");
                return;
            }

            if (cbUserIDFace.Text.Trim() == "")
            {
                MessageBox.Show("Please input the UserID and Face first!", "Error");
                return;
            }
            int idwErrorCode = 0;

            string sUserID = cbUserIDFace.Text.Trim();
            int iFaceIndex = 50;

            Cursor = Cursors.WaitCursor;
            if (axCZKEM1.DelUserFace(iMachineNumber, sUserID, iFaceIndex))
            {
                axCZKEM1.RefreshData(iMachineNumber);
                MessageBox.Show("DelUserFace,UserID=" + sUserID, "Success");

            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error");
            }
            Cursor = Cursors.Default;
        }

        private void BtnGetUserFace_Click(object sender, EventArgs e)
        {
           
            int idwErrorCode = 0;

            string sUserID = cbUserIDFace.Text.Trim();
            int iFaceIndex = 50;//the only possible parameter value
            int iLength = 128 * 1024;//initialize the length(cannot be zero)
            byte[] byTmpData = new byte[iLength];

            Cursor = Cursors.WaitCursor;
            axCZKEM1.EnableDevice(iMachineNumber, false);

            if (axCZKEM1.GetUserFace(iMachineNumber, sUserID, iFaceIndex, ref byTmpData[0], ref iLength))
            {
                //Here you can manage the information of the face templates according to your request.(for example,you can sava them to the database)
                MessageBox.Show("GetUserFace,the  length of the bytes array byTmpData is " + iLength.ToString(), "Success");
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error");
            }

            axCZKEM1.EnableDevice(iMachineNumber, true);
            Cursor = Cursors.Default;
        }

        

        private void BtnClearAdmin_Click(object sender, EventArgs e)
        {
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first", "Error");
                return;
            }
            int idwErrorCode = 0;

            Cursor = Cursors.WaitCursor;
            if (axCZKEM1.ClearAdministrators(iMachineNumber))
            {
                axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                MessageBox.Show("Successfully clear administrator privilege from teiminal!", "Success");


            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error");
            }
            Cursor = Cursors.Default;
        }

        private void BtnDeleteFaceTm_Click(object sender, EventArgs e)
        {
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error");
                return;
            }

            UICO obj = new UICO();
            obj.DeleteAllEmpTmIFACE_FaceTm();
            MessageBox.Show("Successfully Data deleted from database");
        }



      

        private void cardemtpy()
        {
            txtCardnumber.Text = ""; cbUserId_Card.Text = ""; txtName.Text = ""; txtcardfindex.Text = ""; txtcardfingerimage.Text = "";
            txtcardfaceimage.Text = ""; chbEnabled.Checked = false;
        }
        private void BtnGetStrCardNumber_Click(object sender, EventArgs e)
        {
            if (combocardreaderipbox.SelectedIndex >= 0)
            {
                listfilter.Items.Clear();
                Cursor = Cursors.WaitCursor; Class.Users.UserTime = 0;
                cardemtpy();
                btnConnect.Enabled = true;
                comboMasterIp.Enabled = true;
                btnConnect.Refresh();
                btnConnect.Text = "Connect && Download ??";
                bIsConnected = false;

                bIsConnected = axCZKEM1.Connect_Net(combocardreaderipbox.Text, Convert.ToInt32(txtPort.Text));
                if (bIsConnected == true)
                {
                    lvCard.Items.Clear();
                    lvCard.BeginUpdate();
                    listfilter.Items.Clear(); lblattcount.Text = "";


                    axCZKEM1.EnableDevice(iMachineNumber, false);
                    cbUserId_Card.Text = "";
                    cbPrivilegecard.Items.Clear();


                    axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory
                    axCZKEM1.IsTFTMachine(iMachineNumber);   // to distingush machines
                    axCZKEM1.ReadAllTemplate(iMachineNumber);//read all the users' fingerprint templates to the memory
                   

                    lblState.Text = "Current State:Connected";

                    int findex = 0;
                    string sss = "";
                    string ssss = "";
                    iIndex = 1;
                    // string uid = "";                   
                    while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sdwEnrollNumber, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
                    {

                       
                        if (axCZKEM1.GetStrCardNumber(out sCardnumber))//get the card number from the memory  
                        {
                           

                        }

                       // for (idwFingerIndex = 0; idwFingerIndex < 10; idwFingerIndex++)
                       // {
                            if (axCZKEM1.GetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, out iFlag, out sTmpData, out iTmpLength))//get the corresponding templates string and length from the memory
                            {
                                sss = sTmpData;

                                if (axCZKEM1.GetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, ref sTmpData1, ref iLength))//get the face templates from the memory
                                {
                                    ssss = sTmpData1;
                                }
                            }
                        
                                findex = idwFingerIndex;




                            ListViewItem list2 = new ListViewItem();
                            list2.SubItems.Add(iIndex.ToString());
                            list2.SubItems.Add(sdwEnrollNumber);
                            list2.SubItems.Add(sName.ToUpper());
                            list2.SubItems.Add(findex.ToString());
                            list2.SubItems.Add(sss);
                            list2.SubItems.Add(ssss);
                            list2.SubItems.Add(sCardnumber);
                            list2.SubItems.Add(iPrivilege.ToString());
                            list2.SubItems.Add(sPassword);
                            if (bEnabled == true) { list2.SubItems.Add("True"); } else { list2.SubItems.Add("False"); }
                            list2.SubItems.Add(iFlag.ToString());
                            this.listfilter.Items.Add((ListViewItem)list2.Clone());
                            lvCard.Items.Add(list2);
                            if (iIndex % 2 == 0)
                            {

                                list2.BackColor = Color.White;
                            }
                            else
                            {

                                list2.BackColor = Color.WhiteSmoke;
                            }
                            sss = "";
                            ssss = "";
                            iIndex++;

                       // }
                    }
                }
                lvCard.EndUpdate();
                axCZKEM1.EnableDevice(iMachineNumber, true); Cursor = Cursors.Default;
                lblattcount.Text = "Total Batch Card Rows Count  :" + lvCard.Items.Count.ToString() + " and IP Addres   :" + combocardreaderipbox.Text;
                if (lvCard.Items.Count == 0)
                {
                    MessageBox.Show("No Data Found");
                }
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Pls select IP Address from Combobox");
            }
            Cursor = Cursors.Default; lblState.Text = "Current State:DisConnected";
        }


        private void Txtfingertempsearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int item0 = 0;
            //    if (txtfingertempsearch.Text.Length >= 1)
            //    {
            //        listViewFingerTemp.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (listfilter.Items[item0].SubItems[0].ToString().Contains(txtfingertempsearch.Text) || listfilter.Items[item0].SubItems[1].ToString().Contains(txtfingertempsearch.Text))
            //            {


            //                list.Text = listfilter.Items[item0].SubItems[0].Text;
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
            //                listViewFingerTemp.Items.Add(list);


            //            }
            //            item0++;
            //        }

            //    }
            //    else
            //    {
            //        ListView ll = new ListView();
            //        listViewFingerTemp.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {


            //            this.listViewFingerTemp.Items.Add((ListViewItem)item.Clone());



            //            item0++;
            //        }

            //    }


            //}
            //catch (Exception ex)
            //{
            //    //MessageBox.Show("---" + ex.ToString());
            //}
           
        }

        private void Txtfacetempsearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int item0 = 0;
            //    if (txtfacetempsearch.Text.Length >= 1)
            //    {
            //        lvFace.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (listfilter.Items[item0].SubItems[0].ToString().Contains(txtfacetempsearch.Text) || listfilter.Items[item0].SubItems[1].ToString().Contains(txtfacetempsearch.Text))
            //            {


            //                list.Text = listfilter.Items[item0].SubItems[0].Text;
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
            //                lvFace.Items.Add(list);


            //            }
            //            item0++;
            //        }

            //    }
            //    else
            //    {
            //        ListView ll = new ListView();
            //        lvFace.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {


            //            this.lvFace.Items.Add((ListViewItem)item.Clone());



            //            item0++;
            //        }

            //    }


            //}
            //catch (Exception ex)
            //{
            //    //MessageBox.Show("---" + ex.ToString());
            //}
           

        }

        private void Txtcardreadersearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
               
                if (txtcardreadersearch.Text.Length >= 1)
                {
                    lvCard.Items.Clear(); 
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtcardreadersearch.Text) || listfilter.Items[item0].SubItems[3].ToString().Contains(txtcardreadersearch.Text))
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

                            //list.SubItems.Add(listfilter.Items[item0].SubItems[0].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);

                            lvCard.Items.Add(list);


                        }
                        item0++;
                    }

                }
                else
                {
                    ListView ll = new ListView();
                    lvCard.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {


                        this.lvCard.Items.Add((ListViewItem)item.Clone());



                        item0++;
                    }

                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }
            //try
            //{
            //    int i;
            //    if (txtcardreadersearch.Text.Length > 1)
            //    {

            //        for (i = 0; i < lvCard.Items.Count; i++)
            //        {

            //            if (lvCard.Items[i].SubItems[0].ToString().Contains(txtcardreadersearch.Text) || lvCard.Items[i].SubItems[1].ToString().Contains(txtcardreadersearch.Text) || lvCard.Items[i].SubItems[5].ToString().Contains(txtcardreadersearch.Text))
            //            {
            //                lvCard.Items[i].BackColor = Color.Navy;
            //                lvCard.Items[i].ForeColor = Color.White;
            //            }
            //            else
            //            {
            //                lvCard.Items[i].ForeColor = Color.Blue;
            //                lvCard.Items[i].BackColor = Color.Gainsboro;
            //            }

            //        }
            //    }
            //    else
            //    {
            //        try
            //        {
            //            i = 0;
            //            for (i = 0; i < lvCard.Items.Count; i++)
            //            {
            //                lvCard.Items[i].ForeColor = Color.Blue;
            //                lvCard.Items[i].BackColor = Color.Gainsboro;
            //            }
            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("---" + ex.ToString());
            //}


        }

      
    

        private void LvCard_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                cardemtpy();
                if (lvCard.Items.Count > 0)
                {
                    ListViewItem item1 = new ListViewItem();
                    item1.SubItems.Clear();
                    for (int c = 0; c < lvCard.SelectedItems[0].SubItems.Count; c++)
                    {
                        item1.SubItems.Add(lvCard.SelectedItems[0].SubItems[c].Text);


                    }

                    cbUserId_Card.Text = item1.SubItems[3].Text;
                    txtName.Text = item1.SubItems[4].Text;
                    txtcardfindex.Text = item1.SubItems[5].Text;
                    if (item1.SubItems[6].Text == "")
                    {
                        txtcardfingerimage.Text = "--";
                    }
                    else
                    {
                        txtcardfingerimage.Text = item1.SubItems[6].Text;
                    }
                    if (item1.SubItems[7].Text == "")
                    {
                        txtcardfaceimage.Text = "--";
                    }
                    else
                    {
                        txtcardfaceimage.Text = item1.SubItems[7].Text;
                    }

                    txtCardnumber.Text = item1.SubItems[8].Text;
                    cbPrivilegecard.Text = item1.SubItems[9].Text;
                    txtPassword.Text = item1.SubItems[10].Text;
                    if (item1.SubItems[11].Text == "True")
                        chbEnabled.Checked = true;
                    else
                        chbEnabled.Checked = false;


                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("---" + ex.ToString());
            }
        }





        private void Butloaddata_Click(object sender, EventArgs e)
        {
            if (comboremoveip.Text == "" || txtPort.Text.Trim() == "")
            {
                MessageBox.Show("IP and Port cannot be null", "Error");
                return;
            }
            Cursor = Cursors.WaitCursor; lblattcount.Refresh(); lblattcount.Text = "";
            bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == true)
            {


                lvremove.Items.Clear(); cbUserIDTmp.Items.Clear(); cbUserIDFace.Items.Clear(); cbUserIDDE.Items.Clear();
                axCZKEM1.EnableDevice(iMachineNumber, false);
                axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory
                axCZKEM1.IsTFTMachine(iMachineNumber);   // to distingush machines
                axCZKEM1.ReadAllTemplate(iMachineNumber);//read all the users' fingerprint templates to the memory     
                lvremove.BeginUpdate();iIndex = 1;
                while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sdwEnrollNumber, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
                {

                    for (idwFingerIndex = 0; idwFingerIndex < 10; idwFingerIndex++)
                    {
                        if (axCZKEM1.GetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, out iFlag, out sTmpData, out iTmpLength))//get the corresponding templates string and length from the memory
                        {
                            if (axCZKEM1.GetStrCardNumber(out sCardnumber))//get the card number from the memory  
                            {

                            }
                            ListViewItem list = new ListViewItem();

                            list.Text = sdwEnrollNumber.ToString();
                            list.SubItems.Add(sName.ToUpper());
                            list.SubItems.Add(idwFingerIndex.ToString());
                            list.SubItems.Add(sTmpData);
                            list.SubItems.Add(sCardnumber);
                            list.SubItems.Add(iPrivilege.ToString());
                            list.SubItems.Add(sPassword);
                            if (bEnabled == true) { list.SubItems.Add("True"); } else { list.SubItems.Add("False"); }
                            list.SubItems.Add(iFlag.ToString());
                            if (iIndex % 2 == 0) { list.BackColor = Color.White; }
                            else { list.BackColor = Color.WhiteSmoke; }
                            lvremove.Items.Add(list);


                            cbUserIDTmp.Items.Add(sdwEnrollNumber);
                            cbUserIDFace.Items.Add(sdwEnrollNumber);
                            cbUserIDDE.Items.Add(sdwEnrollNumber);

                        }
                    }

                }

                lvremove.EndUpdate();
                axCZKEM1.EnableDevice(iMachineNumber, true);
                Cursor = Cursors.Default;
                lblattcount.Text = "Total Employee Finger/Face Rows Count  :" + lvremove.Items.Count.ToString() + " and Connected IP Addres   :" + comboremoveip.Text;
                //}
            }
            else
            {

                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Unable to connect the device,ErrorCode=" + idwErrorCode.ToString(), "Error");
            }
            Cursor = Cursors.Default;
        }

        private void Lvremove_ItemActivate(object sender, EventArgs e)
        {


            try
            {
                if (lvremove.Items.Count > 0)
                {
                    ListViewItem item1 = new ListViewItem();
                    item1.SubItems.Clear();
                    for (int c = 0; c < lvremove.SelectedItems[0].SubItems.Count; c++)
                    {
                        item1.SubItems.Add(lvremove.SelectedItems[0].SubItems[c].Text);


                    }

                    cbUserIDDE.Text = item1.SubItems[1].Text;
                    cbUserIDTmp.Text = item1.SubItems[1].Text;
                    cbFingerIndex.Text = item1.SubItems[3].Text;
                    cbUserIDFace.Text = item1.SubItems[1].Text;


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("=---" + ex.ToString());
            }
        }





        private void Butcarddwnfrmdatabase_Click(object sender, EventArgs e)
        {
            if (Class.Users.HCompcode == "AGF")
            {
                try
                {
                    if (combocardreaderipbox.SelectedIndex >= 0)
                    {

                        lvCard.Items.Clear();
                        UICO uti = new UICO();
                        DataTable dt = uti.UploadDataTFT(combocardreaderipbox.Text);
                        if (dt.Rows.Count > 0)
                        {
                            cbUserId_Card.Text = "";
                            cbPrivilegecard.Items.Clear();
                            int i = 0;
                            for (i = 0; i < dt.Rows.Count; i++)
                            {
                                sdwEnrollNumber = dt.Rows[i]["User_Id"].ToString();
                                sName = dt.Rows[i]["Name"].ToString();
                                idwFingerIndex = Convert.ToInt32(dt.Rows[i]["FINGER_INDEX"].ToString());
                                sTmpData = dt.Rows[i]["Finger_Image"].ToString();
                                sTmpData1 = dt.Rows[i]["Face_Image"].ToString();
                                sCardnumber = dt.Rows[i]["cardnumber"].ToString();
                                iPrivilege = Convert.ToInt32(dt.Rows[i]["Privilege"].ToString());
                                sPassword = dt.Rows[i]["Passwords"].ToString();
                                if (dt.Rows[i]["Enabled"].ToString() == "True")
                                {
                                    sEnabled = "True";
                                }
                                else
                                {
                                    sEnabled = "False";
                                }

                                iFlag = Convert.ToInt32(dt.Rows[i]["Flag"].ToString());
                                MacIP = Convert.ToString(dt.Rows[i]["MACIP"].ToString());
                                ListViewItem list = new ListViewItem();
                                list.Text = sdwEnrollNumber.ToString();
                                list.SubItems.Add(sName);
                                list.SubItems.Add(idwFingerIndex.ToString());
                                list.SubItems.Add(sTmpData);
                                list.SubItems.Add(sTmpData1);
                                list.SubItems.Add(sCardnumber);
                                list.SubItems.Add(iPrivilege.ToString());
                                list.SubItems.Add(sPassword);
                                list.SubItems.Add(sEnabled.ToString());
                                list.SubItems.Add(iFlag.ToString());
                                list.SubItems.Add(MacIP);
                                lvCard.Items.Add(list);
                                cbUserId_Card.Text = sdwEnrollNumber.ToString();
                                cbPrivilegecard.Items.Add(idwFingerIndex.ToString());
                            }

                            lblcardreadercount.Text = "Total Rows Count from Database:  " + lvCard.Items.Count.ToString();
                            lblattcount.Text = "Total Batch Card Rows Count  :" + lvCard.Items.Count.ToString() + " and IP Addres   :" + combocardreaderipbox.Text;
                        }
                        else
                        {
                            MessageBox.Show("No Data Found");
                            lblattcount.Text = "Total Rows Count:  " + listViewFingerTemp.Items.Count.ToString();
                        }
                    }
                    else
                    {
                        MessageBox.Show("pls select Ip List from Combo box");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("--" + ex.ToString());
                }
            }
        }


     

     
        private void Txtremovesearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int i;
            //    if (txtremovesearch.Text.Length > 1)
            //    {

            //        for (i = 0; i < lvremove.Items.Count; i++)
            //        {

            //            if (lvremove.Items[i].SubItems[0].ToString().Contains(txtremovesearch.Text) || lvremove.Items[i].SubItems[1].ToString().Contains(txtremovesearch.Text))
            //            {
            //                lvremove.Items[i].BackColor = Color.Navy;
            //                lvremove.Items[i].ForeColor = Color.White;
            //            }
            //            else
            //            {
            //                lvremove.Items[i].ForeColor = Color.Blue;
            //                lvremove.Items[i].BackColor = Color.Gainsboro;
            //            }

            //        }
            //    }
            //    else
            //    {
            //        try
            //        {
            //            i = 0;
            //            for (i = 0; i < lvremove.Items.Count; i++)
            //            {
            //                lvremove.Items[i].ForeColor = Color.Blue;
            //                lvremove.Items[i].BackColor = Color.Gainsboro;
            //            }

            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("---" + ex.ToString());
            //}
           
        }

    

        private void Listviewchecklistip_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                Class.Users.UserTime = 0;
                if (listViewupload.Items.Count >= 0)
                {
                    ListViewItem it2 = new ListViewItem();
                    if (e.Item.Checked == true)
                    {


                        Cursor = Cursors.WaitCursor;
                        bIsConnected = axCZKEM1.Connect_Net(e.Item.SubItems[1].Text, Convert.ToInt32(txtPort.Text));

                        if (bIsConnected == true)
                        {

                            e.Item.SubItems[2].Text = "Connected";
                            it2.SubItems.Add(e.Item.SubItems[1].Text);
                            it2.SubItems.Add(e.Item.SubItems[2].Text);
                            it2.SubItems.Add(e.Item.Checked.ToString());
                            allip.Items.Add(it2);

                        }
                        else
                        {
                            MessageBox.Show("This IP   :" + e.Item.SubItems[1].Text + "     Not a BioMetric Machine.");
                        }
                        Cursor = Cursors.Default;
                    }
                    if (e.Item.Checked == false && e.Item.SubItems[2].Text == "Connected")
                    {
                        Cursor = Cursors.WaitCursor;
                        bIsConnected = false;


                        e.Item.SubItems[2].Text = "DisConnected";
                        //it2.SubItems.Add(e.Item.SubItems[1].Text);
                        //it2.SubItems.Add(e.Item.SubItems[2].Text);
                        //it2.SubItems.Add(e.Item.Checked.ToString());
                      //  allip.Items.Add(it2);
                        for (int c = 0; c < allip.Items.Count; c++)
                        {
                            if (listviewchecklistip.SelectedItems[0].SubItems[1].Text == allip.Items[c].SubItems[1].Text)
                            {
                                allip.Items[c].Remove();
                                c--;
                            }
                        }
                        Cursor = Cursors.Default;
                    }
                }
                
            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
           
        }



     

        private void Butdownfrmdb_Click(object sender, EventArgs e)
        {

        }
        private void Lvdownall_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListViewItem itt = new ListViewItem(); Class.Users.UserTime = 0;
            iIndex = Lvdownremove.Items.Count;progressBar1.Value = 0;
            if (e.Item.Checked == true)
            {

                itt.SubItems.Add(e.Item.SubItems[1].Text);
                itt.SubItems.Add(e.Item.SubItems[2].Text);
                itt.SubItems.Add(e.Item.SubItems[3].Text);
                itt.SubItems.Add(e.Item.SubItems[4].Text);
                itt.SubItems.Add(e.Item.SubItems[5].Text);
                itt.SubItems.Add(e.Item.SubItems[6].Text);
                itt.SubItems.Add(e.Item.SubItems[7].Text);
                itt.SubItems.Add(e.Item.SubItems[8].Text);
                itt.SubItems.Add(e.Item.SubItems[9].Text);
                itt.SubItems.Add(e.Item.SubItems[10].Text);
                itt.SubItems.Add(e.Item.SubItems[11].Text);
                itt.SubItems.Add(e.Item.SubItems[12].Text);

               // removeuserid.Items.Add(itt);
                if (iIndex % 2 == 0)
                {
                    itt.BackColor = Color.White;
                }
                else
                {
                    itt.BackColor = Color.WhiteSmoke;
                }
                Lvdownremove.Items.Add(itt);
                iIndex++;
            }
        }

        private void Lvdownall_ItemActivate(object sender, EventArgs e)
        {
            //progressBar1.Value = 0; iIndex = Lvdownremove.Items.Count;
            //if (Lvdownall.Items.Count > 0)
            //{
            //    ListViewItem item1 = new ListViewItem();
            //    for (int c = 0; c< Lvdownall.SelectedItems[0].SubItems.Count; c++)
            //    {
            //        item1.SubItems.Add(Lvdownall.SelectedItems[0].SubItems[c].Text);
                    
            //        removeuserid.Items.Add(Lvdownall.SelectedItems[0].SubItems[c].Text);
            //    }
            //    if (iIndex % 2 == 0)
            //    {
            //        item1.BackColor = Color.White;
            //    }
            //    else
            //    {
            //        item1.BackColor = Color.WhiteSmoke;
            //    }
            //    iIndex++;
            //    Lvdownremove.Items.Add(item1);
            //}
        }

        private void Butremoveall_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Do You want to Remove from Machine Finger Index??", "Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {

                    lblattcount.Text = ""; Class.Users.UserTime = 0;
                    if (Lvdownremove.Items.Count > 0)
                    {
                        Cursor = Cursors.WaitCursor;bool va = false; int cnt = 0;
                        for (int j = 0; j < allip2.Items.Count; j++)
                        {
                            if (allip2.Items.Count > 0)
                            {
                                if (allip2.Items[j].SubItems[1].Text.Length > 10)
                                {
                                    progressBar1.Minimum = 0;
                                    progressBar1.Maximum = Lvdownremove.Items.Count; lblprogress1.Text = "";
                                    bIsConnected = false; iMachineNumber = 1;
                                    bIsConnected = axCZKEM1.Connect_Net(allip2.Items[j].SubItems[1].Text, Convert.ToInt32(txtPort.Text));
                                    if (bIsConnected == true)
                                    {
                                        lblState.Text = "Current State:Connected";

                                        for (int i = 0; i < Lvdownremove.Items.Count; i++)
                                        {

                                            idwErrorCode = 0;
                                            iBackupNumber = 11; Class.Users.UserTime = 0;
                                            string dwEnrollNumber = Convert.ToString(Lvdownremove.Items[i].SubItems[2].Text);

                                            sUserID = Lvdownremove.Items[i].SubItems[2].Text;
                                            idwFingerIndex = Convert.ToInt32(Lvdownremove.Items[i].SubItems[4].Text);
                                            axCZKEM1.EnableDevice(iMachineNumber, true);

                                            if (axCZKEM1.SSR_DelUserTmpExt(iMachineNumber, sUserID, idwFingerIndex))
                                            {
                                                axCZKEM1.RefreshData(iMachineNumber);         
                                                decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(Lvdownremove.Items.Count)) * (i + 1);
                                                lblprogress1.Text = "Data Removing From Machine : " + (per).ToString("N0") + " %" + "ID Card No:-" + sUserID;
                                                va = true;
                                                lblprogress1.Refresh();
                                                progressBar1.Value = i + 1;
                                                axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                                                cbUserIDDE.SelectedIndex = -1; cbUserIDTmp.SelectedIndex = -1;
                                                cnt = +1;

                                            }
                                            if (axCZKEM1.DeleteEnrollData(iMachineNumber, Convert.ToInt32(sUserID), 1, iBackupNumber))
                                            {
                                                axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                                               
                                            }
                                        }


                                    }
                                 
                                }

                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                MessageBox.Show("Invalid");
                            }
                        }
                        if (allip2.Items.Count == 0)
                        {
                            listremovechecklistip.Items.Clear();
                            combo_RemoveIPload(); Cursor = Cursors.Default;
                            MessageBox.Show("IP list is empty.pls select any Ip from Listview", "Information", MessageBoxButtons.OK, MessageBoxIcon.Question);
                            return;
                        }
                        listViewupload.Items.Clear();
                        if (va == true)
                        {
                            MessageBox.Show("Data Removed Successfully. Total Count  : "+ cnt.ToString(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Question);
                        }
                        listViewupload.Items.Clear();
                        Cursor = Cursors.Default;
                    }
                    else
                    {
                        Cursor = Cursors.Default;
                        MessageBox.Show("Pls Select Data from ListView", "error");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("pls Connect Device", "error");
            }
            
            lblattcount.Text = "Total Employee Finger Rows Count  :" + Lvdownremove.Items.Count.ToString() + " and IP Addres   :" + comboMasterIp.Text;
            axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed }
            Cursor = Cursors.Default;
            lblState.Text = "Current State:DisConnected";
            progressBar1.Value = 0;
            Lvdownremove.Items.Clear();
            allip2.Items.Clear();
            lblprogress1.Text = "";
        }

       
        private void Listremovechecklistip_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                if (Lvdownremove.Items.Count >= 0)
                {
                    Class.Users.UserTime = 0;
                    ListViewItem itt = new ListViewItem();

                    if (e.Item.Checked == true)
                    {


                        Cursor = Cursors.WaitCursor;
                        bIsConnected = axCZKEM1.Connect_Net(e.Item.SubItems[1].Text, Convert.ToInt32(txtPort.Text));

                        if (bIsConnected == true)
                        {
                           
                            e.Item.SubItems[2].Text = "Connected";
                            itt.SubItems.Add(e.Item.SubItems[1].Text);
                            itt.SubItems.Add(e.Item.SubItems[2].Text);
                            itt.SubItems.Add(e.Item.Checked.ToString());
                            allip2.Items.Add(itt);

                        }
                        else
                        {
                            MessageBox.Show("This IP   :" + e.Item.SubItems[1].Text + "     Not a BioMetric Machine.");
                        }
                        Cursor = Cursors.Default;
                    }
                    if (e.Item.Checked == false && e.Item.SubItems[2].Text == "Connected")
                    {
                        Cursor = Cursors.WaitCursor;
                        bIsConnected = false;


                        e.Item.SubItems[2].Text = "DisConnected";
                        //itt.SubItems.Add(e.Item.SubItems[1].Text);
                        //itt.SubItems.Add(e.Item.SubItems[2].Text);
                        //itt.SubItems.Add(e.Item.Checked.ToString());
                        //allip2.Items.Add(itt);
                        for (int c = 0; c < allip2.Items.Count; c++)
                        {
                            if (listremovechecklistip.SelectedItems[0].SubItems[1].Text == allip2.Items[c].SubItems[1].Text)
                            {
                                allip2.Items[c].Remove();
                                c--;
                            }
                        }
                        Cursor = Cursors.Default;
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
           
        }

        private void Butfacedownload_Click(object sender, EventArgs e)
        {
            //if (Class.Users.HCompcode == "AGF")
            //{
            //    comboMasterIp.SelectedIndex = -1;
            //    comboMasterIp.Enabled = false;
            //    btnConnect.Enabled = false;
            //    lvFace.Items.Clear();
            //    UICO objupload = new UICO();

            //    DataTable dt = objupload.UploadDataIFACE_FingerTm(combofaceboxip.Text);
            //    if (dt.Rows.Count > 0)
            //    {
            //        Cursor = Cursors.WaitCursor;
            //        for (int i = 0; i < dt.Rows.Count; i++)
            //        {
            //            sdwEnrollNumber = string.IsNullOrEmpty(dt.Rows[i]["User_Id"].ToString()) ? " " : dt.Rows[i]["User_Id"].ToString();
            //            sName = string.IsNullOrEmpty(dt.Rows[i]["Name"].ToString()) ? " " : dt.Rows[i]["Name"].ToString();
            //            iFaceIndex = string.IsNullOrEmpty(dt.Rows[i]["Face_Index"].ToString()) ? 0 : Convert.ToInt32(dt.Rows[i]["Face_Index"].ToString());
            //            sTmpData = string.IsNullOrEmpty(dt.Rows[i]["Face_Image"].ToString()) ? " " : dt.Rows[i]["Face_Image"].ToString();
            //            iPrivilege = string.IsNullOrEmpty(dt.Rows[i]["Privilege"].ToString()) ? 0 : Convert.ToInt32(dt.Rows[i]["Privilege"].ToString());
            //            sPassword = string.IsNullOrEmpty(dt.Rows[i]["Passwords"].ToString()) ? null : dt.Rows[i]["Passwords"].ToString();
            //            sEnabled = string.IsNullOrEmpty(dt.Rows[i]["Enabled"].ToString()) ? " " : dt.Rows[i]["Enabled"].ToString();
            //            iLength = string.IsNullOrEmpty(dt.Rows[i]["Face_Length"].ToString()) ? 0 : Convert.ToInt32(dt.Rows[i]["Face_Length"].ToString());
            //            ListViewItem list = new ListViewItem();
            //            list.Text = sdwEnrollNumber;
            //            list.SubItems.Add(sName);
            //            list.SubItems.Add(iFaceIndex.ToString());
            //            list.SubItems.Add(sTmpData.ToString());
            //            list.SubItems.Add(iPrivilege.ToString());
            //            list.SubItems.Add(sPassword);

            //            if (sEnabled == "True")
            //            {
            //                list.SubItems.Add("True");
            //            }
            //            else
            //            {
            //                list.SubItems.Add("False");
            //            }
            //            list.SubItems.Add(iLength.ToString());

            //            lvFace.Items.Add(list);
            //        }

            //        Cursor = Cursors.Default;
            //        MessageBox.Show("Successfully upload Face templates , " + "total:" + dt.Rows.Count.ToString(), "Success");
            //    }
            //}
        }

        private void Butreset_Click(object sender, EventArgs e)
        {

            Cursor = Cursors.WaitCursor;
            if (comboremoveip.Text == "")
            {
                MessageBox.Show("pls Select IPAddress", "Error"); Cursor = Cursors.Default;
                return;
            }
            bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error"); Cursor = Cursors.Default;
                return;
            }
            if (bIsConnected == true)
            {
                axCZKEM1.RestartDevice(MyCount);

                MyCount = 1;

            }
            Cursor = Cursors.Default;


           
        }

      

        private void Combofaceboxip_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (combofaceboxip.SelectedIndex > 0)
            //{
            //    comboMasterIp.SelectedIndex = -1;
            //    btnConnect.Enabled = true;
            //    comboMasterIp.Enabled = true;
            //    btnConnect.Refresh();
            //    btnConnect.Text = "Connect && Download ??";
            //    bIsConnected = false;
            //}
        }

      

        private void Lvdownremove_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (Lvdownremove.Items.Count > 0)
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        int i = 0;
                        for (i = 0; i < Lvdownremove.Items.Count; i++)
                        {

                            if (Lvdownremove.Items[i].Selected)
                            {
                                MessageBox.Show("UserID:   " + Lvdownremove.Items[i].SubItems[1].Text + "      Name:  " + Lvdownremove.Items[i].SubItems[2].Text, "Delete this Record");

                                Lvdownremove.Items[i].Remove();
                                i--;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("pls upload data from Master IP", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

      
        private void Checkallrows_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                int i = 0;
                if (checkallrows.Checked == true)
                {
                    listViewupload.Items.Clear();
                    foreach (ListViewItem item in LvDownload.Items)
                    {
                        item.Selected = true;



                        ListViewItem item1 = new ListViewItem();
                        for (int c = 1; c < LvDownload.SelectedItems[i].SubItems.Count; c++)
                        {
                            item1.SubItems.Add(LvDownload.SelectedItems[i].SubItems[c].Text);


                           
                           
                        }
                      
                        if (i % 2 == 0)
                        {
                            item1.BackColor = Color.White;

                        }
                        else
                        {
                            item1.BackColor = Color.WhiteSmoke;


                        }
                        listViewupload.Items.Add(item1);
                        i++;
                    }
                }
                else
                {
                    foreach (ListViewItem item in LvDownload.Items)
                    {
                        item.Selected = false;

                    }
                    listViewupload.Items.Clear();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
  

        private void Checkremoveall_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                int i = 0;
                if (checkremoveall.Checked == true)
                {
                    Lvdownremove.Items.Clear();
                    //listremovechecklistip.Items.Clear();
                    foreach (ListViewItem item in Lvdownall.Items)
                    {
                        item.Selected = true;

                        

                        ListViewItem item1 = new ListViewItem();
                        for (int c = 1; c < Lvdownall.SelectedItems[i].SubItems.Count; c++)
                        {
                            item1.SubItems.Add(Lvdownall.SelectedItems[i].SubItems[c].Text);
                           
                        }
                        if (i % 2 == 0)
                        {
                            item1.BackColor = Color.White;

                        }
                        else
                        {
                            item1.BackColor = Color.WhiteSmoke;


                        }
                        Lvdownremove.Items.Add(item1);
                        i++;
                    }
                }
                else
                {
                    foreach (ListViewItem item in Lvdownall.Items)
                    {
                        item.Selected = false;

                    }
                    Lvdownremove.Items.Clear();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void Txtremovelog_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txtremovelog.Text.Length >= 1)
                {
                    Lvdownall.Items.Clear(); 
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtremovelog.Text) || listfilter.Items[item0].SubItems[3].ToString().Contains(txtremovelog.Text))
                        {


                          //  list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[10].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[11].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[12].Text);

                            if (item0 % 2 == 0)
                            {
                                list.BackColor = Color.White;

                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;

                            }
                            Lvdownall.Items.Add(list);

                

                        }
                        item0++;
                    }

                }
                else
                {

                    Lvdownall.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {


                        this.Lvdownall.Items.Add((ListViewItem)item.Clone());
                        if (item0 % 2 == 0)
                        {
                            item.BackColor = Color.White;

                        }
                        else
                        {
                            item.BackColor = Color.WhiteSmoke;

                        }


                        item0++;
                    }

                }   
            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }
      
        }
        
        private void Listviewattdown_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {

                Class.Users.UserTime = 0;
                ListViewItem it = new ListViewItem();
                if (e.Item.Checked == true)
                {                  
                    Cursor = Cursors.WaitCursor; Class.Users.UserTime = 0;
                    bIsConnected = axCZKEM1.Connect_Net(e.Item.SubItems[1].Text, Convert.ToInt32(txtPort.Text));
                    if (bIsConnected == true)
                    {
                        e.Item.SubItems[2].Text = "Connected";
                        it.SubItems.Add(e.Item.SubItems[1].Text);
                        it.SubItems.Add(e.Item.SubItems[2].Text);
                        it.SubItems.Add(e.Item.Checked.ToString());
                        allip1.Items.Add(it);

                    }
                    else
                    {
                        MessageBox.Show("This IP   :" + e.Item.SubItems[1].Text + "     Not a BioMetric Machine.");
                    }
                    Cursor = Cursors.Default;
                }
                if (e.Item.Checked == false && e.Item.SubItems[2].Text == "Connected")
                {
                    Cursor = Cursors.WaitCursor;
                    bIsConnected = false;


                    e.Item.SubItems[2].Text = "DisConnected";
                    //it.SubItems.Add(e.Item.SubItems[1].Text);
                    //it.SubItems.Add(e.Item.SubItems[2].Text);
                    //it.SubItems.Add(e.Item.Checked.ToString());
                    //allip1.Items.Add(it);
                    for (int c = 0; c < allip1.Items.Count; c++)
                    {
                        if (listviewattdown.SelectedItems[0].SubItems[1].Text == allip1.Items[c].SubItems[1].Text)
                        {
                            allip1.Items[c].Remove();
                            c--;
                        }
                    }
                    Cursor = Cursors.Default;
                }

            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

     

        private void ComboMasterIp_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboMasterIp.SelectedIndex > 0)
            {
                progressBar1.Value = 0; lblattcount.Text = ""; lblprogress1.Text = ""; listviewchecklistip.Items.Clear();
                Lvdownremove.Items.Clear();  listremovechecklistip.Items.Clear(); allip2.Items.Clear();
                Lvdownall.Items.Clear(); allip.Items.Clear(); allip1.Items.Clear();
                listviewchecklistip1.Items.Clear(); listviewchecklistip2.Items.Clear(); Class.Users.UserTime = 0;
            }
            
        }

        private void Comboremoveip_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboremoveip.SelectedIndex == -1)
            //{
            //}
            //else
            //{
            //    bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
            //    if (bIsConnected == true)
            //    {

                   
            //        lvremove.Items.Clear();
            //        axCZKEM1.EnableDevice(iMachineNumber, false);
            //        axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory
            //        axCZKEM1.IsTFTMachine(iMachineNumber);   // to distingush machines
            //        axCZKEM1.ReadAllTemplate(iMachineNumber);//read all the users' fingerprint templates to the memory     
            //        lvremove.BeginUpdate();
            //        while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sdwEnrollNumber, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
            //        {
            //            for (idwFingerIndex = 0; idwFingerIndex < 10; idwFingerIndex++)
            //            {
            //                if (axCZKEM1.GetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, out iFlag, out sTmpData, out iTmpLength))//get the corresponding templates string and length from the memory
            //                {
            //                    ListViewItem list = new ListViewItem();
            //                    list.Text = sdwEnrollNumber.ToString();
            //                    list.SubItems.Add(sName);
            //                    list.SubItems.Add(idwFingerIndex.ToString());
            //                    list.SubItems.Add(sTmpData);
            //                    list.SubItems.Add(iPrivilege.ToString());
            //                    list.SubItems.Add(sPassword);
            //                    if (bEnabled == true) { list.SubItems.Add("True"); } else { list.SubItems.Add("False"); }
            //                    list.SubItems.Add(iFlag.ToString());
            //                    lvremove.Items.Add(list);

            //                    cbFingerIndex.Items.Add(idwFingerIndex);
            //                    cbUserIDTmp.Items.Add(sdwEnrollNumber);
            //                    cbUserIDFace.Items.Add(sdwEnrollNumber);
            //                    cbUserIDDE.Items.Add(sdwEnrollNumber);
            //                }
            //            }

            //        }

            //        lvremove.EndUpdate();
            //        axCZKEM1.EnableDevice(iMachineNumber, true);
            //        Cursor = Cursors.Default;
            //        lblattcount.Text = "Total Employee Finger/Face Rows Count  :" + lvremove.Items.Count.ToString() + " and IP Addres   :" + comboremoveip.Text;
            //        //}
            //    }
            //    else
            //    {
            //        MessageBox.Show("Unable to connect the device", "Error");
            //    }
            //}

        }

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeviceCommunication_Load(sender,e); Class.Users.UserTime = 0;
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }

        private void RefreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                allip.Items.Clear(); allip1.Items.Clear(); allip2.Items.Clear();
                listviewchecklistip.Items.Clear();
                listremovechecklistip.Items.Clear(); 
                listviewattdown.Items.Clear();
                lvLogs.Items.Clear();comboMasterIp.SelectedIndex = -1;
                AttIPLoad();
                axCZKEM1.Disconnect();
                bIsConnected = false;
                btnConnect.Text = "Connect && Download ??";
                lblState.Text = "Current State:DisConnected";
                Cursor = Cursors.Default;
            }
            catch(Exception EX)
            {
                MessageBox.Show(EX.Message.ToString());
            }
        }

        private void Listviewuplaod_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void DownLoads_Click(object sender, EventArgs e)
        {
            
           Export2Excel();
        }

        private void Export2Excel()
        {
            try
            {
                string[] st = new string[lvLogs.Columns.Count];
                DirectoryInfo di = new DirectoryInfo(@"c:\Pinnacle");
                if (di.Exists == false)
                    di.Create();
                StreamWriter sw = new StreamWriter(@"c:\Pinnacle\'" + Class.Users.HCompcode + "'TodayAttLogs.xls", false);
                sw.AutoFlush = true;
                for (int col = 0; col < lvLogs.Columns.Count; col++)
                {
                    sw.Write("\t" + lvLogs.Columns[col].Text.ToString());
                }

                int rowIndex = 1;
                int row = 0;
                string st1 = "";
                for (row = 0; row < lvLogs.Items.Count; row++)
                {
                    if (rowIndex <= lvLogs.Items.Count)
                    
                        rowIndex++;
                    if (rowIndex == 2)
                    {
                        st1 = "\n";
                    }
                    else
                    {
                        st1 = ""; 
                    }
                    for (int col = 0; col < lvLogs.Columns.Count; col++)
                    {
                        st1 = st1 + "\t" + "" + lvLogs.Items[row].SubItems[col].Text.ToString();
                    }
                    sw.WriteLine(st1);
                }
                sw.Close();
                FileInfo fil = new FileInfo(@"c:\Pinnacle\'" + Class.Users.HCompcode + "'TodayAttLogs.xls");
                if (fil.Exists == true)
                    MessageBox.Show("DownLoad Completed. \n Folder-Name is :c:\\Pinnacle\\'" + Class.Users.HCompcode + "'TodayAttLogs.xls","Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
            }
        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Btncarddownload_Click(object sender, EventArgs e)
        {
            if (comboMasterIp.Text.Trim() == "" || txtPort.Text.Trim() == "")
            {
                MessageBox.Show("IP and Port cannot be null", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            allip.Items.Clear();checkallrows.Checked = false; checkremoveall.Checked = false;
            SecondtabControl2.SelectTab(tab2card); Class.Users.UserTime = 0;
            listremovechecklistip.Items.Clear(); checkCard.Checked = false; Class.Users.UserTime = 0;
            lblprogress1.Text = ""; listviewchecklistip.Items.Clear(); Lvdownremove.Items.Clear();
            int idwErrorCode = 0; lblattcount.Text = "";
            Cursor = Cursors.WaitCursor;
            string macip = "";
            if (btncarddownload.Text == "Card DisConnect")
            {
                axCZKEM1.Disconnect();
                bIsConnected = false;
                btncarddownload.Text = "Card Download ??";
                lblState.Text = "Current State:DisConnected";
                Cursor = Cursors.Default;
                return;
            }
            axCZKEM1.PullMode = 1;
            bIsConnected = axCZKEM1.Connect_Net(comboMasterIp.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == true)
            {
                LvDownload1.Items.Clear();
                allip.Items.Add(comboMasterIp.Text);                
                Lvdownall.Items.Clear(); lvCard.Items.Clear(); listfilter.Items.Clear();
                axCZKEM1.EnableDevice(iMachineNumber, false);
                Cursor = Cursors.WaitCursor;
                btncarddownload.Text = "Card DisConnect";
                lblState.Text = "Current State:Connected";
                axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory
                axCZKEM1.IsTFTMachine(iMachineNumber);   // to distingush machines
                axCZKEM1.ReadAllTemplate(iMachineNumber);//read all the users' fingerprint templates to the memory
                int r = 1;
                string sss = "";
                string ssss = "";
                string card = "";
               
                progressBar1.Minimum = 0;int i = 0;
                progressBar1.Maximum = i; 
                while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sdwEnrollNumber, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
                {
                   
                    macip = comboMasterIp.Text;
                    LvDownload1.BeginUpdate();
                    ListViewItem list = new ListViewItem();

                    if (axCZKEM1.GetStrCardNumber(out sCardnumber))//get the card number from the memory  
                    {

                    }

                    //if (axCZKEM1.GetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, out iFlag, out sTmpData, out iTmpLength))//get the corresponding templates string and length from the memory
                    //{
                    //    ssss = sTmpData;
                    //}



                    card = sCardnumber;



                    list.SubItems.Add(r.ToString());
                    list.SubItems.Add(sdwEnrollNumber);
                    //  list.Text = sdwEnrollNumber.ToString();
                    list.SubItems.Add(sName.ToUpper());
                    list.SubItems.Add(idwFingerIndex.ToString());
                    list.SubItems.Add(ssss);
                    list.SubItems.Add(sss);
                    list.SubItems.Add(card);
                    list.SubItems.Add(iPrivilege.ToString());
                    list.SubItems.Add(sPassword);
                    if (bEnabled == true) { list.SubItems.Add("True"); } else { list.SubItems.Add("False"); }
                    list.SubItems.Add(iFlag.ToString());
                    list.SubItems.Add(macip.ToString());

                    LvDownload1.Items.Add(list);
                    iIndex = LvDownload1.Items.Count;
                    this.listfilter.Items.Add((ListViewItem)list.Clone());

                    ListViewItem list1 = new ListViewItem();
                    list1.SubItems.Add(r.ToString());
                    list1.SubItems.Add(sdwEnrollNumber.ToUpper());

                    // list1.Text = sdwEnrollNumber.ToString();
                    list1.SubItems.Add(sName.ToUpper());
                    list1.SubItems.Add(idwFingerIndex.ToString());
                    list1.SubItems.Add(ssss);
                    list1.SubItems.Add(sss);
                    list1.SubItems.Add(card);
                    list1.SubItems.Add(iPrivilege.ToString());
                    list1.SubItems.Add(sPassword);
                    if (bEnabled == true) { list1.SubItems.Add("True"); } else { list1.SubItems.Add("False"); }
                    list1.SubItems.Add(iFlag.ToString());
                    list1.SubItems.Add(macip.ToString());
                    Lvdownall.Items.Add(list1);


                    ListViewItem list2 = new ListViewItem();

                    list2.SubItems.Add(r.ToString());
                    list2.SubItems.Add(sdwEnrollNumber.ToString());

                    list2.SubItems.Add(sName.ToUpper());
                    list2.SubItems.Add(idwFingerIndex.ToString());
                    list2.SubItems.Add(ssss);
                    list2.SubItems.Add(sss);
                    list2.SubItems.Add(card);
                    list2.SubItems.Add(iPrivilege.ToString());
                    list2.SubItems.Add(sPassword);
                    if (bEnabled == true) { list2.SubItems.Add("True"); } else { list2.SubItems.Add("False"); }
                    list2.SubItems.Add(iFlag.ToString());
                    list2.SubItems.Add(macip.ToString());
                    lvCard.Items.Add(list2);


                    if (r % 2 == 0)
                    {
                        list.BackColor = Color.White;
                        list1.BackColor = Color.White;
                        list2.BackColor = Color.White;
                    }
                    else
                    {
                        list.BackColor = Color.WhiteSmoke;
                        list1.BackColor = Color.WhiteSmoke;
                        list2.BackColor = Color.WhiteSmoke;
                    }




                    card = "";

                    //  decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(i)) * (i + 1);
                    lblprogress1.Text = "Data DownLoading From Machine '" + r.ToString() + "' ID Card No:-" + sdwEnrollNumber;
                    lblprogress1.Refresh();
                    LvDownload1.EndUpdate();
                    r++; 
                }
               
                Lvdownall.EndUpdate();
                lvCard.EndUpdate();
                axCZKEM1.EnableDevice(iMachineNumber, true);              
                lblattcount.Text = "Total Employee Finger Rows Count  :" + LvDownload1.Items.Count.ToString() + " and IP Addres   :" + comboMasterIp.Text;
                 allip1.Items.Clear();
                

            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode); Cursor = Cursors.Default;
                MessageBox.Show("Unable to connect the device,ErrorCode=" + idwErrorCode.ToString(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            listviewchecklistip1.Items.Clear();
            axCZKEM1.RefreshData(iMachineNumber);
            combo_ToIPload();
            Cursor = Cursors.Default; bIsConnected = false;
        }

        private void BtnfaceDownload_Click(object sender, EventArgs e)
        {
            if (comboMasterIp.Text.Trim() == "" || txtPort.Text.Trim() == "")
            {
                MessageBox.Show("IP and Port cannot be null", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SecondtabControl2.SelectTab(tab3face); listremovechecklistip.Items.Clear();
            lblprogress1.Text = ""; checkallrows.Checked = false; checkremoveall.Checked = false;
            int idwErrorCode = 0; lblattcount.Text = "";
            Cursor = Cursors.WaitCursor;
            string macip = ""; Class.Users.UserTime = 0;
            if (btnfaceDownload.Text == "Face DisConnect")
            {
                axCZKEM1.Disconnect();
                bIsConnected = false;
                btnfaceDownload.Text = "Face Download ??";
                lblState.Text = "Current State:DisConnected";

                Cursor = Cursors.Default;
                return;
            }
            axCZKEM1.PullMode = 1;
            bIsConnected = axCZKEM1.Connect_Net(comboMasterIp.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == true)
            {
                LvDownload2.Items.Clear(); Class.Users.UserTime = 0;
                LvDownload2.BeginUpdate(); iIndex = 1;
                Lvdownall.Items.Clear(); lvCard.Items.Clear(); listfilter.Items.Clear();
                axCZKEM1.EnableDevice(iMachineNumber, false);
                Cursor = Cursors.WaitCursor;
                btnfaceDownload.Text = "Face DisConnect";
                lblState.Text = "Current State:Connected";
                axCZKEM1.ReadAllUserID(iMachineNumber);//read all the user information to the memory
                axCZKEM1.IsTFTMachine(iMachineNumber);   // to distingush machines
                axCZKEM1.ReadAllTemplate(iMachineNumber);//read all the users' fingerprint templates to the memory          
                while (axCZKEM1.SSR_GetAllUserInfo(iMachineNumber, out sdwEnrollNumber, out sName, out sPassword, out iPrivilege, out bEnabled))//get all the users' information from the memory
                {
                    string sss = "";
                    string ssss = "";
                    string card = "";

                    macip = comboMasterIp.Text; 
                    for (idwFingerIndex = 0; idwFingerIndex < 9; idwFingerIndex++)
                    {
                        if (axCZKEM1.GetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, out iFlag, out sTmpData, out iTmpLength))//get the corresponding templates string and length from the memory
                        {
                            ssss = sTmpData;

                            if (axCZKEM1.GetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, ref sTmpData1, ref iLength))//get the face templates from the memory
                            {
                                sss = sTmpData1;

                                if (axCZKEM1.GetStrCardNumber(out sCardnumber))//get the card number from the memory  
                                {


                                }
                                ListViewItem list = new ListViewItem();

                                list.Text = sdwEnrollNumber.ToString();
                                list.SubItems.Add(sName.ToUpper());
                                list.SubItems.Add(idwFingerIndex.ToString());
                                list.SubItems.Add(sTmpData);
                                list.SubItems.Add(sTmpData1);
                                list.SubItems.Add(sCardnumber);
                                list.SubItems.Add(iPrivilege.ToString());
                                list.SubItems.Add(sPassword);
                                if (bEnabled == true) { list.SubItems.Add("True"); } else { list.SubItems.Add("False"); }
                                list.SubItems.Add(iFlag.ToString());
                                list.SubItems.Add(macip.ToString());
                                LvDownload2.Items.Add(list);

                                ListViewItem list1 = new ListViewItem();
                                list1.Text = sdwEnrollNumber.ToString();
                                list1.SubItems.Add(sName.ToUpper());
                                list1.SubItems.Add(idwFingerIndex.ToString());
                                list1.SubItems.Add(sTmpData);
                                list1.SubItems.Add(sTmpData1);
                                list1.SubItems.Add(sCardnumber);
                                list1.SubItems.Add(iPrivilege.ToString());
                                list1.SubItems.Add(sPassword);
                                if (bEnabled == true) { list1.SubItems.Add("True"); } else { list1.SubItems.Add("False"); }
                                list1.SubItems.Add(iFlag.ToString());
                                Lvdownall.Items.Add(list1);
                                this.listfilter.Items.Add((ListViewItem)list1.Clone());

                                if (iIndex % 2 == 0)
                                {
                                    list.BackColor = Color.White;
                                    list1.BackColor = Color.White;
                                }
                                else
                                {
                                    list.BackColor = Color.WhiteSmoke;
                                    list1.BackColor = Color.WhiteSmoke;
                                }
                                sss = "";
                                ssss = "";
                                card = ""; iIndex++;
                            }
                        }
                    }
                }
                LvDownload2.EndUpdate();
                Lvdownall.EndUpdate();
                lvCard.EndUpdate();
                axCZKEM1.EnableDevice(iMachineNumber, true);
                Cursor = Cursors.Default;
                lblattcount.Text = "Total Employee Finger Rows Count  :" + LvDownload2.Items.Count.ToString() + " and IP Addres   :" + comboMasterIp.Text;
                allip.Items.Clear(); allip1.Items.Clear();


            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Unable to connect the device,ErrorCode=" + idwErrorCode.ToString(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            listviewchecklistip2.Items.Clear();
           
            combo_ToIPload();
            Cursor = Cursors.Default; bIsConnected = false;
        }

        private void CheckCard_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                int i = 0;
                if (checkCard.Checked == true)
                {
                    listViewupload1.Items.Clear();
                    foreach (ListViewItem item in LvDownload1.Items)
                    {
                        item.Selected = true;



                        ListViewItem item1 = new ListViewItem();
                        for (int c = 1; c < LvDownload1.SelectedItems[i].SubItems.Count; c++)
                        {
                            item1.SubItems.Add(LvDownload1.SelectedItems[i].SubItems[c].Text);

                        }

                        listViewupload1.Items.Add(item1);
                        i++;
                    }
                }
                else
                {
                    foreach (ListViewItem item in LvDownload1.Items)
                    {
                        item.Selected = false;

                    }
                    listViewupload1.Items.Clear();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Checkface_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                int i = 0;
                if (checkface.Checked == true)
                {
                    listViewupload2.Items.Clear();
                    foreach (ListViewItem item in LvDownload2.Items)
                    {
                        item.Selected = true;
                        ListViewItem item1 = new ListViewItem();
                        for (int c = 1; c < LvDownload2.SelectedItems[i].SubItems.Count; c++)
                        {
                            item1.SubItems.Add(LvDownload2.SelectedItems[i].SubItems[c].Text);
                        }
                        listViewupload2.Items.Add(item1);
                        i++;
                    }
                }
                else
                {
                    foreach (ListViewItem item in LvDownload2.Items)
                    {
                        item.Selected = false;

                    }
                    listViewupload2.Items.Clear();


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void LvDownload1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListViewItem itt = new ListViewItem();
            iIndex = listViewupload1.Items.Count + 1; progressBar1.Value = 0; allip.Items.Clear();
            if (e.Item.Checked == true)
            {



                // itt.SubItems.Add(e.Item.SubItems[0].Text);

             //   itt.Text = iIndex.ToString();
                itt.SubItems.Add(iIndex.ToString());
                itt.SubItems.Add(e.Item.SubItems[2].Text);
                itt.SubItems.Add(e.Item.SubItems[3].Text);
                itt.SubItems.Add(e.Item.SubItems[4].Text);
                itt.SubItems.Add(e.Item.SubItems[5].Text);
                itt.SubItems.Add(e.Item.SubItems[6].Text);
                itt.SubItems.Add(e.Item.SubItems[7].Text);
                itt.SubItems.Add(e.Item.SubItems[8].Text);
                itt.SubItems.Add(e.Item.SubItems[9].Text);
                itt.SubItems.Add(e.Item.SubItems[10].Text);
                itt.SubItems.Add(e.Item.SubItems[11].Text);
                itt.SubItems.Add(e.Item.SubItems[12].Text);
                //  itt.SubItems.Add(e.Item.SubItems[13].Text);
                //  removeuserid.Items.Add(itt);
                if (iIndex % 2 == 0)
                {
                    itt.BackColor = Color.White;
                }
                else
                {
                    itt.BackColor = Color.WhiteSmoke;
                }
                listViewupload1.Items.Add(itt);
                iIndex++;
                lblattcount.Text = "Total Employee Finger Rows Count  :" + listViewupload1.Items.Count.ToString() + " and IP Addres   :" + comboMasterIp.Text;

            }
        }
        private void LvDownload1_ItemActivate(object sender, EventArgs e)
        {
            //try
            //{
            //    iIndex = listViewupload1.Items.Count;
            //    if (LvDownload1.Items.Count > 0)
            //    {
            //        ListViewItem item1 = new ListViewItem();
            //        for (int c = 0; c < LvDownload1.SelectedItems[0].SubItems.Count; c++)
            //        {
            //            item1.SubItems.Add(LvDownload1.SelectedItems[0].SubItems[c].Text);
            //            // item1.BackColor = Color.WhiteSmoke;


            //        }
            //        if (iIndex % 2 == 0)
            //        {
            //            item1.BackColor = Color.White;
            //        }
            //        else
            //        {
            //            item1.BackColor = Color.WhiteSmoke;
            //        }
            //        iIndex++;
            //        listViewupload1.Items.Add(item1);
            //    }


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        private void LvDownload2_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (LvDownload2.Items.Count > 0)
                {
                    ListViewItem itemlv2 = new ListViewItem();
                    for (int c = 0; c < LvDownload2.SelectedItems[0].SubItems.Count; c++)
                    {
                        itemlv2.SubItems.Add(LvDownload2.SelectedItems[0].SubItems[c].Text);
                       
                            itemlv2.BackColor = Color.WhiteSmoke;
                           
                        
                    }
                    listViewupload2.Items.Add(itemlv2);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Butcardtransfer_Click(object sender, EventArgs e)
        {
            try
            {
                txtcardtransfer.Text = "";
                if (listviewchecklistip1.CheckedItems.Count >= 0)
                {

                    DialogResult result = MessageBox.Show("Do You want to Export  Card Index??", "Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result.Equals(DialogResult.OK))
                    {
                        for (int j = 0; j < allip.Items.Count; j++)
                        {

                            if (allip.Items.Count >= 0)
                            {

                                Cursor = Cursors.WaitCursor;
                                if (allip.Items[j].SubItems[1].Text.Length > 10)
                                {
                                    lblattcount.Text = ""; bIsConnected = false;lblprogress1.Text = "";
                                    bIsConnected = axCZKEM1.Connect_Net(allip.Items[j].SubItems[1].Text, Convert.ToInt32(txtPort.Text));
                                    if (bIsConnected == true)
                                    {
                                        axCZKEM1.EnableDevice(iMachineNumber, false);
                                        lblState.Text = "Current State:Connected";
                                        int idwErrorCode = 0;
                                        int iFlag = 1;

                                        axCZKEM1.EnableDevice(iMachineNumber, false);

                                        int i = 0;
                                      
                                        progressBar1.Minimum = 0;
                                        progressBar1.Maximum = listViewupload1.Items.Count;
                                        for (i = 0; i < listViewupload1.Items.Count; i++)
                                        {

                                            txtcardfindex.Text = "6";

                                            sdwEnrollNumber = listViewupload1.Items[i].SubItems[2].Text;
                                            sName = "";
                                            idwFingerIndex = Convert.ToInt32("0" + txtcardfindex.Text);
                                            sTmpData = txtcardfingerimage.Text;
                                            sTmpData1 = txtcardfaceimage.Text;
                                            sCardnumber = listViewupload1.Items[i].SubItems[7].Text.Trim();
                                            iPrivilege = Convert.ToInt32("0" + cbPrivilegecard.Text);
                                            sPassword = txtPassword.Text;
                                          

                                            MacIP = allip.Items[j].SubItems[1].Text;

                                            if (sEnabled == "True")
                                            {
                                                bEnabled = true;

                                            }
                                            else
                                            {
                                                bEnabled = false;
                                            }

                                            sTmpData = "NO=Finger";
                                            axCZKEM1.SetStrCardNumber(sCardnumber);//Before you using function SetUserInfo,set the card number to make sure you can upload it to the device
                                            if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//face templates are part of users' information
                                            {
                                                axCZKEM1.SetUserTmpExStr(iMachineNumber, sdwEnrollNumber, idwFingerIndex, iFlag, sTmpData);//upload templates information to the memory

                                                //axCZKEM1.SetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, sTmpData1, iLength);//upload face templates information to the device
                                                lblattcount.Text = "Total Employee Finger Rows Count  : " + listViewupload1.Items.Count.ToString() + " and IP Addres   :" + allip.Items[j].SubItems[1].Text;

                                                decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(listViewupload1.Items.Count)) * (i + 1);
                                                lblprogress1.Text = " Data Transfer Machine to Machine : " + (per).ToString("N0") + " %";
                                                lblprogress1.Refresh();

                                                progressBar1.Value = i + 1;


                                            }
                                            //else
                                            //{
                                            //    Cursor = Cursors.Default;
                                            //    axCZKEM1.GetLastError(ref idwErrorCode);
                                            //    MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            //    axCZKEM1.EnableDevice(iMachineNumber, true); 
                                            //    return;
                                            //}


                                        }
                                        axCZKEM1.BatchUpdate(iMachineNumber);//upload all the information in the memory
                                        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed

                                        axCZKEM1.EnableDevice(iMachineNumber, true);

                                       


                                        MessageBox.Show("Successfully upload fingerprint, " + "total:" + listViewupload1.Items.Count.ToString() + "IP      :" + allip.Items[j].SubItems[0].Text, "Success");
                                        Cursor = Cursors.Default;
                                        progressBar1.Value = 0;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Machine DisConnected" + allip.Items[j].SubItems[1].Text);
                                    }

                                }


                                //j--;
                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                MessageBox.Show("Invalid");
                            }


                            Cursor = Cursors.Default;
                        }
                    }

                }
                else
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("Machine not connected.pls send IP Address", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                //Cursor = Cursors.Default;
                comboMasterIp.Enabled = true;

                btncarddownload.Enabled = true;
                lblState.Text = "Current State:DisConnected";
                //listviewchecklistip.Items.Clear(); 
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("pls Connect Device", "error");
            }
     
             listViewupload1.Items.Clear();
        }

        private void Listviewchecklistip1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {

                if (listViewupload1.Items.Count >= 0)
                {
                    ListViewItem it2 = new ListViewItem();
                    if (e.Item.Checked == true)
                    {


                        Cursor = Cursors.WaitCursor;
                        bIsConnected = axCZKEM1.Connect_Net(e.Item.SubItems[1].Text, Convert.ToInt32(txtPort.Text));

                        if (bIsConnected == true)
                        {

                            e.Item.SubItems[2].Text = "Connected";
                            it2.SubItems.Add(e.Item.SubItems[1].Text);
                            it2.SubItems.Add(e.Item.SubItems[2].Text);
                            it2.SubItems.Add(e.Item.Checked.ToString());
                            allip.Items.Add(it2);

                        }
                        else
                        {
                            MessageBox.Show("This IP   :" + e.Item.SubItems[1].Text + "     Not a BioMetric Machine.");
                        }
                        Cursor = Cursors.Default;
                    }
                    if (e.Item.Checked == false && e.Item.SubItems[2].Text == "Connected")
                    {
                        Cursor = Cursors.WaitCursor;
                        bIsConnected = false;


                        e.Item.SubItems[2].Text = "DisConnected";
                        //it2.SubItems.Add(e.Item.SubItems[1].Text);
                        //it2.SubItems.Add(e.Item.SubItems[2].Text);
                        //it2.SubItems.Add(e.Item.Checked.ToString());
                        //  allip.Items.Add(it2);
                        for (int c = 0; c < allip.Items.Count; c++)
                        {
                            if (listviewchecklistip1.SelectedItems[0].SubItems[1].Text == allip.Items[c].SubItems[1].Text)
                            {
                                allip.Items[c].Remove();
                                c--;
                            }
                        }
                        Cursor = Cursors.Default;
                    }
                }

            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

        private void Butfacetransfer_Click(object sender, EventArgs e)
        {
            try
            {
                txtfactransferesearch.Text = ""; Class.Users.UserTime = 0;
                if (listviewchecklistip2.CheckedItems.Count >= 0)
                {

                    DialogResult result = MessageBox.Show("Do You want to Export  Face Index??", "Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result.Equals(DialogResult.OK))
                    {                     

                        for (int j = 0; j < allip.Items.Count; j++)
                        {

                            if (allip.Items.Count >= 0)
                            {

                                Cursor = Cursors.WaitCursor;
                                if (allip.Items[j].SubItems[1].Text.Length > 10)
                                {
                                    lblattcount.Text = ""; bIsConnected = false;
                                    bIsConnected = axCZKEM1.Connect_Net(allip.Items[j].SubItems[1].Text, Convert.ToInt32(txtPort.Text));
                                    if (bIsConnected == true)
                                    {
                                       
                                        lblState.Text = "Current State:Connected";
                                        int idwErrorCode = 0;
                                        int iFlag = 1;

                                        axCZKEM1.EnableDevice(iMachineNumber, false);

                                        int i = 0;
                                        UICO uti = new UICO();
                                        progressBar1.Minimum = 0;
                                        progressBar1.Maximum = listViewupload2.Items.Count;
                                        for (i = 0; i < listViewupload2.Items.Count; i++)
                                        {

                                            //string ss = "";
                                            //string finger = "";
                                            //string face = "";
                                          
                                            string c = "";

                                            sdwEnrollNumber = listViewupload2.Items[i].SubItems[1].Text;
                                            sName = listViewupload2.Items[i].SubItems[2].Text;
                                            idwFingerIndex = Convert.ToInt32(listViewupload2.Items[i].SubItems[3].Text);
                                            sTmpData = listViewupload2.Items[i].SubItems[4].Text;
                                            sTmpData1 = listViewupload2.Items[i].SubItems[5].Text;
                                            sCardnumber = listViewupload2.Items[i].SubItems[6].Text;
                                            iPrivilege = Convert.ToInt32(listViewupload2.Items[i].SubItems[7].Text);
                                            sPassword = listViewupload2.Items[i].SubItems[8].Text;
                                            sEnabled = listViewupload2.Items[i].SubItems[9].Text;
                                            iFlag = Convert.ToInt32("0" + listViewupload2.Items[i].SubItems[10].Text);
                                            MacIP = allip.Items[j].SubItems[1].Text;    
                                            if (sEnabled == "True")
                                            {
                                                bEnabled = true;

                                            }
                                            else
                                            {
                                                bEnabled = false;
                                            }
                                            axCZKEM1.SetStrCardNumber(sCardnumber);//Before you using function SetUserInfo,set the card number to make sure you can upload it to the device

                                            if (axCZKEM1.SSR_SetUserInfo(iMachineNumber, sdwEnrollNumber, sName, sPassword, iPrivilege, bEnabled))//face templates are part of users' information
                                            {

                                                axCZKEM1.SetUserFaceStr(iMachineNumber, sdwEnrollNumber, iFaceIndex, sTmpData1, iLength);//upload face templates information to the device

                                            

                                                lblattcount.Text = "Total Employee Finger Rows Count  : " + listViewupload2.Items.Count.ToString() + " and IP Addres   :" + allip.Items[j].SubItems[1].Text;

                                                decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(listViewupload2.Items.Count)) * (i + 1);
                                                lblprogress1.Text = " Data Transfer Machine to Machine : " + (per).ToString("N0") + " %";
                                                lblprogress1.Refresh();

                                                progressBar1.Value = i + 1;

                                            }
                                            else
                                            {
                                                axCZKEM1.GetLastError(ref idwErrorCode);
                                                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error");
                                                Cursor = Cursors.Default;
                                                axCZKEM1.EnableDevice(iMachineNumber, true);
                                                return;
                                            }                                            
                                        }
                                        
                                        axCZKEM1.BatchUpdate(iMachineNumber);//upload all the information in the memory
                                        axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed

                                        axCZKEM1.EnableDevice(iMachineNumber, true);


                                        MessageBox.Show("Successfully upload fingerprint, " + "total:" + listViewupload2.Items.Count.ToString() + "IP      :" + allip.Items[j].SubItems[1].Text, "Success");
                                        Cursor = Cursors.Default;
                                        progressBar1.Value = 0;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Machine DisConnected" + allip.Items[j].SubItems[1].Text);
                                    }

                                }
                                //j--;
                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                MessageBox.Show("Invalid");
                            }


                            Cursor = Cursors.Default;
                        }
                    }

                }
                else
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("Machine not connected.pls send IP Address", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                //Cursor = Cursors.Default;
                comboMasterIp.Enabled = true;
                
                btnfaceDownload.Enabled = true;
                lblState.Text = "Current State:DisConnected";
                listviewchecklistip2.Items.Clear(); 
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("pls Connect Device", "error");
            }
            listViewupload2.Items.Clear();
        }

        private void Listviewchecklistip2_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                Class.Users.UserTime = 0;
                if (listViewupload2.Items.Count >= 0)
                {
                    ListViewItem it2 = new ListViewItem();
                    if (e.Item.Checked == true)
                    {


                        Cursor = Cursors.WaitCursor;
                        bIsConnected = axCZKEM1.Connect_Net(e.Item.SubItems[1].Text, Convert.ToInt32(txtPort.Text));

                        if (bIsConnected == true)
                        {

                            e.Item.SubItems[2].Text = "Connected";
                            it2.SubItems.Add(e.Item.SubItems[1].Text);
                            it2.SubItems.Add(e.Item.SubItems[2].Text);
                            it2.SubItems.Add(e.Item.Checked.ToString());
                            allip.Items.Add(it2);

                        }
                        else
                        {
                            MessageBox.Show("This IP   :" + e.Item.SubItems[1].Text + "     Not a BioMetric Machine.");
                        }
                        Cursor = Cursors.Default;
                    }
                    if (e.Item.Checked == false && e.Item.SubItems[2].Text == "Connected")
                    {
                        Cursor = Cursors.WaitCursor;
                        bIsConnected = false;


                        e.Item.SubItems[2].Text = "DisConnected";
                        //it2.SubItems.Add(e.Item.SubItems[1].Text);
                        //it2.SubItems.Add(e.Item.SubItems[2].Text);
                        //it2.SubItems.Add(e.Item.Checked.ToString());
                        //  allip.Items.Add(it2);
                        for (int c = 0; c < allip.Items.Count; c++)
                        {
                            if (listviewchecklistip2.SelectedItems[0].SubItems[1].Text == allip.Items[c].SubItems[1].Text)
                            {
                                allip.Items[c].Remove();
                                c--;
                            }
                        }
                        Cursor = Cursors.Default;
                    }
                }

            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

        private void Txtcardtransfer_TextChanged(object sender, EventArgs e)
        {
            try
            {
                
                int item0 = 0; Class.Users.UserTime = 0;
                if (txtcardtransfer.Text.Length >= 1)
                {
                     LvDownload1.Items.Clear();
                  
                    foreach (ListViewItem item in listfilter.Items)
                    {
                       
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtcardtransfer.Text) || listfilter.Items[item0].SubItems[3].ToString().Contains(txtcardtransfer.Text))
                        {

                            ListViewItem list = new ListViewItem();

                          
                            list.Text = item0.ToString();
                            // list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[10].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[11].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[12].Text);
                            if (item0 % 2 == 0)
                            {
                                list.BackColor = Color.White;

                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;

                            }
                            LvDownload1.Items.Add(list);

                        }
                        item0++; 
                       
                    }

                }
                else
                {
                    ListView ll = new ListView();
                    LvDownload1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {


                        this.LvDownload1.Items.Add((ListViewItem)item.Clone());

                        if (item0 % 2 == 0)
                        {
                            item.BackColor = Color.White;

                        }
                        else
                        {
                            item.BackColor = Color.WhiteSmoke;

                        }

                        item0++;
                    }

                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }


        }

        private void Txtfactransferesearch_TextChanged(object sender, EventArgs e)
        {
            try
            {

                int item0 = 0;
                if (txtfactransferesearch.Text.Length >= 1)
                {
                    LvDownload2.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {

                        if (listfilter.Items[item0].SubItems[0].ToString().Contains(txtfactransferesearch.Text) || listfilter.Items[item0].SubItems[1].ToString().Contains(txtfactransferesearch.Text))
                        {

                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
                            LvDownload2.Items.Add(list);


                        }
                        item0++;
                    }

                }
                else
                {
                    ListView ll = new ListView();
                    LvDownload2.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {


                        this.LvDownload2.Items.Add((ListViewItem)item.Clone());



                        item0++;
                    }

                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }
        }

        private void Tab5CardReader_Click(object sender, EventArgs e)
        {

        }

        private void listViewupload1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listViewupload1.Items.Count > 0)
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        int i = 0;
                        for (i = 0; i < listViewupload1.Items.Count; i++)
                        {

                            if (listViewupload1.Items[i].Selected)
                            {
                                MessageBox.Show("UserID:   " + listViewupload1.Items[i].SubItems[1].Text + "      Name:  " + listViewupload1.Items[i].SubItems[2].Text, "Delete this Record");

                                listViewupload1.Items[i].Remove();
                                i--;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("pls upload data from Master IP", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void listViewupload2_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listViewupload2.Items.Count > 0)
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        int i = 0;
                        for (i = 0; i < listViewupload2.Items.Count; i++)
                        {

                            if (listViewupload2.Items[i].Selected)
                            {
                                MessageBox.Show("UserID:   " + listViewupload2.Items[i].SubItems[1].Text + "      Name:  " + listViewupload2.Items[i].SubItems[2].Text, "Delete this Record");

                                listViewupload2.Items[i].Remove();
                                i--;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("pls upload data from Master IP", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error"); Cursor = Cursors.Default;
                return;
            }

            if (cbUserIDTmp.Text.Trim() == "" || cbFingerIndex.Text.Trim() == "")
            {
                MessageBox.Show("Please input the UserID and FingerIndex first!", "Error"); Cursor = Cursors.Default;
                return;
            }
            if (bIsConnected == true)
            {
                int idwErrorCode = 0; iMachineNumber = 1;
                string sUserID = cbUserIDTmp.Text.Trim();
                int iFingerIndex = Convert.ToInt32(cbFingerIndex.Text.Trim());
                axCZKEM1.EnableDevice(iMachineNumber, true);
                Cursor = Cursors.WaitCursor;
                if (axCZKEM1.SSR_DelUserTmpExt(iMachineNumber, sUserID, iFingerIndex))
                {
                    axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                    axCZKEM1.Disconnect(); Cursor = Cursors.Default;
                    MessageBox.Show("SSR_DelUserTmpExt,UserID:" + sUserID + " FingerIndex:" + iFingerIndex.ToString(), "Success");
                }
                else
                {

                    axCZKEM1.GetLastError(ref idwErrorCode); Cursor = Cursors.Default;
                    MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error");
                }
            }
            Cursor = Cursors.Default;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error");
                return;
            }

            if (cbUserIDDE.Text.Trim() == "" || cbBackupDE.Text.Trim() == "")
            {
                MessageBox.Show("Please input the UserID and BackupNumber first!", "Error");
                return;
            }
            int idwErrorCode = 0; 

            string sUserID = cbUserIDDE.Text.Trim();
            int iBackupNumber = Convert.ToInt32(cbBackupDE.Text.Trim());

           
            if (axCZKEM1.SSR_DeleteEnrollData(iMachineNumber, sUserID, iBackupNumber))
            {
                axCZKEM1.RefreshData(iMachineNumber);//the data in the device should be refreshed
                axCZKEM1.Disconnect();
                MessageBox.Show("DeleteEnrollData,UserID=" + sUserID + " BackupNumber=" + iBackupNumber.ToString(), "Success"); Cursor = Cursors.Default;
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                MessageBox.Show("Operation failed,ErrorCode=" + idwErrorCode.ToString(), "Error"); Cursor = Cursors.Default;
            }
            Cursor = Cursors.Default;
        }

        private void GetDeviceTime_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor; 
            if (comboremoveip.Text == "")
            {
                MessageBox.Show("pls Select IPAddress", "Error"); Cursor = Cursors.Default;
                return;
            }
            bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error"); Cursor = Cursors.Default;
                return;
            }
            if (bIsConnected == true)
            {
                if (axCZKEM1.GetDeviceTime(iMachineNumber, ref  idwYear, ref  idwMonth, ref  idwDay, ref  idwHour, ref idwMinute, ref idwSecond))
                {
                    labeltime.Text = "";

                    labeltime.Text = "Now Machine Date Time  " + comboremoveip.Text + "   --    " + idwDay + "-" + idwMonth + "-" + idwYear + "   -  " + idwHour + "-" + idwMinute + "-" + idwSecond;
                    axCZKEM1.Disconnect();
                }
            }
            Cursor = Cursors.Default;
        }



        private void poweroffdevice_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (comboremoveip.Text == "")
            {
                MessageBox.Show("pls Select IPAddress", "Error"); Cursor = Cursors.Default;
                return;
            }
            bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error"); Cursor = Cursors.Default;
                return;
            }
            if (bIsConnected == true)
            {
                axCZKEM1.PowerOffDevice(iMachineNumber); axCZKEM1.Disconnect();
                //  MessageBox.Show(dd.ToString());
            }
            Cursor = Cursors.Default;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
          
            
                axCZKEM1.PowerOnAllDevice();
               
            
           
            Cursor = Cursors.Default; return;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor; Class.Users.UserTime = 0;
            if (comboremoveip.Text == "")
            {
                MessageBox.Show("pls Select IPAddress", "Error"); Cursor = Cursors.Default;
                return;
            }
            bIsConnected = axCZKEM1.Connect_Net(comboremoveip.Text, Convert.ToInt32(txtPort.Text));
            if (bIsConnected == false)
            {
                MessageBox.Show("Please connect the device first!", "Error"); Cursor = Cursors.Default;
                return;
            }
            if (bIsConnected == true)
            {
                axCZKEM1.SetDeviceTime(iMachineNumber);
                
                  
                labeltime.Text = "Now Machine Date Time  " + comboremoveip.Text + "   --    " + System.DateTime.Now.ToString();
                axCZKEM1.Disconnect();
               

            }
            Cursor = Cursors.Default;
        }

        private void BtnDatabase_Click1_Click(object sender, EventArgs e)
        {

        }

        private void Searchs_Click(object sender, EventArgs e)
        {

        }

        private void txtfingertempsearch_TextChanged_1(object sender, EventArgs e)
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

        

        public void GridLoad()
        {
           
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //reportdocument.ExportToDisk(ExportFormatType.Excel, folderLocation + "-" + dtprint1.Rows[0]["FNAME"].ToString() + "-" + dtprint1.Rows[0]["PAYPERIOD"].ToString() + " PaySlip.xls");
        }

        private void SecondtabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
        }

        public void Searchs(int id)
        {
            throw new NotImplementedException();
        }

        public void Deletes(int id)
        {
            throw new NotImplementedException();
        }
    }

}
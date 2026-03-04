using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using CrystalDecisions.CrystalReports.Engine;
using System.Configuration;
using CrystalDecisions.Shared;
using System.Windows.Forms;
using System.IO.Ports;
using System.Text.RegularExpressions;
using System.Net.Sockets;
using SuperSimpleTcp;
using System.Text;

namespace Pinnacle.Transactions.CFM
{
    public partial class RawMaterialWeightCertificate : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        private static RawMaterialWeightCertificate _instance; string s = "";
        bool tabcheck = false;
        public static RawMaterialWeightCertificate Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new RawMaterialWeightCertificate();
                _instance.Font = Class.Users.FontName; 
                GlobalVariables.CurrentForm = _instance; return _instance;
               
            }
           
        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
       

        public RawMaterialWeightCertificate()
        {
            InitializeComponent(); 
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToShortTimeString().ToString());
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            tabcheck = false; Class.Users.PayPeriod = ""; Class.Users.Paramid = 0;
        }
        public void usercheck(string s, string ss, string sss)
        {
            try
            {
                DataTable dt1 = sm.headerdropdowns(s, ss, sss);
                if (dt1.Rows.Count > 0)
                {
                    if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                    {
                        for (int r = 0; r < dt1.Rows.Count; r++)
                        {

                           if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") 
                            { 
                                GlobalVariables.TreeButtons.Visible = false; txtgrossweight.Enabled = true; txttareweight.Enabled = true; } 
                            else {
                                GlobalVariables.TreeButtons.Visible = false; txttareweight.Enabled = false; txtgrossweight.Enabled = false;
                            }
                      

                        }
                    }


                }
                else
                {
                    MessageBox.Show("Invalid");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("usercheck: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private string readserialvalue;
        private Decimal wt = 0;
        bool valid = false; bool validprint = false;
        Models.Validate va = new Models.Validate();

        // ThulliamMdi mdi = new ThulliamMdi();

              SimpleTcpClient client; string[] words1; SimpleTcpServer server;
        private void RawMaterialWeightCertificate_Load(object sender, EventArgs e)
        {



            tabControl1.SelectTab(tabPageraw2);
            Models.Serial.PortIP = Class.Users.PortIP + ":" + Class.Users.PortNo;
            Models.Serial.PortNo = Class.Users.PortNo;
            server = new SimpleTcpServer(Models.Serial.PortIP);
            //server.Events.ClientConnected += Events_ClientConnected;
            //server.Events.ClientDisconnected += Events_ClientDisconnected;
            //server.Events.DataReceived += Events_DataReceived;
            client = new SimpleTcpClient(Models.Serial.PortIP);
            client.Events.DataReceived += Events_DataReceived;
           
            serialPort2.Close();
            txtsearch.Select();
            //txtgrossweight.Text = "0.00";
            //txttareweight.Text = "0.00";
            if (server.IsListening)
            {
                server.Start();
                //  server.Send(Models.Serial.PortIP, "0");

            }

        }
        private void Events_DataReceived(object sender, DataReceivedEventArgs e)
        {
            s = "";
            this.Invoke((MethodInvoker)delegate
           {

               s += Encoding.UTF8.GetString(e.Data.Array).Trim();
           });

        }

   
        public void News()
        {
            empty();
            Models.Serial.PortIP = Class.Users.PortIP + ":" + Class.Users.PortNo;
            Models.Serial.PortNo = Class.Users.PortNo;
            
            GridLoad();
            Godownmload(); Varityitemload(); ReceivedFormload();
            companyload();

            this.BackColor = Class.Users.BackColors;
            buttdelete.BackColor = Class.Users.BackColors;
            buttsave.BackColor = Class.Users.BackColors;
            buttprint.BackColor = Class.Users.BackColors;
            buttnew.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;

            buttsearch.Font = Class.Users.FontName;
            butfirstweight.Font = Class.Users.FontName;
            butgetall.Font = Class.Users.FontName;
            butnetweight.Font = Class.Users.FontName;
           
            try
            {
                tabControl1.SelectTab(tabPageraw1);  txtrawmetid.Text = "";
                txtvechileno.Focus(); combovisualstatus.SelectedIndex = 0; 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (tabcheck == false)
            {
                tabControl1.SelectTab(tabPageraw2);
                tabcheck = true;
            }
            else
            {
                tabControl1.SelectTab(tabPageraw1);
                tabcheck = true;
            }
            txtvechileno.Text = "";
        }


        protected override bool ProcessCmdKey(ref Message message, Keys keys)
        {
            switch (keys)
            {
                case Keys.N | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    News(); tabcheck = true; txtvechileno.Text = ""; tabControl1.SelectTab(tabPageraw2);
                    return true; // signal that we've processed this key
                case Keys.S | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Saves();
                    return true; // signal that we've processed this key
                case Keys.P | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Saves();
                    Prints();

                    return true; // signal that we've processed this key
                //case Keys.E | Keys.Control:
                //    // ... Process Shift+Ctrl+Alt+B ...
                //    updating = true;
                //    adding = false;
                //    EnableText();
                //    return true; // signal that we've processed this key
                case Keys.D | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Deletes();
                    tabControl1.SelectTab(tabPageraw2);
                    return true; // signal that we've processed this key
            }
            // run base implementation
            return base.ProcessCmdKey(ref message, keys);
        }


        public void ReceivedFormload()
        {
            try
            {
                string sel = "select a.asptblpartymasid,a.partycode,a.partyname from asptblpartymas a order by 2;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                DataTable dt = ds.Tables["asptblpartymas"];
                comboreceivedfrom.DataSource = dt;
                comboreceivedfrom.DisplayMember = "partyname";
                comboreceivedfrom.ValueMember = "asptblpartymasid";
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReceivedFormload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
      
        public void companyload()
        {
            try
            {
                string sel = "select a.gtcompmastid,a.compcode from  gtcompmast a  where a.ptransaction ='COMPANY' order by 2 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid"; combocompcode.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Finyearload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
               
                    autonumberload();
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void autonumberload()
        {
            try
            {

                string sel = "select max(a.asptblrawmaterialid1)+1 as id,b.compname from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  where b.ptransaction='COMPANY' and a.finyear='" + combofinyear.Text + "'  and b.compcode='" + combocompcode.Text + "' GROUP BY B.compname;  ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblrawmaterial");
                DataTable dt = ds.Tables["asptblrawmaterial"];
                int cnt = dt.Rows.Count;
                if (cnt == 0)
                {
                    string sel1 = "select b.gtcompmastid, b.compname from  gtcompmast b  where b.ptransaction='COMPANY'  and b.compcode='" + combocompcode.Text + "'; ";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                    DataTable dt1 = ds1.Tables["gtcompmast"];
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt1;

                    txtcertificateno.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + 1;
                    txtrawmetid1.Text = "1";
                }
                else
                {
                    combocompname.Text = dt.Rows[0]["compname"].ToString();
                    txtcertificateno.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                    txtrawmetid1.Text = dt.Rows[0]["id"].ToString();
                    // }
                }
                
            }
            catch (Exception ex)
            {
               // MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
       
        public void Varityitemload()
        {
            string sel = "select a.asptblitemmastid,a.itemname from asptblitemmast a where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblitemmast");
            DataTable dt = ds.Tables["asptblitemmast"];
          
            combovarietyitem.DisplayMember = "itemname";
            combovarietyitem.ValueMember = "asptblitemmastid";
            combovarietyitem.DataSource = dt;

        }
        public void Godownmload()
        {
            string sel = "select a.asptblgodwonmasid,a.godownname from asptblgodwonmas a where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblgodwonmas");
            DataTable dt = ds.Tables["asptblgodwonmas"];
            
            combogodown.DisplayMember = "godownname";
            combogodown.ValueMember = "asptblgodwonmasid"; 
            combogodown.DataSource = dt;
        }

        public void Saves()
        {
            try
            {
                Models.Validate va = new Models.Validate(); Int64 maxid = 0;
                Class.Users.PayPeriod = "";Class.Users.Paramid = 0;
                string sw = Convert.ToString(txttareweight.Text);
               
                string status1 = sw == "0.00" ? "P" : "C";
                if (txtcertificateno.Text == "")
                {
                    MessageBox.Show("'Certificate Number  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.txtcertificateno.Focus(); 
                    return;
                }
                if (combocompcode.SelectedValue == null)
                {
                    MessageBox.Show("'Company Name Field is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    combocompcode.Select(); 
                    return;
                }                
                if (va.IsStringNumberic(txtvechileno.Text) == false)
                {
                    MessageBox.Show("'VechileNo  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    txtvechileno.Select(); 
                    return;
                }
                if (va.IsInteger(txtnoofbags.Text) == false)
                {
                    MessageBox.Show("'No Of Bags  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    txtnoofbags.Select(); 
                    return;
                }
                if (va.IsInteger(combovarietyitem.SelectedValue.ToString()) == false)
                {
                    MessageBox.Show("'Variety/ItemName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    combovarietyitem.Select(); 
                    return;
                }
                if (txttripwagonno.Text == "")
                {
                    MessageBox.Show("'Trip/WagonNo Field  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txttripwagonno.Select(); 
                    return;
                }
                if (va.IsDecimal(txtgrossweight.Text) == false)
                {
                    MessageBox.Show("'Gross Weight Field  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtgrossweight.Focus();
                   
                }
                else
                {
                    validprint = false;
                    string words = txtvechileno.Text;
                    words = Regex.Replace(words, @" ", "");
                    txtvechileno.Text = words;
                    string dat = "";
                    dat = dateTimePicker1.Value.ToString("yyyy-MM-dd").Substring(0,10);
                    string dat1 = "";
                    dat1 = Convert.ToDateTime(System.DateTime.Now.ToString()).ToString("yyyy-MM-dd");
                    string times = Convert.ToDateTime(System.DateTime.Now.ToString()).ToString("hh-mm-ss");
                    string sel = "select asptblrawmaterialid    from  asptblrawmaterial   WHERE  finyear='" + combofinyear.Text + "' and compcode='" + combocompcode.SelectedValue.ToString() + "' and compname='" + combocompcode.SelectedValue.ToString() + "' and certificateno='" + txtcertificateno.Text + "' and  vechileno='" + txtvechileno.Text.ToUpper() + "' and datetime1='" + dat + "' and receivedFrom='" + comboreceivedfrom.SelectedValue.ToString() + "' and itemname='" + comboitem.Text + "' and itemnamevarity='" + combovarietyitem.SelectedValue.ToString() + "' and grossweight='" + txtgrossweight.Text + "'  and tareweight='" + txttareweight.Text + "' and netweight='" + txtnetweight.Text + "' and  noofbag='" + txtnoofbags.Text + "' and godownname='" + combogodown.SelectedValue.ToString() + "' and thirdpartyweight='" + txtthirdpartywt.Text + "' and tripwagonno='" + txttripwagonno.Text + "' and lotno='" + txtlotno.Text + "' and sampledby='" + txtsampledby.Text + "' and certifiedby='" + txtcertifiedby.Text + "' and visualstatus='" + combovisualstatus.Text + "' and delayreason='" + txtremarks.Text + "' and remarks='" + txtremarks.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblrawmaterial");
                    DataTable dt = ds.Tables["asptblrawmaterial"];

                    Class.Users.Paramid = 0; Class.Users.PayPeriod = "";
                  
                    if (dt.Rows.Count != 0)
                    {
                         validprint = true;
                        mas.pop("Already Exits", "Certificate No :" + txtcertificateno.Text, "NetWight :" + txtnetweight.Text);
                         tabControl1.SelectTab(tabPageraw2);
                        Class.Users.Paramid = Convert.ToInt64("0" + txtrawmetid.Text);
                        Class.Users.PayPeriod = txtvechileno.Text;
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtrawmetid.Text) == 0 || Convert.ToInt32("0" + txtrawmetid.Text) == 0)
                    {
                        string ins = "insert into asptblrawmaterial(asptblrawmaterialid1,finyear,compcode,compname,certificateno, vechileno,datetime1,receivedFrom,itemname,itemnamevarity,grossweight,tareweight,netweight, noofbag,godownname,thirdpartyweight,tripwagonno,lotno,sampledby,certifiedby,visualstatus,delayreason,compcode1, username,createdby, modifiedby,ipaddress,createdon)values('" + txtrawmetid1.Text + "','" + combofinyear.Text + "','" + combocompcode.SelectedValue.ToString() + "','" + combocompcode.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "', '" + txtvechileno.Text.ToUpper() + "','" + dat + "','" + comboreceivedfrom.SelectedValue.ToString() + "','" + comboitem.Text + "','" + combovarietyitem.SelectedValue.ToString() + "','" + Convert.ToString(txtgrossweight.Text) + "','" + Convert.ToDecimal("0" + txttareweight.Text).ToString() + "','" + Convert.ToString(txtnetweight.Text) + "', '" + txtnoofbags.Text + "','" + combogodown.SelectedValue.ToString() + "','" + Convert.ToDecimal("0" + txtthirdpartywt.Text) + "','" + txttripwagonno.Text.ToUpper() + "','" + txtlotno.Text + "','" + txtsampledby.Text.ToUpper() + "','" + txtcertifiedby.Text.ToUpper() + "','" + combovisualstatus.Text + "','" + txtremarks.Text.ToUpper() + "','" + Class.Users.COMPCODE + "', '" + Class.Users.USERID + "','" + Class.Users.HUserName + "', '','" + Class.Users.IPADDRESS + "' ,'" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' );";
                        Utility.ExecuteNonQuery(ins);

                        string inssap = "insert into saptable(purchaseid,Compcode,FinYear,TransType,LoginUser,Date,Time,WeighBridgeSlipNo,VehicleNo,TripNoWagonNo,FirstWeight,FirstWeightDate,FirstWeightTime,ThirdPartyWeight,NetWeight,Status)" +
                        "values('" + txtrawmetid1.Text + "','" + combocompcode.SelectedValue.ToString() + "','" + combofinyear.Text + "','Purchase','" + Class.Users.HUserName + "','" + dat + "','" + times + "','" + txtcertificateno.Text.ToUpper() + "','" + txtvechileno.Text.ToUpper() + "','" + txttripwagonno.Text.ToUpper() + "','" + Convert.ToInt64("0" + txtgrossweight.Text) + "','" + dat + "','" + times + "','" + Convert.ToDecimal("0" + txtthirdpartywt.Text) + "','" +  txtnetweight.Text + "','P' );";

                        Utility.ExecuteNonQuery(inssap);

                        MessageBox.Show("Record Saved Successfully " + txtcertificateno.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        validprint = true; Class.Users.Paramid = 0;
                  

                        string sel2 = "select max(asptblrawmaterialid) as asptblrawmaterialid   from  asptblrawmaterial   WHERE  finyear='" + combofinyear.Text + "' and compname='" + combocompcode.SelectedValue.ToString() + "' ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblrawmaterial");
                        DataTable dt2 = ds2.Tables["asptblrawmaterial"]; maxid = 0;
                        maxid = Convert.ToInt64(dt2.Rows[0]["asptblrawmaterialid"].ToString());
                        string ins1 = "insert into asptblrawmaterial2(asptblrawmaterialid,asptblrawmaterialid1,finyear,compcode,compname,certificateno, vechileno,datetime1,datetime2,receivedFrom,itemname,itemnamevarity,grossweight,tareweight,netweight, noofbag,godownname,thirdpartyweight,tripwagonno,lotno,sampledby,certifiedby,visualstatus,delayreason,compcode1, username,createdby, modifiedby,ipaddress,createdon,transaction, operation,remarks)values('" + maxid + "','" + txtrawmetid1.Text + "','" + combofinyear.Text + "','" + combocompcode.SelectedValue.ToString() + "','" + combocompcode.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "', '" + txtvechileno.Text.ToUpper() + "','" + Convert.ToDateTime(dat).ToString("dd-MM-yyyy") + " " + times + "','"+ dat1 + "','" + comboreceivedfrom.SelectedValue.ToString() + "','" + comboitem.Text + "','" + combovarietyitem.SelectedValue.ToString() + "','" + Convert.ToString(txtgrossweight.Text) + "','" + Convert.ToDecimal("0" + txttareweight.Text).ToString() + "','" + Convert.ToString(txtnetweight.Text) + "', '" + txtnoofbags.Text + "','" + combogodown.SelectedValue.ToString() + "','" + Convert.ToDecimal("0" + txtthirdpartywt.Text) + "','" + txttripwagonno.Text.ToUpper() + "','" + txtlotno.Text + "','" + txtsampledby.Text.ToUpper() + "','" + txtcertifiedby.Text.ToUpper() + "','" + combovisualstatus.Text + "','" + txtremarks.Text.ToUpper() + "','" + Class.Users.COMPCODE + "', '" + Class.Users.USERID + "','" + Class.Users.HUserName+ " : " + Class.Users.IPADDRESS  + "', '','" + Class.Users.IPADDRESS + "' ,'" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','" + Class.Users.ScreenName + "','INSERTED','"+txtremarks.Text+ "');";
                        Utility.ExecuteNonQuery(ins1);

                       

                        Class.Users.Paramid = Convert.ToInt64("0" + maxid);
                        Class.Users.PayPeriod = txtvechileno.Text;
                        mas.pop("Record Saved Successfully", "Certificate No :" + txtcertificateno.Text, "NetWight :" + txtnetweight.Text);
                        mas.pop("Record Saved Successfully", "Certificate No :" + txtcertificateno.Text, "NetWight :" + txtnetweight.Text);
                        Class.Users.Paramid = Convert.ToInt64("0" + txtrawmetid.Text);
                        GridLoad();
                        empty();txtvechileno.Text = "";
                        autonumberload(); tabControl1.SelectTab(tabPageraw2);
                    }
                    else
                    {
                        string up = "update  asptblrawmaterial  set asptblrawmaterialid1='" + txtrawmetid1.Text + "',finyear='" + combofinyear.Text + "',compcode='" + combocompcode.SelectedValue.ToString() + "',compname='" + combocompcode.SelectedValue.ToString() + "',certificateno='" + txtcertificateno.Text.ToUpper() + "', vechileno='" + txtvechileno.Text.ToUpper() + "',datetime1='" + dat + "',receivedFrom='" + comboreceivedfrom.SelectedValue.ToString() + "',itemname='" + comboitem.Text + "',itemnamevarity='" + combovarietyitem.SelectedValue.ToString() + "',grossweight='" + txtgrossweight.Text + "',tareweight='" + Convert.ToDecimal("0" + txttareweight.Text).ToString() + "',netweight='" + txtnetweight.Text + "', noofbag='" + txtnoofbags.Text + "',godownname='" + combogodown.SelectedValue.ToString() + "',thirdpartyweight='" + txtthirdpartywt.Text + "',tripwagonno='" + txttripwagonno.Text.ToUpper() + "',lotno='" + txtlotno.Text + "',sampledby='" + txtsampledby.Text.ToUpper() + "',certifiedby='" + txtcertifiedby.Text.ToUpper() + "',visualstatus='" + combovisualstatus.Text.ToUpper() + "',delayreason='" + txtremarks.Text.ToUpper() + "',compcode1='" + Class.Users.COMPCODE + "', username='" + Class.Users.USERID + "', modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' , modifiedon='" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "',remarks='" + txtremarks.Text + "'  where asptblrawmaterialid='" + txtrawmetid.Text + "';";
                        Utility.ExecuteNonQuery(up);

                        string ins = "insert into asptblrawmaterial2(asptblrawmaterialid,asptblrawmaterialid1,finyear,compcode,compname,certificateno, vechileno,datetime1,datetime2,receivedFrom,itemname,itemnamevarity,grossweight,tareweight,netweight, noofbag,godownname,thirdpartyweight,tripwagonno,lotno,sampledby,certifiedby,visualstatus,delayreason,compcode1, username,createdby, modifiedby,ipaddress,createdon,transaction, operation,remarks)values('" + txtrawmetid.Text + "','" + txtrawmetid1.Text + "','" + combofinyear.Text + "','" + combocompcode.SelectedValue.ToString() + "','" + combocompcode.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "', '" + txtvechileno.Text.ToUpper() + "','" +  Convert.ToDateTime(dat).ToString("dd-MM-yyyy") + " "+ times + "','" + dat1 + "','" + comboreceivedfrom.SelectedValue.ToString() + "','" + comboitem.Text + "','" + combovarietyitem.SelectedValue.ToString() + "','" + Convert.ToString(txtgrossweight.Text) + "','" + Convert.ToDecimal("0" + txttareweight.Text).ToString() + "','" + Convert.ToString(txtnetweight.Text) + "', '" + txtnoofbags.Text + "','" + combogodown.SelectedValue.ToString() + "','" + Convert.ToDecimal("0" + txtthirdpartywt.Text) + "','" + txttripwagonno.Text.ToUpper() + "','" + txtlotno.Text + "','" + txtsampledby.Text.ToUpper() + "','" + txtcertifiedby.Text.ToUpper() + "','" + combovisualstatus.Text + "','" + txtremarks.Text.ToUpper() + "','" + Class.Users.COMPCODE + "', '" + Class.Users.USERID + "','" + Class.Users.HUserName+" : " + Class.Users.IPADDRESS  + "', '','" + Class.Users.IPADDRESS + "' ,'" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','"+Class.Users.ScreenName+ "','UPDATED','" + txtremarks.Text + "' );";
                        Utility.ExecuteNonQuery(ins);
                        string inssap1 = "update  saptable set LoginUser='" + Class.Users.HUserName + "' , WeighBridgeSlipNo='" + txtcertificateno.Text.ToUpper() + "' , VehicleNo='" + txtvechileno.Text.ToUpper() + "',TripNoWagonNo='" + txttripwagonno.Text.ToUpper() + "',FirstWeight='" + txtgrossweight.Text + "',SecondWeight='" + Convert.ToDecimal(txttareweight.Text).ToString() + "',SecondWeightDate='" + dat + "',SecondWeightTime='" + times + "',ThirdPartyWeight='" + Convert.ToDecimal(txtthirdpartywt.Text) + "',NetWeight='" + txtnetweight.Text + "' ,Status='"+ status1 + "'  where purchaseid='" + txtrawmetid1.Text + "'";
                        Utility.ExecuteNonQuery(inssap1);

                        MessageBox.Show("Record Updated Successfully " + txtcertificateno.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        validprint = true;
                        GridLoad();
                        mas.pop("Record Updated Successfully", "Certificate No :" + txtcertificateno.Text, "NetWight :" + txtnetweight.Text);
                        Class.Users.Paramid = Convert.ToInt64("0"+txtrawmetid.Text);
                        Class.Users.PayPeriod = txtvechileno.Text;
                        empty();txtvechileno.Text = "";
                        autonumberload(); tabControl1.SelectTab(tabPageraw2);

                    }


                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Saves_Click " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RawMaterialWeightCertificate_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty(); tabcheck = false;
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }


        private void empty()
        {

            
            txtrawmetid.Text = "";s = "";words1 = null;
            txtrawmetid1.Text = "";
            combocompcode.Enabled = true;
            combofinyear.Text = Class.Users.Finyear;
          
            dateTimePicker1.Value = System.DateTime.Now;

            txtgrossweight.Text = "";
            txttareweight.Text = "";
            txtnetweight.Text = "";
            txtnoofbags.Text = "";
            txtthirdpartywt.Text = "";
            txttripwagonno.Text = "";
            txtlotno.Text = "";
            txtcertifiedby.Text = "";
            combovisualstatus.SelectedIndex = 0;
           
            
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptblrawmaterialid,a.vechileno,a.certificateno,c.partyname, date_format(a.datetime1,'%d-%m-%Y') as datetime1,a.grossweight,a.tareweight," +
                    "a.netweight,a.thirdpartyweight,a.tripwagonno,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid  where  a.datetime1='" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "'  order by a.asptblrawmaterialid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblrawmaterial");
                DataTable dt = ds.Tables["asptblrawmaterial"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblrawmaterialid"].ToString());
                        list.SubItems.Add(myRow["vechileno"].ToString());                     
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["partyname"].ToString());
                        list.SubItems.Add(myRow["datetime1"].ToString());
                        list.SubItems.Add(myRow["grossweight"].ToString());
                        list.SubItems.Add(myRow["tareweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());
                        list.SubItems.Add(myRow["thirdpartyweight"].ToString());
                        list.SubItems.Add(myRow["tripwagonno"].ToString());
                        list.SubItems.Add(myRow["delayreason"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                empty(); Class.Users.PayPeriod = ""; Class.Users.Paramid = 0;

                if (listView1.Items.Count >0)
                {
                   Class.Users.Paramid = 0;
                    txtrawmetid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    txtvechileno.Text = listView1.SelectedItems[0].SubItems[3].Text;                   
                    string sel1 = "select a.asptblrawmaterialid,a.asptblrawmaterialid1,a.finyear,b.compcode,b.compname,a.certificateno, a.vechileno,a.datetime1,substr(a.createdon,11,19) as times,c.partyname as  receivedFrom,a.itemname,d.itemname as itemnamevarity,a.grossweight,a.tareweight,a.netweight, a.noofbag,e.godownname,a.thirdpartyweight,a.tripwagonno,a.lotno,a.sampledby,a.certifiedby,a.visualstatus,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid  where a.asptblrawmaterialid='" + txtrawmetid.Text + "' and a.vechileNo='"+ txtvechileno.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblrawmaterial");
                    DataTable dt = ds.Tables["asptblrawmaterial"];

                    if (dt.Rows.Count > 0)
                    {
                       
                        txtrawmetid.Text = Convert.ToString(dt.Rows[0]["asptblrawmaterialid"].ToString());                      
                       txtrawmetid1.Text = Convert.ToString(dt.Rows[0]["asptblrawmaterialid1"].ToString());
                        txtcertificateno.Text= Convert.ToString(dt.Rows[0]["certificateno"].ToString());
                        combofinyear.Text= Convert.ToString(dt.Rows[0]["finyear"].ToString());
                        combocompcode.Text= Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        txtcertificateno.Text = Convert.ToString(dt.Rows[0]["certificateno"].ToString());
                        txtvechileno.Text = Convert.ToString(dt.Rows[0]["vechileno"].ToString());
                        dateTimePicker1.Text = Convert.ToString(dt.Rows[0]["datetime1"].ToString());
                        dateTimePicker2.Text = Convert.ToString(dt.Rows[0]["times"].ToString());
                        comboreceivedfrom.Text = Convert.ToString(dt.Rows[0]["receivedFrom"].ToString());
                        combovarietyitem.Text = Convert.ToString(dt.Rows[0]["itemnamevarity"].ToString());                      
                        txtgrossweight.Text = Convert.ToString(dt.Rows[0]["grossweight"].ToString());
                        txttareweight.Text = Convert.ToString(dt.Rows[0]["tareweight"].ToString());
                        txtnetweight.Text = Convert.ToString(dt.Rows[0]["netweight"].ToString());
                        txtnoofbags.Text = Convert.ToString(dt.Rows[0]["noofbag"].ToString());
                        combogodown.Text = Convert.ToString(dt.Rows[0]["godownname"].ToString());
                        txtthirdpartywt.Text = Convert.ToString(dt.Rows[0]["thirdpartyweight"].ToString());
                        txttripwagonno.Text = Convert.ToString(dt.Rows[0]["tripwagonno"].ToString());
                        txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());                      
                        txtsampledby.Text = Convert.ToString(dt.Rows[0]["sampledby"].ToString());
                        txtcertifiedby.Text = Convert.ToString(dt.Rows[0]["certifiedby"].ToString());
                        combovisualstatus.Text = Convert.ToString(dt.Rows[0]["visualstatus"].ToString());
                        txtremarks.Text = Convert.ToString(dt.Rows[0]["delayreason"].ToString());
                        combocompcode.Enabled = false; Class.Users.PayPeriod = txtvechileno.Text;
                        tabControl1.SelectTab(tabPageraw1);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
           

            try
            {
                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length >= 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[5].ToString().Contains(txtsearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[11].ToString().Contains(txtsearch.Text.ToUpper())  || listfilter.Items[item0].SubItems[12].ToString().Contains(txtsearch.Text.ToUpper()))
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
                            list.SubItems.Add(listfilter.Items[item0].SubItems[10].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[11].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[12].Text);
                
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {
                    GridLoad();
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }

        }



        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReceivedFormload(); usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }

        private void ContextMenuStrip2_Click(object sender, EventArgs e)
        {
            Godownmload(); Varityitemload(); 
        }

        private void butserialportsetting_Click(object sender, EventArgs e)
        {
        

        }
        public  void  portconnection(string porttype ,SerialPort _serialPort2)
        {
            try
            {
                porttype = Models.Serial.PortType;
                if (porttype == "TCP/IP")
                {
         
                }
                else
                {
                    _serialPort2.Close();
                    _serialPort2.PortName = Models.Serial.PortName;
                    _serialPort2.BaudRate = Convert.ToInt32("0" + Models.Serial.BaudRate);
                    _serialPort2.DataBits = Convert.ToInt32("0" + Models.Serial.DataBits);
                    _serialPort2.StopBits = (StopBits)Enum.Parse(typeof(StopBits), Models.Serial.StopBits);
                    _serialPort2.Parity = (Parity)Enum.Parse(typeof(Parity), Models.Serial.Parity);
                    _serialPort2.Open();
                    _serialPort2.DiscardInBuffer();
                    _serialPort2.DiscardOutBuffer();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Digitalizer not Connected." + ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            
        }
        private void butstart_Click(object sender, EventArgs e)
        {
            //portconnection(Models.Serial.PortType, this.serialPort2);
            try
            {
                if (Models.Serial.PortType != "TCP/IP")
                {


                    serialPort2.Close();
                    serialPort2.PortName = Models.Serial.PortName;
                    serialPort2.BaudRate = Convert.ToInt32("0" + Models.Serial.BaudRate);
                    serialPort2.DataBits = Convert.ToInt32("0" + Models.Serial.DataBits);
                    serialPort2.StopBits = (StopBits)Enum.Parse(typeof(StopBits), Models.Serial.StopBits);
                    serialPort2.Parity = (Parity)Enum.Parse(typeof(Parity), Models.Serial.Parity);
                    serialPort2.Open();
                    serialPort2.DiscardInBuffer();
                    serialPort2.DiscardOutBuffer();
                }
                else
                {

                  
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Digitalizer not Connected." + ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }
       
        private void btntateweight_Click(object sender, EventArgs e)
        {
            try
            {
                txttareweight.Text = "0";

                if (string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Vehicle Number is Empty");
                    return;
                }

                string[] words = null;

                if (Models.Serial.PortType == "TCP/IP")
                {
                    if (!client.IsConnected)
                        client.Connect();

                    client.Send("0");

                    string cleaned = Regex.Replace(s ?? "", @"Gross", "").Trim();
                    words = cleaned.Split(' ', (char)StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {
                    butstart_Click(sender, e);
                    string readValue = serialPort2.ReadLine()?.Trim();
                    serialPort2.Close();

                    words = readValue?.Split(' ', (char)StringSplitOptions.RemoveEmptyEntries);
                }

                // Only set tare when raw material exists
                if (txtrawmetid.Text != "" && TryGetTareWeight(words, out long tare))
                {
                    txttareweight.Text = tare.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Tare Weight Read Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }

           
            serialPort2.Close(); s = "";  words1 = null; 
        }
       
       
     

    

     

        private void btngetweight_Click(object sender, EventArgs e)
        {
            try
            {
                txtgrossweight.Text = "0";

                if (string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Vehicle Number is Empty");
                    return;
                }

                string[] words = null;

                if (Models.Serial.PortType == "TCP/IP")
                {
                    if (!client.IsConnected)
                        client.Connect();

                    client.Send("0");

                    string cleaned = Regex.Replace(s ?? "", @"Gross", "").Trim();
                    words = cleaned.Split(' ', (char)StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {
                    butstart_Click(sender, e);
                    string readValue = serialPort2.ReadLine()?.Trim();
                    words = readValue?.Split(' ', (char)StringSplitOptions.RemoveEmptyEntries);
                }

                if (txtrawmetid.Text == "" && TryGetWeight(words, out long grossWeight))
                {
                    txtgrossweight.Text = grossWeight.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Weight Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

           
            serialPort2.Close(); words1 = null;s = "";
        }
        private bool TryGetTareWeight(string[] words, out long tare)
        {
            tare = 0;
            if (words == null) return false;

            foreach (string word in words)
            {
                if (long.TryParse(word, out long value))
                {
                    tare = value;
                    return true;
                }
            }
            return false;
        }

        private bool TryGetWeight(string[] words, out long weight)
        {
            weight = 0;

            if (words == null) return false;

            foreach (string word in words)
            {
                if (long.TryParse(word, out long value) && value >= 0)
                {
                    weight = value;
                    return true;
                }
            }
            return false;
        }


        public void Deletes()
        {
            try
            {

                if (txtrawmetid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        string times = Convert.ToDateTime(System.DateTime.Now.ToString()).ToString("HH-mm-ss");
                        string dat = "";
                        dat = dateTimePicker1.Value.ToString("dd-MM-yyyy").Substring(0, 10);
                        string dat1 = "";
                        dat1 = Convert.ToDateTime(System.DateTime.Now.ToString()).ToString("yyyy-MM-dd");
                        string del = "delete from asptblrawmaterial  where asptblrawmaterialid='" + txtrawmetid.Text + "'";
                        Utility.ExecuteNonQuery(del);
                        string ins = "insert into asptblrawmaterial2(asptblrawmaterialid,asptblrawmaterialid1,finyear,compcode,compname,certificateno, vechileno,datetime1,datetime2,receivedFrom,itemname,itemnamevarity,grossweight,tareweight,netweight, noofbag,godownname,thirdpartyweight,tripwagonno,lotno,sampledby,certifiedby,visualstatus,delayreason,compcode1, username,createdby, modifiedby,ipaddress,createdon,transaction, operation,remarks,modifiedon)values('" + txtrawmetid.Text + "','" + txtrawmetid1.Text + "','" + combofinyear.Text + "','" + combocompcode.SelectedValue.ToString() + "','" + combocompcode.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "', '" + txtvechileno.Text.ToUpper() + "','" + Convert.ToDateTime(dat).ToString("dd-MM-yyyy") + " "+ times+ "','"+ dat1 + "','" + comboreceivedfrom.SelectedValue.ToString() + "','" + comboitem.Text + "','" + combovarietyitem.SelectedValue.ToString() + "','" + Convert.ToString(txtgrossweight.Text) + "','" + Convert.ToDecimal("0" + txttareweight.Text).ToString() + "','" + Convert.ToString(txtnetweight.Text) + "', '" + txtnoofbags.Text + "','" + combogodown.SelectedValue.ToString() + "','" + Convert.ToDecimal("0" + txtthirdpartywt.Text) + "','" + txttripwagonno.Text.ToUpper() + "','" + txtlotno.Text + "','" + txtsampledby.Text.ToUpper() + "','" + txtcertifiedby.Text.ToUpper() + "','" + combovisualstatus.Text + "','" + txtremarks.Text.ToUpper() + "','" + Class.Users.COMPCODE + "', '" + Class.Users.USERID + "','" + Class.Users.HUserName + "', '','" + Class.Users.IPADDRESS + "' ,'" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','"+ Class.Users.ScreenName + "','DELETED','" + txtremarks.Text + "','" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' );";
                        Utility.ExecuteNonQuery(ins);


                        string del2 = "delete from saptable  where WeighBridgeSlipNo='" + txtcertificateno.Text+"' and  purchaseid='" + txtrawmetid1.Text + "';";
                        Utility.ExecuteNonQuery(del2);
                        MessageBox.Show("Record Deleted Successfully " + txtrawmetid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        mas.pop("Record Deleted Successfully", "ID :" + txtrawmetid.Text, "");

                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txtgrossweight_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (txtgrossweight.Text != "")
                {
                   
                   // valid = va.IsIntegerdot(txtgrossweight.Text);                   
                    //if (valid == true)
                    //{
                        decimal wt = Convert.ToDecimal("0" + txtgrossweight.Text) - Convert.ToDecimal("0" + txttareweight.Text);
                        txtnetweight.Text = wt.ToString("0.00");
                   // }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //try
            //{
            //    decimal wt = Convert.ToDecimal("0" + txtgrossweight.Text) - Convert.ToDecimal("0" + txttareweight.Text);
            //    txtnetweight.Text = wt.ToString("0.000");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }
        private void txttareweight_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (txttareweight.Text != "")
                {
                    //valid = va.IsIntegerdot(txttareweight.Text);
                    //if (valid == false)
                    //{
                    //    MessageBox.Show("Numeric Only Allowed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    txttareweight.Focus(); txttareweight.Text = ""; return;
                    //}
                    //if (valid == true)
                    // {
                    //if (Convert.ToDecimal("0" + txtgrossweight.Text) <= Convert.ToDecimal("0" + txttareweight.Text))
                    //{
                    //    MessageBox.Show("Invalid Weight in TareWeight Field", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}
                    //if (Convert.ToDecimal("0" + txtgrossweight.Text) > Convert.ToDecimal("0" + txttareweight.Text))
                    //{

                        decimal wt = Convert.ToDecimal("0" + txtgrossweight.Text) - Convert.ToDecimal("0" + txttareweight.Text);
                        txtnetweight.Text = wt.ToString("0.00");
                    //}
                    //}
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //try
            //{
            //    decimal wt = Convert.ToDecimal("0" + txtgrossweight.Text) - Convert.ToDecimal("0" + txttareweight.Text);
            //    txtnetweight.Text = wt.ToString("0.000");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }


        public void Searchs()
        {
            tabControl1.SelectTab(tabPageraw2);
        }
        public void Prints()
        {
            try
            {
                
       
                if (Convert.ToInt64("0" + txtrawmetid.Text) == 0 || txtrawmetid.Text=="")
                {
                    string sel2 = " select max(a.asptblrawmaterialid) as asptblrawmaterialid from asptblrawmaterial a ;";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblrawmaterial");
                    DataTable dt2 = ds2.Tables["asptblrawmaterial"];
                    Class.Users.Paramid = Convert.ToInt64("0" + dt2.Rows[0]["asptblrawmaterialid"].ToString());
                    Pinnacle.Report.CFM.DosPrintRawMeterial p = new Report.CFM.DosPrintRawMeterial();
                    p.Show();
                }
                if (Convert.ToInt64("0" + txtrawmetid.Text) >= 1)
                {

                    Pinnacle.Report.CFM.DosPrintRawMeterial p = new Report.CFM.DosPrintRawMeterial();
                    p.Show();
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:'"+ txtrawmetid.Text.ToString() + "' " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
           
        }


        private void combocompname_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

 

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPageraw1"])//your specific tabname
            {
                if (txtrawmetid.Text == "")
                {
                    empty();
                    autonumberload();
                }
                txtvechileno.Select(); combovisualstatus.SelectedIndex = 0;
                Class.Users.UserTime = 0;
             

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPageraw2"])//your specific tabname
            {
               txtsearch.Select();
                Class.Users.UserTime = 0;
            }


            
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void combovarietyitem_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

   

        private void txtvechileno_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void txtvechileno_KeyDown(object sender, KeyEventArgs e)
        {
           if(e.KeyCode==Keys.Enter)
            {
                comboreceivedfrom.Focus();
         
            }
        }

        private void comboreceivedfrom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combovarietyitem.Select();
      
            }
        }

        private void txtgrossweight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtnoofbags.Select();
        
            }
        }

        private void combocompcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtvechileno.Focus();
        
            }
        }

        private void dateTimePicker1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void combovarietyitem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
               
         

                txtnoofbags.Select();
            }
        }

    
        private void btngetweight_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    txtnoofbags.Select();  
            //}
        }

        private void txtnoofbags_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
             
                txttripwagonno.Select();
               // txttareweight.Select(); txttareweight.BackColor = Color.Yellow;
            }
        }

        private void dateTimePicker2_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void txtgrossweight_KeyDown_1(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    txtnoofbags.Focus(); txtnoofbags.BackColor = Color.Yellow;
            //}
        }

        private void txttareweight_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    txttripwagonno.Focus();  
            //}
        }
        private void txttripwagonno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtthirdpartywt.Focus();
         
            }
        }

        private void txtthirdpartywt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combogodown.Focus();
              
            }
        }

        private void combogodown_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
               
                txtlotno.Focus();
            }
        }

        private void txtlotno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtsampledby.Focus();
         
            }
        }

        private void txtsampledby_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtcertifiedby.Focus();
        
            }
        }

        private void txtcertifiedby_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combovisualstatus.Focus();

            }
        }

        private void combovisualstatus_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtremarks.Focus();
         
            }
        }

        private void txtremarks_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                buttprint.Focus();
            }
        }
        private void buttsave_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                buttsave.Focus();
            }
        }
        private void buttsave_Click(object sender, EventArgs e)
        {
            Saves();
           
        }
        private void buttprint_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {            
                buttprint.Focus();
            }
        }
        private void buttsearch_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptblrawmaterialid,a.vechileno,a.certificateno,c.partyname, date_format(a.datetime1,'%d-%m-%Y') as datetime1,a.grossweight,a.tareweight," +
                    "a.netweight,a.thirdpartyweight,a.tripwagonno,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid   where a.datetime1=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')  order by a.asptblrawmaterialid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblrawmaterial");
                DataTable dt = ds.Tables["asptblrawmaterial"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblrawmaterialid"].ToString());
                        list.SubItems.Add(myRow["vechileno"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["partyname"].ToString());
                        list.SubItems.Add(myRow["datetime1"].ToString());
                        list.SubItems.Add(myRow["grossweight"].ToString());
                        list.SubItems.Add(myRow["tareweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());                      
                        list.SubItems.Add(myRow["thirdpartyweight"].ToString());
                        list.SubItems.Add(myRow["tripwagonno"].ToString());
                        list.SubItems.Add(myRow["delayreason"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());

                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


          
        }

        private void buttprint_Click(object sender, EventArgs e)
        {
           
            tabcheck = true;
            Saves();
            if (validprint == true)
            {
                Prints();
                Class.Users.Paramid = 0; Class.Users.PayPeriod = "";
                empty(); txtvechileno.Text = "";
            }
        }

    
        private void butfirstweight_Click(object sender, EventArgs e)
        {
            try
            {
               
                    listView1.Items.Clear(); listfilter.Items.Clear();
                int iGLCount = 1;

                    string sel1 = "select a.asptblrawmaterialid,a.vechileno,a.certificateno,c.partyname, date_format(a.datetime1,'%d-%m-%Y') as datetime1,a.grossweight,a.tareweight,a.netweight, a.tripwagonno,a.thirdpartyweight,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid  where  a.tareweight = 0 and a.datetime1=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptblrawmaterialid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblrawmaterial");
                    DataTable dt = ds.Tables["asptblrawmaterial"];
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(iGLCount.ToString());
                        list.SubItems.Add(myRow["asptblrawmaterialid"].ToString());
                        list.SubItems.Add(myRow["vechileno"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["partyname"].ToString());
                        list.SubItems.Add(myRow["datetime1"].ToString());
                        list.SubItems.Add(myRow["grossweight"].ToString());
                        list.SubItems.Add(myRow["tareweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());                     
                        list.SubItems.Add(myRow["thirdpartyweight"].ToString());
                        list.SubItems.Add(myRow["tripwagonno"].ToString());
                        list.SubItems.Add(myRow["delayreason"].ToString());
                       

                        list.BackColor = iGLCount % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        iGLCount++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                //else
                //{
                //    listView1.Items.Clear();

                //    GridLoad();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void butnetweight_Click(object sender, EventArgs e)
        {
            try
            {
              
                    listView1.Items.Clear(); listfilter.Items.Clear();
                int iGLCount = 1;

                    string sel1 = "select a.asptblrawmaterialid,a.vechileno,a.certificateno, c.partyname,date_format(a.datetime1,'%d-%m-%Y') as datetime1,a.grossweight,a.tareweight,a.netweight, a.tripwagonno,a.thirdpartyweight,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid where a.tareweight >=1  and a.datetime1=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptblrawmaterialid desc ;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblrawmaterial");
                    DataTable dt = ds.Tables["asptblrawmaterial"];
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(iGLCount.ToString());
                        list.SubItems.Add(myRow["asptblrawmaterialid"].ToString());
                        list.SubItems.Add(myRow["vechileno"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["partyname"].ToString());
                        list.SubItems.Add(myRow["datetime1"].ToString());
                        list.SubItems.Add(myRow["grossweight"].ToString());
                        list.SubItems.Add(myRow["tareweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());

                        list.SubItems.Add(myRow["thirdpartyweight"].ToString());
                        list.SubItems.Add(myRow["tripwagonno"].ToString());
                        list.SubItems.Add(myRow["delayreason"].ToString());
                        listView1.Items.Add(list);

                        list.BackColor = iGLCount % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        iGLCount++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void butgetall_Click(object sender, EventArgs e)
        {
            txtsearch.Text = "";

            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();

                string sel1 = "select a.asptblrawmaterialid,a.vechileno,a.certificateno,c.partyname, date_format(a.datetime1,'%d-%m-%Y') as datetime1,a.grossweight,a.tareweight," +
                   "a.netweight,a.thirdpartyweight,a.tripwagonno,a.delayreason from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid  where a.datetime1=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptblrawmaterialid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblrawmaterial");
                DataTable dt = ds.Tables["asptblrawmaterial"];


                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblrawmaterialid"].ToString());
                        list.SubItems.Add(myRow["vechileno"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["partyname"].ToString());
                        list.SubItems.Add(myRow["datetime1"].ToString());
                        list.SubItems.Add(myRow["grossweight"].ToString());
                        list.SubItems.Add(myRow["tareweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());                       
                        list.SubItems.Add(myRow["thirdpartyweight"].ToString());
                        list.SubItems.Add(myRow["tripwagonno"].ToString());
                        list.SubItems.Add(myRow["delayreason"].ToString());
                        
                        this.listfilter.Items.Add((ListViewItem)list.Clone());

                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void buttsearch_Click_1(object sender, EventArgs e)
        {
            Txtsearch_TextChanged(sender,e);
        }

        private void txtcertifiedby_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtlotno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtnoofbags_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }
        private void txtsampledby_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtvechileno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtremarks_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txttripwagonno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '/' || e.KeyChar == (char)Keys.Back);
        }

        private void txtthirdpartywt_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtgrossweight_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txttareweight_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtthirdpartywt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void enableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            combocompcode.Enabled = true;
        }

        private void previousCertificateNoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (txtrawmetid.Text != "")
            {
                string sel1 = "select a.asptblrawmaterialid,a.asptblrawmaterialid1,a.finyear,b.compcode,b.compname,a.certificateno from asptblrawmaterial a join gtcompmast b on a.compname=b.gtcompmastid  join asptblpartymas c on a.receivedFrom =c.asptblpartymasid join asptblitemmast d on a.itemnamevarity=d.asptblitemmastid join asptblgodwonmas e on a.godownname=e.asptblgodwonmasid   where a.asptblrawmaterialid='" + txtrawmetid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblrawmaterial");
                DataTable dt = ds.Tables["asptblrawmaterial"];

                if (dt.Rows.Count > 0)
                {
                   
                    txtrawmetid1.Text = Convert.ToString(dt.Rows[0]["asptblrawmaterialid1"].ToString());
                    txtcertificateno.Text = Convert.ToString(dt.Rows[0]["certificateno"].ToString());
                    combofinyear.Text = Convert.ToString(dt.Rows[0]["finyear"].ToString());
                }
            }
        }

        private void tabPageraw1_Click(object sender, EventArgs e)
        {

        }

       

       
        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

       
        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        public void Imports()
        {
            throw new NotImplementedException();
        }

        public void Pdfs()
        {
            throw new NotImplementedException();
        }

        public void ChangePasswords()
        {
            throw new NotImplementedException();
        }

        public void DownLoads()
        {
            throw new NotImplementedException();
        }

        public void ChangeSkins()
        {
            throw new NotImplementedException();
        }

        public void Logins()
        {
            throw new NotImplementedException();
        }

        public void GlobalSearchs()
        {
            throw new NotImplementedException();
        }

        public void TreeButtons()
        {
            throw new NotImplementedException();
        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Models.Serial.PortIP = Class.Users.PortIP + ":" + Class.Users.PortNo;
            Models.Serial.PortNo = Class.Users.PortNo;
            server = new SimpleTcpServer(Models.Serial.PortIP);
            client = new SimpleTcpClient(Models.Serial.PortIP);
            client.Events.DataReceived += Events_DataReceived;
            serialPort2.Close();
            txtsearch.Select();
            if (server.IsListening)
            {
                server.Start();
            }
        }

        private void buttdelete_Click(object sender, EventArgs e)
        {
            Deletes();  tabControl1.SelectTab(tabPageraw2); txtvechileno.Text = "";
        }

        private void buttnew_Click(object sender, EventArgs e)
        {
            News(); tabControl1.SelectTab(tabPageraw2); txtvechileno.Text = "";
        }
    }
}

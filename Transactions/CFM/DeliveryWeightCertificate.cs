using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net.Sockets;
using SuperSimpleTcp;

namespace Pinnacle.Transactions.CFM
{
    public partial class DeliveryWeightCertificate : Form,ToolStripAccess
    {
        private static DeliveryWeightCertificate _instance;
        ListView listfilter = new ListView();
        public static DeliveryWeightCertificate Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DeliveryWeightCertificate();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }

        }
        SimpleTcpClient client; string [] words;SimpleTcpServer server;
        bool tabcheck = false;
        public DeliveryWeightCertificate()
        {
            InitializeComponent();
            tabcheck = false; Class.Users.PayPeriod = ""; Class.Users.Paramid = 0;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
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

                            if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { GlobalVariables.TreeButtons.Visible = false; txtfirstweight.Enabled = true; txtsecondweight.Enabled = true; } else { GlobalVariables.TreeButtons.Visible = false; txtfirstweight.Enabled = false; txtsecondweight.Enabled = false; }


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
        decimal listview2totalweight = 0;
        Models.Validate va = new Models.Validate();
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        int i = 1;
        int invalid = 0;
        Int32 gridid = 0;
        bool validprint = false;
        protected override bool ProcessCmdKey(ref Message message, Keys keys)
        {
            switch (keys)
            {
                case Keys.N | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...

                    News(); tabcheck = true; tabControl1.SelectTab(tabPagedel2);
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
                    tabControl1.SelectTab(tabPagedel2);
                    return true; // signal that we've processed this key
            }
            // run base implementation
            return base.ProcessCmdKey(ref message, keys);
        }

        private void DeliveryWeightCertificate_Load(object sender, EventArgs e)
        {
           
            toolTip1.SetToolTip(butrefres, "Refresh"); 



            tabControl1.SelectTab(tabPagedel2); dateTimePicker1.Value.AddDays(0);
            txtsearch.Select(); dateTimePicker2.Value.ToLongDateString();
            Models.Serial.PortIP = Class.Users.PortIP + ":" + Class.Users.PortNo;
            Models.Serial.PortNo = Class.Users.PortNo;
            server = new SimpleTcpServer(Models.Serial.PortIP);
            //server.Events.ClientConnected += Events_ClientConnected;
            //server.Events.ClientDisconnected += Events_ClientDisconnected;
            //server.Events.DataReceived += Events_DataReceived;
            client = new SimpleTcpClient(Models.Serial.PortIP);
            client.Events.Connected += Events_Connected;
            client.Events.Disconnected += Events_Disconnected;
            client.Events.DataReceived += Events_DataReceived;
            //txtsecondweight.Text = "0.00";
            //txtfirstweight.Text = "0.00";
            if (server.IsListening)
            {
                server.Start();
                server.Send(Models.Serial.PortIP, "0");
            }
            
        }
        string s = "";
        private void Events_DataReceived(object sender, DataReceivedEventArgs e)
        {
            //  this.Invoke((MethodInvoker)delegate
            s = "";
            this.Invoke((MethodInvoker)delegate
            {
                s += Encoding.UTF8.GetString(e.Data.Array).Trim();
            });


        }

        private void Events_Disconnected(object sender, ConnectionEventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            {
               // lblcon.Text = "Sever DisConnected ";

            });
        }

        private void Events_Connected(object sender, ConnectionEventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            {
               // lblcon.Text = "Sever Connected ";
            });
        }


       
        public void News()
        {
          
            GridLoad();
            companyload(); comboproductload(); serialPort2.Close();
            toolTip1.SetToolTip(butrefres, "Refresh"); empty();
            dateTimePicker1.Value.AddDays(0);
            txtsearch.Select(); dateTimePicker2.Value.ToLongDateString();
            this.BackColor = Class.Users.BackColors;
            buttdelete.BackColor = Class.Users.BackColors;
            buttsave.BackColor = Class.Users.BackColors;
            buttprint.BackColor = Class.Users.BackColors;
            buttnew.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            panel4.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            panel5.BackColor = Class.Users.BackColors;
            panel6.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            Models.Serial.PortIP = Class.Users.PortIP + ":" + Class.Users.PortNo;
            Models.Serial.PortNo = Class.Users.PortNo;
            server = new SimpleTcpServer(Models.Serial.PortIP);           
            client = new SimpleTcpClient(Models.Serial.PortIP);
            client.Events.Connected += Events_Connected;
            client.Events.Disconnected += Events_Disconnected;
            client.Events.DataReceived += Events_DataReceived;          
            if (server.IsListening)
            {
                server.Start();
                server.Send(Models.Serial.PortIP, "0");
            }            
            txtdeltid.Text = "";
            if (tabcheck==false)
            {
                 tabControl1.SelectTab(tabPagedel2);
                tabcheck = true;
            }
            else
            {
                tabControl1.SelectTab(tabPagedel1);
                tabcheck = true;
            }
            empty(); 
            txtvechileno.Focus();
        }

       
        public void companyload()
        {
            try
            {
                string sel = "select a.gtcompmastid,a.compcode, a.compname from  gtcompmast a  where a.ptransaction ='COMPANY' order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;

                combocompname.DisplayMember = "compname";
                combocompname.ValueMember = "gtcompmastid";
                combocompname.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("companyload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }



        public void autonumberload()
        {
            try
            {
                //txtdeltid1.Text = ""; txtcertificateno.Text = "";
                //if(txtdeltid.Text != "") { }

                string sel = "select max(a.asptbldeliveryid1)+1 as id,b.compname from asptbldelivery a join gtcompmast b on a.compname = b.gtcompmastid  where b.ptransaction = 'COMPANY'  and a.finyear='" + Class.Users.Finyear + "' and b.compcode ='" + combocompcode.Text + "' group by b.compname; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbldelivery");
                DataTable dt = ds.Tables["asptbldelivery"];
                int cnt = dt.Rows.Count;
                if (cnt==0)
                {
                    string sel1 = "select b.gtcompmastid,b.compname from gtcompmast b  where b.ptransaction = 'COMPANY' and b.compcode ='" + combocompcode.Text + "'; ";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                    DataTable dt1 = ds1.Tables["asptbldelivery"];
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt1;

                    txtcertificateno.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + 1;
                    txtdeltid1.Text = "1";
                }
                else
                {
                    combocompname.Text = dt.Rows[0]["compname"].ToString();
                    txtcertificateno.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                    txtdeltid1.Text = dt.Rows[0]["id"].ToString();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void comboproductload()
        {
            try
            {
                string sel = "select a.asptblproductweightmasid,a.productname1 from asptblproductweightmas a where a.active='T';";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproductweightmas");
                DataTable dt = ds.Tables["asptblproductweightmas"];

                comboproduct.DisplayMember = "productname1";
                comboproduct.ValueMember = "asptblproductweightmasid";
                comboproduct.DataSource = dt;
                txtgrossweight1.Text = "";comboproduct.SelectedIndex= - 1;comboproduct.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("comboproductload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void comboproduct_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (invalid >= 1) { }
                else
                {
                    txtbags.Text = "";
                }
                txtkgs.Text = ""; txtnetweight1.Text = ""; txtgrossweight1.Text = "";
                string sel = "select a.asptblproductweightmasid,a.productname1,a.netweight,a.grossweight from asptblproductweightmas a where a.asptblproductweightmasid='" + comboproduct.SelectedValue + "'; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproductweightmas");
                DataTable dt = ds.Tables["asptblproductweightmas"];
                if (dt.Rows.Count >= 1)
                {
                    txtnetweight1.Text = Convert.ToString(dt.Rows[0]["netweight"].ToString());
                    txtgrossweight1.Text = dt.Rows[0]["grossweight"].ToString();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("comboproduct_SelectedIndexChanged: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void txtbags_TextChanged(object sender, EventArgs e)
        {

            try
            {



                if (txtbags.Text != "" && comboproduct.Text != "")
                {
                    txtkgs.Text = "";
                    decimal tot1 = Convert.ToInt64(txtbags.Text) * Convert.ToDecimal(txtgrossweight1.Text);
                    txtkgs.Text = tot1.ToString();


                }

            }
            catch (Exception ex)
            {
                // MessageBox.Show("txtbags_TextChanged: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void Saves()
        {
            Int64 maxid = 0; Int64 j = 0; validprint = false;
            Class.Users.Paramid = 0; Class.Users.PayPeriod = "";
            try
            {

                Models.Validate va = new Models.Validate();
                if (va.IsStringNumbericspace(txtvechileno.Text) == false)
                {
                    MessageBox.Show("'VechileNo Field is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.txtvechileno.Select();
                    return;
                }
                if (va.IsDecimal(txtfirstweight.Text) == false)
                {
                    MessageBox.Show("'First Weight Field  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.txtfirstweight.Select();

                    return;

                }
                else
                {

                 
                    string dat = "";
                    dat = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string dat1 = "";
                    dat1 = System.DateTime.Now.ToString("hh:mm:ss");
                    string dat2 = "";
                    dat2 = Convert.ToDateTime(System.DateTime.Now.ToString()).ToString("yyyy-MM-dd");
                    txtproductweight.Text = listview2totalweight.ToString();
                    string sel = "select asptbldeliveryid,certifiedby    from  asptbldelivery   WHERE   finyear='" + combofinyear.Text + "' and compname='" + combocompname.SelectedValue.ToString() + "' and certificateno='" + txtcertificateno.Text + "' and sendto='" + txtsendto.Text.ToUpper() + "'  and vechileNo='" + txtvechileno.Text + "' and deliverydate='" + dat + "' and deliverytime='" + dat1 + "' and firstweight='" + txtfirstweight.Text + "' and secondweight='" + txtsecondweight.Text + "' and netweight='" + txtnetweight.Text + "'  and productweight='" + txtproductweight.Text + "' ;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbldelivery");
                    DataTable dt = ds.Tables["asptbldelivery"];                   
                    string[] cer = txtcertifiedby.Text.Split(new[] {' '}, StringSplitOptions.RemoveEmptyEntries);
                    txtcertifiedby.Text = "";                    
                    txtcertifiedby.Text += string.Join(" ",cer);
                    Class.Users.PayPeriod = "";
                    Class.Users.PayPeriod = txtvechileno.Text;

                    string sw = Convert.ToString(txtsecondweight.Text);                    
                    char status1 = sw=="0.00" ? 'P' : 'C';


                    if (dt.Rows.Count != 0)
                    {
                        validprint = true;

                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt64("0" + txtdeltid.Text) == 0 || Convert.ToInt64("0" + txtdeltid.Text) == 0)
                    {

                        string ins = "insert into asptbldelivery(asptbldeliveryid1,finyear,compcode,compname,certificateno ,sendto, vechileNo, deliverydate,deliverytime , firstweight, secondweight, netweight , productweight , certifiedby ,compcode1,  username,createdby, modifiedby,ipaddress,weightdeffer,createdon) values('" + txtdeltid1.Text + "','" + combofinyear.Text + "','" + combocompname.SelectedValue.ToString() + "','" + combocompname.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "','" + txtsendto.Text.ToUpper().Trim() + "','" + txtvechileno.Text.ToUpper() + "','" + dat + "','" + dat1 + "','" + txtfirstweight.Text + "','" + Convert.ToDecimal("0" + txtsecondweight.Text) + "','" + txtnetweight.Text + "','" + Convert.ToDecimal("0" + txtproductweight.Text) + "','" + txtcertifiedby.Text.ToUpper() + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','','" + Class.Users.IPADDRESS + "','" + txtdifference.Text + "','" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "');";
                        Utility.ExecuteNonQuery(ins);

                        string inssap = "insert into saptable(salesid,TransType,Compcode,FinYear,LoginUser,Date,Time,WeighBridgeSlipNo,VehicleNo,FirstWeight,FirstWeightDate,FirstWeightTime,NetWeight,WeightDifference,Status)" +
                            "values('" + txtdeltid1.Text + "','Sales','" + combocompcode.SelectedValue.ToString() + "','" + combofinyear.Text + "','" + Class.Users.HUserName + "','" + dat2 + "','" + dat1 + "','" + txtcertificateno.Text.ToUpper() + "','" + txtvechileno.Text.ToUpper() + "','" + Convert.ToInt64("0" + txtfirstweight.Text) + "','" + dat + "','" + dat1 + "','" +  txtnetweight.Text + "','" + txtdifference.Text + "','P' );";
                        Utility.ExecuteNonQuery(inssap);
                      
                        Class.Users.Paramid = 0;
                        string sel2 = "select max(asptbldeliveryid) as asptbldeliveryid   from  asptbldelivery   WHERE  finyear='" + combofinyear.Text + "' and compname='" + combocompcode.SelectedValue.ToString() + "' ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldelivery");
                        DataTable dt2 = ds2.Tables["asptbldelivery"]; maxid = 0;
                        maxid = Convert.ToInt64(dt2.Rows[0]["asptbldeliveryid"].ToString());
                        string ins1 = "insert into asptbldelivery2(asptbldeliveryid,asptbldeliveryid1,finyear,compcode,compname,certificateno ,sendto, vechileNo, deliverydate,deliverydate1, deliverytime , firstweight, secondweight, netweight , productweight , certifiedby ,compcode1,  username,createdby, modifiedby,ipaddress,weightdeffer,createdon,operation) values('" + maxid + "','" + txtdeltid1.Text + "','" + combofinyear.Text + "','" + combocompname.SelectedValue.ToString() + "','" + combocompname.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "','" + txtsendto.Text.ToUpper() + "','" + txtvechileno.Text.ToUpper() + "','" + Convert.ToDateTime(dat).ToString("dd-MM-yyyy")+ " "+ dat1 + "','" + dat + "','" + dat1 + "','" + txtfirstweight.Text + "','" + Convert.ToDecimal("0" + txtsecondweight.Text) + "','" + txtnetweight.Text + "','" + Convert.ToDecimal("0" + txtproductweight.Text) + "','" + txtcertifiedby.Text.ToUpper().Trim() + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + " : " + Class.Users.IPADDRESS + "','','" + Class.Users.IPADDRESS + "','" + txtdifference.Text + "','" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','INSERTED');";
                        Utility.ExecuteNonQuery(ins1);
                        Class.Users.Paramid = Convert.ToInt64("0" + maxid);
                        validprint = true;
                    }
                    else
                    {
                       
                        string up = "update  asptbldelivery  set asptbldeliveryid1='" + txtdeltid1.Text + "',finyear='" + combofinyear.Text + "',compcode='" + combocompcode.SelectedValue.ToString() + "',certificateno='" + txtcertificateno.Text.ToUpper() + "' ,sendto='" + txtsendto.Text.ToUpper().Trim() + "', vechileNo='" + txtvechileno.Text.ToUpper() + "', deliverydate='" + dat + "', deliverytime='" + dat1 + "' , firstweight='" + txtfirstweight.Text + "', secondweight='" +  txtsecondweight.Text + "', netweight='" + txtnetweight.Text + "' , productweight='" + Convert.ToDecimal("0" + txtproductweight.Text) + "' , certifiedby='" + txtcertifiedby.Text.ToUpper().Trim() + "' , compcode1='" + Class.Users.COMPCODE + "',  username='" + Class.Users.USERID + "',modifiedby='" + Class.Users.HUserName + "', modifiedon='" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "',ipaddress='" + Class.Users.IPADDRESS + "',weightdeffer='" + txtdifference.Text + "' where asptbldeliveryid='" + txtdeltid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                       
                        
                        string upsap = "update  saptable set TransType='Sales',Compcode='" + combocompcode.SelectedValue.ToString() + "',FinYear='" + combofinyear.Text + "',LoginUser='" + Class.Users.HUserName + "',WeighBridgeSlipNo='" + txtcertificateno.Text.ToUpper() + "',VehicleNo='" + txtvechileno.Text.ToUpper() + "',FirstWeight='"+txtfirstweight.Text+"',SecondWeight='" + txtsecondweight.Text + "',SecondWeightDate='" + dat + "',SecondWeightTime='" + dat1 + "',NetWeight='" + txtnetweight.Text + "',WeightDifference='" + txtdifference.Text + "',TheoreticalWeight='" +txtproductweight.Text + "' ,Status='" + status1 + "' WHERE salesid='" + txtdeltid1.Text + "'";
                        Utility.ExecuteNonQuery(upsap);

                 
                        maxid = 0; validprint = true;
                        maxid = Convert.ToInt64(txtdeltid.Text);
                        if (listView2.Items.Count == 0)
                        {
                            string ins1 = "insert into asptbldelivery2(asptbldeliveryid,asptbldeliveryid1,finyear,compcode,compname,certificateno ,sendto, vechileNo, deliverydate,deliverydate1, deliverytime , firstweight, secondweight, netweight , productweight , certifiedby ,compcode1,  username,createdby, modifiedby,ipaddress,weightdeffer,modifiedon,operation) values('" + maxid + "','" + txtdeltid1.Text + "','" + combofinyear.Text + "','" + combocompname.SelectedValue.ToString() + "','" + combocompname.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "','" + txtsendto.Text.ToUpper() + "','" + txtvechileno.Text.ToUpper() + "','" + Convert.ToDateTime(dat).ToString("dd-MM-yyyy") + " " + dat1 + "','" + dat + "','" + dat1 + "','" + txtfirstweight.Text + "','" + Convert.ToDecimal("0" + txtsecondweight.Text) + "','" + txtnetweight.Text + "','" + "0.00" + "','" + txtcertifiedby.Text.ToUpper() + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + " : " + Class.Users.IPADDRESS + "','','" + Class.Users.IPADDRESS + "','" + txtdifference.Text + "','" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','No Row in List');";
                            Utility.ExecuteNonQuery(ins1);
                        }
                        else
                        {
                            string ins1 = "insert into asptbldelivery2(asptbldeliveryid,asptbldeliveryid1,finyear,compcode,compname,certificateno ,sendto, vechileNo, deliverydate,deliverydate1, deliverytime , firstweight, secondweight, netweight , productweight , certifiedby ,compcode1,  username,createdby, modifiedby,ipaddress,weightdeffer,modifiedon,operation) values('" + maxid + "','" + txtdeltid1.Text + "','" + combofinyear.Text + "','" + combocompname.SelectedValue.ToString() + "','" + combocompname.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "','" + txtsendto.Text.ToUpper() + "','" + txtvechileno.Text.ToUpper() + "','" + Convert.ToDateTime(dat).ToString("dd-MM-yyyy") + " " + dat1 + "','" + dat + "','" + dat1 + "','" + txtfirstweight.Text + "','" + Convert.ToDecimal("0" + txtsecondweight.Text) + "','" + txtnetweight.Text + "','" + Convert.ToDecimal("0" + txtproductweight.Text) + "','" + txtcertifiedby.Text.ToUpper() + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + " : " + Class.Users.IPADDRESS + "','','" + Class.Users.IPADDRESS + "','" + txtdifference.Text + "','" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','UPDATED');";
                            Utility.ExecuteNonQuery(ins1);
                        }
                        Class.Users.Paramid = Convert.ToInt64("0" + txtdeltid.Text);

                    }
                    int i = 0;
                    if (listView2.Items.Count >= 0)
                    {

                        for (i = 0; i < listView2.Items.Count; i++)
                        {
                            Int64 maxid1 = 0;
                            string sel1 = "select asptbldeliverydetid    from  asptbldeliverydet   where  asptbldeliveryid='" + listView2.Items[i].SubItems[3].Text + "'  and productname='" + listView2.Items[i].SubItems[4].Text + "' and productweight='" + listView2.Items[i].SubItems[6].Text + "';";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbldeliverydet");
                            DataTable dt1 = ds1.Tables["asptbldeliverydet"];
                            if (dt1.Rows.Count != 0)
                            {
                                validprint = true;
                                mas.pop(combocompcode.Text, txtvechileno.Text, txtnetweight.Text);
                                tabControl1.SelectTab(tabPagedel2);
                            }
                            else if (dt1.Rows.Count != 0 && Convert.ToInt64("0" + listView2.Items[i].SubItems[2].Text) == 0 || Convert.ToInt64("0" + listView2.Items[i].SubItems[2].Text) == 0)
                            {

                                string ins1 = "insert into asptbldeliverydet(asptbldeliveryid ,asptbldeliveryid1,productname , productweight , productkgs,compcode,finyear) values('" + maxid.ToString() + "' ,'" + txtdeltid1.Text + "' , '" + listView2.Items[i].SubItems[4].Text + "' ,'" + Convert.ToString(listView2.Items[i].SubItems[6].Text) + "' , '" + Convert.ToString(listView2.Items[i].SubItems[7].Text) + "','"+combocompcode.SelectedValue+ "' ,'" + combofinyear.Text + "');";
                                Utility.ExecuteNonQuery(ins1);

                                string sel2 = "select max(asptbldeliverydetid) as asptbldeliverydetid   from  asptbldeliverydet   WHERE  asptbldeliveryid='" + maxid.ToString() + "' and finyear='" + combofinyear.Text + "' and compcode='" + combocompcode.SelectedValue.ToString() + "' ;";
                                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldeliverydet");
                                DataTable dt2 = ds2.Tables["asptbldeliverydet"]; 

                                maxid1 = Convert.ToInt64(dt2.Rows[0]["asptbldeliverydetid"].ToString());
                                string ins2 = "insert into asptbldeliverydet2(asptbldeliverydetid,asptbldeliveryid ,asptbldeliveryid1,productname , productweight , productkgs,operation,compcode,finyear) values('" + maxid1.ToString() + "' ,'" + maxid.ToString() + "' ,'" + txtdeltid1.Text + "' , '" + listView2.Items[i].SubItems[4].Text + "' ,'" + Convert.ToString(listView2.Items[i].SubItems[6].Text) + "' , '" + Convert.ToString(listView2.Items[i].SubItems[7].Text) + "','INSERTED','" + combocompcode.SelectedValue + "'  ,'" + combofinyear.Text + "' );";
                                Utility.ExecuteNonQuery(ins2);
                                validprint = true;
                            }
                            else
                            {
                                if (listView2.Items[i].SubItems[2].Text != "")
                                {
                                    string up1 = "update  asptbldeliverydet  set asptbldeliveryid='" + txtdeltid.Text + "' ,asptbldeliveryid1='" + txtdeltid1.Text + "', productname='" + Convert.ToString(listView2.Items[i].SubItems[4].Text) + "' , productweight='" + listView2.Items[i].SubItems[6].Text + "' , productkgs='" + listView2.Items[i].SubItems[7].Text + "'  where asptbldeliverydetid='" + listView2.Items[i].SubItems[2].Text + "';";
                                    Utility.ExecuteNonQuery(up1);
                                }
                                string ins2 = "insert into asptbldeliverydet2(asptbldeliverydetid,asptbldeliveryid ,asptbldeliveryid1,productname , productweight , productkgs,operation,compcode,finyear) values('"+ listView2.Items[i].SubItems[2].Text+"','" + txtdeltid.Text.ToString() + "' ,'" + txtdeltid1.Text + "' , '" + listView2.Items[i].SubItems[4].Text + "' ,'" + Convert.ToString(listView2.Items[i].SubItems[6].Text) + "' , '" + Convert.ToString(listView2.Items[i].SubItems[7].Text) + "','UPDATED','"+combocompcode.SelectedValue+ "'  ,'" + combofinyear.Text + "');";
                                Utility.ExecuteNonQuery(ins2);
                                validprint = true;
                            }

                        }
                    }
                    if (txtdeltid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtcertificateno.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); 
                        mas.pop("Record Saved Successfully ","Vehicle No: " +txtvechileno.Text, "Net Wight :"+txtnetweight.Text);
                        autonumberload(); txtdeltid1.Text = "";
                       
                        empty(); txtvechileno.Text = "";
                        tabControl1.SelectTab(tabPagedel2);
                    }
                    else
                    {
                         MessageBox.Show("Record Updated Successfully " + txtcertificateno.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        mas.pop("Record Updated Successfully ", "Vehicle No: " + txtvechileno.Text, "Net Wight :" + txtnetweight.Text);
                         autonumberload(); txtdeltid1.Text = "";
                        empty(); txtvechileno.Text = "";
                        tabControl1.SelectTab(tabPagedel2);
                    }
                }
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private void DeliveryWeightCertificate_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            empty();
            //if (serialPort2.IsOpen) serialPort2.Close();
            tabcheck = false;
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }


        private void empty()
        {
            
            txtdeltid.Text = ""; invalid = 0;
            combocompcode.Enabled = true; ; combocompcode.Enabled = true;
            listfilter.Items.Clear();
            if (serialPort2.IsOpen) serialPort2.Close();
            lbltotalfooterkgs.Text = ""; listview2totalweight = 0; txtproductweight.Text = "";
            combofinyear.Text = Class.Users.Finyear;
            dateTimePicker1.Value = System.DateTime.Now;
            txtcertificateno.Text = "";
         
            txtfirstweight.Text = "";
            comboproduct.SelectedIndex = -1; comboproduct.Text = "";
            txtsecondweight.Text = ""; txtdifference.Text = "";
            txtnetweight.Text = ""; lbltotal1.Text = "Total Count: ";
            comboproduct.SelectedIndex = -1; comboproduct.Text = "";
            txtproductweight.Text = "";
            txtsendto.Text = "";
            //txtcertifiedby.Text = "";
            txtbags.Text = "";
            txtkgs.Text = ""; txtgrossweight1.Text = ""; txtdifference.Text = "";
            lbltotalfooterkgs.Text = ""; listView2.Items.Clear();autonumberload();

        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptbldeliveryid,a.vechileNo,a.certificateno,a.sendto,date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,a.firstweight,a.secondweight,a.netweight,a.productweight from asptbldelivery a  where a.deliverydate='" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "' order by a.asptbldeliveryid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                DataTable dt = ds.Tables["asptbldelivery"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldeliveryid"].ToString());
                        list.SubItems.Add(myRow["vechileNo"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["sendto"].ToString());
                        list.SubItems.Add(myRow["deliverydate"].ToString());
                        list.SubItems.Add(myRow["firstweight"].ToString());
                        list.SubItems.Add(myRow["secondweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());
                        list.SubItems.Add(myRow["productweight"].ToString());
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
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                listview2totalweight = 0; Class.Users.PayPeriod = ""; Class.Users.Paramid = 0;
                // empty();
                if (listView1.Items.Count > 0)
                {
                    listView2.Items.Clear(); Class.Users.Paramid = 0;
                    txtdeltid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    txtvechileno.Text = listView1.SelectedItems[0].SubItems[3].Text;
                    Class.Users.Paramid = Convert.ToInt64(txtdeltid.Text);
                    string sel1 = "select a.asptbldeliveryid,a.asptbldeliveryid1 ,a.finyear,d.compcode, d.compname,a.certificateno , a.sendto,a.vechileNo, date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,a.deliverytime,a.firstweight,a.secondweight, a.netweight,a.productweight,a.certifiedby  from asptbldelivery a  join gtcompmast d on a.compname=d.gtcompmastid  where a.asptbldeliveryid='" + txtdeltid.Text + "' and a.vechileNo='"+ txtvechileno.Text +"'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                    DataTable dt = ds.Tables["asptbldelivery"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtdeltid.Text = Convert.ToString(dt.Rows[0]["asptbldeliveryid"].ToString());
                        txtdeltid1.Text = Convert.ToString(dt.Rows[0]["asptbldeliveryid1"].ToString());
                        combofinyear.Text = Convert.ToString(dt.Rows[0]["finyear"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        txtcertificateno.Text = Convert.ToString(dt.Rows[0]["certificateno"].ToString());
                        txtsendto.Text = Convert.ToString(dt.Rows[0]["sendto"].ToString());
                        txtvechileno.Text = Convert.ToString(dt.Rows[0]["vechileno"].ToString());
                        dateTimePicker1.Text = Convert.ToString(dt.Rows[0]["deliverydate"].ToString());
                        dateTimePicker2.Text = Convert.ToString(dt.Rows[0]["deliverytime"].ToString());
                        txtfirstweight.Text = Convert.ToDecimal("0" + dt.Rows[0]["firstweight"].ToString()).ToString();

                        txtsecondweight.Text = Convert.ToString(dt.Rows[0]["secondweight"].ToString());
                        txtnetweight.Text = Convert.ToString(dt.Rows[0]["netweight"].ToString());
                        txtproductweight.Text = Convert.ToString(dt.Rows[0]["productweight"].ToString());
             



                       

                    }
                    // combosendto.Text = dt.Rows[0]["sendid"].ToString();
                    //combosendto.DataSource = dt;
                    txtcertifiedby.Text = Convert.ToString(dt.Rows[0]["certifiedby"].ToString());
                    string sel2 = "select b.asptbldeliverydetid, b.asptbldeliveryid,e.asptblproductweightmasid, e.productname1, b.productweight, b.productkgs  from asptbldelivery a join asptbldeliverydet  b on a.asptbldeliveryid=b.asptbldeliveryid  join gtcompmast d on a.compname=d.gtcompmastid join asptblproductweightmas e on e.asptblproductweightmasid=b.productname where a.asptbldeliveryid='" + txtdeltid.Text + "' and a.vechileNo='" + txtvechileno.Text + "';";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldeliverydet");
                    DataTable dt2 = ds2.Tables["asptbldeliverydet"];
                    if (dt2.Rows.Count >= 1)
                    {
                        foreach (DataRow myRow in dt2.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());

                            list.SubItems.Add(myRow["asptbldeliverydetid"].ToString());
                            list.SubItems.Add(myRow["asptbldeliveryid"].ToString());
                            list.SubItems.Add(myRow["asptblproductweightmasid"].ToString());
                            list.SubItems.Add(myRow["productname1"].ToString());
                            list.SubItems.Add(myRow["productweight"].ToString());
                            list.SubItems.Add(myRow["productkgs"].ToString());
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView2.Items.Add(list);
                            i++;
                            listview2totalweight += Convert.ToDecimal("0" + myRow["productkgs"].ToString());

                        }
                        txtnetweight_TextChanged(sender, e);
                    }
                    decimal nw = 0; txtdifference.Text = "";
                    decimal net = Convert.ToDecimal("0" + txtnetweight.Text);
                    decimal prdqt = Convert.ToDecimal("0" + txtproductweight.Text);
                    nw = net- prdqt;
                    txtdifference.Text = nw.ToString();
                    lbltotal1.Text = "Total Count: " + listView2.Items.Count;
                    lbltotalfooterkgs.Text = listview2totalweight.ToString();
                    txtproductweight.Text = listview2totalweight.ToString(); combocompcode.Enabled = false;
                   
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("ListView1_ItemActivate: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            tabControl1.SelectTab(tabPagedel1); txtvechileno.Focus();
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

                        if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[5].ToString().Contains(txtsearch.Text))
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
                            
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);

                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    listView1.Items.Clear();

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
           
        }

        private void ContextMenuStrip2_Click(object sender, EventArgs e)
        {
            
        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad(); comboproductload(); usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            companyload(); combocompcode.Select();

        }

        private void butAddListView_Click(object sender, EventArgs e)
        {
            try
            {

                

                bool addbool = false;
                Models.Validate va = new Models.Validate();

                if (comboproduct.Text == "")
                {
                    MessageBox.Show("product name is Empty ", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    comboproduct.Select();
                    return;
                }
                if (va.IsInteger(txtbags.Text) == false)
                {
                    MessageBox.Show("No of Bags is empty  ", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtbags.Select();
                    return;
                }
                if (va.IsInteger(txtbags.Text) == true)
                {

                    int item0 = 0; int j = listView2.Items.Count;
                    if (comboproduct.SelectedValue != null && Convert.ToInt64(txtbags.Text) >= 1 && Convert.ToDecimal(txtkgs.Text) >= 1)
                    {
                        if (invalid >= 1)
                        {
                           
                            for (i = 0; i < listView2.Items.Count; i++)
                            {
                                if (listView2.Items[i].Selected)
                                {
                                    decimal invalidtoal1 = Convert.ToDecimal("0" + listView2.Items[i].SubItems[7].Text);
                                    listview2totalweight -= Convert.ToDecimal("0" + invalidtoal1);
                                    listView2.Items[i].Remove();
                                    invalid = 0; invalidtoal1 = 0;
                                }
                            }

                        }
                        if (invalid == 0)
                        {
                            ListViewItem list1 = new ListViewItem();
                            foreach (ListViewItem item in listView2.Items)
                            {
                                if (item.SubItems[5].Text.Contains(comboproduct.Text.ToString()) && invalid == 0)
                                {
                                    addbool = false;
                                    MessageBox.Show("This ItemName Already Added:-   " + comboproduct.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Question);
                                    comboproduct.SelectedIndex = -1; comboproduct.Select();
                                    return;
                                }
                                else
                                {
                                    addbool = true;

                                }
                                item0++;
                            }

                            if (addbool == true)
                            {


                                txtproductweight.Text = ""; lbltotalfooterkgs.Text = "";
                                list1.SubItems.Add(i.ToString());
                                list1.SubItems.Add(gridid.ToString());
                                list1.SubItems.Add(txtdeltid.Text);
                                list1.SubItems.Add(comboproduct.SelectedValue.ToString());
                                list1.SubItems.Add(comboproduct.Text);
                                list1.SubItems.Add(txtbags.Text);
                                list1.SubItems.Add(Convert.ToDecimal("0" + txtkgs.Text).ToString());
                                list1.BackColor = j % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                                listView2.Items.Add(list1);
                                listview2totalweight += Convert.ToDecimal("0" + list1.SubItems[7].Text);
                                lbltotalfooterkgs.Text = listview2totalweight.ToString();
                                txtproductweight.Text = listview2totalweight.ToString();
                                lbltotal1.Text = "Total Count: " + listView2.Items.Count.ToString();
                                comboproduct.Text = "";
                                txtnetweight1.Text = "";
                                txtgrossweight1.Text = "";
                                comboproduct.Select();
                                txtbags.Text = ""; txtkgs.Text = "";
                                txtnetweight_TextChanged(sender, e);
                                j++;
                                gridid = 0;
                            }

                            if (listView2.Items.Count == 0)
                            {
                                txtproductweight.Text = ""; txtdifference.Text = ""; listview2totalweight = 0; lbltotalfooterkgs.Text = "";
                                lbltotalfooterkgs.Text = "";
                                // i += listView2.Items.Count;
                                // list1.Text = i.ToString();
                                list1.SubItems.Add(i.ToString());
                                list1.SubItems.Add(gridid.ToString());
                                list1.SubItems.Add(txtdeltid1.Text);
                                list1.SubItems.Add(comboproduct.SelectedValue.ToString());
                                list1.SubItems.Add(comboproduct.Text);
                                list1.SubItems.Add(txtbags.Text);
                                list1.SubItems.Add(Convert.ToDecimal("0" + txtkgs.Text).ToString());
                                list1.BackColor = j % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                                listView2.Items.Add(list1);
                                listview2totalweight += Convert.ToDecimal("0" + list1.SubItems[7].Text);
                                lbltotalfooterkgs.Text = listview2totalweight.ToString();
                                txtproductweight.Text = listview2totalweight.ToString();
                                lbltotal1.Text = "Total Count: " + listView2.Items.Count.ToString();
                                comboproduct.Text = "";
                                txtnetweight1.Text = "";
                                txtgrossweight1.Text = "";
                                comboproduct.Select();
                                txtbags.Text = ""; txtkgs.Text = "";
                                //  txtnetweight_TextChanged(sender, e);
                                j++;

                                i++;
                                gridid = 0;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Pls Enter No of Bag Field ", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        comboproduct.Select(); return;
                    }
                }

                //   txtbags.Text = ""; butAddListView.BackColor = Color.White; txtbags.BackColor = Color.White; comboproduct.BackColor = Color.Yellow;

            }
            catch (Exception ex)
            {
                // MessageBox.Show("butAddListView_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                comboproduct.Select();
            }
            txtnetweight_TextChanged(sender, e);
            comboproductload(); comboproduct.Text = "";
        }



        private void Searchs_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPagedel2);

        }

        private void listView2_ItemActivate(object sender, EventArgs e)
        {


            invalid = 0;

            for (i = 0; i < listView2.Items.Count; i++)
            {

                if (listView2.Items[i].Selected)
                {
                    comboproduct.Text = listView2.Items[i].SubItems[5].Text;
                    txtbags.Text = listView2.Items[i].SubItems[6].Text;
                    gridid = Convert.ToInt32("0" + listView2.Items[i].SubItems[2].Text);
                    invalid = Convert.ToInt32(listView2.Items[i].SubItems[4].Text);

                    comboproduct_SelectedIndexChanged(sender, e);
                    txtbags_TextChanged(sender, e);
                }


            }
        }
        private void butsecondweight_Click(object sender, EventArgs e)
        {

            //try
            //{

            //    string readserialvalue1 = ""; 
            //    if (Models.Serial.PortType == "TCP/IP")
            //    {
            //        if (client.IsConnected == false)
            //        {
            //            client.Connect();
            //        }
            //        client.Send("0");
            //    }
            //    if (Models.Serial.PortType != "TCP/IP")
            //    {
            //        butstart_Click(sender, e);
            //        words = null;
            //        readserialvalue1 = serialPort2.ReadLine().Trim();
            //        btngetweight.ForeColor = Color.Green;
            //        words = readserialvalue1.Split(' ');
            //    }               
            //    if (textBox1.Text != "")
            //    {
            //        Int64 ss = 0;

            //        if (Models.Serial.PortType == "TCP/IP")
            //        {

            //            string words2 = Regex.Replace(s, @"Gross", "").Trim();
            //            words = words2.Split();
            //            if (words.Length >= 1)
            //            {

            //                if (txtdeltid.Text != "")
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    if (Convert.ToInt64("0" + words[0].ToString()) >= 1)
            //                    {
            //                        ss = 0; txtsecondweight.Text = "";
            //                        ss = Convert.ToInt32("0" + words[0].ToString());
            //                        txtsecondweight.Text = ss.ToString();
            //                        s = "";words2 = "";words = null;words = null;
            //                    }
            //                }
            //                words2 = null;
            //            }
            //            else
            //            {
            //                txtsecondweight.Text = "0"; s = ""; words2 = ""; words = null; words = null;
            //            }


            //        }
            //        if (Models.Serial.PortType != "TCP/IP")
            //        {
            //            if (txtdeltid.Text != "" && txtsecondweight.Text == "0.00")
            //            {
            //                if (Convert.ToInt64("0" + words[2].ToString()) >= 1)
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[2].ToString().Trim());
            //                    txtsecondweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[3].ToString()) >= 1)
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[3].ToString().Trim());
            //                    txtsecondweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[4].ToString()) >= 1)
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[4].ToString().Trim());
            //                    txtsecondweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[5].ToString()) >= 1)
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[5].ToString().Trim());
            //                    txtsecondweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[6].ToString()) >= 1)
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[6].ToString().Trim());
            //                    txtsecondweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[7].ToString()) >= 1)
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[7].ToString().Trim());
            //                    txtsecondweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[8].ToString()) >= 1)
            //                {
            //                    ss = 0; txtsecondweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[8].ToString().Trim());
            //                    txtsecondweight.Text = ss.ToString();
            //                    //return;
            //                }

            //            }
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("Vehicle Number is Empty");
            //    }
            //}
            //catch (Exception ex)
            //{
            //}

            try
            {
                if (string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Vehicle Number is Empty");
                    return;
                }

                long weight = 0;
                string[] words = null;

                /* ---------------- TCP/IP ---------------- */
                if (Models.Serial.PortType == "TCP/IP")
                {
                    if (!client.IsConnected)
                        client.Connect();

                    client.Send("0");

                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        string clean = Regex.Replace(s, @"Gross", "", RegexOptions.IgnoreCase).Trim();
                        words = clean.Split(' ', (char)StringSplitOptions.RemoveEmptyEntries);

                        if (words.Length > 0 && txtdeltid.Text != "")
                        {
                            if (long.TryParse(words[0], out weight) && weight > 0)
                            {
                                txtsecondweight.Text = weight.ToString();
                            }
                        }
                    }

                    s = string.Empty;
                    return;
                }

                /* ---------------- SERIAL PORT ---------------- */
                butstart_Click(sender, e);

                string readValue = serialPort2.ReadLine()?.Trim();
                if (string.IsNullOrEmpty(readValue))
                    return;

                btngetweight.ForeColor = Color.Green;
                words = readValue.Split(' ',(char)StringSplitOptions.RemoveEmptyEntries);

                if (txtdeltid.Text != "" && txtsecondweight.Text == "0.00")
                {
                    foreach (string w in words)
                    {
                        if (long.TryParse(w.Trim(), out weight) && weight > 0)
                        {
                            txtsecondweight.Text = weight.ToString();
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Weight Read Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            serialPort2.Close(); s = "";   words = null; words = null;
        }
       
        private void btngetweight_Click(object sender, EventArgs e)
        {

            //try
            //{


            //    string readserialvalue1 = ""; 

            //    if (Models.Serial.PortType == "TCP/IP")
            //    {
            //        if (client.IsConnected == false)
            //        {
            //            client.Connect();
            //        }
            //        client.Send("0");
            //    }
            //    if (Models.Serial.PortType != "TCP/IP")
            //    {
            //        butstart_Click(sender, e);
            //        words = null;
            //        readserialvalue1 = serialPort2.ReadLine().Trim();
            //        words = readserialvalue1.Split(' ');


            //    }

            //    if (textBox1.Text != null)
            //    {
            //        Int64 ss = 0;
            //        if (Models.Serial.PortType == "TCP/IP")
            //        {

            //            string words2 = Regex.Replace(s, @"Gross", "").Trim();
            //            words = words2.Split();

            //            if (words.Length >= 1)
            //            {
            //                if (txtdeltid.Text == "")
            //                {
            //                    ss = 0; txtfirstweight.Text = "0";
            //                    if (Convert.ToInt64("0" + words[0].ToString()) >= 1 || Convert.ToInt64("0" + words[0].ToString()) == 00)
            //                    {
            //                        ss = 0; txtfirstweight.Text = "";
            //                        ss = Convert.ToInt32("0" + words[0].ToString());
            //                        txtfirstweight.Text = ss.ToString();
            //                        s = ""; words2 = ""; words = null; words = null;
            //                    }
            //                }

            //            }
            //            else
            //            {
            //                txtfirstweight.Text = "0"; s = ""; words2 = ""; words = null; words = null;
            //            }
            //        }
            //        if (Models.Serial.PortType != "TCP/IP")
            //        {

            //            if (txtdeltid.Text == "")
            //            {
            //                if (Convert.ToInt64("0" + words[2].ToString()) >= 1)
            //                {
            //                    ss = 0; txtfirstweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[2].ToString().Trim());
            //                    txtfirstweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[3].ToString()) >= 1)
            //                {
            //                    ss = 0; txtfirstweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[3].ToString().Trim());
            //                    txtfirstweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[4].ToString()) >= 1)
            //                {
            //                    ss = 0; txtfirstweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[4].ToString().Trim());
            //                    txtfirstweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[5].ToString()) >= 1 || Convert.ToInt64("0" + words[5].ToString()) == 00)
            //                {
            //                    ss = 0; txtfirstweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[5].ToString().Trim());
            //                    txtfirstweight.Text = ss.ToString();
            //                    //return;
            //                }
            //                if (Convert.ToInt64("0" + words[6].ToString()) >= 1 || Convert.ToInt64("0" + words[6].ToString()) == 00)
            //                {
            //                    ss = 0; txtfirstweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[6].ToString().Trim());
            //                    txtfirstweight.Text = ss.ToString();
            //                    // return;
            //                }
            //                if (Convert.ToInt64("0" + words[7].ToString()) >= 1)
            //                {
            //                    ss = 0; txtfirstweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[7].ToString().Trim());
            //                    txtfirstweight.Text = ss.ToString();
            //                    // return;
            //                }
            //                if (Convert.ToInt64("0" + words[8].ToString()) >= 1)
            //                {
            //                    ss = 0; txtfirstweight.Text = "";
            //                    ss = Convert.ToInt64("0" + words[8].ToString().Trim());
            //                    txtfirstweight.Text = ss.ToString();
            //                    //return;
            //                }

            //            }
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("Vehicle Number is Empty");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    // MessageBox.Show("Error: " + ex.ToString(), " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}

            try
            {
                string readserialvalue1 = "";
                long weight = 0;

                // ================= TCP / IP =================
                if (Models.Serial.PortType == "TCP/IP")
                {
                    if (!client.IsConnected)
                        client.Connect();

                    client.Send("0");

                    string clean = Regex.Replace(s ?? "", @"Gross", "").Trim();
                    string[] words = clean.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    if (txtdeltid.Text == "" && words.Length > 0)
                    {
                        if (TryGetWeight(words[0], out weight))
                            txtfirstweight.Text = weight.ToString();
                    }
                }
                // ================= SERIAL PORT =================
                else
                {
                    butstart_Click(sender, e);

                    readserialvalue1 = serialPort2.ReadLine()?.Trim();
                    string[] words = readserialvalue1?.Split(' ');

                    if (txtdeltid.Text == "" && words != null)
                    {
                        // Check index 2 to 8 safely
                        for (int i = 2; i <= 8 && i < words.Length; i++)
                        {
                            if (TryGetWeight(words[i], out weight))
                            {
                                txtfirstweight.Text = weight.ToString();
                                break;
                            }
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("Vehicle Number is Empty");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            serialPort2.Close(); words = null; s = "";
        }


        private bool TryGetWeight(string value, out long weight)
        {
            return long.TryParse(value?.Trim(), out weight);
        }




        private void button2_Click(object sender, EventArgs e)
        {

        }

        public void Deletes()
        {
            try
            {
                if (txtdeltid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        string dat = "";
                        dat = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                        string dat1 = "";
                        dat1 = System.DateTime.Now.ToString("hh:mm:ss");
                        Int64 maxid = 0;
                       
                        string del = "delete from  asptbldelivery where compcode='" + combocompcode.SelectedValue + "' and finyear='" + combofinyear.Text + "'  and asptbldeliveryid='" + txtdeltid.Text + "';";
                        Utility.ExecuteNonQuery(del);
                        string ins1 = "insert into asptbldelivery2(asptbldeliveryid,asptbldeliveryid1,finyear,compcode,compname,certificateno ,sendto, vechileNo, deliverydate,deliverydate1, deliverytime , firstweight, secondweight, netweight , productweight , certifiedby ,compcode1,  username,createdby, modifiedby,ipaddress,weightdeffer,modifiedon,operation) values('" + txtdeltid.Text + "','" + txtdeltid1.Text + "','" + combofinyear.Text + "','" + combocompname.SelectedValue.ToString() + "','" + combocompname.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "','" + txtsendto.Text.ToUpper() + "','" + txtvechileno.Text.ToUpper() + "','" + Convert.ToDateTime(dat).ToString("dd-MM-yyyy") + " " + dat1 + "','" + dat + "','" + dat1 + "','" + txtfirstweight.Text + "','" + Convert.ToDecimal("0" + txtsecondweight.Text) + "','" + txtnetweight.Text + "','" + Convert.ToDecimal("0" + txtproductweight.Text) + "','" + txtcertifiedby.Text.ToUpper() + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + " : "+ Class.Users.IPADDRESS + "','','" + Class.Users.IPADDRESS + "','" + txtdifference.Text + "','" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','DELETED');";
                        Utility.ExecuteNonQuery(ins1);


                        if (listView2.Items.Count >= 0)
                        {
                            for (i = 0; i < listView2.Items.Count; i++)
                            {
                                string ins2 = "insert into asptbldeliverydet2(asptbldeliverydetid,asptbldeliveryid ,asptbldeliveryid1,productname , productweight , productkgs,operation,compcode,finyear) values('" + listView2.Items[i].SubItems[2].Text + "' ,'" + txtdeltid.Text.ToString() + "' ,'" + txtdeltid1.Text + "' , '" + listView2.Items[i].SubItems[4].Text + "' ,'" + Convert.ToString(listView2.Items[i].SubItems[6].Text) + "' , '" + Convert.ToString(listView2.Items[i].SubItems[7].Text) + "','DELETED','" + combocompcode.SelectedValue + "','" + combofinyear.Text + "');";
                                Utility.ExecuteNonQuery(ins2);
                                string del1 = "delete from asptbldeliverydet  where  asptbldeliveryid='" + txtdeltid.Text + "';";
                                Utility.ExecuteNonQuery(del1);
                                string del2 = "delete from saptable  where   WeighBridgeSlipNo='" + txtcertificateno.Text + "' and  salesid='" + txtdeltid1.Text + "';";
                                Utility.ExecuteNonQuery(del2);
                            }
                        }
                       
                        MessageBox.Show("Record Deleted Successfully " + txtdeltid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        mas.pop("Record Deleted Successfully ", "ID: " + txtdeltid.Text,"");
                        maxid = 0;
                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletes_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtsecondweight_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (txtsecondweight.Text != "")
                {

                    //if (Convert.ToDecimal("0" + txtfirstweight.Text) < Convert.ToDecimal("0" + txtsecondweight.Text))
                    //{
                    //    MessageBox.Show("Invalid Weight in SecondWeight Field", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;
                    //}
                    //if (Convert.ToDecimal("0" + txtfirstweight.Text) <= Convert.ToDecimal("0" + txtsecondweight.Text))
                    //{

                    decimal wt = Convert.ToDecimal("0" + txtsecondweight.Text) - Convert.ToDecimal("0" + txtfirstweight.Text);
                    txtnetweight.Text = wt.ToString("0.00");
                    txtnetweight_TextChanged(sender, e);
                    //}


                    // }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("txtsecondweight_TextChanged: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //try
            //{
            //    decimal wt = Convert.ToDecimal("0" + txtfirstweight.Text) - Convert.ToDecimal("0" + txtsecondweight.Text);
            //    txtnetweight.Text = wt.ToString("0.000");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("txtsecondweight_TextChanged: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void txtfirstweight_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtfirstweight.Text != "")
                {
                    //valid = va.IsIntegerdot(txtfirstweight.Text);
                    //if (valid == false)
                    //{
                    //    MessageBox.Show("Numeric Only Allowed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    txtfirstweight.Focus(); txtfirstweight.Text = ""; return;
                    //}
                    //if (valid == true)
                    //{
                    decimal wt = Convert.ToDecimal("0" + txtsecondweight.Text) - Convert.ToDecimal("0" + txtfirstweight.Text);
                    txtnetweight.Text = wt.ToString();

                    txtnetweight_TextChanged(sender, e);

                    // }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("txtfirstweight_TextChanged" + ex.Message);
            }

        }
        //DataTable dtprint1 = new DataTable();
        //DataTable dtprint2 = new DataTable();
        //ReportFormate.RawPrint_Formate rd = new ReportFormate.RawPrint_Formate();
        public void Prints()
        {
            Pinnacle.Report.CFM.DosPrint_DeliveryCertificatecs p = new Pinnacle.Report.CFM.DosPrint_DeliveryCertificatecs();
            try
            {
                if (Convert.ToInt64("0" + txtdeltid.Text) == 0 || txtdeltid.Text == "")
                {

                    Class.Users.Paramid = 0;
                    string sel1 = "select max(a.asptbldeliveryid) as asptbldeliveryid  from asptbldelivery a;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                    DataTable dt = ds.Tables["asptbldelivery"];
                    Class.Users.Paramid = Convert.ToInt64(dt.Rows[0]["asptbldeliveryid"].ToString());

                    p.Show();
                }
                if (Convert.ToInt64("0" + txtdeltid.Text) >= 1)
                {

                    p.Show();


                }
            }catch(Exception ex)
            {

            }
            finally
            {
                empty(); txtvechileno.Text = ""; p.Dispose();
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
        bool valid = false;
     
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPagedel1"])//your specific tabname
            {
                if (txtdeltid.Text == "")
                {
                    empty();
                    autonumberload();
                }
                txtvechileno.Select(); comboproduct.SelectedIndex = -1;


            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPagedel2"])//your specific tabname
            {
                txtsearch.Select();

            }
        }

        private void combocompname_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void combocompcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtvechileno.Select();


            }
        }



        private void txtvechileno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtsendto.Focus();

            }
        }

        private void dateTimePicker2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtfirstweight.Select();

            }
        }

        private void txtfirstweight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtsecondweight.Select();

            }
        }

        private void txtsecondweight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                txtsendto.Select();

            }
        }

        private void txtsendto_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtcertifiedby.Select();

            }
        }

        private void txtcertifiedby_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                comboproduct.Select();

            }
        }

        private void comboproduct_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtbags.Select();

            }
        }

        private void txtbags_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                butAddListView.Select();

            }
        }



        private void buttsave_Click(object sender, EventArgs e)
        {
           
            Saves();
            
        }

        private void butAddListView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                comboproduct.Select();
            }
        }
        private void buttsearch_Click(object sender, EventArgs e)
        {
           
        }

        private void buttprint_Click(object sender, EventArgs e)
        {
            tabcheck = true;
            Saves();
            if (validprint == true)
            {
                Prints();
                Class.Users.Paramid = 0; Class.Users.PayPeriod = "";
            }
        }

        private void buttdelete_Click(object sender, EventArgs e)
        {
            Deletes();tabcheck = true; tabControl1.SelectTab(tabPagedel2);
        }

        private void buttnew_Click(object sender, EventArgs e)
        {
            News(); tabcheck = true; tabControl1.SelectTab(tabPagedel2);
        }

        private void butgetall_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptbldeliveryid,a.asptbldeliveryid1,a.vechileNo,a.certificateno,a.sendto,date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,a.firstweight,a.secondweight,a.netweight,a.productweight from asptbldelivery a   where   a.deliverydate=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptbldeliveryid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                DataTable dt = ds.Tables["asptbldelivery"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldeliveryid"].ToString());

                        list.SubItems.Add(myRow["vechileNo"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["sendto"].ToString());
                        list.SubItems.Add(myRow["deliverydate"].ToString());
                        list.SubItems.Add(myRow["firstweight"].ToString());
                        list.SubItems.Add(myRow["secondweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());
                        list.SubItems.Add(myRow["productweight"].ToString());

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
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void butfirstweight_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptbldeliveryid,a.asptbldeliveryid1,a.vechileNo,a.certificateno,a.sendto,date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,a.firstweight,a.secondweight,a.netweight,a.productweight from asptbldelivery a where  a.secondweight = 0 and a.deliverydate=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptbldeliveryid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                DataTable dt = ds.Tables["asptbldelivery"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldeliveryid"].ToString());

                        list.SubItems.Add(myRow["vechileNo"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["sendto"].ToString());
                        list.SubItems.Add(myRow["deliverydate"].ToString());
                        list.SubItems.Add(myRow["firstweight"].ToString());
                        list.SubItems.Add(myRow["secondweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());
                        list.SubItems.Add(myRow["productweight"].ToString());
                        
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
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void butnetweight_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptbldeliveryid,a.asptbldeliveryid1,a.vechileNo,a.certificateno,a.sendto,date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,a.firstweight,a.secondweight,a.netweight,a.productweight from asptbldelivery a  where   a.secondweight >=1 and a.deliverydate=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptbldeliveryid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                DataTable dt = ds.Tables["asptbldelivery"];
                if (dt.Rows.Count >= 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldeliveryid"].ToString());

                        list.SubItems.Add(myRow["vechileNo"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["sendto"].ToString());
                        list.SubItems.Add(myRow["deliverydate"].ToString());
                        list.SubItems.Add(myRow["firstweight"].ToString());
                        list.SubItems.Add(myRow["secondweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());
                        list.SubItems.Add(myRow["productweight"].ToString());
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
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttsearch_Click_1(object sender, EventArgs e)
        {

        }

        private void txtsendto_TextChanged(object sender, EventArgs e)
        {
            //if (txtsendto.Text != "")
            //{
            //    if (va.IsStringSpace(txtsendto.Text) == false)
            //    {
            //        //words = Regex.Replace(words, @".", "");
            //        txtsendto.Text = "";
            //        MessageBox.Show("Send To Field Only Allowed Charectors: " + txtsendto.Text, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //    }
            //}

            //if (!System.Text.RegularExpressions.Regex.IsMatch(txtsendto.Text, "^[a-zA-Z ]"))


            //    txtsendto.Text.Remove(txtsendto.Text.Length - 1);
            //}
        }

        private void txtsendto_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || char.IsWhiteSpace(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back);
        }

        private void txtcertifiedby_KeyPress(object sender, KeyPressEventArgs e)
        {
           // e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == ','  || e.KeyChar == (char)Keys.Back);

        }

        private void txtfirstweight_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txtbags_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        }

        private void txtsecondweight_KeyPress(object sender, KeyPressEventArgs e)
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
        public  void  portconnection(string porttype, SerialPort _serialPort2)
        {
            //try
            //{
            //    porttype = Models.Serial.PortType;
            //    if (porttype == "TCP/IP")
            //    {
            //        _serialPort2.Close();
            //        string szPort = Class.Users.PortNo;
            //        int alPort = System.Convert.ToInt16(szPort, 10);
            //        _serialPort2.PortName = Class.Users.PortNo.ToString();
            //        _serialPort2.BaudRate = Convert.ToInt32("0" + Models.Serial.BaudRate);
            //        _serialPort2.DataBits = Convert.ToInt32("0" + Models.Serial.DataBits);
            //        _serialPort2.StopBits = (StopBits)Enum.Parse(typeof(StopBits), Models.Serial.StopBits);
            //        _serialPort2.Parity = (Parity)Enum.Parse(typeof(Parity), Models.Serial.Parity);
            //        Socket m_socClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            //        System.Net.IPAddress remoteIPAddress = System.Net.IPAddress.Parse(Class.Users.PortIP);
            //        System.Net.IPEndPoint remoteEndPoint = new System.Net.IPEndPoint(remoteIPAddress, alPort);
            //        m_socClient.Connect(remoteEndPoint);
            //        _serialPort2.Open();
            //        _serialPort2.DiscardInBuffer();
            //        _serialPort2.DiscardOutBuffer();
            //    }
            //    else
            //    {
            //        _serialPort2.Close();
            //        _serialPort2.PortName = Models.Serial.PortName;
            //        _serialPort2.BaudRate = Convert.ToInt32("0" + Models.Serial.BaudRate);
            //        _serialPort2.DataBits = Convert.ToInt32("0" + Models.Serial.DataBits);
            //        _serialPort2.StopBits = (StopBits)Enum.Parse(typeof(StopBits), Models.Serial.StopBits);
            //        _serialPort2.Parity = (Parity)Enum.Parse(typeof(Parity), Models.Serial.Parity);
            //        _serialPort2.Open();
            //        _serialPort2.DiscardInBuffer();
            //        _serialPort2.DiscardOutBuffer();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Digitalizer not Connected." + ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //}
           
        }
        private void butstart_Click(object sender, EventArgs e)
        {
           
            try
            {
                 if(Models.Serial.PortType !="TCP/IP")
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
        private void serialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //readserialvalue = serialPort2.ReadExisting();
        }


        private void txtvechileno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtnetweight_TextChanged(object sender, EventArgs e)
        {
            try
            {
                decimal d1 = Convert.ToDecimal("0" + txtnetweight.Text);

                decimal d2 = Convert.ToDecimal("0" + txtproductweight.Text);
                decimal d3 = d1 - d2;
                txtdifference.Text = d3.ToString();
            }
            catch (Exception ex) { }

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }



        private void tabPagedel1_Click(object sender, EventArgs e)
        {

        }

        private void butrefres_Click(object sender, EventArgs e)
        {
            txtbags.Text = ""; comboproductload(); comboproduct.SelectedIndex = -1; comboproduct.Text = ""; invalid = 0; gridid = 0;
        }

        private void enableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            combocompcode.Enabled = true;
        }

        private void previousCertificateNoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (txtdeltid.Text != "")
            {
                string sel1 = "select a.asptbldeliveryid,a.asptbldeliveryid1 ,a.finyear,d.compcode, d.compname,a.certificateno   from asptbldelivery a     join gtcompmast d on a.compname=d.gtcompmastid  where a.asptbldeliveryid='" + txtdeltid.Text + "' ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                DataTable dt = ds.Tables["asptbldelivery"];

                if (dt.Rows.Count > 0)
                {
                    txtdeltid.Text = Convert.ToString(dt.Rows[0]["asptbldeliveryid"].ToString());
                    txtdeltid1.Text = Convert.ToString(dt.Rows[0]["asptbldeliveryid1"].ToString());
                    combofinyear.Text = Convert.ToString(dt.Rows[0]["finyear"].ToString());
                    combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                    combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                    txtcertificateno.Text = Convert.ToString(dt.Rows[0]["certificateno"].ToString());
                }
            }
        }

        private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
           
            try
            {
                if (listView2.Items.Count == 0)
                {
                    MessageBox.Show(
                        "If you want to Remove, Double Click a Specific Row in ListView.",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                if (MessageBox.Show(
                    "Do You want Delete this Record ?",
                    "Information",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    return;
                }

                decimal totalWeight = 0;
                decimal.TryParse(txtproductweight.Text, out totalWeight);

                // Collect selected items first
                var selectedItems = listView2.SelectedItems.Cast<ListViewItem>().ToList();

                foreach (ListViewItem item in selectedItems)
                {
                    long detId = 0;
                    long.TryParse(item.SubItems[2].Text, out detId);

                    decimal itemKgs = 0;
                    decimal.TryParse(item.SubItems[7].Text, out itemKgs);

                    if (detId > 1)
                    {
                        string del1 = "delete from asptbldeliverydet where asptbldeliverydetid='" + detId + "';";
                        Utility.ExecuteNonQuery(del1);
                        // INSERT HISTORY
                        string ins2 = "insert into asptbldeliverydet2(asptbldeliverydetid,asptbldeliveryid ,asptbldeliveryid1,productname , productweight , productkgs,operation,compcode,finyear) values('" + detId + "' ,'" + txtdeltid.Text.ToString() + "' ,'" + txtdeltid1.Text + "' , '" + item.SubItems[4].Text + "' ,'" + item.SubItems[6].Text + "' , '" + item.SubItems[7].Text + "','DELETED','" + combocompcode.SelectedValue + "','" + combofinyear.Text + "' );";
                        Utility.ExecuteNonQuery(ins2);
                    }
                    totalWeight -= itemKgs;
                    MessageBox.Show($"Index: {item.SubItems[1].Text}   Name: {item.SubItems[5].Text}\nDeleted Successfully", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    listView2.Items.Remove(item);
                }

                // Update UI once
                txtproductweight.Text = totalWeight > 0 ? totalWeight.ToString() : "";
                lbltotalfooterkgs.Text = txtproductweight.Text;
                listview2totalweight = totalWeight;
                comboproduct.Text = "";

                if (listView2.Items.Count == 0)
                {
                    string dat = "";
                    dat = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string dat1 = "";
                    dat1 = System.DateTime.Now.ToString("hh:mm:ss");
                    string dat2 = "";
                    dat2 = Convert.ToDateTime(System.DateTime.Now.ToString()).ToString("yyyy-mm-dd");
                    comboproduct.SelectedIndex = -1;
                    txtdifference.Text = "";
                    lbltotal1.Text = "Total Count:";
                    string ins1 = "insert into asptbldelivery2(asptbldeliveryid,asptbldeliveryid1,finyear,compcode,compname,certificateno ,sendto, vechileNo, deliverydate,deliverydate1, deliverytime , firstweight, secondweight, netweight , productweight , certifiedby ,compcode1,  username,createdby, modifiedby,ipaddress,weightdeffer,modifiedon,operation) values('" + txtdeltid.Text + "','" + txtdeltid1.Text + "','" + combofinyear.Text + "','" + combocompname.SelectedValue.ToString() + "','" + combocompname.SelectedValue.ToString() + "','" + txtcertificateno.Text.ToUpper() + "','" + txtsendto.Text.ToUpper() + "','" + txtvechileno.Text.ToUpper() + "','" + Convert.ToDateTime(dat).ToString("dd-MM-yyyy") + " " + dat1 + "','" + dat + "','" + dat1 + "','" + txtfirstweight.Text + "','" + Convert.ToDecimal("0" + txtsecondweight.Text) + "','" + txtnetweight.Text + "','" + Convert.ToDecimal("0") + "','" + txtcertifiedby.Text.ToUpper() + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + " : "+ Class.Users.IPADDRESS +  "','"+ Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "','" + txtdifference.Text + "','" + System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','DELETED');";
                    Utility.ExecuteNonQuery(ins1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void txtvechileno_TextChanged(object sender, EventArgs e)
        {

        }
  

        public void Searchs()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptbldeliveryid,a.asptbldeliveryid1,a.vechileNo,a.certificateno,a.sendto,date_format(a.deliverydate,'%d-%m-%Y') as deliverydate,a.firstweight,a.secondweight,a.netweight,a.productweight from asptbldelivery a where  a.deliverydate=date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.asptbldeliveryid desc;";
                //substr(a.deliverydate,1,10)='" + dateTimePicker3.Value.ToShortDateString() + "' ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldelivery");
                DataTable dt = ds.Tables["asptbldelivery"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldeliveryid"].ToString());

                        list.SubItems.Add(myRow["vechileNo"].ToString());
                        list.SubItems.Add(myRow["certificateno"].ToString());
                        list.SubItems.Add(myRow["sendto"].ToString());
                        list.SubItems.Add(myRow["deliverydate"].ToString());
                        list.SubItems.Add(myRow["firstweight"].ToString());
                        list.SubItems.Add(myRow["secondweight"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());
                        list.SubItems.Add(myRow["productweight"].ToString());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Searchs(int EditID)
        {
            
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
    }
}

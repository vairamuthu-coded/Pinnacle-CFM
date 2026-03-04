
using ExcelDataReader;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace Pinnacle.Transactions.Bank
{
    public partial class AdvancePayment : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        public AdvancePayment()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(System.DateTime.Now.ToShortTimeString().ToString());
            Class.Users.UserTime = 0;

        }

        private static AdvancePayment _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.Validate va = new Models.Validate();
        string filepath =  System.IO.Directory.GetCurrentDirectory(); FileStream FS = null;
        protected override bool ProcessCmdKey(ref Message message, Keys keys)
        {
            switch (keys)
            {
                case Keys.N | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    News();
                    return true; // signal that we've processed this key
                case Keys.S | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Saves();
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
                    return true; // signal that we've processed this key
            }
            // run base implementation
            return base.ProcessCmdKey(ref message, keys);
        }
        public static AdvancePayment Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new AdvancePayment();
                GlobalVariables.CurrentForm = _instance; return _instance;

            }
        }


     
        private void AdvancePayment_Load(object sender, EventArgs e)
        {
          
            int MON = Convert.ToInt32(txtadvDate.Value.ToString().Substring(3, 2));
            string getmonth = getFullName(MON);
            lblmonth.Text = getmonth;
            int year = Convert.ToInt32(txtadvDate.Value.ToString().Substring(6, 4));
            string getyear = getFullName1(year);
            lblyear.Text = getyear;
        }
        System.Text.StringBuilder sb = new System.Text.StringBuilder();

        public void PartyLoad()
        {
            sb.Clear();
            sb.Append("SELECT  a.asptblpartymasID,a.partyname from asptblpartymas a where a.active='T'");           
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblpartymas");
            DataTable dt = ds.Tables["asptblpartymas"];           
            comboparty.DisplayMember = "partyname";
            comboparty.ValueMember = "asptblpartymasID";
            comboparty.DataSource = dt;
           
        }

        public void PersonLoad()
        {
            sb.Clear();
            sb.Append("SELECT  a.asptblresmasid,a.resonseperson from asptblresmas a where a.active='T'");
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblresmas");
            DataTable dt = ds.Tables["asptblresmas"];
            comboresponsibleperson.DisplayMember = "resonseperson";
            comboresponsibleperson.ValueMember = "asptblresmasid";
            comboresponsibleperson.DataSource = dt;

        }

        public void DeptLoad()
        {
            sb.Clear();
            sb.Append("SELECT  a.asptbldeptmasid,a.department from asptbldeptmas a where a.active='T'");
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptbldeptmas");
            DataTable dt = ds.Tables["asptbldeptmas"];
            combodept.DisplayMember = "department";
            combodept.ValueMember = "asptbldeptmasid";
            combodept.DataSource = dt;

        }
        string maxid = ""; string maxid1 = "";
        public void Saves()
        {
            try
            {
                MySqlCommand cmd;
                int MON = Convert.ToInt32(txtadvDate.Value.ToString().Substring(3, 2));
                string getmonth = getFullName(MON);
                if (combodept.SelectedValue == null)
                {
                    MessageBox.Show("'Department  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combodept.Focus();
                    return;
                }
                if (comboparty.SelectedValue == null)
                {
                    MessageBox.Show("'Supplier Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.comboparty.Focus();
                    return;
                }
                if (txtorderno.Text == "")
                {

                    MessageBox.Show("'Order Number  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtorderno.Focus(); return;
                }
                if (txtitemdesc.Text == "")
                {

                    MessageBox.Show("'Item Description  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtitemdesc.Focus(); return;
                }
                if (comboresponsibleperson.SelectedValue == null)
                {
                    MessageBox.Show("'Response Person  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.comboresponsibleperson.Focus();
                    return;
                }

                if (txtinvoiceamt.Text == "")
                {
                    MessageBox.Show("'Invoice Amount  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtinvoiceamt.Focus();
                    return;
                }
                if (txtadvterms.Text == "" || txtadvterms.Text=="0")
                {
                    MessageBox.Show("'Advance Terms  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtadvterms.Focus();
                    return;
                }
                if (txtadvamount.Text == "")
                {
                    MessageBox.Show("'Advance Amount  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtadvamount.Focus();
                    return;
                }
                if (txtgst.Text == "")
                {
                    MessageBox.Show("'Enter GST %  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtgst.Focus();
                    return;
                }
                if (txttds.Text == "")
                {
                    MessageBox.Show("'Enter TDS %  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txttds.Focus();
                    return;
                }
                if (combomodeofpayment.SelectedIndex == -1)
                {
                    MessageBox.Show("'Mode of Payment  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combomodeofpayment.Focus();
                    return;
                }
               
                else
                {
                    string sel = "select asptbladvpaymasid    from  asptbladvpaymas    WHERE department='" + combodept.SelectedValue + "' and partyname='" + comboparty.SelectedValue + "' and orderno='" + txtorderno.Text.ToUpper().Trim() + "' and itemdesc='" + txtitemdesc.Text.ToUpper().Trim() + "' and responseperson='" + comboresponsibleperson.SelectedValue + "' and invoiceamt='" + txtinvoiceamt.Value.ToString().Replace(",", "").Trim() + "' and advanceterms='" + txtadvterms.Text.ToUpper().Trim() + "' and advanceamount=round('" + txtadvamount.Value.ToString().Replace(",", "").Trim() + "') and modeofpayment='" + combomodeofpayment.Text + "' and dateofpayment='" + txtadvDate.Value.ToString("yyyy-MM-dd") + "'  and amount='" + txtamount.Value.ToString().Replace(",", "").Trim() + "' and gst='" + txtgst.Text+ "' and tds='" + txttds.Text + "' and tdsvalue='" + txttdsvalue.Value.ToString().Replace(",", "").Trim() + "'and deductionamt='" + txtdedutionamt.Value.ToString().Replace(",", "").Trim() + "' ";//and a.ledgerDate='" + txtledgerDate.Value.ToString("yyyy-MM-dd") + "'
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbladvpaymas");
                    DataTable dt = ds.Tables["asptbladvpaymas"];
                    if (dt.Rows.Count != 0)
                    {
                        maxid = txtadvpayid.Text;
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtadvpayid.Text) == 0 || Convert.ToInt32("0" + txtadvpayid.Text) == 0)
                    {
                        string ins = "insert into asptbladvpaymas(department,partyname,orderno,itemdesc,responseperson,invoiceamt,gst,amount,tds,tdsvalue,deductionamt,advanceterms,advanceamount,modeofpayment,dateofpayment,finyear,COMPCODE,USERNAME,MODIFIEDON,CREATEDBY,CREATEDON,IPADDRESS)  VALUES('" + combodept.SelectedValue + "','" + comboparty.SelectedValue + "','" + txtorderno.Text.ToUpper().Trim() + "','" + txtitemdesc.Text.ToUpper().Trim() + "','" + comboresponsibleperson.SelectedValue + "','" + txtinvoiceamt.Value.ToString().Replace(",", "").Trim() + "','" + txtgst.Text + "','" + txtamount.Value.ToString().Replace(",", "").Trim() + "','" + txttds.Text + "', '" + txttdsvalue.Value.ToString().Replace(",", "").Trim() + "','" + txtdedutionamt.Value.ToString().Replace(",", "").Trim() + "','" + txtadvterms.Text.ToUpper().Trim() + "',round('" + txtadvamount.Value.ToString().Replace(",", "").Trim() + "'),'" + combomodeofpayment.Text + "','" + txtadvDate.Value.ToString("yyyy-MM-dd") + "','"+lblyear.Text+ "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.IPADDRESS + "')";
                        Utility.ExecuteNonQuery(ins);

                        string sel2 = "select  max(asptbladvpaymasid) id  from  asptbladvpaymas   where  compcode='" + Class.Users.COMPCODE + "' and finyear='" + lblyear.Text + "'";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbladvpaymas");
                        DataTable dt2 = ds2.Tables["asptbladvpaymas"];
                        maxid = dt2.Rows[0]["id"].ToString();
                    }
                    else
                    {
                        string up = "update  asptbladvpaymas  set department='" + combodept.SelectedValue + "',partyname='" + comboparty.SelectedValue + "',orderno='" + txtorderno.Text.ToUpper().Trim() + "',itemdesc='" + txtitemdesc.Text.ToUpper().Trim() + "',responseperson='" + comboresponsibleperson.SelectedValue + "',invoiceamt='" + txtinvoiceamt.Value.ToString().Replace(",", "").Trim() + "',advanceterms='" + txtadvterms.Text.ToUpper().Trim() + "',advanceamount=round('" + txtadvamount.Value.ToString().Replace(",", "").Trim() + "'),modeofpayment='" + combomodeofpayment.Text + "',dateofpayment='" + txtadvDate.Value.ToString("yyyy-MM-dd") + "', amount='" + txtamount.Value.ToString().Replace(",", "").Trim() + "',gst='" + txtgst.Text + "' , tds='" + txttds.Text + "',tdsvalue='" + txttdsvalue.Value.ToString().Replace(",", "").Trim() + "',deductionamt='" + txtdedutionamt.Value.ToString().Replace(",", "").Trim() + "',COMPCODE='" + Class.Users.COMPCODE + "',USERNAME='" + Class.Users.USERID + "',MODIFIEDON='" + System.DateTime.Now.ToString() + "',CREATEDBY='" + Class.Users.HUserName + "',CREATEDON='" + System.DateTime.Now.ToString() + "',IPADDRESS='" + Class.Users.IPADDRESS + "',finyear='" + lblyear.Text + "' where asptbladvpaymasid='" + txtadvpayid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        maxid = txtadvpayid.Text;
                    }

                    Models.Bank.AdvancePayment pp = new Models.Bank.AdvancePayment();
                    int i = 0, j = 1; int cnt = dataGridView1.Rows.Count-1;
                    if (cnt >= 0)
                    {
                       
                        for (i = 0; i < cnt; i++)
                        {
                            
                            if (dataGridView1.Rows[i].Cells[0].Value == null)
                            {
                                pp.asptbladvpaydetid= Convert.ToInt64("0");
                            }
                            else
                            {
                                pp.asptbladvpaydetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                            }                        
                         
                            pp.asptbladvpaymasid = Convert.ToInt64("0" + maxid);
                            pp.compcode = Class.Users.COMPCODE; 
                            pp.compcode = Convert.ToInt64("0" + Class.Users.COMPCODE);
                            pp.department = Convert.ToInt64("0" + combodept.SelectedValue);
                            pp.partyname = Convert.ToInt64("0" + comboparty.SelectedValue);
                            pp.invoicetype = dataGridView1.Rows[i].Cells[1].EditedFormattedValue.ToString();
                            pp.invoice = dataGridView1.Rows[i].Cells[2].EditedFormattedValue.ToString();
                            try
                            {
                                
                                if (stdbytes != null)
                                {
                                    pp.invoicebyte = Convert.ToInt64("0" + stdbytes.Length);
                                }
                                if (stdbytes2 != null)
                                {
                                    pp.proforminvoicebyte = Convert.ToInt64("0" + stdbytes2.Length);
                                }
                                if (stdbytes3 != null)
                                {
                                    pp.quatationbyte = Convert.ToInt64("0" + stdbytes3.Length);
                                }
                                else
                                {
                                    pp.quatationbyte = 0;
                                }
                                if (stdbytes4 != null)
                                {
                                    pp.otherbyte = Convert.ToInt64("0" + stdbytes4.Length.ToString());
                                }
                                else
                                {
                                    pp.otherbyte = 0;
                                }
                                if (stdbytes5 != null)
                                {
                                    pp.powobyte = Convert.ToInt64("0" + stdbytes5.Length.ToString());
                                }
                                else
                                {
                                    pp.powobyte = 0;
                                }
                            }
                            catch(Exception ex) { }
                                if (pp.invoicetype  != "" && pp.invoice !="")
                            {
                                Models.CommonClass CC = new Models.CommonClass();
                                DataTable dt1 = CC.select("select asptbladvpaydetid  from  asptbladvpaydet   where  asptbladvpaymasid='" + pp.asptbladvpaymasid + "' and  compcode='" + pp.compcode + "' and  department='" + pp.department + "'  and  partyname='" + pp.partyname + "' and invoicetype='" + pp.invoicetype + "' and invoice='" + pp.invoice + "'and invoicebyte='"+pp.invoicebyte+ "' and proforminvoicebyte='" + pp.proforminvoicebyte + "' and quatationbyte='" + pp.quatationbyte + "' and otherbyte='" + pp.otherbyte + "' and powobyte='" + pp.powobyte+ "'", "asptbladvpaydet");
                                if (dt1.Rows.Count != 0) { }
                                else if (dt1.Rows.Count !=0 && pp.asptbladvpaydetid == 0 || pp.asptbladvpaydetid == 0) //,&& pp.asptbladvpaydetid == 0 || pp.asptbladvpaydetid == 0asptbladvpaymasid,compcode,department,partyname,invoicetype,invoice 
                                {
                                    string ins1 = "insert into asptbladvpaydet(asptbladvpaymasid,compcode,department,partyname,invoicetype,invoice,INVBLOB,INVPROBLOB,QUABLOB,powoblob,OTHBLOB,invoicebyte,proforminvoicebyte,quatationbyte,otherbyte,powobyte) values('" + pp.asptbladvpaymasid + "' ,'" + pp.compcode + "' ,'" + pp.department + "' ,'" + pp.partyname + "','" + pp.invoicetype + "','" + pp.invoice + "',@INVBLOB,@INVPROBLOB,@QUABLOB,@powoblob,@OTHBLOB,'" + pp.invoicebyte+ "','" + pp.proforminvoicebyte + "','" + pp.quatationbyte + "','" + pp.powobyte + "','" + pp.otherbyte + "');";
                                    cmd = new MySqlCommand(ins1, Utility.con);
                                    cmd.Parameters.Add("@INVBLOB", MySqlDbType.LongBlob);
                                    cmd.Parameters.Add("@INVPROBLOB", MySqlDbType.LongBlob);
                                    cmd.Parameters.Add("@QUABLOB", MySqlDbType.LongBlob);                                    
                                    cmd.Parameters.Add("@powoblob", MySqlDbType.LongBlob);
                                    cmd.Parameters.Add("@OTHBLOB", MySqlDbType.LongBlob);
                                    cmd.Parameters["@INVBLOB"].Value= stdbytes;
                                    cmd.Parameters["@INVPROBLOB"].Value = stdbytes2;
                                    cmd.Parameters["@QUABLOB"].Value = stdbytes3;                                  
                                    cmd.Parameters["@powoblob"].Value = stdbytes4;
                                    cmd.Parameters["@OTHBLOB"].Value = stdbytes5;
                                    cmd.ExecuteNonQuery();                                   
                                    
                                }
                                else
                                {
                                    string query = "update   asptbladvpaydet set  asptbladvpaymasid='" + pp.asptbladvpaymasid + "',compcode='" + pp.compcode + "',department='" + pp.department + "',partyname='" + pp.partyname + "',invoicetype='" + pp.invoicetype + "',invoice='" + pp.invoice + "',INVBLOB=@INVBLOB,INVPROBLOB=@INVPROBLOB,QUABLOB=@QUABLOB,powoblob=@powoblob,OTHBLOB=@OTHBLOB,invoicebyte='" + pp.invoicebyte + "' , proforminvoicebyte='" + pp.proforminvoicebyte + "' , quatationbyte='" + pp.quatationbyte + "' , powobyte='" + pp.powobyte + "' , otherbyte='" + pp.otherbyte + "' where  asptbladvpaydetid=" + pp.asptbladvpaydetid;
                                    cmd = new MySqlCommand(query, Utility.con);
                                    cmd.Parameters.Add("@INVBLOB", MySqlDbType.LongBlob);
                                    cmd.Parameters.Add("@INVPROBLOB", MySqlDbType.LongBlob);
                                    cmd.Parameters.Add("@QUABLOB", MySqlDbType.LongBlob);                                    
                                    cmd.Parameters.Add("@powoblob", MySqlDbType.LongBlob);
                                    cmd.Parameters.Add("@OTHBLOB", MySqlDbType.LongBlob);
                                    cmd.Parameters["@INVBLOB"].Value = stdbytes;
                                    cmd.Parameters["@INVPROBLOB"].Value = stdbytes2;
                                    cmd.Parameters["@QUABLOB"].Value = stdbytes3;
                                    cmd.Parameters["@powoblob"].Value = stdbytes4;
                                    cmd.Parameters["@OTHBLOB"].Value = stdbytes5;
                                   
                                    cmd.ExecuteNonQuery();                       
                                    
                                }
                               
                            }
                        }
           
                    }
                    if (txtadvpayid.Text == "")
                    {
                        mas.pop(comboparty.Text, "Department : "+ combodept.Text, "Amount : "+txtadvamount.Value.ToString());
                        MessageBox.Show("Record Saved  Successfully", "Success");

                        GridLoad();empty();
                    }
                    else
                    {
                        mas.pop(comboparty.Text, "Department : " + combodept.Text, "Amount : " + txtadvamount.Value.ToString());
                        MessageBox.Show("Record Updated  Successfully", "Updated");
                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(" " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void LedgerMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        static string getFullName(int month)
        {
            DateTime date = new DateTime(System.DateTime.Now.Year, month,System.DateTime.Now.Day);

            return date.ToString("MMMM");
        }
        static string getFullName1(int year)
        {
            DateTime date = new DateTime(year, System.DateTime.Now.Month, System.DateTime.Now.Day);

            return date.ToString("yyyy");
        }
        public void News()
        {
            GridLoad();
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            butheader.BackColor = Class.Users.BackColors;
            button1.BackColor = Class.Users.BackColors;
            dataGridView1.Rows.Clear(); checkall.Checked = false;
            CommonFunctions.SetRowNumber(dataGridView1);
            PartyLoad(); PersonLoad();DeptLoad(); empty();

        }
        private void empty()
        {
            txtadvpayid.Text = ""; checkall.BackColor = Class.Users.BackColors;
            filepath = ""; txtgstvalue.Text = "";
           Class.Users.UserTime = 0;
            combodept.SelectedIndex = -1;  combodept.Select(); txtitemdesc.Text = "";txtinvoiceamt.Text = "";
            comboparty.Text = ""; comboparty.Text = ""; txtadvDate.Text = ""; 
            combomodeofpayment.SelectedIndex = -1;txtadvterms.Text = "";
            comboresponsibleperson.SelectedIndex = -1;txtadvamount.Text = "";combomodeofpayment.Text = "";txtinvoiceamt.Text = "";
            txtorderno.Text = "";txtamount.Text = "";txtgst.Text = "";txttds.Text = "";txttdsvalue.Text = ""; txtdedutionamt.Value = 0; txtinvoiceamt.Value = 0;
            
            //do
            //{
            //    int i = 0;
            //    for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }

            //}
            //while (dataGridView1.Rows.Count > 0);
            dataGridView1.Rows.Clear();
            dataGridView1.AllowUserToAddRows = true;
        }
        public void GridLoad()
        {
            try
            {
                if (checkall.Checked == true)
                {
                    Class.Users.SearchQuery = " SELECT a.asptbladvpaymasid as ID, b.department AS Department ,c.partyname as Supplier,a.orderno as OrderNo,d.resonseperson as  ResponsePerson,a.gst as Gst,a.tds as TDS,a.modeofpayment as PaymentType, a.advanceamount as InvAmount,a.dateofpayment as Date,a.approval as Manager,a.mdapproval as MD from asptbladvpaymas a join asptbldeptmas b on b.asptbldeptmasID=a.department join asptblpartymas c on c.asptblpartymasid=a.partyname join asptblresmas d on d.asptblresmasid=a.responseperson   order by a.asptbladvpaymasid desc;";

                    Class.Users.HideCols = new string[] { "ID" };
                    uccListView2.Load_Details();
                }
                else
                {
                    Class.Users.SearchQuery = " SELECT a.asptbladvpaymasid as ID, b.department AS Department ,c.partyname as Supplier,a.orderno as OrderNo,d.resonseperson as  ResponsePerson,a.gst as Gst,a.tds as TDS,a.modeofpayment as PaymentType, a.advanceamount as InvAmount,a.dateofpayment as Date,a.approval as Manager,a.mdapproval as MD from asptbladvpaymas a join asptbldeptmas b on b.asptbldeptmasID=a.department join asptblpartymas c on c.asptblpartymasid=a.partyname join asptblresmas d on d.asptblresmasid=a.responseperson where a.dateofpayment='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' order by a.asptbladvpaymasid desc;";

                    Class.Users.HideCols = new string[] { "ID" };
                    uccListView2.Load_Details();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }
        public void Searchs(int EditID)
        {
            try
            {
                if (EditID > 0)
                {

                     stdbytes = null; stdbytes2 = null; stdbytes3 = null; stdbytes4 = null; stdbytes5 = null;

                    txtadvpayid.Text = EditID.ToString();
                    DataTable dt = Utility.SQLQuery("SELECT a.asptbladvpaymasid as ID, b.department  ,c.partyname ,a.orderno ," +
                        "a.itemdesc, d.resonseperson ,a.invoiceamt,a.advanceterms,a.advanceamount,a.modeofpayment,a.amount,a.gst,a.tds,a.approval from asptbladvpaymas a join asptbldeptmas b on b.asptbldeptmasID=a.department join asptblpartymas c on c.asptblpartymasid=a.partyname join asptblresmas d on d.asptblresmasid=a.responseperson     where a.asptbladvpaymasid='" + txtadvpayid.Text + "'");
                    if (dt.Rows.Count > 0)
                    {

                        txtadvpayid.Text = Convert.ToString(dt.Rows[0]["ID"].ToString());
                        combodept.Text = Convert.ToString(dt.Rows[0]["department"].ToString());
                        comboparty.Text = dt.Rows[0]["partyname"].ToString();
                        txtorderno.Text = dt.Rows[0]["orderno"].ToString();
                        txtitemdesc.Text = dt.Rows[0]["itemdesc"].ToString();
                        comboresponsibleperson.Text = dt.Rows[0]["resonseperson"].ToString();
                        txtinvoiceamt.Text = dt.Rows[0]["invoiceamt"].ToString();
                        txtadvterms.Text = dt.Rows[0]["advanceterms"].ToString();
                        txtadvamount.Text = dt.Rows[0]["advanceamount"].ToString();
                        combomodeofpayment.Text = dt.Rows[0]["modeofpayment"].ToString();
                        txtamount.Text = dt.Rows[0]["amount"].ToString();
                        txtgst.Text = dt.Rows[0]["gst"].ToString();
                        txttds.Text = dt.Rows[0]["tds"].ToString();
                        DataTable dt1 = Utility.SQLQuery("SELECT  a.asptbladvpaydetid,a.asptbladvpaymasid,a.compcode,b.department,c.partyname,a.invoicetype,a.invoice,a.INVBLOB,a.INVPROBLOB,a.QUABLOB,a.OTHBLOB,a.powoblob from asptbladvpaydet a join asptbldeptmas b on b.asptbldeptmasid=a.department join asptblpartymas c on c.asptblpartymasid=a.partyname   where a.asptbladvpaymasid='" + txtadvpayid.Text + "'");
                        dataGridView1.Rows.Clear();
                        int i = 0;
                        foreach (DataRow myRow in dt1.Rows)
                        {
                            string filepath1 = "";
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[i].Cells[0].Value = myRow["asptbladvpaydetid"].ToString();
                            dataGridView1.Rows[i].Cells[1].Value = myRow["invoicetype"].ToString();
                            dataGridView1.Rows[i].Cells[2].Value = myRow["invoice"].ToString();
                            string[] ext = myRow["invoice"].ToString().Split('.');
                            if (i == 0)
                            {
                                stdbytes = (byte[])myRow["INVBLOB"];
                               
                                //filepath1 = "D:\\temp1."+ext[1].ToString();
                               // FS = new FileStream(filepath1, System.IO.FileMode.Create);
                               // FS.Write(stdbytes, 0, stdbytes.Length);
                               // FS.Close();
                            }
                            if (i == 1)
                            {
                                stdbytes2 = (byte[])myRow["INVPROBLOB"];
                                //filepath1 = "";
                                //filepath1 = "D:\\temp1." + ext[1].ToString();
                                //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                                //FS.Write(stdbytes2, 0, stdbytes2.Length);
                                //FS.Close();
                            }
                            if (i == 2)
                            {
                                stdbytes3 = (byte[])myRow["QUABLOB"];

                                //filepath1 = "D:\\temp1." + ext[1].ToString();
                                //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                                //FS.Write(stdbytes3, 0, stdbytes3.Length);
                                //FS.Close();
                            }
                            if (i == 3)
                            {
                                stdbytes4 = (byte[])myRow["powoblob"];
                                //filepath1 = "";
                                //filepath1 = "D:\\temp1." + ext[1].ToString();
                                //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                                //FS.Write(stdbytes4, 0, stdbytes4.Length);
                                //FS.Close();
                            }
                            if (i == 4)
                            {
                                stdbytes5 = (byte[])myRow["OTHBLOB"];
                                //filepath1 = "";
                                //filepath1 = "D:\\temp1." + ext[1].ToString();
                                //FS = new FileStream(filepath1, System.IO.FileMode.Create);
                                //FS.Write(stdbytes5, 0, stdbytes5.Length);
                                //FS.Close();
                            }
                            i++;
                           
                        }

                    }
                    amountinwords(); 
                    CommonFunctions.SetRowNumber(dataGridView1);
                    tabControl1.SelectTab(tabPage1);
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
           
        }
      
        //private void ListView1_ItemActivate(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (listView1.Items.Count >= 0)
        //        {

        //            txtledgerid.Text = listView1.SelectedItems[0].SubItems[2].Text;
        //            string sel1 = "select a.asptbladvpaymasid,b.partyname, a.ledgermonth, a.amount , a.active,a.partyname    from  asptbladvpaymas a  join asptbladvpaymas b on a.partyname=b.asptbladvpaymasid where a.asptbladvpaymasid=" + txtledgerid.Text;
        //            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbladvpaymas");
        //            DataTable dt = ds.Tables["asptbladvpaymas"];

        //            if (dt.Rows.Count > 0)
        //            {
        //                txtledgerid.Text = Convert.ToString(dt.Rows[0]["asptbladvpaymasid"].ToString());
        //                comboparty.Text = Convert.ToString(dt.Rows[0]["partyname"].ToString());
        //                txtledgermonth.Text = dt.Rows[0]["month"].ToString();
        //                txtamt.Text = dt.Rows[0]["amount"].ToString();
        //                if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
        //              //  txtpartyname.Text = Convert.ToString(dt.Rows[0]["partyname"].ToString())+;

        //                txtledgerDate.Text = comboparty.Text + "-" + txtmonth.Text;
        //            }
        //          //  PartyLoad();
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        //private void Txtsearch_TextChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        int item0 = 0;
        //        if (txtsearch.Text.Length > 1)
        //        {
        //            listView1.Items.Clear();
        //            foreach (ListViewItem item in listfilter.Items)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                if (item.SubItems[3].ToString().Contains(txtsearch.Text.ToUpper()) || item.SubItems[4].ToString().Contains(txtsearch.Text.ToUpper()))
        //                {
        //                    list.Text = item.SubItems[0].Text;
        //                    list.SubItems.Add(item.SubItems[1].Text);
        //                    list.SubItems.Add(item.SubItems[2].Text);
        //                    list.SubItems.Add(item.SubItems[3].Text);
        //                    list.SubItems.Add(item.SubItems[4].Text);
        //                    list.SubItems.Add(item.SubItems[5].Text);
        //                    list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
        //                    listView1.Items.Add(list);
        //                }
        //                item0++;
        //            }
        //        }
        //        else
        //        {

        //            listView1.Items.Clear();
        //            GridLoad();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    lbltotal.Refresh();
        //    lbltotal.Text = "Total "+listView1.Items.Count.ToString();
        //}


        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PartyLoad(); PersonLoad(); DeptLoad(); GridLoad(); Class.Users.UserTime = 0;
        }

        private void Txtamount_TextChanged(object sender, EventArgs e)
        {
          
            if (va.IsIntegerdot(comboparty.Text) == false)
            {
               // txtamount.Text = "";
            }
            if (va.IsIntegerdot(comboparty.Text) == true)
            {
                comboparty.Text = "";
                
            }
        }
        //private void Txtmonth_TextChanged(object sender, EventArgs e)
        //{
        //    txtledgerDate.Text = ""; 
        // //   txtpartyname.Text = combopartyname.Text + " " + txtmonth.Text;
        //    txtledgerDate.Text = comboparty.Text + " " + txtledgermonth.Text + " " + "KG"; ;
        //}

        public void Deletes()
        {
            if (txtadvpayid.Text != "")
            {
                string sel1 = "select a.asptbladvpaymasid from asptbladvpaymas a join asptblledmas b on a.asptbladvpaymasid=b.asptbladvpaymasid where a.asptbladvpaymasid='" + txtadvpayid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbladvpaymas");
                DataTable dt = ds.Tables["asptbladvpaymas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + combodept.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;

                }
                else
                {

                    string del = "delete from asptbladvpaymas where asptbladvpaymasid='" + Convert.ToInt64("0" + txtadvpayid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    string del1 = "delete from asptbladvpaydet where asptbladvpaymasid='" + Convert.ToInt64("0" + txtadvpayid.Text) + "'";
                    Utility.ExecuteNonQuery(del1);
                    MessageBox.Show("Record Deleted Successfully " + combodept.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void combopartyname_KeyDown(object sender, KeyEventArgs e)
        {
           
        }

        private void txtmonth_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        private void txtmonth_KeyPress(object sender, KeyPressEventArgs e)
        {
          
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
         //   e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtamount_KeyPress(object sender, KeyPressEventArgs e)
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

        private void combopartyname_MouseClick(object sender, MouseEventArgs e)
        {
            //this.combodept.BackColor = Color.Yellow;
        }

        private void txtamount_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                
            }
        }

        public void Prints()
        {
            
        }

        public void Searchs()
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
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void combobank_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboresponsibleperson.SelectedValue != null && comboresponsibleperson.SelectedValue.ToString() != "System.Data.DataRowView" && comboresponsibleperson.Text !="")
            {
               
                string sel1 = "select distinct b.bankname from gtcompmast a  join asptblbanmas b on b.asptblbanmasid=a.bankname  where a.active='T'  and a.compcode='" + Class.Users.HCompcode + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbladvpaymas");
                DataTable dt = ds.Tables["asptbladvpaymas"];
                if (comboresponsibleperson.Text == dt.Rows[0]["bankname"].ToString())
                {

                    combomodeofpayment.SelectedIndex = 1;
                }
                else
                {

                    combomodeofpayment.SelectedIndex = 2;
                }
            }
        }

        private void txtamt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == ',' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back);
        }

        private void comboparty_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combodept.Text != "System.Data.DataRowView" && combodept.Text != "")
            {
                sb.Clear();
                sb.Append("SELECT  a.asptblbanmasid,a.bankname from asptblbanmas a join asptblpartymas b on a.asptblbanmasid=b.bankname  where b.active='T' and b.asptblpartymasid='"+combodept.SelectedValue+"'");
                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblbanmas");
                DataTable dt = ds.Tables["asptblbanmas"];
                comboresponsibleperson.DataSource = dt;
                comboresponsibleperson.DisplayMember = "bankname";
                comboresponsibleperson.ValueMember = "asptblbanmasid";
                comboparty.Focus();
            }
        }

        private void txtledgerDate_ValueChanged(object sender, EventArgs e)
        {
            int MON = Convert.ToInt32(txtadvDate.Value.ToString().Substring(3, 2));
            string getmonth = getFullName(MON);
            lblmonth.Text = getmonth;
            int year = Convert.ToInt32(txtadvDate.Value.ToString().Substring(6, 4));
            string getyear = getFullName1(year);
            lblyear.Text = getyear;
        }


        private void txtamt_TextChanged_1(object sender, EventArgs e)
        {
          
        }

        private void txtamt_ValueChanged(object sender, EventArgs e)
        {
           
        }
        public string Rupees(Int64 rup)
        {
            string result = "";
            Int64 res;

            if ((rup / 10000000) > 0)
            {
                res = rup / 10000000;
                rup = rup % 10000000;
                result = result + ' ' + RupeesToWords(res) + " Crore";
            }

            if ((rup / 100000) > 0)
            {
                res = rup / 100000;
                rup = rup % 100000;
                result = result + ' ' + RupeesToWords(res) + " Lack";
            }


            if ((rup / 1000) > 0)
            {
                res = rup / 1000;
                rup = rup % 1000;
                result = result + ' ' + RupeesToWords(res) + " Thousand";
            }

            if ((rup / 100) > 0)
            {
                res = rup / 100;
                rup = rup % 100;
                result = result + ' ' + RupeesToWords(res) + " Hundred";
            }
            if ((rup % 10) >= 0)
            {
                res = rup % 100;
                result = result + " " + RupeesToWords(res);
            }
            result = result + ' ' + " Rupees only";
            return result;
        }
        public string RupeesToWords(Int64 rup)
        {
            string result = "";
            if ((rup >= 1) && (rup <= 10))
            {
                if ((rup % 10) == 1) result = "One";
                if ((rup % 10) == 2) result = "Two";
                if ((rup % 10) == 3) result = "Three";
                if ((rup % 10) == 4) result = "Four";
                if ((rup % 10) == 5) result = "Five";
                if ((rup % 10) == 6) result = "Six";
                if ((rup % 10) == 7) result = "Seven";
                if ((rup % 10) == 8) result = "Eight";
                if ((rup % 10) == 9) result = "Nine";
                if ((rup % 10) == 0) result = "Ten";
            }
            if (rup > 9 && rup < 20)
            {
                if (rup == 11) result = "Eleven";
                if (rup == 12) result = "Twelve";
                if (rup == 13) result = "Thirteen";
                if (rup == 14) result = "Forteen";
                if (rup == 15) result = "Fifteen";
                if (rup == 16) result = "Sixteen";
                if (rup == 17) result = "Seventeen";
                if (rup == 18) result = "Eighteen";
                if (rup == 19) result = "Nineteen";
            }
            if (rup > 20 && (rup / 10) == 2 && (rup % 10) == 0) result = "Twenty";
            if (rup > 20 && (rup / 10) == 3 && (rup % 10) == 0) result = "Thirty";
            if (rup > 20 && (rup / 10) == 4 && (rup % 10) == 0) result = "Forty";
            if (rup > 20 && (rup / 10) == 5 && (rup % 10) == 0) result = "Fifty";
            if (rup > 20 && (rup / 10) == 6 && (rup % 10) == 0) result = "Sixty";

            if (rup > 20 && (rup / 10) == 7 && (rup % 10) == 0) result = "Seventy";

            if (rup > 20 && (rup / 10) == 8 && (rup % 10) == 0) result = "Eighty";
            if (rup > 20 && (rup / 10) == 9 && (rup % 10) == 0) result = "Ninty";

            if (rup > 20 && (rup / 10) == 2 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Twenty One";
                if ((rup % 10) == 2) result = "Twenty Two";
                if ((rup % 10) == 3) result = "Twenty Three";
                if ((rup % 10) == 4) result = "Twenty Four";
                if ((rup % 10) == 5) result = "Twenty Five";
                if ((rup % 10) == 6) result = "Twenty Six";
                if ((rup % 10) == 7) result = "Twenty Seven";
                if ((rup % 10) == 8) result = "Twenty Eight";
                if ((rup % 10) == 9) result = "Twenty Nine";
            }
            if (rup > 20 && (rup / 10) == 3 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Thirty One";
                if ((rup % 10) == 2) result = "Thirty Two";
                if ((rup % 10) == 3) result = "Thirty Three";
                if ((rup % 10) == 4) result = "Thirty Four";
                if ((rup % 10) == 5) result = "Thirty Five";
                if ((rup % 10) == 6) result = "Thirty Six";
                if ((rup % 10) == 7) result = "Thirty Seven";
                if ((rup % 10) == 8) result = "Thirty Eight";
                if ((rup % 10) == 9) result = "Thirty Nine";
            }
            if (rup > 20 && (rup / 10) == 4 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Forty One";
                if ((rup % 10) == 2) result = "Forty Two";
                if ((rup % 10) == 3) result = "Forty Three";
                if ((rup % 10) == 4) result = "Forty Four";
                if ((rup % 10) == 5) result = "Forty Five";
                if ((rup % 10) == 6) result = "Forty Six";
                if ((rup % 10) == 7) result = "Forty Seven";
                if ((rup % 10) == 8) result = "Forty Eight";
                if ((rup % 10) == 9) result = "Forty Nine";
            }
            if (rup > 20 && (rup / 10) == 5 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Fifty One";
                if ((rup % 10) == 2) result = "Fifty Two";
                if ((rup % 10) == 3) result = "Fifty Three";
                if ((rup % 10) == 4) result = "Fifty Four";
                if ((rup % 10) == 5) result = "Fifty Five";
                if ((rup % 10) == 6) result = "Fifty Six";
                if ((rup % 10) == 7) result = "Fifty Seven";
                if ((rup % 10) == 8) result = "Fifty Eight";
                if ((rup % 10) == 9) result = "Fifty Nine";
            }
            if (rup > 20 && (rup / 10) == 6 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Sixty One";
                if ((rup % 10) == 2) result = "Sixty Two";
                if ((rup % 10) == 3) result = "Sixty Three";
                if ((rup % 10) == 4) result = "Sixty Four";
                if ((rup % 10) == 5) result = "Sixty Five";
                if ((rup % 10) == 6) result = "Sixty Six";
                if ((rup % 10) == 7) result = "Sixty Seven";
                if ((rup % 10) == 8) result = "Sixty Eight";
                if ((rup % 10) == 9) result = "Sixty Nine";
            }
            if (rup > 20 && (rup / 10) == 7 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Seventy One";
                if ((rup % 10) == 2) result = "Seventy Two";
                if ((rup % 10) == 3) result = "Seventy Three";
                if ((rup % 10) == 4) result = "Seventy Four";
                if ((rup % 10) == 5) result = "Seventy Five";
                if ((rup % 10) == 6) result = "Seventy Six";
                if ((rup % 10) == 7) result = "Seventy Seven";
                if ((rup % 10) == 8) result = "Seventy Eight";
                if ((rup % 10) == 9) result = "Seventy Nine";
            }
            if (rup > 20 && (rup / 10) == 8 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Eighty One";
                if ((rup % 10) == 2) result = "Eighty Two";
                if ((rup % 10) == 3) result = "Eighty Three";
                if ((rup % 10) == 4) result = "Eighty Four";
                if ((rup % 10) == 5) result = "Eighty Five";
                if ((rup % 10) == 6) result = "Eighty Six";
                if ((rup % 10) == 7) result = "Eighty Seven";
                if ((rup % 10) == 8) result = "Eighty Eight";
                if ((rup % 10) == 9) result = "Eighty Nine";
            }
            if (rup > 20 && (rup / 10) == 9 && (rup % 10) != 0)
            {
                if ((rup % 10) == 1) result = "Ninty One";
                if ((rup % 10) == 2) result = "Ninty Two";
                if ((rup % 10) == 3) result = "Ninty Three";
                if ((rup % 10) == 4) result = "Ninty Four";
                if ((rup % 10) == 5) result = "Ninty Five";
                if ((rup % 10) == 6) result = "Ninty Six";
                if ((rup % 10) == 7) result = "Ninty Seven";
                if ((rup % 10) == 8) result = "Ninty Eight";
                if ((rup % 10) == 9) result = "Ninty Nine";
            }
            return result;
        }
        private void txtadvamount_TextChanged(object sender, EventArgs e)
        {
            Int64 len = txtadvamount.Value.ToString().Length;
            if (len >= 9)
            {
                MessageBox.Show("Invalid Value");
                txtadvamount.Value = 0;
            }
            else
            {
                //Int64 ss = Convert.ToInt64(txtadvamount.Value.ToString());
                //txtorderno.Text = NumbertoWords(ss);
            }
        }

        private void combodept_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                comboparty.Focus();
        }

        private void comboparty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtorderno.Focus();
        }

        private void txtorderno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtitemdesc.Focus();
        }

        private void txtitemdesc_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                comboresponsibleperson.Focus();
        }

        private void comboresponsibleperson_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                combomodeofpayment.Focus();
        }

        private void txtinvoiceamt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtadvterms.Focus();
        }

        private void txtadvterms_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                combomodeofpayment.Focus();
        }

        private void txtadvamount_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtgst.Focus();
        }

        private void txtorderno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == ' ' || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back);

        }

        private void txtitemdesc_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == ' ' || e.KeyChar == (char)Keys.Back);

        }

        private void txtadvterms_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == '%' || e.KeyChar == (char)Keys.Back);

        }
        
        void amountinwords()
        {
            try
            {
                decimal per = Convert.ToDecimal("0" + txtadvterms.Text.ToString());
                if (per <= 100)
                {


                    decimal amt1 = Convert.ToDecimal("0" + txtinvoiceamt.Value.ToString().Replace(",", "").Trim());
                    decimal amt2 = Convert.ToDecimal("0" + txttdsvalue.Value.ToString().Replace(",", "").Trim());
                    decimal tot = amt1 - amt2;
                    decimal per1 = (tot * per) / 100;
                    txtadvamount.Value = per1;



                    txtamtinwords.Text = Rupees(Convert.ToInt64(per1));
                }
                else
                {
                    txtadvamount.Text = "";
                    MessageBox.Show("Invalid Advance Terms");
                    txtadvamount.Text = "";
                }
            }
            catch (Exception ex) { }
        }
        private void txtadvterms_TextChanged(object sender, EventArgs e)
        {
            amountinwords();

        }

        private void txtinvoiceamt_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                decimal per = Convert.ToDecimal("0" + txtadvterms.Text.ToString());
                if (per <= 100)
                {
                    decimal amt1 = Convert.ToDecimal("0" + txtinvoiceamt.Value.ToString().Replace(",", "").Trim());
                    decimal amt2 = Convert.ToDecimal("0" + txttdsvalue.Value.ToString().Replace(",", "").Trim());
                    decimal tot = amt1 - amt2;

                    decimal per1 = (tot * per) / 100;
                    txtadvamount.Value = per1;

                }
                else
                {
                    txtadvamount.Text = "";
                    MessageBox.Show("Invalid Advance Terms");
                    txtadvamount.Text = "";
                }
            }
            catch(Exception ex) { }
        }
        byte[] stdbytes; byte[] stdbytes2; byte[] stdbytes3; byte[] stdbytes4; byte[] stdbytes5; Int64 std;
        OpenFileDialog open = new OpenFileDialog();
        

        private void butclose_Click(object sender, EventArgs e)
        {
          
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void txtadvamount_ValueChanged(object sender, EventArgs e)
        {

        }
        private void txtgst_TextChanged(object sender, EventArgs e)
        {
            try
            {
                decimal per = Convert.ToDecimal("0" + txtgst.Text.ToString());
                if (per <= 28)
                {
                    
                    decimal amt = Convert.ToDecimal("0" + txtamount.Value.ToString().Replace(",", "").Trim());

                    decimal per1 = (amt * per) / 100;
                    decimal tot = per1 + amt;
                    txtinvoiceamt.Value = tot;
                    
                    txtgstvalue.Text = per1.ToString();
                }
                else
                {
                    txtinvoiceamt.Text = "";
                    MessageBox.Show("Invalid Advance Terms");

                }
            }
            catch(Exception EX) { }
        }

        private void txtamount_ValueChanged(object sender, EventArgs e)
        {
            try {

                if (txtamount.Value.ToString().Length > 12)
                {
                    MessageBox.Show("Invalid Value");
                    txtamount.Value = 0; txtamount.Focus();
                    return;
                }
                else
                {

                    decimal per = Convert.ToDecimal("0" + txtgst.Text.ToString());
                    if (per <= 28)
                    {
                        decimal amt = Convert.ToDecimal("0" + txtamount.Value.ToString().Replace(",", "").Trim());

                        decimal per1 = (amt * per) / 100;
                        txtinvoiceamt.Value = per1;
                        if (txtamount.Text.Length > 0)
                        {
                            txtgst_TextChanged(sender, e);
                            txttds_TextChanged(sender, e);
                            txtadvterms_TextChanged(sender, e);
                        }
                    }
                    else
                    {
                        txtinvoiceamt.Text = "";
                        MessageBox.Show("Invalid Advance Terms");

                    }
                }
            }catch(Exception ex) { }
            }

        private void txttds_TextChanged(object sender, EventArgs e)
        {
            try
            {
                decimal per = Convert.ToDecimal("0" + txttds.Text.ToString());
                if (per <= 28)
                {
                    
                        

                        decimal amt = Convert.ToDecimal("0" + txtamount.Value.ToString().Replace(",", "").Trim());
                        decimal gstamt = Convert.ToDecimal("0" + txtinvoiceamt.Value.ToString().Replace(",", "").Trim());
                        decimal per1 = (amt * per) / 100;
                        txttdsvalue.Text = per1.ToString();
                        decimal tot = gstamt - per1;
                        decimal dedtot = gstamt - per1;
                        txtdedutionamt.Value = dedtot;
                   
                }
                else
                {
                    txttdsvalue.Text = "";
                    MessageBox.Show("Invalid Advance Terms");

                }
            }
            catch(Exception EX) { }
        }

        private void txtamount_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtgst.Focus();
        }

        private void txtgst_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txttds.Focus();
        }

        private void txttds_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtadvterms.Focus();
        }

        private void combomodeofpayment_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtamount.Focus();
        }
        //private void butproform_Click(object sender, EventArgs e)
        //{

        //    Class.Users.Picture = pictureBox2;
        //    Master.Bank.PopUp pop = new Master.Bank.PopUp();
        //    pop.Show();
        //}



        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo("xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES';"; //for above excel 2007  
            using (System.Data.OleDb.OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    System.Data.OleDb.OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }
        string pdfFilePath = ""; string[] ext;
        public static string ByteArrayToStr(byte[] barr)
        {
            UTF8Encoding encoding = new UTF8Encoding();
            return encoding.GetString(barr, 0, barr.Length);
        }
        public static DataTable Exceltodatatable(string path)
        {
            DataTable dt = new DataTable();
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                using (OleDbCommand comm = new OleDbCommand())
                {
                    string sheetName = "sheet1";
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }

            }
        }
        public void checkcellvalue(int index, DataGridView dgrid)
        {
            if (dgrid.CurrentCell.ColumnIndex.Equals(2) && index != -1)
            {
                if (dgrid.CurrentCell != null && dgrid.CurrentRow.Cells[1].FormattedValue.ToString() != "")
                {

                    try
                    {
                       
                        //open.Filter = "(*.jpg;*.jpeg;*.bmp;*.pdf;)| *.png;*.jpg; *.jpeg; *.bmp;*.pdf;";
                      
                        if (open.ShowDialog() == DialogResult.OK)
                        {
                            string extension = Path.GetExtension(open.FileName);
                            if (extension == ".xls" || extension == ".xlsx" || extension == ".pdf" || extension == ".jpg" || extension == ".JPG" || extension == ".jpeg" || extension == ".png")
                            {
                                if (index == 0)
                                {
                                    Class.Users.Extension = "";
                                    Class.Users.Extension = extension;
                                    dgrid.CurrentRow.Cells[2].Value = dgrid.CurrentRow.Cells[1].FormattedValue.ToString().Substring(0, 3) + extension;
                                    pdfFilePath = "";
                                    pdfFilePath = open.FileName;
                                    stdbytes = System.IO.File.ReadAllBytes(pdfFilePath);
                                }
                                if (index == 1)
                                {
                                    Class.Users.Extension = "";
                                    Class.Users.Extension = extension;
                                    dgrid.CurrentRow.Cells[2].Value = dgrid.CurrentRow.Cells[1].FormattedValue.ToString().Substring(0, 3) + extension;
                                    pdfFilePath = "";
                                    pdfFilePath = open.FileName;
                                    stdbytes2 = System.IO.File.ReadAllBytes(pdfFilePath);
                                }
                                if (index == 2)
                                {
                                    Class.Users.Extension = "";
                                    Class.Users.Extension = extension;
                                    dgrid.CurrentRow.Cells[2].Value = dgrid.CurrentRow.Cells[1].FormattedValue.ToString().Substring(0, 3) + extension;
                                    pdfFilePath = "";
                                    pdfFilePath = open.FileName;
                                    stdbytes3 = System.IO.File.ReadAllBytes(pdfFilePath);


                                }
                                if (index == 3)
                                {
                                    Class.Users.Extension = "";
                                    Class.Users.Extension = extension;
                                    dgrid.CurrentRow.Cells[2].Value = dgrid.CurrentRow.Cells[1].FormattedValue.ToString().Substring(0,3) + extension;
                                    pdfFilePath = "";
                                    pdfFilePath = open.FileName;
                                    stdbytes4 = System.IO.File.ReadAllBytes(pdfFilePath);
                                    Class.Users.StaticByts = stdbytes4;
                                }
                                if (index == 4)
                                {
                                    Class.Users.Extension = "";
                                    Class.Users.Extension = extension;
                                    dgrid.CurrentRow.Cells[2].Value = dgrid.CurrentRow.Cells[1].FormattedValue.ToString().Substring(0, 3) + extension;
                                    pdfFilePath = "";
                                    pdfFilePath = open.FileName;
                                    stdbytes5 = System.IO.File.ReadAllBytes(pdfFilePath);
                                    Class.Users.StaticByts = stdbytes5;


                                }
                            }
                            else
                            {
                                MessageBox.Show("Pls Upload (.pdf,.xls,.xlsx,jpeg) File Only");
                                this.Focus();
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            if (dgrid.CurrentCell.ColumnIndex.Equals(3) && index != -1)
            {
                if (dgrid.CurrentCell != null  && dgrid.CurrentRow.Cells[2].FormattedValue.ToString() != "")
                {
                    DataTable dt = new DataTable();
                    try
                    {
                        ext = dgrid.CurrentRow.Cells[2].FormattedValue.ToString().Split('.');
                        Class.Users.Extension = ext[1];                       
                        //filepath = "";
                        //filepath = "D:\\temp1." + Class.Users.Extension;
                        //FS = new FileStream(filepath, System.IO.FileMode.Create);
                        if (index == 0)
                        {
                            Class.Users.StaticByts = null;
                            Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            //FS.Write(stdbytes, 0, stdbytes.Length);
                            //FS.Close();
                            Class.Users.StaticByts = stdbytes;
                            Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();
                        }
                        if (index == 1)
                        {
                            Class.Users.StaticByts = null;
                            Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            //FS.Write(stdbytes2, 0, stdbytes2.Length);
                            //FS.Close();
                            Class.Users.StaticByts = stdbytes2;


                            Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();
                        }
                        if (index == 2)
                        {
                            Class.Users.StaticByts = null;
                            Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            //FS.Write(stdbytes3, 0, stdbytes3.Length);
                            //FS.Close();
                            Class.Users.StaticByts = stdbytes3;

                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();

                        }
                        if (index == 3)
                        {
                            Class.Users.StaticByts = null;
                            Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            //FS.Write(stdbytes4, 0, stdbytes4.Length);
                            //FS.Close();
                            Class.Users.StaticByts = stdbytes4;


                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();
                        }
                        if (index == 4)
                        {
                            Class.Users.StaticByts = null;
                            Class.Users.Paramid = Convert.ToInt64("0" + txtadvpayid.Text);
                            Class.Users.TableName = dgrid.CurrentRow.Cells[2].FormattedValue.ToString();
                            //FS.Write(stdbytes5, 0, stdbytes5.Length);
                            //FS.Close();
                            Class.Users.StaticByts = stdbytes5;

                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    Class.Users.StaticByts = null;
                }
            }
           
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            checkcellvalue(e.RowIndex,dataGridView1);
            

        }
        private void txttds_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txtgst_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtdedutionamt_ValueChanged(object sender, EventArgs e)
        {
            Int64 len = txttdsvalue.Value.ToString().Length;
            if (len >= 12)
            {
                MessageBox.Show("Invalid Value");
                txttdsvalue.Value = 0;
                return;
            }
            else
            {
            }
        }



        
        public void checkduplicate(int index, DataGridView dgrid)
        {
            // dgrid = new DataGridView();
            foreach (DataGridViewRow row in dgrid.Rows)
            {
                if (row.Index == dgrid.CurrentCell.RowIndex)
                {
                    continue;
                }
                if (dgrid.CurrentCell.Value == null)
                {
                    continue;
                }
                if (row.Cells[index].Value != null && row.Cells[index].Value.ToString() == dgrid.CurrentCell.Value.ToString())
                {
                    MessageBox.Show("Duplicate not allowed "); 
                    dgrid.CurrentCell.Value = null;
                }
            }
        }
        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                Class.Users.UserTime = 0;
                mas.checkduplicate(e.ColumnIndex, dataGridView1);
            }
        }

        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rowDeleteToolStripMenuItem.Text == "Row Delete")
                {
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        Class.Users.UserTime = 0;
                        dataGridView1.Rows.RemoveAt(row.Index);
                        dataGridView1.Refresh();
                    }
                }
            }
            catch(Exception ex) { }
        }

        private void txttdsvalue_ValueChanged(object sender, EventArgs e)
        {

        }

       private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            //   // using ceTe.DynamicPDF.Rasterizer;
            //   // open.Filter = "(*.jpg;*.jpeg;*.bmp;*.pdf;)| *.png;*.jpg; *.jpeg; *.bmp;*.pdf;";
            //    if (open.ShowDialog() == DialogResult.OK)
            //    {
            //        string strText = string.Empty; byte[] ba2;
            //        string pdfFilePath = open.FileName;
            //        byte[] bytes = Encoding.UTF8.GetBytes(open.FileName);
            //        File.WriteAllBytes(pdfFilePath, bytes);     
                    
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        private void uccListView2_MouseClick(object sender, MouseEventArgs e)
        {
            tabControl1.SelectTab(tabPage1);
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void combodept_SelectedIndexChanged(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Class.Users.UserTime = 0;
        }

        private void checkall_CheckedChanged(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }
    }
 
}

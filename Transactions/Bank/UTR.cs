using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions.Bank
{
    public partial class UTR : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        public UTR()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(System.DateTime.Now.ToShortTimeString().ToString());
            Class.Users.UserTime = 0;
        }

        private static UTR _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.Validate va = new Models.Validate();
       
        public static UTR Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new UTR();
                GlobalVariables.CurrentForm = _instance; return _instance;

            }
        }

        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        private void UTR_Load(object sender, EventArgs e)
        {

        }

        public void loadproduct()
        {
            sb.Clear();
            sb.Append("SELECT  a.asptblpartymasID,a.partyname from asptblpartymas a where a.active='T'");           
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblpartymas");
            DataTable dt = ds.Tables["asptblpartymas"];           
            comboparty.DisplayMember = "partyname";
            comboparty.ValueMember = "asptblpartymasID";
            comboparty.DataSource = dt;
            empty();
        }
      
        string[] s;

      
        private void approvalCancelToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }
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

        public void Saves()
        {
            try
            {
                int MON = Convert.ToInt32(txtutrdate.Value.ToString().Substring(3, 2));
                string getmonth = getFullName(MON);string conform= "";
                if(txtapprovalid.Text != "")
                {
                    conform = "T";
                }
                else
                {
                     conform = "F";
                }
                if (comboparty.SelectedValue == null)
                {
                    MessageBox.Show("'Party Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtutrno.Focus();
                    return;
                }
                if (txtamt.Text == "")
                {

                    MessageBox.Show("'Amount  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtutrno.Focus(); return;
                }
                if (combobank.Text == "")
                {
                    MessageBox.Show("'BankName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtutrno.Focus();
                    return;
                }
                if (combopaymenttype.Text == "")
                {
                    MessageBox.Show("'Payment Type  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtutrno.Focus();
                    return;
                }
                else
                {

                   
                  
                    string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select asptblutrid    from  asptblutr    WHERE asptblledmasid='" + txtpaymentid.Text + "' and partyname='" + comboparty.SelectedValue + "' and ledgermonth='" + lblmonth.Text + lblyear.Text + "' and bankname='"+combobank.SelectedValue+"' and amount='" + txtamt.Value.ToString().Replace(",", "") + "'  and active='" + chk + "' and utrno='" + txtutrno.Text.ToUpper() + "' and paymenttype='"+combopaymenttype.Text+ "' and billno='" + txtbillno.Text.ToUpper() + "' and remarks='" + txtremarks.Text.ToUpper() + "'";//and a.ledgerDate='" + txtledgerDate.Value.ToString("yyyy-MM-dd") + "'
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblutr");
                        DataTable dt = ds.Tables["asptblutr"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + comboparty.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtutrid.Text) == 0 || Convert.ToInt32("0" + txtutrid.Text) == 0)
                        {
                        string ins = "insert into asptblutr(asptblledmasid,partyname, ledgermonth,bankname,amount,active,utrdate,utrno,createdby,CREATEDON,MODIFIEDON,ipaddress,compcode,paymenttype,billno,remarks,asptbladvpaymasid,conformed,finyear,utrref)  VALUES('" + txtpaymentid.Text + "','" + comboparty.SelectedValue + "','" + lblmonth.Text + lblyear.Text + "','" + combobank.SelectedValue + "','" + txtamt.Value.ToString().Replace(",", "") + "','" + chk + "','" + txtutrdate.Value.ToString("yyyy-MM-dd") + "','"+txtutrno.Text.ToUpper()+"','" + Class.Users.HUserName + "','" + System.DateTime.Now.ToString() + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.IPADDRESS + "','" + Class.Users.COMPCODE + "','" + combopaymenttype.Text.ToUpper() + "','" + txtbillno.Text.ToUpper() + "','" + txtremarks.Text.ToUpper() + "','" + Convert.ToInt64("0"+txtapprovalid.Text) + "','" + conform + "','" + System.DateTime.Now.Year + "','T')";
                            Utility.ExecuteNonQuery(ins);
                        string up = "update asptblledmas set utrref='T' where asptblledmasid='" + txtpaymentid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Saved Successfully " + comboparty.SelectedValue, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                        string up = "update  asptblutr  set asptblledmasid='" + txtpaymentid.Text + "',partyname='" + comboparty.SelectedValue + "', ledgermonth='" + lblmonth.Text + lblyear.Text + "',bankname='" + combobank.SelectedValue + "',amount='" + txtamt.Value.ToString().Replace(",", "") + "' ,utrdate='" + txtutrdate.Value.ToString("yyyy-MM-dd") + "', active='" + chk + "' , MODIFIEDON='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "', utrno='" + txtutrno.Text.ToUpper() + "' , compcode='" + Class.Users.COMPCODE + "',paymenttype='" + combopaymenttype.Text + "',billno='" + txtbillno.Text.ToUpper() + "' , remarks='" + txtremarks.Text.ToUpper() + "',asptbladvpaymasid='" + Convert.ToInt64("0" + txtapprovalid.Text) + "',conformed='"+conform+ "', finyear='" + System.DateTime.Now.Year + "',utrref='T' where asptblutrid='" + txtutrid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                        string up1 = "update asptblledmas set utrref='T' where asptblledmasid='" + txtpaymentid.Text + "'";
                        Utility.ExecuteNonQuery(up1);
                        MessageBox.Show("Record Updated Successfully " + comboparty.SelectedValue, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                   
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("" + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        private void UTR_FormClosed(object sender, FormClosedEventArgs e)
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
           
            GridLoad(); loadproduct();
           
            tabControl1.SelectTab(tabPage1); empty();

        }
        private void empty()
        {
            txtutrid.Text = "";txtpaymentid.Text = "";
            txtaccholdername.Text = "";txtapprovalid.Text = "";
            comboparty.SelectedIndex = -1;  
            txtamt.Text = "";            txtamt.Text = ""; txtutrdate.Text = ""; Class.Users.UserTime = 0;
            combopaymenttype.SelectedIndex = -1;txtutrno.Text = "";
            combobank.SelectedIndex = -1;txtbillno.Text = "";txtremarks.Text = "";
            textBox1.Text = "";
            this.BackColor = Class.Users.BackColors; 
            this.Font = Class.Users.FontName; Class.Users.UserTime = 0;
            butheader.BackColor = Class.Users.BackColors; checkall.Checked = false;
            checkall.BackColor = Class.Users.BackColors;
            txtutrno.Focus();

        }
        DataTable dtapproval = new DataTable();
        public void GridLoad()
        {
            try
            {
                if (checkall.Checked == true) {
                    Class.Users.SearchQuery = " select a.asptblledmasid as ID,b.partyname as PartyName,b.accountholdername as AccName , a.Amount,c.bankname as BankName from  asptblledmas a  join asptblpartymas b on a.partyname=b.asptblpartymasid   join asptblbanmas c on c.asptblbanmasid=a.bankname   where a.conformed='T' and a.utrref='F' order by  a.asptblledmasid desc";
                    Class.Users.HideCols = new string[] { "ID" };
                    uccListView1.Load_Details();
                }
                else
                {
                    Class.Users.SearchQuery = " select a.asptblledmasid as ID,b.partyname as PartyName, b.accountholdername as AccName,a.Amount,c.bankname as BankName  from  asptblledmas a  join asptblpartymas b on a.partyname=b.asptblpartymasid   join asptblbanmas c on c.asptblbanmasid=a.bankname  where a.conformed='T'  and a.utrref='F' and a.ledgerDate='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "'  order by  a.asptblledmasid desc";
                    Class.Users.HideCols = new string[] { "ID" };
                    uccListView1.Load_Details();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
       
        public void Searchs(int EditID)
        {
            try
            {
                if (EditID > 0)
                {
                    Class.Users.UserTime = 0;
                    txtpaymentid.Text = EditID.ToString();
                    DataTable dt = Utility.SQLQuery("select  a.asptblledmasid AS ID, b.partyname,a.amount,a.active,a.paymenttype,c.bankname,a.conformed,a.asptbladvpaymasid   from  asptblledmas a join  asptblpartymas b on a.partyname=b.asptblpartymasid join asptblbanmas c on c.asptblbanmasid=a.bankname   where a.asptblledmasid='" + txtpaymentid.Text + "' order by a.asptblledmasid desc");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["conformed"].ToString() == "T")
                        {
                            txtamt.Enabled = false;
                            comboparty.Enabled = false;
                            combopaymenttype.Enabled = false;
                            txtapprovalid.Text = Convert.ToString(dt.Rows[0]["asptbladvpaymasid"].ToString());
                            txtpaymentid.Text = Convert.ToString(dt.Rows[0]["ID"].ToString());
                            comboparty.Text = Convert.ToString(dt.Rows[0]["partyname"].ToString());
                            txtamt.Text = dt.Rows[0]["amount"].ToString();                                          
                            combopaymenttype.Text = dt.Rows[0]["paymenttype"].ToString();
                            combobank.Text = dt.Rows[0]["bankname"].ToString();                           
                            if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                            textBox1.Text = Rupees(Convert.ToInt64(txtamt.Value));
                            txtutrno.Focus();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void butdatewise_Click(object sender, EventArgs e)
        {
               
        }
        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            loadproduct();
        }

        private void Txtamount_TextChanged(object sender, EventArgs e)
        {
          
            if (va.IsIntegerdot(txtamt.Text) == false)
            {
               // txtamount.Text = "";
            }
            if (va.IsIntegerdot(txtamt.Text) == true)
            {
                txtamt.Text = "";
                
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
            if (txtutrid.Text != "")
            {
                if (txtapprovalid.Text != "")
                {
                    string update = "update asptbladvpaymas set approval='T',mdapproval='T',conformed='F' where asptbladvpaymasid='" + txtapprovalid.Text + "'";
                    Utility.ExecuteNonQuery(update); 
                }
                string del = "delete from asptblledmas where  asptblledmasid='" + Convert.ToInt64("0" + txtutrid.Text) + "' and finyear='"+System.DateTime.Now.Year+"' ";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + comboparty.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
               
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
            this.comboparty.BackColor = Color.Yellow;
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
            if (combobank.SelectedValue != null && combobank.SelectedValue.ToString() != "System.Data.DataRowView" && combobank.Text !="")
            {
               
                string sel1 = "select distinct b.bankname from gtcompmast a  join asptblbanmas b on b.asptblbanmasid=a.bankname  where a.active='T'  and a.compcode='" + Class.Users.HCompcode + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblledmas");
                DataTable dt = ds.Tables["asptblledmas"];
                if (combobank.Text == dt.Rows[0]["bankname"].ToString())
                {
                   
                    combopaymenttype.SelectedIndex = 1;
                }
                else
                {
                   
                    combopaymenttype.SelectedIndex = 2;
                }
            }
        }

        private void txtamt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == ',' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back);
        }

        private void comboparty_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboparty.Text != "System.Data.DataRowView" && comboparty.Text != "")
            {
                sb.Clear();
                sb.Append("SELECT   a.asptblbanmasid,a.bankname,b.accountholdername from asptblbanmas a join asptblpartymas b on a.asptblbanmasid=b.bankname  where b.active='T' and b.asptblpartymasid='"+comboparty.SelectedValue+"'");
                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblbanmas");
                DataTable dt = ds.Tables["asptblbanmas"];
                combobank.DataSource = dt;
                combobank.DisplayMember = "bankname";
                combobank.ValueMember = "asptblbanmasid";
                txtaccholdername.Text = dt.Rows[0]["accountholdername"].ToString();
                txtamt.Focus();
            }
        }

        private void txtledgerDate_ValueChanged(object sender, EventArgs e)
        {
            int MON = Convert.ToInt32(txtutrdate.Value.ToString().Substring(3, 2));
            string getmonth = getFullName(MON);
            lblmonth.Text = getmonth;
            int year = Convert.ToInt32(txtutrdate.Value.ToString().Substring(6, 4));
            string getyear = getFullName1(year);
            lblyear.Text = getyear;
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



        private void txtamt_TextChanged_1(object sender, EventArgs e)
        {
          
        }

        private void txtamt_ValueChanged(object sender, EventArgs e)
        {
            
            if (txtamt.Value.ToString().Length > 12) {
               
                 txtamt.Focus();
             
            }
            else
            {

               
             
                textBox1.Text = Rupees(Convert.ToInt64(txtamt.Value));
            }
        }

        private void comboparty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtamt.Focus();
        }

        private void txtamt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtutrno.Focus();
        }

        private void combobank_KeyDown(object sender, KeyEventArgs e)
        {
          
        }

        private void combopaymenttype_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtutrno.Focus();
        }

        private void txtremarks_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void txtremarks2_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void uccListView1_Click(object sender, EventArgs e)
        {
          
        }

        private void checkall_CheckedChanged(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

       
      
     
    


     

       

        private void txtbillno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtremarks.Focus();
        }

        private void txtutrno_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtbillno.Focus();
        }

        private void txtutrno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back);

        }

        private void txtbillno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back);

        }
    }
}

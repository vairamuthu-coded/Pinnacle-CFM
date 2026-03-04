using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Transactions.SKL
{
    public partial class HRPayDetails : Form,ToolStripAccess
    {
        private static HRPayDetails _instance;
        public static HRPayDetails Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new HRPayDetails();
                GlobalVariables.CurrentForm = _instance; return _instance;

            }
        }
        public HRPayDetails()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();            
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToShortTimeString().ToString());
            combofinyear.Text = Convert.ToString(System.DateTime.Now.Year.ToString());
            Class.Users.Finyear = System.DateTime.Now.Year.ToString();
        }
        private string readserialvalue;
        private Decimal wt = 0;
        bool valid = false; bool validprint = false;
        Models.Validate va = new Models.Validate();
        ListView listfilter = new ListView();
        private object worksheet;
        private void empty()
        {
            txthrpaydetailsid.Text = ""; combocompcode.Text = ""; combocompname.Text = ""; combocompcode.SelectedIndex = -1;
            combocompcode.Enabled = true; combofinyear.Text = Class.Users.Finyear; txtdate.Value = System.DateTime.Now;
            txtidcard.Text = "";
            txtempname.Text = "";
            txtuanno.Text = "0";
            txtesino.Text = "0";
            txtfathername.Text = "0";
            txtunited.Text = "";
            txtdepartment.Text = "";
            txtdesignation.Text = "";
            txtorjpayabledays.Text = "0";
            txtnhdays.Text = "0";
            txtpayabledays.Text = "0";
            txtgovtdaysal.Text = "0";
            txtotwages.Text = "0";
            txtbasicda.Text = "0";
            txtbasic.Text = "0";
            txtda.Text = "0";
            txthra.Text = "0";
            txtothers.Text = "0";
            txtebasic.Text = "0";
            txteda.Text = "0";
            txtebasicda.Text = "0";
            txtehra.Text = "0";
            txteothers.Text = "0";
            txtpayableothours.Text = "0";
            txtotamount.Text = "0";
            txtincentive.Text = "0";
            txtgovtgross.Text = "0";
            txtpfamount.Text = "0";
            txtesiamount.Text = "0";
            txtmessamount.Text = "0";
            txtdeduction.Text = "0";
            txtnetamount.Text = "0";
            txtbankaccountno.Text = "0";
            txtbankname.Text = "0";
            txtifsccode.Text = "0";
            txtpayperiod.Text = "";
            txtcategory.Text = "0";
            txtotherexp.Text = "0";
            txtadvance.Text = "0";
            //comboperiodsearh.SelectedIndex = -1; comboperiodsearh.Text = "";
            txtcreditdate.Text = "0"; dataGridView1.AllowUserToAddRows = true;
            //  combounitsearch.SelectedIndex = -1;combounitsearch.Text = ""; comboperiodsearh.SelectedIndex = -1; comboperiodsearh.Text = "";
            do
            {
                int i = 0;
                for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }

            }
            while (dataGridView1.Rows.Count > 0);
            dataGridView1.AllowUserToAddRows = false; checkall.Checked = false;
            listView1.Items.Clear();
            this.Font = Class.Users.FontName;
           
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
           // panel1.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
          butheader.BackColor= Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
        }
        public void companyload()
        {
            try
            {
                string sel = "select a.gtcompmastid,a.compcode from  gtcompmast a  where a.ptransaction ='COMPANY' and a.active='T'  and a.compcode='" + Class.Users.HCompcode + "' order by 2 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;

                try
                {

                    string sel1 = "select '-1' as  gtcompmastid,'-------' as compcode from dual union all select distinct a.gtcompmastid,a.compcode from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ptransaction ='COMPANY' and a.active='T' and a.compcode='" + Class.Users.HCompcode + "' order by 2";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];

                    combounitsearch.DisplayMember = "compcode";
                    combounitsearch.ValueMember = "gtcompmastid";
                    combounitsearch.DataSource = dt1;

                    combounitreport.DisplayMember = "compcode";
                    combounitreport.ValueMember = "gtcompmastid";
                    combounitreport.DataSource = dt1;
                    //combounitsearch.Text = ""; combounitsearch.SelectedIndex = -1;
                    //combounitreport.Text = ""; combounitreport.SelectedIndex = -1;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("companyload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("companyload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (txthrpaydetailsid.Text == "" && Class.Users.HCompcode != "")
                {
                    //string sel = "select a.gtcompmastid,a.compcode, a.compname from  gtcompmast a  where a.ptransaction ='COMPANY' and a.compcode='" + combocompcode.Text + "' ;";
                    //DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                    //DataTable dt = ds.Tables["gtcompmast"];
                    //combocompname.Text = dt.Rows[0]["compname"].ToString();
                    autonumberload();
                }
                //if (txthrpaydetailsid.Text != "")
                //{

                //    txtdocid.Text = ""; txthrpaydetailsid1.Text = "";
                //    string sel = "select max(HRPayDetailsid)+1 as id,b.compname from HRPayDetails a join gtcompmast b on a.compname=b.gtcompmastid where b.ptransaction='COMPANY'  and b.compcode='" + combocompcode.Text + "'; ";
                //    DataSet ds = Utility.ExecuteSelectQuery(sel, "HRPayDetails");
                //    DataTable dt = ds.Tables["HRPayDetails"];
                //    combocompname.Text = dt.Rows[0]["compname"].ToString();
                //    txtdocid.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                //    txthrpaydetailsid1.Text = dt.Rows[0]["id"].ToString();

                //}
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

                string sel = "select max(a.HRPayDetailsid1)+1 as id,b.compname from HRPayDetails a join gtcompmast b on a.compname=b.gtcompmastid  where b.ptransaction='COMPANY' and a.active='T' and a.finyear='" + Class.Users.Finyear + "' and b.compcode='" + Class.Users.HCompcode + "' group by b.compname; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "HRPayDetails");
                DataTable dt = ds.Tables["HRPayDetails"];
                int cnt = dt.Rows.Count;
                if (cnt == 0)
                {

                    string sel1 = "select a.gtcompmastid,a.compcode, a.compname from  gtcompmast a where a.compcode='" + Class.Users.HCompcode + "'  order by 2 ;";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                    DataTable dt1 = ds1.Tables["gtcompmast"];
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt1;
                    combocompname.Text = dt1.Rows[0]["compname"].ToString();
                    txtdocid.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + 1;
                    txthrpaydetailsid1.Text = "1";
                }
                else
                {
                    combocompname.Text = dt.Rows[0]["compname"].ToString();
                    txtdocid.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                    txthrpaydetailsid1.Text = dt.Rows[0]["id"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public void News()
        {
            tabControl1.SelectTab(tabPageraw2);
            GridLoad(); companyload(); autonumberload();
            empty();


            txtsearch.Focus();
        }

        public void Saves()
        {
            try
            {


                if (dataGridView1.Rows.Count > 0 && Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                {
                    Cursor = Cursors.WaitCursor; lblprogress2.Visible = true; lblprogress2.Refresh(); label48.Refresh();
                    Models.HrPayModel c = new Models.HrPayModel();
                    int cc = 0;
                    cc = dataGridView1.Rows.Count;
                    c.docdate = System.DateTime.Now.ToString("yyyy-MM-dd");
                    c.finyear = Convert.ToString(Class.Users.Finyear);
                    c.compcode = Convert.ToInt64("0" + Class.Users.COMPCODE);
                    c.compname = Convert.ToInt64("0" + Class.Users.COMPCODE);
                    c.compcode1 = Class.Users.COMPCODE;
                    c.username = Class.Users.USERID;
                    c.createdby = Convert.ToString(Class.Users.CREATED);
                    c.createdon = Convert.ToString(Class.Users.CREATED);
                    c.modifiedby = Convert.ToString(Class.Users.HUserName);
                    c.ipaddress = Class.Users.IPADDRESS;
                    if (checkactive.Checked == true)
                        c.active = "T";
                    else c.active = "F";
                    if (cc >= 0)
                    {
                        progressBar3.Minimum = 0; progressBar3.Refresh();
                        progressBar3.Maximum = dataGridView1.Rows.Count;

                        for (int i = 0; i < cc; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value.ToString() != "")
                            {
                                combocompcode.Text = Class.Users.HCompcode;
                                combofinyear.Text = Class.Users.Finyear;
                                c.compcode = Class.Users.COMPCODE;
                                c.compname = Class.Users.COMPCODE;
                                c.finyear = Class.Users.Finyear;
                                if (txthrpaydetailsid.Text == "") { c.hrpaydetailsid = Convert.ToInt64("0" + txthrpaydetailsid.Text); autonumberload(); c.hrpaydetailsid1 = Convert.ToInt64("0" + txthrpaydetailsid1.Text); }
                                else { c.hrpaydetailsid = Convert.ToInt64("0" + txthrpaydetailsid.Text); c.hrpaydetailsid1 = Convert.ToInt64("0" + txthrpaydetailsid1.Text); }
                                c.docid = Convert.ToString(txtdocid.Text);
                                c.idcardno = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                                c.midcard = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                                c.empname = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
                                c.doj = Convert.ToDateTime(dataGridView1.Rows[i].Cells[3].Value.ToString().Substring(0, 10));

                                if (dataGridView1.Rows[i].Cells[4].Value.ToString() != "")
                                {
                                    c.dol = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value.ToString().Substring(0, 10));

                                }
                                else
                                {
                                    c.dol = "";
                                }
                                c.uanno = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                                c.esino = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);
                                c.fathername = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value);
                                c.united = Convert.ToString(Class.Users.HCompcode);
                                c.category = Convert.ToString(dataGridView1.Rows[i].Cells[9].Value);
                                c.department = Convert.ToString(dataGridView1.Rows[i].Cells[10].Value);
                                c.designation = Convert.ToString(dataGridView1.Rows[i].Cells[11].Value);
                                c.orjpayabledays = Convert.ToString(dataGridView1.Rows[i].Cells[12].Value);
                                c.nhdays = Convert.ToString(dataGridView1.Rows[i].Cells[13].Value);
                                c.payabledays = Convert.ToString(dataGridView1.Rows[i].Cells[14].Value);
                                c.govtdaysalary = Convert.ToString(dataGridView1.Rows[i].Cells[15].Value);
                                c.otwages = Convert.ToString(dataGridView1.Rows[i].Cells[16].Value);
                                c.basicda = Convert.ToString(dataGridView1.Rows[i].Cells[17].Value);
                                c.basic = Convert.ToString(dataGridView1.Rows[i].Cells[18].Value);
                                c.da = Convert.ToString(dataGridView1.Rows[i].Cells[19].Value);
                                c.hra = Convert.ToString(dataGridView1.Rows[i].Cells[20].Value);
                                c.others = Convert.ToString(dataGridView1.Rows[i].Cells[21].Value);
                                c.ebasic = Convert.ToString(dataGridView1.Rows[i].Cells[22].Value);
                                c.ebasicda = Convert.ToString(dataGridView1.Rows[i].Cells[23].Value);
                                c.eda = Convert.ToString(dataGridView1.Rows[i].Cells[24].Value);
                                c.ehra = Convert.ToString(dataGridView1.Rows[i].Cells[25].Value);
                                c.eothers = Convert.ToString(dataGridView1.Rows[i].Cells[26].Value);
                                c.payableothours = Convert.ToString(dataGridView1.Rows[i].Cells[27].Value);
                                c.otamount = Convert.ToString(dataGridView1.Rows[i].Cells[28].Value);
                                c.incentive = Convert.ToString(dataGridView1.Rows[i].Cells[29].Value);
                                c.govtgross = Convert.ToString(dataGridView1.Rows[i].Cells[30].Value);
                                c.pfamount = Convert.ToString(dataGridView1.Rows[i].Cells[31].Value);
                                c.esiamount = Convert.ToString(dataGridView1.Rows[i].Cells[32].Value);
                                c.messamount = Convert.ToString(dataGridView1.Rows[i].Cells[33].Value);
                                c.othersexp = Convert.ToDecimal("0" + dataGridView1.Rows[i].Cells[34].Value).ToString();
                                c.advance = Convert.ToDecimal("0" + dataGridView1.Rows[i].Cells[35].Value).ToString();
                                c.deduction = Convert.ToString(dataGridView1.Rows[i].Cells[36].Value);
                                c.netamount = Convert.ToString(dataGridView1.Rows[i].Cells[37].Value);
                                c.bankaccountno = Convert.ToString(dataGridView1.Rows[i].Cells[38].Value);
                                c.bankname = Convert.ToString(dataGridView1.Rows[i].Cells[39].Value);
                                c.ifsccode = Convert.ToString(dataGridView1.Rows[i].Cells[40].Value);
                                c.payperiod = Convert.ToString(dataGridView1.Rows[i].Cells[41].Value);
                                c.fromdate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[42].Value).ToString("dd-MMM-yyyy");
                                c.todate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[43].Value).ToString("dd-MMM-yyyy");
                                if (dataGridView1.Rows[i].Cells[44].Value.ToString() != "")
                                {
                                    c.creditdate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[44].Value).ToString("dd-MM-yyyy");                         //
                                }                                                                                                      //
                                                                                                                                       //  c.creditdate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[34].Value).ToString("dd-MMM-yyyy");                          and  doj = '" + c.doj + "'  and uanno = '" + c.uanno + "'  and esino = '" + c.esino + "'  and fathername = '" + c.fathername + "' and united = '" + c.united + "'  and  category='" + c.category + "' and   department='" + c.department + "'  and   designation='" + c.designation + "'  and    orjpayabledays='" + c.orjpayabledays + "'  and  nhdays='" + c.nhdays + "'  and  payabledays='" + c.payabledays + "'  and  govtdaysalary='" + c.govtdaysalary + "' and  otwages='" + c.otwages + "' and  basicda='" + c.basicda + "' and   basic='" + c.basic + "' and  da='" + c.da + "' and  hra='" + c.hra + "' and  others='" + c.others + "' and  ebasic='" + c.ebasic + "' and   eda='" + c.eda + "' and    ebasicda='" + c.ebasicda + "' and  ehra='" + c.ehra + "' and  eothers='" + c.eothers + "' and  payableothrs='" + c.payableothrs + "' and  otamount='" + c.otamount + "' and  incentive='" + c.incentive + "' and   govtgross='" + c.govtgross + "' and    pfamount='" + c.pfamount + "' and  esiamount='" + c.esiamount + "' and  messamount='" + c.messamount + "' and  deduction='" + c.deduction + "' and   netamount='" + c.netamount + "' and    bankaccountno='" + c.bankaccountno + "'  and  bankname='" + c.bankname + "'  and  ifsccode='" + c.ifsccode + "' and  payperiod='" + c.payperiod + "'  and  fromdate='" + c.fromdate + "' and  todate='" + c.todate + "' and active='" + c.active + "'

                                string sel = "select  hrpaydetailsid    from  HRPayDetails   WHERE   compcode='" + c.compcode + "' and midcard='" + c.midcard + "' and  idcardno='" + c.idcardno + "'  and  empname='" + c.empname + "' and  doj='" + Convert.ToDateTime(c.doj).ToString("yyyy-MM-dd").Substring(0, 10) + "'  and  united='" + combocompcode.Text + "' and payperiod='" + c.payperiod + "' ;";
                                DataSet ds = Utility.ExecuteSelectQuery(sel, "HRPayDetails");
                                DataTable dt = ds.Tables["HRPayDetails"];
                                if (dt.Rows.Count != 0)
                                {

                                }
                                if (dt.Rows.Count != 0)
                                {

                                }
                                else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txthrpaydetailsid.Text) == 0 || Convert.ToInt32("0" + txthrpaydetailsid.Text) == 0)
                                {
                                    string ins = "";
                                    if (c.dol == "")
                                    {
                                        ins = "insert into HRPayDetails(hrpaydetailsid1,  docid,  docdate ,  finyear,  compcode,  compname,midcard,  idcardno ,  empname,  doj,uanno,  esino,  fathername ,  united,  category,  department, designation, orjpayabledays ,  nhdays ,  payabledays ,  govtdaysalary,  otwages,  basicda ,  basic,  da ,  hra,  others, ebasic, ebasicda,  eda,  ehra,  eothers,  payableothrs ,  otamount,  incentive,  govtgross,  pfamount,  esiamount,  messamount,  othersexp,  advance, deduction,  netamount ,  bankaccountno,  bankname , ifsccode,payperiod, fromdate ,  todate,  active,creditdate,  compcode1,  username,  createdby,  createdon,  modifiedby,  ipaddress)values('" + c.hrpaydetailsid1 + "', '" + c.docid + "',  '" + c.docdate + "' ,  '" + c.finyear + "',  '" + c.compcode + "',  '" + c.compname + "', '" + c.midcard + "' ,  '" + c.idcardno + "' ,  '" + c.empname + "',  '" + Convert.ToDateTime(c.doj).ToString("yyyy-MM-dd").Substring(0, 10) + "','" + c.uanno + "',  '" + c.esino + "',  '" + c.fathername + "' ,  '" + c.united + "',  '" + c.category + "',  '" + c.department + "', '" + c.designation + "', '" + c.orjpayabledays + "' ,  '" + c.nhdays + "' ,  '" + c.payabledays + "' ,  '" + c.govtdaysalary + "',  '" + c.otwages + "',  '" + c.basicda + "' ,  '" + c.basic + "',  '" + c.da + "' ,  '" + c.hra + "',  '" + c.others + "', '" + c.ebasic + "', '" + c.ebasicda + "',  '" + c.eda + "',  '" + c.ehra + "', '" + c.eothers + "',  '" + c.payableothours + "' ,  '" + c.otamount + "',  '" + c.incentive + "',  '" + c.govtgross + "',  '" + c.pfamount + "',  '" + c.esiamount + "',  '" + c.messamount + "',  '" + c.othersexp + "',  '" + c.advance + "', '" + c.deduction + "',  '" + c.netamount + "' ,  '" + c.bankaccountno + "',  '" + c.bankname + "' ,  '" + c.ifsccode + "',  '" + c.payperiod + "', '" + c.fromdate + "',  '" + c.todate + "',  '" + c.active + "','" + c.creditdate + "',  '" + c.compcode1 + "',  '" + c.username + "' , '" + c.createdby + "',  '" + c.createdon + "',  '" + c.modifiedby + "',  '" + c.ipaddress + "' );";
                                        Utility.ExecuteNonQuery(ins);
                                    }
                                    else
                                    {
                                        ins = "insert into HRPayDetails(hrpaydetailsid1,  docid,  docdate ,  finyear,  compcode,  compname,midcard,  idcardno ,  empname,  doj,dol,  uanno,  esino,  fathername ,  united,  category,  department, designation, orjpayabledays ,  nhdays ,  payabledays ,  govtdaysalary,  otwages,  basicda ,  basic,  da ,  hra,  others, ebasic, ebasicda,  eda,  ehra,  eothers,  payableothrs ,  otamount,  incentive,  govtgross,  pfamount,  esiamount,  messamount,  othersexp,  advance, deduction,  netamount ,  bankaccountno,  bankname , ifsccode,payperiod, fromdate ,  todate,  active,creditdate,  compcode1,  username,  createdby,  createdon,  modifiedby,  ipaddress)values('" + c.hrpaydetailsid1 + "', '" + c.docid + "',  '" + c.docdate + "' ,  '" + c.finyear + "',  '" + c.compcode + "',  '" + c.compname + "', '" + c.midcard + "' ,  '" + c.idcardno + "' ,  '" + c.empname + "',  '" + Convert.ToDateTime(c.doj).ToString("yyyy-MM-dd").Substring(0, 10) + "','" + Convert.ToDateTime(c.dol).ToString("yyyy-MM-dd").Substring(0, 10) + "',  '" + c.uanno + "',  '" + c.esino + "',  '" + c.fathername + "' ,  '" + c.united + "',  '" + c.category + "',  '" + c.department + "', '" + c.designation + "', '" + c.orjpayabledays + "' ,  '" + c.nhdays + "' ,  '" + c.payabledays + "' ,  '" + c.govtdaysalary + "',  '" + c.otwages + "',  '" + c.basicda + "' ,  '" + c.basic + "',  '" + c.da + "' ,  '" + c.hra + "',  '" + c.others + "', '" + c.ebasic + "', '" + c.ebasicda + "',  '" + c.eda + "',  '" + c.ehra + "', '" + c.eothers + "',  '" + c.payableothours + "' ,  '" + c.otamount + "',  '" + c.incentive + "',  '" + c.govtgross + "',  '" + c.pfamount + "',  '" + c.esiamount + "',  '" + c.messamount + "',  '" + c.othersexp + "',  '" + c.advance + "', '" + c.deduction + "',  '" + c.netamount + "' ,  '" + c.bankaccountno + "',  '" + c.bankname + "' ,  '" + c.ifsccode + "',  '" + c.payperiod + "', '" + c.fromdate + "',  '" + c.todate + "',  '" + c.active + "','" + c.creditdate + "',  '" + c.compcode1 + "',  '" + c.username + "' , '" + c.createdby + "',  '" + c.createdon + "',  '" + c.modifiedby + "',  '" + c.ipaddress + "' );";
                                        Utility.ExecuteNonQuery(ins);
                                    }
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(dataGridView1.Rows.Count)) * (i + 1);
                                    lblprogress3.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %";
                                    label48.Text = " Total Rows : " + i.ToString() + "  " + (per).ToString("N0") + " %" + "Emp Id" + c.midcard + " Name : " + c.empname;
                                    lblprogress3.Refresh(); label48.Refresh();

                                    progressBar3.Value = i + 1;
                                }
                                else
                                {
                                    string up = "update  HRPayDetails  set hrpaydetailsid1='" + c.hrpaydetailsid1 + "' ,docid='" + c.docid + "' , docdate='" + c.docdate + "',finyear='" + c.finyear + "',compcode='" + c.compcode + "', compname='" + c.compname + "', midcard='" + c.midcard + "',idcardno='" + c.idcardno + "' , empname='" + c.empname + "', doj='" + c.doj + "',dol='" + c.dol + "', uanno='" + c.uanno + "' , esino='" + c.esino + "' , fathername='" + c.fathername + "', united='" + c.united + "' , category='" + c.category + "',  department='" + c.department + "', designation='" + c.designation + "' , orjpayabledays='" + c.orjpayabledays + "' , nhdays='" + c.nhdays + "' , payabledays='" + c.payabledays + "' , govtdaysalary='" + c.govtdaysalary + "', otwages='" + c.otwages + "', basicda='" + c.basicda + "',  basic='" + c.basic + "', da='" + c.da + "', hra='" + c.hra + "', others='" + c.others + "',ebasicda='" + c.ebasicda + "', ebasic='" + c.ebasic + "',  eda='" + c.eda + "', ehra='" + c.ehra + "', eothers='" + c.eothers + "', payableothrs='" + c.payableothours + "', otamount='" + c.otamount + "', incentive='" + c.incentive + "',  govtgross='" + c.govtgross + "',   pfamount='" + c.pfamount + "', esiamount='" + c.esiamount + "', messamount='" + c.messamount + "', othersexp='" + c.othersexp + "', advance='" + c.advance + "', deduction='" + c.deduction + "',  netamount='" + c.netamount + "',   bankaccountno='" + c.bankaccountno + "' , bankname='" + c.bankname + "' , ifsccode='" + c.ifsccode + "', payperiod='" + c.payperiod + "' , fromdate='" + c.fromdate + "', todate='" + c.todate + "',active='" + c.active + "',creditdate='" + c.creditdate + "',compcode1='" + Class.Users.COMPCODE + "', username='" + Class.Users.USERID + "',createdby='" + Class.Users.CREATED + "', modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where hrpaydetailsid='" + txthrpaydetailsid.Text + "';";
                                    Utility.ExecuteNonQuery(up);


                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(dataGridView1.Rows.Count)) * (i + 1);
                                    lblprogress3.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %";
                                    label48.Text = " Total Rows : " + i.ToString() + "  " + (per).ToString("N0") + " %" + "Emp Id" + c.midcard + " Name : " + c.empname;
                                    lblprogress3.Refresh(); label48.Refresh();

                                    progressBar3.Value = i + 1;
                                }

                            }
                        }

                        if (txthrpaydetailsid.Text == "")
                        {
                            Cursor = Cursors.Default;
                            MessageBox.Show("Record Saved Successfully.Toal Record are:," + cc.ToString(), " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            tabControl1.SelectTab(tabPageraw2);
                            GridLoad(); companyload(); progressBar3.Value = 0; lblprogress3.Text = "";
                            empty();
                        }
                        else
                        {
                            Cursor = Cursors.Default;
                            MessageBox.Show("Record Updated Successfully.Toal Record are:," + cc.ToString(), " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            tabControl1.SelectTab(tabPageraw2);
                            GridLoad(); companyload(); progressBar3.Value = 0; lblprogress3.Text = "";
                            empty();
                        }
                        Cursor = Cursors.Default;
                    }
                }
                else
                {
                    if (dataGridView1.Rows.Count == 0)
                    {

                        if (combocompcode.Text == "")
                        {
                            MessageBox.Show("CompCode is Empty." + combocompcode.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            combocompcode.Select();
                            return;

                        }

                        if (txtempname.Text == "")
                        {
                            MessageBox.Show("Pls Enter EmpName", "Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        if (txtidcard.Text == "")
                        {
                            MessageBox.Show("Pls Enter IDCardNo", "Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        if (txtunited.Text == "")
                        {
                            MessageBox.Show("Pls Enter UnitName", "Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        if (txtempname.Text != "" && txtidcard.Text != "" && txtunited.Text != "")
                        {
                            Models.HrPayModel c = new Models.HrPayModel();

                            if (txthrpaydetailsid.Text == "") { c.hrpaydetailsid = Convert.ToInt64("0" + txthrpaydetailsid.Text); autonumberload(); c.hrpaydetailsid1 = Convert.ToInt64("0" + txthrpaydetailsid1.Text); }
                            else { c.hrpaydetailsid = Convert.ToInt64("0" + txthrpaydetailsid.Text); c.hrpaydetailsid1 = Convert.ToInt64("0" + txthrpaydetailsid1.Text); }
                            c.hrpaydetailsid = Convert.ToInt64("0" + txthrpaydetailsid.Text);
                            c.hrpaydetailsid1 = Convert.ToInt64("0" + txthrpaydetailsid1.Text);
                            c.docid = Convert.ToString(txtdocid.Text);
                            c.docdate = txtdate.Value.ToString("yyyy-MM-dd");
                            c.finyear = Convert.ToString(combofinyear.Text);
                            c.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                            c.compname = Convert.ToInt64("0" + combocompname.SelectedValue);
                            c.idcardno = Convert.ToString(txtsequenceid.Text);
                            c.midcard = Convert.ToString(txtidcard.Text);
                            c.empname = Convert.ToString(txtempname.Text);
                            c.doj = Convert.ToDateTime(txtdoj.Value.ToString().Substring(0, 10));

                            if (txtdol.Value.ToString() != "")
                            {
                                c.dol = Convert.ToString(txtdol.Text.ToString().Substring(0, 10));

                            }
                            else
                            {
                                c.dol = "";
                            }
                            //string s = ""; c.dol = "";
                            //if (txtdol.Value.ToString().Substring(0, 10) != "")
                            //{
                            //    s = txtdol.Value.ToString().Substring(0, 10);
                            //    if (s.Substring(0, 4) == "0001") { txtdol.CustomFormat = ""; c.dol = txtdol.Text; }
                            //    else
                            //    {
                            //        txtdol.Text = Convert.ToDateTime(s).ToString();
                            //        c.dol = txtdol.Text;
                            //    }
                            //}
                            //else
                            //{
                            //    txtdol.CustomFormat = ""; c.dol = txtdol.Text;
                            //}
                            c.uanno = Convert.ToString(txtuanno.Text);
                            c.esino = Convert.ToString(txtesino.Text);
                            c.fathername = Convert.ToString(txtfathername.Text);
                            c.united = Convert.ToString(txtunited.Text);
                            c.category = Convert.ToString(txtcategory.Text);
                            c.department = Convert.ToString(txtdepartment.Text);
                            c.designation = Convert.ToString(txtdesignation.Text);
                            c.orjpayabledays = Convert.ToString(txtorjpayabledays.Text);
                            c.nhdays = Convert.ToString(txtnhdays.Text);
                            c.payabledays = Convert.ToString(txtpayabledays.Text);
                            c.govtdaysalary = Convert.ToString(txtgovtdaysal.Text);
                            c.otwages = Convert.ToString(txtotwages.Text);
                            c.basicda = Convert.ToString(txtbasicda.Text);
                            c.basic = Convert.ToString(txtbasic.Text);
                            c.da = Convert.ToString(txtda.Text);
                            c.hra = Convert.ToString(txthra.Text);
                            c.others = Convert.ToString(txtothers.Text);
                            c.ebasic = Convert.ToString(txtebasic.Text);
                            c.eda = Convert.ToString(txteda.Text);
                            c.ebasicda = Convert.ToString(txtebasicda.Text);
                            c.ehra = Convert.ToString(txtehra.Text);
                            c.eothers = Convert.ToString(txteothers.Text);
                            c.payableothours = Convert.ToString(txtpayableothours.Text);
                            c.otamount = Convert.ToString(txtotamount.Text);
                            c.incentive = Convert.ToString(txtincentive.Text);
                            c.govtgross = Convert.ToString(txtgovtgross.Text);
                            c.pfamount = Convert.ToString(txtpfamount.Text);
                            c.esiamount = Convert.ToString(txtesiamount.Text);
                            c.messamount = Convert.ToString(txtmessamount.Text);
                            c.othersexp = Convert.ToString(txtotherexp.Text);
                            c.advance = Convert.ToString(txtadvance.Text);
                            c.deduction = Convert.ToString(txtdeduction.Text);
                            c.netamount = Convert.ToString(txtnetamount.Text);
                            c.bankaccountno = Convert.ToString(txtbankaccountno.Text);
                            c.bankname = Convert.ToString(txtbankname.Text);
                            c.ifsccode = Convert.ToString(txtifsccode.Text);
                            c.payperiod = Convert.ToString(txtpayperiod.Text);
                            c.fromdate = txtfromdate.Value.ToString("dd-MMM-yyyy");
                            c.todate = txttodate.Value.ToString("dd-MMM-yyyy");
                            c.creditdate = Convert.ToString(txtcreditdate.Text);
                            string sel = "select hrpaydetailsid  from  HRPayDetails   WHERE  compcode='" + c.compcode + "' and  compname='" + c.compname + "' and midcard='" + c.midcard + "' and  idcardno='" + c.idcardno + "'  and  empname='" + c.empname + "' and  doj='" + Convert.ToDateTime(c.doj).ToString("yyyy-MM-dd").Substring(0, 10) + "' and  uanno='" + c.uanno + "'  and  esino='" + c.esino + "'  and  fathername='" + c.fathername + "' and  united='" + c.united + "'  and  category='" + c.category + "' and   department='" + c.department + "'  and   designation='" + c.designation + "'  and    orjpayabledays='" + c.orjpayabledays + "'  and  nhdays='" + c.nhdays + "'  and  payabledays='" + c.payabledays + "'  and  govtdaysalary='" + c.govtdaysalary + "' and  otwages='" + c.otwages + "' and  basicda='" + c.basicda + "' and   basic='" + c.basic + "' and  da='" + c.da + "' and  hra='" + c.hra + "' and  others='" + c.others + "' and    ebasicda='" + c.ebasicda + "' and  ebasic='" + c.ebasic + "' and   eda='" + c.eda + "'  and  ehra='" + c.ehra + "' and  eothers='" + c.eothers + "' and  payableothrs='" + c.payableothours + "' and  otamount='" + c.otamount + "' and  incentive='" + c.incentive + "' and   govtgross='" + c.govtgross + "' and    pfamount='" + c.pfamount + "' and  esiamount='" + c.esiamount + "' and  messamount='" + c.messamount + "' and  othersexp='" + c.othersexp + "' and  advance='" + c.advance + "'   and  deduction='" + c.deduction + "' and   netamount='" + c.netamount + "' and    bankaccountno='" + c.bankaccountno + "'  and  bankname='" + c.bankname + "'  and  ifsccode='" + c.ifsccode + "' and  payperiod='" + c.payperiod + "'  and  fromdate='" + c.fromdate + "' and  todate='" + c.todate + "' and active='" + c.active + "' and  creditdate='" + c.creditdate + "' ;";
                            DataSet ds = Utility.ExecuteSelectQuery(sel, "HRPayDetails");
                            DataTable dt = ds.Tables["HRPayDetails"];
                            if (dt.Rows.Count != 0)
                            {
                                empty(); tabControl1.SelectTab(tabPageraw2);
                            }
                            else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txthrpaydetailsid.Text) == 0 || Convert.ToInt32("0" + txthrpaydetailsid.Text) == 0)
                            {
                                string ins = "";
                                if (c.dol == "")
                                {
                                    ins = "insert into HRPayDetails(hrpaydetailsid1, docid,  docdate ,finyear,compcode,compname, midcard,idcardno , empname, doj,uanno, esino, fathername, united,  category,  department, designation, orjpayabledays , nhdays , payabledays , govtdaysalary, otwages, basicda , basic, da ,hra, others, ebasic, ebasicda,  eda,  ehra,  eothers, payableothrs ,otamount,incentive,  govtgross,  pfamount, esiamount,messamount,othersexp,advance, deduction,  netamount ,  bankaccountno,  bankname , ifsccode,payperiod, fromdate ,todate,active,creditdate,compcode1,username, createdby, createdon,  modifiedby,  ipaddress)values('" + c.hrpaydetailsid1 + "', '" + c.docid + "',  '" + c.docdate + "' ,  '" + c.finyear + "',  '" + c.compcode + "',  '" + c.compname + "', '" + c.midcard + "', '" + c.idcardno + "' ,  '" + c.empname + "',  '" + c.doj + "','" + c.uanno + "',  '" + c.esino + "',  '" + c.fathername + "' ,  '" + c.united + "',  '" + c.category + "',  '" + c.department + "', '" + c.designation + "', '" + c.orjpayabledays + "' ,  '" + c.nhdays + "' ,  '" + c.payabledays + "' ,  '" + c.govtdaysalary + "',  '" + c.otwages + "',  '" + c.basicda + "' ,  '" + c.basic + "',  '" + c.da + "' ,  '" + c.hra + "',  '" + c.others + "', '" + c.ebasic + "', '" + c.ebasicda + "',  '" + c.eda + "',  '" + c.ehra + "', '" + c.eothers + "',  '" + c.payableothours + "' ,  '" + c.otamount + "',  '" + c.incentive + "',  '" + c.govtgross + "',  '" + c.pfamount + "',  '" + c.esiamount + "',  '" + c.messamount + "',  '" + c.othersexp + "',  '" + c.advance + "', '" + c.deduction + "',  '" + c.netamount + "' ,  '" + c.bankaccountno + "',  '" + c.bankname + "' ,  '" + c.ifsccode + "',  '" + c.payperiod + "', '" + c.fromdate + "',  '" + c.todate + "',  '" + c.active + "','" + c.creditdate + "',  '" + c.compcode1 + "',  '" + c.username + "' , '" + c.createdby + "',  '" + c.createdon + "',  '" + c.modifiedby + "',  '" + c.ipaddress + "' );";
                                    Utility.ExecuteNonQuery(ins);
                                }
                                else
                                {
                                    ins = "insert into HRPayDetails(hrpaydetailsid1, docid,  docdate ,finyear,compcode,compname, midcard,idcardno , empname, doj,dol, uanno, esino, fathername, united,  category,  department, designation, orjpayabledays , nhdays , payabledays , govtdaysalary, otwages, basicda , basic, da ,hra, others, ebasic, ebasicda,  eda,  ehra,  eothers, payableothrs ,otamount,incentive,  govtgross,  pfamount, esiamount,messamount,othersexp,advance, deduction,  netamount ,  bankaccountno,  bankname , ifsccode,payperiod, fromdate ,todate,active,creditdate,compcode1,username, createdby, createdon,  modifiedby,  ipaddress)values('" + c.hrpaydetailsid1 + "', '" + c.docid + "',  '" + c.docdate + "' ,  '" + c.finyear + "',  '" + c.compcode + "',  '" + c.compname + "', '" + c.midcard + "', '" + c.idcardno + "' ,  '" + c.empname + "',  '" + c.doj + "','" + c.dol + "',  '" + c.uanno + "',  '" + c.esino + "',  '" + c.fathername + "' ,  '" + c.united + "',  '" + c.category + "',  '" + c.department + "', '" + c.designation + "', '" + c.orjpayabledays + "' ,  '" + c.nhdays + "' ,  '" + c.payabledays + "' ,  '" + c.govtdaysalary + "',  '" + c.otwages + "',  '" + c.basicda + "' ,  '" + c.basic + "',  '" + c.da + "' ,  '" + c.hra + "',  '" + c.others + "', '" + c.ebasic + "', '" + c.ebasicda + "',  '" + c.eda + "',  '" + c.ehra + "', '" + c.eothers + "',  '" + c.payableothours + "' ,  '" + c.otamount + "',  '" + c.incentive + "',  '" + c.govtgross + "',  '" + c.pfamount + "',  '" + c.esiamount + "',  '" + c.messamount + "',  '" + c.othersexp + "',  '" + c.advance + "', '" + c.deduction + "',  '" + c.netamount + "' ,  '" + c.bankaccountno + "',  '" + c.bankname + "' ,  '" + c.ifsccode + "',  '" + c.payperiod + "', '" + c.fromdate + "',  '" + c.todate + "',  '" + c.active + "','" + c.creditdate + "',  '" + c.compcode1 + "',  '" + c.username + "' , '" + c.createdby + "',  '" + c.createdon + "',  '" + c.modifiedby + "',  '" + c.ipaddress + "' );";
                                    Utility.ExecuteNonQuery(ins);
                                }
                                MessageBox.Show("Record Saved Successfully.Toal Record are:," + c.hrpaydetailsid1.ToString(), " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                empty(); GridLoad();
                            }
                            else
                            {
                                string up = "update  HRPayDetails  set hrpaydetailsid1='" + c.hrpaydetailsid1 + "' ,docid='" + c.docid + "' , docdate='" + c.docdate + "',finyear='" + c.finyear + "',compcode='" + c.compcode + "', compname='" + c.compname + "',  midcard='" + c.midcard + "',idcardno='" + c.idcardno + "' , empname='" + c.empname + "', doj='" + c.doj + "',dol='" + Convert.ToDateTime(c.dol).ToString("yyyy-MM-dd").Substring(0, 10) + "', uanno='" + c.uanno + "' , esino='" + c.esino + "' , fathername='" + c.fathername + "', united='" + c.united + "' , category='" + c.category + "',  department='" + c.department + "', designation='" + c.designation + "' , orjpayabledays='" + c.orjpayabledays + "' , nhdays='" + c.nhdays + "' , payabledays='" + c.payabledays + "' , govtdaysalary='" + c.govtdaysalary + "', otwages='" + c.otwages + "', basicda='" + c.basicda + "',  basic='" + c.basic + "', da='" + c.da + "', hra='" + c.hra + "', others='" + c.others + "',ebasicda='" + c.ebasicda + "', ebasic='" + c.ebasic + "',  eda='" + c.eda + "',    ehra='" + c.ehra + "', eothers='" + c.eothers + "', payableothrs='" + c.payableothours + "', otamount='" + c.otamount + "', incentive='" + c.incentive + "',  govtgross='" + c.govtgross + "',   pfamount='" + c.pfamount + "', esiamount='" + c.esiamount + "', messamount='" + c.messamount + "', othersexp='" + c.othersexp + "', advance='" + c.advance + "', deduction='" + c.deduction + "',  netamount='" + c.netamount + "',   bankaccountno='" + c.bankaccountno + "' , bankname='" + c.bankname + "' , ifsccode='" + c.ifsccode + "', payperiod='" + c.payperiod + "' , fromdate='" + c.fromdate + "', todate='" + c.todate + "',active='" + c.active + "',creditdate='" + c.creditdate + "',compcode1='" + Class.Users.COMPCODE + "', username='" + Class.Users.USERID + "',createdby='" + Class.Users.HUserName + "', modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where hrpaydetailsid='" + txthrpaydetailsid.Text + "';";
                                Utility.ExecuteNonQuery(up);

                                MessageBox.Show("Record Updated Successfully.Toal Record are:," + c.hrpaydetailsid1.ToString(), " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                empty(); GridLoad();
                            }
                        }
                    }
                    else
                    {

                        MessageBox.Show("InvalidDate.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); //custom messageBox to show error  
                        this.Dispose();
                    }
                }
            }

            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Saves_Click " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Prints()
        {
            Saves();
            if (validprint == true)
            {
               
                empty();
            }
        }

        private void buttsearch_Click_1(object sender, EventArgs e)
        {
            Txtsearch_TextChanged(sender, e);
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

        private void txtdelayreason_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txttripwagonno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
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
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
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

     

     
        private void button1_Click(object sender, EventArgs e)
        {

            if (dataGridView2.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);
                for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {

                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        xcelApp.Cells[i + 2, j + 1] = Convert.ToString(dataGridView2.Rows[i].Cells[j].Value);
                    }
                }
                xcelApp.Columns.AutoFit();
                xcelApp.Visible = true;
            }
        }




        private void button2_Click(object sender, EventArgs e)
        {
            string sel2 = "select  a.idcardno ,  a.empname,  a.doj,  a.uanno,  a.esino,  a.fathername ,  a.united,  a.category,  a.department, a.designation, a.orjpayabledays ,  a.nhdays ,  a.payabledays ,  a.govtdaysalary,  a.otwages,  a.basicda ,  a.basic,  a.da ,  a.hra,  a.others, a.ebasic, a.ebasicda,  a.eda,  a.ehra,  a.eothers,  a.payableothrs ,  a.otamount,  a.incentive,  a.govtgross,  a.pfamount,  a.esiamount,  a.messamount,  a.othersexp,  a.advance, a.deduction,  a.netamount ,  a.bankaccountno,  a.bankname , a.ifsccode,a.payperiod, a.fromdate ,  a.todate,  a.active,a.creditdate from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + combounitreport.Text + "' and a.payperiod='" + comboperiodreport.Text + "' order by 1;";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "HRPayDetails");
            DataTable dt2 = ds2.Tables["HRPayDetails"];
            dataGridView2.DataSource = dt2;
            int cnt = dataGridView2.Rows.Count - 1;
            lbltotalcount.Text = "Total Count  :" + cnt.ToString();
        }

       

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //ListViewItem item = e.Item as ListViewItem;

            //if (item.Checked==true)
            //{
            //    try
            //    {

            //        if (item.SubItems[2].Text != "")
            //        {
            //            var confirmation = MessageBox.Show("Do You want Delete this Record ?. IDCARD:"+ item.SubItems[2].Text+"=="+ item.SubItems[6].Text, "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            //            if (confirmation == DialogResult.Yes)
            //            {
            //                string del = "delete from HRPayDetails  where HRPayDetailsid='" + item.SubItems[2].Text + "';";
            //                Utility.ExecuteNonQuery(del);
            //                period = item.SubItems[6].Text;
            //                gridload();
            //                empty();
            //                MessageBox.Show("Record Deleted Successfully " + item.SubItems[2].Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);


            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }


            //}
        }

        private void HRPayDetails_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.S)
            {
                Saves();
            }

        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();
            companyload();
        }



        private void combounitsearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combounitsearch.SelectedValue) > 0)
                {
                    checkall.Checked = true;
                    string sel1 = "select  distinct '------' as  payperiod from dual union all  select distinct b.payperiod from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ACTIVE ='T' and  a.compcode='" + combounitsearch.Text + "'";

                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];
                    comboperiodsearh.DisplayMember = "payperiod";
                    comboperiodsearh.ValueMember = "payperiod";
                    comboperiodsearh.DataSource = dt1;
                }
            }
            catch (Exception ex) { }

        }

        private void comboperiodsearh_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combounitsearch.SelectedValue) > 0 && comboperiodsearh.Text != "------" && comboperiodsearh.Text != "")
                {
                    Cursor = Cursors.WaitCursor; lbltotal.Text = "";
                    listView1.Items.Clear(); listfilter.Items.Clear();
                    string sel1 = "select  a.hrpaydetailsid,a.docid ,b.compcode, a.idcardno ,a.midcard, a.payperiod,a.empname, a.doj,a.fathername,a.department,a.united from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + combounitsearch.Text + "' and a.payperiod='" + comboperiodsearh.Text + "' order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRPayDetails");
                    DataTable dt = ds.Tables["HRPayDetails"];
                    if (dt.Rows.Count >= 0)
                    {
                        int i = 1;
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(myRow["hrpaydetailsid"].ToString());
                            list.SubItems.Add(myRow["docid"].ToString());
                            list.SubItems.Add(myRow["compcode"].ToString());
                            list.SubItems.Add(myRow["idcardno"].ToString());
                            list.SubItems.Add(myRow["midcard"].ToString());
                            list.SubItems.Add(myRow["payperiod"].ToString());
                            list.SubItems.Add(myRow["empname"].ToString());
                            list.SubItems.Add(myRow["doj"].ToString());
                            list.SubItems.Add(myRow["fathername"].ToString());
                            list.SubItems.Add(myRow["department"].ToString());
                            list.SubItems.Add(myRow["united"].ToString());
                            this.listfilter.Items.Add((ListViewItem)list.Clone());
                            if (i % 2 == 0)
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            else
                            {
                                list.BackColor = Color.White;
                            }
                            listView1.Items.Add(list);
                            i++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }
                    Cursor = Cursors.Default;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void combounitreport_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(combounitreport.SelectedValue) > 0)
                {
                    string sel1 = "select '------' as  payperiod from dual union all select distinct b.payperiod from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ptransaction ='COMPANY' and  a.compcode='" + combounitreport.Text + "'";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];
                    comboperiodreport.DisplayMember = "payperiod";
                    comboperiodreport.ValueMember = "payperiod";
                    comboperiodreport.DataSource = dt1;
                }
            }
            catch (Exception ex) { }
        }

      

        private void checkall_CheckedChanged(object sender, EventArgs e)
        {
            if (txthrpaydetailsid.Text != "")
            {
                //var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                //if (confirmation == DialogResult.Yes)
                //{

                //    string sel1 = "select b.hrpaydetailsid from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ptransaction ='COMPANY' and  a.compcode='" + Class.Users.HCompcode + "' AND B.payperiod='"+ comboperiodsearh.Text + "';";

                //    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                //    DataTable dt1 = ds1.Tables["hrpaydetails"];
                //    if (dt1.Rows.Count > 0)
                //    {
                //        for (int i = 0; i < dt1.Rows.Count; i++)
                //        {

                //            string del = "delete from HRPayDetails  where compcode='" + combocompcode.SelectedValue + "' and payperiod='" + comboperiodsearh.Text + "' and  HRPayDetailsid='" + dt1.Rows[i]["hrpaydetailsid"].ToString() + "';";
                //            Utility.ExecuteNonQuery(del);
                //        }
                //        MessageBox.Show("Record Deleted Successfully " + txthrpaydetailsid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //        gridload(); empty();
                //    }
                //    else
                //    {
                //        MessageBox.Show("Invalid  Delete","Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //    }
                //}
            }
        }

        private void tabPageraw3_Click(object sender, EventArgs e)
        {

        }
        public void Searchs()
        {
            throw new NotImplementedException();
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void Deletes()
        {
            try
            {

                if (txthrpaydetailsid.Text != "" && checkall.Checked == false)
                {
                    string sel1 = "select b.empid from  gtcompmast a join pldatta b on a.compcode=b.compcode  where   a.compcode='" + combounitsearch.Text + "' AND B.payperiod='" + comboperiodsearh.Text + "';";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];
                    if (dt1.Rows.Count <= 0)
                    {
                        var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (confirmation == DialogResult.Yes)
                        {
                            Cursor = Cursors.WaitCursor;
                            string del = "delete from HRPayDetails where compcode='" + combocompcode.SelectedValue + "' and  HRPayDetailsid='" + txthrpaydetailsid.Text + "'";
                            Utility.ExecuteNonQuery(del);
                            if (Convert.ToInt64(combounitsearch.SelectedValue) > 0)
                            {
                                string sel2 = "select  distinct '------' as  payperiod from dual union all  select distinct b.payperiod from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ACTIVE ='T' and  a.compcode='" + combounitsearch.Text + "'";
                                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                                DataTable dt2 = ds2.Tables["hrpaydetails"];
                                comboperiodsearh.DisplayMember = "payperiod";
                                comboperiodsearh.ValueMember = "payperiod";
                                comboperiodsearh.DataSource = dt2;
                                comboperiodsearh.Refresh();
                            }
                            MessageBox.Show("Record Deleted Successfully " + txthrpaydetailsid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            empty(); Cursor = Cursors.Default;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Child Record Found.Can Not Delete.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }
                if (txthrpaydetailsid.Text == "" && checkall.Checked == true)
                {
                    string sel1 = "select b.empid from  gtcompmast a join pldatta b on a.compcode=b.compcode  where   a.compcode='" + combounitsearch.Text + "' AND B.payperiod='" + comboperiodsearh.Text + "';";

                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "hrpaydetails");
                    DataTable dt1 = ds1.Tables["hrpaydetails"];
                    if (dt1.Rows.Count <= 0)
                    {
                        var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (confirmation == DialogResult.Yes)
                        {
                            Cursor = Cursors.WaitCursor;
                            string sel0 = "select b.hrpaydetailsid from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where   a.compcode='" + combounitsearch.Text + "' AND B.payperiod='" + comboperiodsearh.Text + "';";

                            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "hrpaydetails");
                            DataTable dt0 = ds0.Tables["hrpaydetails"]; int tot = 0;
                            if (dt0.Rows.Count > 0)
                            {
                                progressBar2.Minimum = 0;
                                progressBar2.Maximum = dt0.Rows.Count;
                                for (int i = 0; i < dt0.Rows.Count; i++)
                                {

                                    string del = "delete from HRPayDetails  where compcode='" + combounitsearch.SelectedValue + "' and payperiod='" + comboperiodsearh.Text + "' and  HRPayDetailsid='" + dt0.Rows[i]["hrpaydetailsid"].ToString() + "';";
                                    Utility.ExecuteNonQuery(del); tot++;
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToDecimal(dt0.Rows.Count)) * (i + 1);
                                    lblprogress2.Text = " Data Transfer to Table : " + (per).ToString("N0") + " %" + dt0.Rows[i]["hrpaydetailsid"].ToString();
                                    lbltotal.Text = "Data Remove from Table" + (per).ToString("N0") + " % " + "EmpID " + dt0.Rows[i]["hrpaydetailsid"].ToString();
                                    lblprogress2.Refresh(); lbltotal.Refresh();
                                    progressBar2.Value = i + 1; tot++;

                                }

                                MessageBox.Show("Record Deleted Successfully. Total:- " + tot.ToString(), "Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                lblprogress2.Text = ""; progressBar2.Value = 0; lbltotal.Text = "";
                                empty();
                                if (Convert.ToInt64(combounitsearch.SelectedValue) > 0)
                                {
                                    string sel2 = "select  distinct '------' as  payperiod from dual union all  select distinct b.payperiod from  gtcompmast a join hrpaydetails b on a.gtcompmastid=b.compcode  where a.ACTIVE ='T' and  a.compcode='" + combounitsearch.Text + "'";
                                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "hrpaydetails");
                                    DataTable dt2 = ds2.Tables["hrpaydetails"];
                                    comboperiodsearh.DisplayMember = "payperiod";
                                    comboperiodsearh.ValueMember = "payperiod";
                                    comboperiodsearh.DataSource = dt2;
                                    comboperiodsearh.Refresh();
                                }
                                Cursor = Cursors.Default;
                            }
                            else
                            {
                                MessageBox.Show("No Data Found.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information); checkall.Checked = false;
                            }
                            Cursor = Cursors.Default;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Child Record Found.Can Not Delete.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPageraw1"])//your specific tabname
            {
                txtidcard.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPageraw2"])//your specific tabname
            {

            }

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPageraw3"])//your specific tabname
            {

            }


            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPageraw4"])//your specific tabname
            {
                combounitreport.Select();
                comboperiodreport.Select();
            }

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }
        private void txtvechileno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtdoj.Focus();
            }
        }

        private void comboreceivedfrom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtpayabledays.Select();

            }
        }

        private void txtgrossweight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtgovtdaysal.Select();

            }
        }

        private void combocompcode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtidcard.Focus();

            }
        }

        private void dateTimePicker1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void combovarietyitem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {



                txtgovtdaysal.Select();
            }
        }

        private void txttripwagonno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtbasicda.Focus();

            }
        }

        private void txtthirdpartywt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtunited.Focus();

            }
        }

        private void combogodown_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                txtbasic.Focus();
            }
        }

        private void txtlotno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtcategory.Focus();

            }
        }

        private void txtsampledby_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtda.Focus();

            }
        }

        private void txtcertifiedby_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtdepartment.Focus();

            }
        }

        private void combovisualstatus_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txthra.Focus();

            }
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
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                string filePath = string.Empty; dataGridView1.AllowUserToAddRows = false;
                string fileExt = string.Empty; combocompcode.Text = ""; combocompcode.SelectedIndex = -1;
                OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
                if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
                {
                    filePath = file.FileName; //get the path of the file  
                    fileExt = Path.GetExtension(filePath); //get the file extension  
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        try
                        {
                            DataTable dtExcel = new DataTable();
                            dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                            dataGridView1.Visible = true;
                            dataGridView1.DataSource = dtExcel;

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                    }
                }
                combounitsearch.SelectedIndex = -1;
                tabControl1.SelectTab(tabPageraw3);
                int cnt = dataGridView1.Rows.Count - 1;
                label48.Text = "Total Count  :" + cnt.ToString();
            }
            else
            {
                MessageBox.Show("pls Contact your Administrator."+ Class.Users.Log.ToString(), "Register Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Dispose();
            }
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

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        public void GridLoad()
        {
            try
            {

                //if (period == "")
                //{
                if (Convert.ToInt64(combounitsearch.SelectedValue) > 0)
                {
                    string sel1 = "select  a.hrpaydetailsid,a.docid ,b.compcode, a.idcardno ,a.midcard, a.payperiod,a.empname, a.doj,a.fathername,a.department,a.united from  hrpaydetails a join gtcompmast b on a.compcode=b.gtcompmastid WHERE b.compcode='" + Class.Users.HCompcode + "'   order by a.hrpaydetailsid desc;";


                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRPayDetails");
                    DataTable dt = ds.Tables["HRPayDetails"];
                    if (dt.Rows.Count >= 0)
                    {
                        int i = 1; listView1.Items.Clear(); listfilter.Items.Clear();
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(myRow["hrpaydetailsid"].ToString());
                            list.SubItems.Add(myRow["docid"].ToString());
                            list.SubItems.Add(myRow["compcode"].ToString());
                            list.SubItems.Add(myRow["idcardno"].ToString());
                            list.SubItems.Add(myRow["midcard"].ToString());
                            list.SubItems.Add(myRow["payperiod"].ToString());
                            list.SubItems.Add(myRow["empname"].ToString());
                            list.SubItems.Add(myRow["doj"].ToString());
                            list.SubItems.Add(myRow["fathername"].ToString());
                            list.SubItems.Add(myRow["department"].ToString());
                            list.SubItems.Add(myRow["united"].ToString());
                            this.listfilter.Items.Add((ListViewItem)list.Clone());
                            if (i % 2 == 0)
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            else
                            {
                                list.BackColor = Color.White;
                            }
                            listView1.Items.Add(list);
                            i++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }
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
                //empty();

                if (listView1.Items.Count >= 0)
                {

                    txthrpaydetailsid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select a.HRPayDetailsid, a.hrpaydetailsid1 ,a.docid , a.docdate,a.finyear,b.compcode, b.compname, a.idcardno , a.midcard,a.empname, a.doj,a.dol, a.uanno , a.esino , a.fathername, a.united , a.category,  a.department, a.designation,  a.orjpayabledays , a.nhdays , a.payabledays , a.govtdaysalary, a.otwages, a.basicda,  a.basic, a.da, a.hra, a.others, a.ebasic,  a.eda,   a.ebasicda, a.ehra, a.eothers, a.payableothrs, a.otamount, a.incentive,  a.govtgross,   a.pfamount, a.esiamount, a.messamount,a.othersexp,a.advance,a.creditdate, a.deduction,  a.netamount,   a.bankaccountno , a.bankname , a.ifsccode, a.payperiod ,   a.fromdate, a.todate,a.active from HRPayDetails a  join gtcompmast b on a.compcode = b.gtcompmastid  where a.HRPayDetailsid='" + txthrpaydetailsid.Text + "'; ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRPayDetails");
                    DataTable dt = ds.Tables["HRPayDetails"];
                    if (dt.Rows.Count > 0)
                    {
                        if (txthrpaydetailsid.Text != "")
                        {
                            txthrpaydetailsid.Text = dt.Rows[0]["HRPayDetailsid"].ToString();
                            txthrpaydetailsid1.Text = dt.Rows[0]["HRPayDetailsid1"].ToString();
                            txtdocid.Text = dt.Rows[0]["docid"].ToString();
                            txtdate.Text = dt.Rows[0]["docdate"].ToString();
                            combofinyear.Text = dt.Rows[0]["finyear"].ToString();
                            combocompcode.Text = dt.Rows[0]["compcode"].ToString();
                            combocompname.Text = dt.Rows[0]["compname"].ToString();
                            txtsequenceid.Text = dt.Rows[0]["idcardno"].ToString();
                            txtidcard.Text = dt.Rows[0]["midcard"].ToString();
                            txtempname.Text = dt.Rows[0]["empname"].ToString();
                            txtdoj.Text = dt.Rows[0]["doj"].ToString(); string s = "";
                            if (dt.Rows[0]["dol"].ToString() != "")
                            {
                                s = dt.Rows[0]["dol"].ToString().Substring(0,10);                               
                                if (s.Substring(6,4) == "0001") { txtdol.CustomFormat = ""; }
                                else
                                {
                                    txtdol.Text = Convert.ToDateTime(s).ToString();
                                }
                            }
                            else
                            {
                                txtdol.CustomFormat = "";
                            }
                            txtuanno.Text = dt.Rows[0]["uanno"].ToString();
                            txtesino.Text = dt.Rows[0]["esino"].ToString();
                            txtfathername.Text = dt.Rows[0]["fathername"].ToString();
                            txtunited.Text = dt.Rows[0]["united"].ToString();
                            txtcategory.Text = dt.Rows[0]["category"].ToString();
                            txtdepartment.Text = dt.Rows[0]["department"].ToString();
                            txtdesignation.Text = dt.Rows[0]["designation"].ToString();
                            txtorjpayabledays.Text = dt.Rows[0]["orjpayabledays"].ToString();
                            txtnhdays.Text = dt.Rows[0]["nhdays"].ToString();
                            txtpayabledays.Text = dt.Rows[0]["payabledays"].ToString();
                            txtgovtdaysal.Text = dt.Rows[0]["govtdaysalary"].ToString();
                            txtotwages.Text = dt.Rows[0]["otwages"].ToString();
                            txtbasicda.Text = dt.Rows[0]["basicda"].ToString();
                            txtbasic.Text = dt.Rows[0]["basic"].ToString();
                            txtda.Text = dt.Rows[0]["da"].ToString();
                            txthra.Text = dt.Rows[0]["hra"].ToString();
                            txtothers.Text = dt.Rows[0]["others"].ToString();
                            txtebasicda.Text = dt.Rows[0]["ebasicda"].ToString();
                            txtebasic.Text = dt.Rows[0]["ebasic"].ToString();
                            txteda.Text = dt.Rows[0]["eda"].ToString();
                            txtehra.Text = dt.Rows[0]["ehra"].ToString();
                            txteothers.Text = dt.Rows[0]["eothers"].ToString();
                            txtpayableothours.Text = dt.Rows[0]["payableothrs"].ToString();
                            txtotamount.Text = dt.Rows[0]["otamount"].ToString();
                            txtincentive.Text = dt.Rows[0]["incentive"].ToString();
                            txtgovtgross.Text = dt.Rows[0]["govtgross"].ToString();
                            txtpfamount.Text = dt.Rows[0]["pfamount"].ToString();
                            txtesiamount.Text = dt.Rows[0]["esiamount"].ToString();
                            txtmessamount.Text = dt.Rows[0]["messamount"].ToString();
                            txtotherexp.Text = dt.Rows[0]["othersexp"].ToString();
                            txtadvance.Text = dt.Rows[0]["advance"].ToString();
                            txtdeduction.Text = dt.Rows[0]["deduction"].ToString();
                            txtnetamount.Text = dt.Rows[0]["netamount"].ToString();
                            txtbankaccountno.Text = dt.Rows[0]["bankaccountno"].ToString();
                            txtbankname.Text = dt.Rows[0]["bankname"].ToString();
                            txtifsccode.Text = dt.Rows[0]["ifsccode"].ToString();
                            txtpayperiod.Text = dt.Rows[0]["payperiod"].ToString();
                            txtfromdate.Value = Convert.ToDateTime(dt.Rows[0]["fromdate"].ToString());
                            txttodate.Value = Convert.ToDateTime(dt.Rows[0]["todate"].ToString());
                            txtcreditdate.Text = dt.Rows[0]["creditdate"].ToString();
                            if (dt.Rows[0]["active"].ToString() == "T")
                                checkactive.Checked = true;
                            else checkactive.Checked = false;
                            combocompcode.Enabled = false;
                            tabControl1.SelectTab(tabPageraw1);
                        }
                        else
                        {
                            return;
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                //   MessageBox.Show("Error: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            tabControl1.SelectTab(tabPageraw1);
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {


            try
            {
                if (Convert.ToInt64(combounitsearch.SelectedValue) > 0 && comboperiodsearh.Text != "")
                {
                    int item0 = 0; listView1.Items.Clear();
                    if (txtsearch.Text.Length > 1)
                    {

                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            if (listfilter.Items[item0].SubItems[6].ToString().Contains(txtsearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[8].ToString().Contains(txtsearch.Text.ToUpper()))
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
                                if (item0 % 2 == 0)
                                {
                                    list.BackColor = Color.WhiteSmoke;
                                }
                                else
                                {
                                    list.BackColor = Color.White;
                                }
                                item0++;
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
                        foreach (ListViewItem item in listfilter.Items)
                        {

                            this.listView1.Items.Add((ListViewItem)item.Clone());
                            if (item0 % 2 == 0)
                            {
                                item.BackColor = Color.WhiteSmoke;
                            }
                            else
                            {
                                item.BackColor = Color.White;
                            }


                            item0++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }

                }
                else
                {
                    MessageBox.Show("pls select CompCode and PayPeriod");
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
        private void HRPayDetails_Load(object sender, EventArgs e)
        {
           
        }
    }
}

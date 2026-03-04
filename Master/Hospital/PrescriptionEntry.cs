using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master.Hospital
{
    public partial class PrescriptionEntry : Form,ToolStripAccess
    {
        private static PrescriptionEntry _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();string std3 = "";
        ListView listfilter3 = new ListView(); ListView listfilter1 = new ListView(); ListView speceficlistfilter1 = new ListView();
        // ListView allip = new ListView();
        byte[] stdbytes; byte[] patientbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
        public static PrescriptionEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PrescriptionEntry();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public PrescriptionEntry()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            GlobalVariables.CurrentForm = this;
            tabControl1.TabPages.Remove(tabPage2);

            // listView3.DrawSubItem += new DrawListViewSubItemEventHandler(lv_DrawSubItem);
        }

       

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void txtstudentid_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboschoolcode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {

                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }

        public void News()
        {
            empty(); GridLoad(); MedicineLoad();
          
            //GlobalVariables.HideCols = new string[] { "SNo" , "MedicineID","txtdiagnosisid","Tokenno" };
            //CommonFunctions.RemoveColumn(dataGridView1, GlobalVariables.HideCols);
            //CommonFunctions.SetRowNumber(dataGridView1);
        }
        Models.CommonClass com= new Models.CommonClass();
        public void Saves()
        {
            try
            {
                if (combotokenno.Text != "" && combopatientname.SelectedValue != null)
                {
                    string chk = "", chk1 = "";
                    if (radiomale.Checked == true) { chk = "M"; radiofemale.Checked = false; radioothers.Checked = false; }
                    if (radiofemale.Checked == true) { chk = "F"; radiomale.Checked = false; radioothers.Checked = false; }
                    if (radioothers.Checked == true) { chk = "O"; radioothers.Checked = true; radiofemale.Checked = false; radiomale.Checked = false; }
                    if (checkBox1.Checked == true) { chk1 = "T"; } else { chk1 = "F"; }
                    CommonFunctions.dt = CommonFunctions.select("SELECT asptbldiagnosismasid FROM asptbldiagnosismas WHERE phoneno='" + txtphoneno.Text + "' and asptbldiagnosismasid='" + txtdiagnosisid.Text + "' and  asptblregistermasid='" + txtregisterid.Text + "'  AND    patientname='" + combopatientname.SelectedValue + "' AND   gender='" + chk + "'  and  symptoms='" + txtsymptoms.Text + "'  and   diagonisis='" + txtdiagnosis.Text + "'  and  tokenno='" + combotokenno.SelectedValue + "'   and listcount='" + Class.Users.Staticallip.Items.Count.ToString() + "' ", "asptbldiagnosismas");
                    if (CommonFunctions.dt.Rows.Count != 0)
                    {
                    }
                    else if (CommonFunctions.dt.Rows.Count != 0 && Convert.ToInt64("0" + txtdiagnosisid.Text) == 0 || Convert.ToInt64("0" + txtdiagnosisid.Text) == 0)
                    {
                        chk1 = "T";
                        com.query = "insert into asptbldiagnosismas(asptblregistermasid,active,patientname,gender,symptoms,diagonisis,tokenno,compcode,username,createdby,createdon,modifiedby,modifiedon,ipaddress,diagnoseddate,listcount,phoneno)  VALUES('" + txtregisterid.Text + "','" + chk1 + "','" + combopatientname.SelectedValue + "','" + chk + "','" + txtsymptoms.Text.ToUpper() + "','" + txtdiagnosis.Text.ToUpper() + "','" + combotokenno.SelectedValue + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.USERID + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "' ,'" + Class.Users.USERID + "','" + System.DateTime.Now.ToString("yyyy-MM-dd") + "','" + Class.Users.IPADDRESS + "','" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + Class.Users.Staticallip.Items.Count.ToString() + "','" + txtphoneno.Text + "')";
                        Utility.ExecuteNonQuery(com.query);
                        com.query = "update  asptblregistermas  set active='T',   majordiagnosis='" + txtdiagnosis.Text + "' ,TOKENNO='" + combotokenno.SelectedValue + "',patientname='" + combopatientname.Text + "' where asptblregistermasid='" + txtregisterid.Text + "'";
                        Utility.ExecuteNonQuery(com.query);
                        CommonFunctions.dt = CommonFunctions.select("select MAX(asptbldiagnosismasid) asptbldiagnosismasid    from  asptbldiagnosismas", "asptbldiagnosismas");// WHERE asptblregistermasid='" + txtregisterid.Text + "' and patientname='" + combopatientname.SelectedValue + "' AND phoneno='" + txtphoneno.Text + "' and tokenno='" + combotokenno.Text + "'", "asptbldiagnosismas");
                        com.query = "update  asptblpatientmas  set active='T' , asptbldiagnosismasid='" + CommonFunctions.dt.Rows[0]["asptbldiagnosismasid"].ToString() + "' where patientname='" + combopatientname.Text + "' and tokenno='" + combotokenno.Text + "'";
                        Utility.ExecuteNonQuery(com.query);
                    }
                    else
                    {
                        chk1 = "T";
                        com.query = "update  asptbldiagnosismas  set  phoneno='" + txtphoneno.Text + "',active='" + chk1 + "' , asptblregistermasid='" + txtregisterid.Text + "' ,patientname='" + combopatientname.SelectedValue + "' ,   gender='" + chk + "'  ,  symptoms='" + txtsymptoms.Text.ToUpper() + "'  ,   diagonisis='" + txtdiagnosis.Text.ToUpper() + "'  ,tokenno='" + combotokenno.Text + "',  compcode='" + Class.Users.COMPCODE + "'  ,    username='" + Class.Users.USERID + "'  , modifiedby='" + Class.Users.USERID + "'  ,     ipaddress='" + Class.Users.IPADDRESS + "' , listcount='" + Class.Users.Staticallip.Items.Count.ToString() + "' where asptbldiagnosismasid='" + txtdiagnosisid.Text + "'";
                        Utility.ExecuteNonQuery(com.query);
                        com.query = "update  asptblregistermas  set active='T', majordiagnosis='" + txtdiagnosis.Text + "' ,  TOKENNO='" + combotokenno.SelectedValue + "',phoneno='" + txtphoneno.Text + "',asptblregistermasid='" + txtregisterid.Text + "',   patientname='" + combopatientname.Text + "'  where asptblregistermasid='" + txtregisterid.Text + "'";
                        Utility.ExecuteNonQuery(com.query);
                        CommonFunctions.dt = CommonFunctions.select("select MAX(asptbldiagnosismasid) asptbldiagnosismasid    from  asptbldiagnosismas", "asptbldiagnosismas");// WHERE asptblregistermasid='" + txtregisterid.Text + "' and patientname='" + combopatientname.SelectedValue + "' AND phoneno='" + txtphoneno.Text + "' and tokenno='" + combotokenno.Text + "'", "asptbldiagnosismas");
                        com.query = "update  asptblpatientmas  set active='T' , asptbldiagnosismasid='" + CommonFunctions.dt.Rows[0]["asptbldiagnosismasid"].ToString() + "' where patientname='" + combopatientname.Text + "' and tokenno='" + combotokenno.Text + "'";
                        Utility.ExecuteNonQuery(com.query);
                    }
                    GridBind();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void GridBind()
        {
            int cc = dataGridView1.Rows.Count;
            int diagnosid=0, mediid=0, tok=0;
            object med,dose,mor, afn, eve, nig, bef, aff = "";
            DataTable dt = CommonFunctions.select("SELECT MAX(asptbldiagnosismasid) AS asptbldiagnosismasid FROM asptbldiagnosismas ", "asptbldiagnosismas");
            //if (txtdiagnosisid.Text != "")
            //{
            //    DataTable dt1 = CommonFunctions.select("SELECT asptblmedmasid, asptbldiagnosismasid,tokenno,asptblregistermasid,patientname  FROM asptblmedmas where asptbldiagnosismasid='" + txtdiagnosisid.Text + "' AND  asptblregistermasid='" + txtregisterid.Text + "'", "asptbldiagnosismas");
            //    diagnosid = Convert.ToInt32("0"+ dt1.Rows[0]["asptbldiagnosismasid"].ToString());
            //    tok = Convert.ToInt32("0" + dt1.Rows[0]["tokenno"].ToString());
            //    mediid = Convert.ToInt32("0" + dt1.Rows[0]["asptblmedmasid"].ToString());
            //}           
            for (int i = 0; i < cc; i++)
            {
                if (dataGridView1.Rows[i].Cells[4].FormattedValue.ToString() != "")
                {
                    if (txtdiagnosisid.Text != "" && dataGridView1.Rows[i].Cells[1].Value != null)
                    {
                       
                    }
                    else
                    {
                        dataGridView1.Rows[i].Cells[1].Value =null;
                        dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt32("0" + dt.Rows[0]["asptbldiagnosismasid"].ToString());
                        dataGridView1.Rows[i].Cells[3].Value = combotokenno.Text;
                    }                 
                   
                    med = dataGridView1.Rows[i].Cells[4].FormattedValue.ToString().ToUpper();
                    dose  = dataGridView1.Rows[i].Cells[5].FormattedValue.ToString().ToUpper();
                    if (dataGridView1.Rows[i].Cells[6].FormattedValue.ToString() == "True") { mor = "T"; } else { mor = "F"; }
                    if (dataGridView1.Rows[i].Cells[7].FormattedValue.ToString() == "True") { afn = "T"; } else { afn = "F"; }
                    if (dataGridView1.Rows[i].Cells[8].FormattedValue.ToString() == "True") { eve = "T"; } else { eve = "F"; }
                    if (dataGridView1.Rows[i].Cells[9].FormattedValue.ToString() == "True") { nig = "T"; } else { nig = "F"; }
                    if (dataGridView1.Rows[i].Cells[10].FormattedValue.ToString() == "True") { bef = "T"; } else { bef = "F"; }
                    if (dataGridView1.Rows[i].Cells[11].FormattedValue.ToString() == "True") { aff = "T"; } else { aff = "F"; }

                    CommonFunctions.dt1 = CommonFunctions.select("SELECT asptblmedmasid from asptblmedmas where phoneno='"+txtphoneno.Text+"' and asptbldiagnosismasid = '" +txtdiagnosisid.Text + "' and asptblregistermasid = '" + txtregisterid.Text + "' and patientname = '" + combopatientname.SelectedValue + "'and tokenno = '" + combotokenno.Text + "' and medicine = '" + med + "' and dose = '" + dose + "' and mor = '" + mor + "' and afn = '" + afn + "' and eve = '" + eve + "' and nig = '" + nig + "' and bef = '" + bef + "' and aff = '" + aff + "' and asptblmedmasid=" + Convert.ToInt32("0" + dataGridView1.Rows[i].Cells[1].FormattedValue.ToString()),"asptblmedmas");
                
                    if (CommonFunctions.dt1.Rows.Count != 0) { }
                    else if (dataGridView1.Rows[i].Cells[1].Value == null && dataGridView1.Rows[i].Cells[4].Value != null)
                    {
                        com.query = "insert into asptblmedmas(asptbldiagnosismasid,asptblregistermasid,patientname,tokenno,medicine,dose,mor,afn,eve,nig,bef,aff,compcode,phoneno)  VALUES('" + dataGridView1.Rows[i].Cells[2].Value + "','" + txtregisterid.Text + "','" + combopatientname.SelectedValue + "','" + dataGridView1.Rows[i].Cells[3].Value.ToString() + "','" + med + "','" + dose + "','" + mor + "','" + afn + "','" + eve + "','" + nig + "','" + bef + "','" + aff + "','" + Class.Users.COMPCODE + "','" + txtphoneno.Text + "')";
                        Utility.ExecuteNonQuery(com.query);
                    }
                    else
                    {

                        com.query = "update  asptblmedmas set phoneno='" + txtphoneno.Text + "' ,asptbldiagnosismasid='" + dataGridView1.Rows[i].Cells[2].Value.ToString() + "',asptblregistermasid='" + txtregisterid.Text + "',patientname='" + combopatientname.SelectedValue + "',tokenno='" + dataGridView1.Rows[i].Cells[3].Value.ToString() + "',medicine='" + med + "',dose='" + dose + "',mor='" + mor + "',afn='" + afn + "',eve='" + eve + "',nig='" + nig + "',bef='" + bef + "',aff='" + aff + "',compcode='" + Class.Users.COMPCODE + "' where asptblmedmasid='" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                        Utility.ExecuteNonQuery(com.query);
                    }
                }
            }
            if (txtdiagnosisid.Text == "")
            {
                if (Class.Users.Staticallip.Items.Count >= 0)
                {
                    foreach (ListViewItem item in Class.Users.Staticallip.Items)
                    {
                        CommonFunctions.dt1 = CommonFunctions.select("SELECT asptblmedmasid from asptbldiatest where phoneno='" + txtphoneno.Text + "' and asptbldiagnosismasid = '" + txtdiagnosisid.Text + "' and asptblregistermasid = '" + txtregisterid.Text + "' and patientname = '" + combopatientname.SelectedValue + "'and tokenno = '" + combotokenno.Text + "'  and asptbllabtestitemmasid='" + item.SubItems[5].Text + "' and asptbllabtestmasid='" + item.SubItems[6].Text + "' and asptbldiagnosismasid='" + txtdiagnosisid.Text + "'", "asptblmedmas");
                        if (CommonFunctions.dt1==null)
                        {
                            diagnosid = Convert.ToInt32("0" + dt.Rows[0]["asptbldiagnosismasid"].ToString());
                           
                            CommonFunctions.query = "insert into asptbldiatest(asptblregistermasid,patientname,asptbllabtestitemmasid,asptbllabtestmasid,asptbldiagnosismasid,tokenno,labtestitem,modifiedon,phoneno)  VALUES('" + txtregisterid.Text + "','" + combopatientname.SelectedValue + "','" + item.SubItems[5].Text + "','" + item.SubItems[6].Text + "','" + diagnosid + "','" + combotokenno.Text + "','" + item.SubItems[4].Text + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + txtphoneno.Text + "')";
                            Utility.ExecuteNonQuery(CommonFunctions.query);
                        }
                    }
                }            
            }
            else
            {
                if (Class.Users.Staticallip.Items.Count >= 0)
                {
                    foreach (ListViewItem item in Class.Users.Staticallip.Items)
                    {
                        CommonFunctions.dt = CommonFunctions.select("SELECT asptbldiatestid FROM asptbldiatest WHERE phoneno='" + txtphoneno.Text + "' and labtestitem='" + item.SubItems[6].Text + "'  and  asptbldiagnosismasid='" + txtdiagnosisid.Text + "' and  asptblregistermasid='" + txtregisterid.Text + "' and  patientname='" + combopatientname.SelectedValue + "' AND    asptbllabtestmasid='" + item.SubItems[6].Text + "' and modifiedon='" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "' ", "asptbldiatest");
                        if (CommonFunctions.dt.Rows.Count == 0)
                        {
                            diagnosid = Convert.ToInt32("0" + dt.Rows[0]["asptbldiagnosismasid"].ToString());
                            CommonFunctions.query = "insert into asptbldiatest(asptblregistermasid,patientname,asptbllabtestmasid,asptbllabtestitemmasid,asptbldiagnosismasid,tokenno,labtestitem,modifiedon,phoneno)  VALUES('" + txtregisterid.Text + "','" + combopatientname.SelectedValue + "','" + item.SubItems[6].Text + "','" + item.SubItems[5].Text + "','" + txtdiagnosisid.Text + "','" + combotokenno.Text + "','" + item.SubItems[6].Text + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','"+txtphoneno.Text+"')";
                            Utility.ExecuteNonQuery(CommonFunctions.query);
                        }                       
                    }
                }               
            }
            if (txtdiagnosisid.Text == "")
            {
                MessageBox.Show("Record Saved Successfully " + txtdiagnosisid.Text, "  Inserted ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                GridLoad(); empty();
            }
            else
            {
                MessageBox.Show("Record Updated Successfully " + txtdiagnosisid.Text, " Updated ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                GridLoad(); empty();

            }
        }
        private void empty()
        {
            listView3.Visible = false;
            txtdiagnosisid.Text = ""; Class.Users.Staticallip.Items.Clear(); listfilter.Items.Clear(); 
            combotokenno.Text = ""; checkall.Checked = false;
            combopatientname.Text = ""; txtweight.Text = ""; txttemp.Text = "";
           txtregisterid.Text = ""; 
            txtpp.Text = ""; txtsuger.Text = ""; txtpules.Text = ""; txtphoneno.Text = ""; txtregisterid.Text = ""; txtregisterid.Text = "";
            checklabtest1.Checked = false;

            txtdiagnosis.Text = ""; txtsymptoms.Text = ""; radiofemale.Checked = false; radiomale.Checked = false;
            pictureBoxPatientphoto.Image = null;
            panel5.BackColor = Class.Users.BackColors;
            butborder.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            this.Font= Class.Users.FontName; 
            listView2.Font = Class.Users.FontName;          
            listView1.Font = Class.Users.FontName;
            butheader.BackColor = Class.Users.BackColors;
            button1.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panelreporbottom.BackColor = Class.Users.BackColors;
            panelreporttop.BackColor = Class.Users.BackColors;
            butreportback.ForeColor = Class.Users.ForeColors;
            panel4.BackColor = Class.Users.Color1;
            crystalReportViewer1.ReportSource = null;dataGridView1.Rows.Clear(); dataGridView2.Rows.Clear();listView1.Items.Clear();
            crystalReportViewer1.Refresh();
            this.Font = Class.Users.FontName;
            GlobalVariables.HideCols = new string[] {"MedicineID", "asptbldiagnosismasid", "Tokenno" };
            CommonFunctions.RemoveColumn(dataGridView1, GlobalVariables.HideCols);
           CommonFunctions.SetRowNumber(dataGridView1);
            GlobalVariables.HideCols = new string[] { "SNo1","ID1","asptbldiagnosismasid1","Tokenno1" };
            CommonFunctions.RemoveColumn(dataGridView2, GlobalVariables.HideCols);
            CommonFunctions.SetRowNumber(dataGridView2);
            butheader.Text = Class.Users.ScreenName; 
            dataGridView1.Font = Class.Users.FontName;         
            tabControl1.SelectTab(tabPage3);
        }
        public void Prints()
        {
            if (txtdiagnosisid.Text != "")
            {
                Class.Users.Paramid = Convert.ToInt64(txtdiagnosisid.Text);
                Master.Hospital.DiagnosisReport p = new DiagnosisReport();
               
                Class.Users.Paramid = Convert.ToInt64(txtdiagnosisid.Text);
                Class.Users.DoctorName = Convert.ToInt64(combotokenno.Text);
                Class.Users.PatientName = Convert.ToInt64(combopatientname.SelectedValue);
                
                p.Show();
            }
        }



        void GridLoad(string s,string ss)
        {

            //string sel1 = "select  A.asptblpatientmasid,b.asptblregistermasid,a.patientname,a.gender,b.majordiagnosis,b.TOKENNO,b.pp,b.suger,b.pules,a.patientphoto, b.weight,b.temp from  asptblpatientmas a   join asptblregistermas b on b.patientname=a.asptblpatientmasid  where  B.asptblregistermasid='" + s + "'  and  b.TOKENNO='" + ss + "'   order by 1";
            //DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpatientmas");
            //DataTable dt = ds.Tables["asptblpatientmas"];
            //if (dt.Rows.Count > 0)
            //{
            //    combopatientname.ValueMember = "asptblpatientmasid";
            //    combopatientname.DisplayMember = "patientname";
            //    combopatientname.DataSource = dt;
            //    combotokenno.ValueMember = "TOKENNO";
            //    combotokenno.DisplayMember = "TOKENNO";
            //    combotokenno.DataSource = dt;
            //    if (dt.Rows[0]["gender"].ToString() == "M") { radiomale.Checked = true; }
            //    if (dt.Rows[0]["gender"].ToString() == "F") { radiofemale.Checked = true; }
            //    if (dt.Rows[0]["gender"].ToString() == "O") { radioothers.Checked = true; }
            //    txtsymptoms.Text = dt.Rows[0]["majordiagnosis"].ToString();
            //    txtregisterid.Text = dt.Rows[0]["asptblregistermasid"].ToString();
            //    txtpatientid.Text = dt.Rows[0]["asptblpatientmasid"].ToString();
            //    txtpp.Text = dt.Rows[0]["pp"].ToString();
            //    txtsuger.Text = dt.Rows[0]["suger"].ToString();
            //    txtpules.Text = dt.Rows[0]["pules"].ToString();
            //    pictureBoxPatientphoto.Image = null; patientbytes = null;
            //    if (dt.Rows[0]["patientphoto"].ToString() != "")
            //    {
            //        patientbytes = (byte[])dt.Rows[0]["patientphoto"];
            //        Image img = Models.Device.ByteArrayToImage(patientbytes);
            //        pictureBoxPatientphoto.Image = img;
            //    }
            //    txtweight.Text = dt.Rows[0]["weight"].ToString();
            //    txttemp.Text = dt.Rows[0]["temp"].ToString();
            //    tabControl1.SelectTab(tabPage1);
            //}

        }
        void patientNameload(string s, string ss)
        {
            CommonFunctions.dt = CommonFunctions.select("select distinct A.asptbldiagnosismasid,a.diagonisis, c.modifiedon,c.asptblregistermasid,c.patientname,c.lastname,c.age,c.gender,c.tokenno,c.phoneno,c.pp,c.suger,c.pules,c.temp,c.weight,a.symptoms ,a.asptblregistermasid,c.patientphoto from asptblpatientmas c  join asptbldiagnosismas a on a.asptblregistermasid=c.asptblregistermasid and c.TOKENNO=a.tokenno  join asptbldocmas d on d.active='T'  where a.asptbldiagnosismasid='" + s + "'  AND C.ACTIVE='T'  order by a.asptbldiagnosismasid desc", "asptbldiagnosismas");

            if (CommonFunctions.dt.Rows.Count > 0)
            {
                txtdiagnosisid.Text = CommonFunctions.dt.Rows[0]["asptbldiagnosismasid"].ToString();
                txtdiagnosis.Text = CommonFunctions.dt.Rows[0]["diagonisis"].ToString();
                combopatientname.ValueMember = "asptblregistermasid";
                combopatientname.DisplayMember = "patientname";
                combopatientname.DataSource = CommonFunctions.dt;
                combotokenno.ValueMember = "tokenno";
                combotokenno.DisplayMember = "tokenno";
                combotokenno.DataSource = CommonFunctions.dt;
                txtphoneno.Text = CommonFunctions.dt.Rows[0]["phoneno"].ToString();
                if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "M") { radiomale.Checked = true; }
                if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "F") { radiofemale.Checked = true; }
                if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "O") { radioothers.Checked = true; }
                txtsymptoms.Text = CommonFunctions.dt.Rows[0]["symptoms"].ToString();
                txtregisterid.Text = CommonFunctions.dt.Rows[0]["asptblregistermasid"].ToString();
                txtpp.Text = CommonFunctions.dt.Rows[0]["pp"].ToString();
                txtsuger.Text = CommonFunctions.dt.Rows[0]["suger"].ToString();
                txtpules.Text = CommonFunctions.dt.Rows[0]["pules"].ToString();
                pictureBoxPatientphoto.Image = null; patientbytes = null;
                if (CommonFunctions.dt.Rows[0]["patientphoto"].ToString() != "")
                {
                    patientbytes = (byte[])CommonFunctions.dt.Rows[0]["patientphoto"];
                    Image img = Models.Device.ByteArrayToImage(patientbytes);
                    pictureBoxPatientphoto.Image = img;
                }

                txtweight.Text = CommonFunctions.dt.Rows[0]["weight"].ToString();
                txttemp.Text = CommonFunctions.dt.Rows[0]["temp"].ToString();
            }

            CommonFunctions.dt = CommonFunctions.select("SELECT asptblmedmasid,asptbldiagnosismasid,tokenno,medicine,dose,mor,afn,eve,nig,bef,aff from asptblmedmas where asptbldiagnosismasid = '" + txtdiagnosisid.Text + "'", "asptblmedmas");

            if (CommonFunctions.dt.Rows.Count > 0)
            {
                int i = 0, j = 1; dataGridView1.Rows.Clear();
                foreach (DataRow row in CommonFunctions.dt.Rows)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = j.ToString();
                    dataGridView1.Rows[i].Cells[1].Value = row.ItemArray[0].ToString();
                    dataGridView1.Rows[i].Cells[2].Value = row.ItemArray[1].ToString();
                    dataGridView1.Rows[i].Cells[3].Value = row.ItemArray[2].ToString();
                    dataGridView1.Rows[i].Cells[4].Value = row.ItemArray[3].ToString();
                    dataGridView1.Rows[i].Cells[5].Value = row.ItemArray[4].ToString();
                    if (row.ItemArray[5].ToString() == "T") { dataGridView1.Rows[i].Cells[6].Value = true; } else { dataGridView1.Rows[i].Cells[6].Value = false; }
                    if (row.ItemArray[6].ToString() == "T") { dataGridView1.Rows[i].Cells[7].Value = true; } else { dataGridView1.Rows[i].Cells[7].Value = false; }
                    if (row.ItemArray[7].ToString() == "T") { dataGridView1.Rows[i].Cells[8].Value = true; } else { dataGridView1.Rows[i].Cells[8].Value = false; }
                    if (row.ItemArray[8].ToString() == "T") { dataGridView1.Rows[i].Cells[9].Value = true; } else { dataGridView1.Rows[i].Cells[9].Value = false; }
                    if (row.ItemArray[9].ToString() == "T") { dataGridView1.Rows[i].Cells[10].Value = true; } else { dataGridView1.Rows[i].Cells[10].Value = false; }
                    if (row.ItemArray[10].ToString() == "T") { dataGridView1.Rows[i].Cells[11].Value = true; } else { dataGridView1.Rows[i].Cells[11].Value = false; }

                    i++; j++;
                }
                CommonFunctions.SetRowNumber(dataGridView1);
            }
            CommonFunctions.dt = CommonFunctions.select("select distinct b.asptbllabtestmasid,b.labtest from asptbldiatest a join asptbllabtestmas b on a.asptbllabtestmasid = b.asptbllabtestmasid where a.asptbldiagnosismasid='" + txtdiagnosisid.Text + "'  order by 1", "asptbldiatest");

            foreach (DataRow myRow in CommonFunctions.dt.Rows)
            {
                ListViewItem list = new ListViewItem();
                list.SubItems.Add("");
            }
        }
   
        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class.Users.Staticallip.Items.Clear();
        }

        private void DiagnosisMaster_Load(object sender, EventArgs e)
        {
            News(); 
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            GridLoad();empty();
        }

       

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int item0 = 0;
            //    if (txtsearch.Text.Length > 0)
            //    {
            //        listView1.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text))
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
            //                list.SubItems.Add(listfilter.Items[item0].SubItems[10].Text);
            //                if (item0 % 2 == 0)
            //                {
            //                    list.BackColor = Color.White;
            //                }
            //                else
            //                {
            //                    list.BackColor = Color.WhiteSmoke;
            //                }
            //                listView1.Items.Add(list);
            //            }
            //            item0++;
            //        }
            //    }
            //    else
            //    {
            //        try
            //        {
            //            listView1.Items.Clear(); item0 = 0;
            //            foreach (ListViewItem item in listfilter.Items)
            //            {
            //                this.listView1.Items.Add((ListViewItem)item.Clone());
            //                if (item0 % 2 == 0)
            //                {
            //                    item.BackColor = Color.White;
            //                }
            //                else
            //                {
            //                    item.BackColor = Color.WhiteSmoke;
            //                }
            //                item0++;
            //            }
            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    // MessageBox.Show("---" + ex.ToString());
            //}
        }


        private void combotokenno_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combotokenno.Text != "System.Data.DataRowView" && combotokenno.Text != "")
            {
                //patientNameload(combotokenno.Text); 
                Class.Users.Staticallip.Items.Clear(); listfilter.Items.Clear();
            }
        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void combotesting_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (allip.Items.Count <= 0)
            //{

            //    listView1.Items.Clear(); listfilter.Items.Clear();
            //    tabControl2.Width = 600;allip.Items.Clear();
            //    tabControl2.Height = 270; checklabtest.Visible = false;


            //    string sel1 = " select a.asptbllabtestitemmasID, b.labtest,a.labtestitem,a.rate, a.active  from  asptbllabtestitemmas a  join asptbllabtestmas b on a.labtest=b.asptbllabtestmasid where b.labtest='" + allip.Items[2].Text + "'  order by 1";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
            //    DataTable dt = ds.Tables["asptbllabtestitemmas"];
            //    if (dt.Rows.Count > 0)
            //    {
            //        int i = 1;
            //        foreach (DataRow myRow in dt.Rows)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            list.SubItems.Add(i.ToString());
            //            list.SubItems.Add(myRow["asptbllabtestitemmasid"].ToString());
            //            list.SubItems.Add(myRow["labtest"].ToString());
            //            list.SubItems.Add(myRow["labtestitem"].ToString());
            //            list.SubItems.Add(myRow["rate"].ToString());
            //            list.SubItems.Add("");
            //            list.SubItems.Add("0");
            //            list.SubItems.Add("0");
            //            list.SubItems.Add("0");
            //            listfilter.Items.Add((ListViewItem)list.Clone());
            //            if (i % 2 == 0)
            //            {
            //                list.BackColor = Color.White;
            //            }
            //            else
            //            {
            //                list.BackColor = Color.WhiteSmoke;
            //            }
            //            i++;
            //            listView1.Items.Add(list);
            //        }
            //        lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            //        label7.Refresh();
            //        label7.Text = "Total Count    :" + listView1.Items.Count;
            //    }
            //    else
            //    {
            //        lbltotal.Refresh();
            //        lbltotal.Text = "No data Found.";
            //        label7.Refresh();
            //        label7.Text = "No data Found.";
            //    }
            //}
           
        }
        void MedicineLoad()
        {
            string sel1 = "select a.asptblmedicalmasid, concat(a.medicine ,'-', a.Description) as medicine from asptblmedicalmas a where a.active='T' order by 1;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmedmas");
            DataTable dt = ds.Tables["asptblmedmas"];
            Medicine.DisplayMember="medicine";
            Medicine.ValueMember = "asptblmedicalmasid"; Medicine.DataSource = dt;
        }
        void MedicineLoad(string s)
        {
            //string sel1 = "select a.asptbllabtestmasid,a.labtest from asptbllabtestmas a where a.active='T' order by 1;";
            //DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestmas");
            //DataTable dt = ds.Tables["asptbllabtestmas"];
            //combotesting.DisplayMember = "labtest";
            //combotesting.ValueMember = "asptbllabtestmasid"; combotesting.DataSource = dt;
        }
        private void checkall_CheckedChanged(object sender, EventArgs e)
        {
            GridLoad();


        }
       void SpecifiGridLoad(string patid,string ph)
        {
            try
            {
                listView1.Items.Clear(); speceficlistfilter1.Items.Clear();


                CommonFunctions.dt = CommonFunctions.select("select distinct A.asptbldiagnosismasid, c.modifiedon,c.patientname,c.lastname,c.age,c.gender,c.tokenno,c.pp,c.suger,c.pules,c.temp,c.weight from asptblpatientmas c  join asptbldiagnosismas a on a.asptblregistermasid=c.asptblregistermasid and c.TOKENNO=a.tokenno  where c.patientname='" + patid + "'  and c.phoneno='" + ph + "'  AND C.ACTIVE='T'  order by A.asptbldiagnosismasid desc", "asptbldiagnosismas");
                // CommonFunctions.dt = CommonFunctions.select("SELECT  a.asptbldiagnosismasid,c.modifiedon,c.patientname,c.lastname,c.age,c.gender,c.tokenno,c.pp,c.suger,c.pules,c.temp,c.weight FROM asptbldiagnosismas a    join asptblregistermas c on c.asptblregistermasid=a.asptblregistermasid   where c.asptblregistermasid='" + patid + "'  and c.phoneno='" + ph + "'   order by a.asptbldiagnosismasid desc", "asptbldiagnosismas");

                if (CommonFunctions.dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in CommonFunctions.dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldiagnosismasid"].ToString());
                        list.SubItems.Add(myRow["modifiedon"].ToString());
                        list.SubItems.Add(myRow["pp"].ToString());
                        list.SubItems.Add(myRow["suger"].ToString());
                        list.SubItems.Add(myRow["pules"].ToString());
                        list.SubItems.Add(myRow["temp"].ToString());
                        list.SubItems.Add(myRow["tokenno"].ToString());
                        list.SubItems.Add(myRow["weight"].ToString());
                        speceficlistfilter1.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        i++;
                        listView1.Items.Add(list);
                    }

                    lbltotal2.Text = "Total Count    :" + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        public void GridLoad()
        {
            
            try
            {
                listView2.Items.Clear(); listfilter1.Items.Clear();
                DataTable dt;
                if (checkall.Checked == true) {
                    CommonFunctions.dt=CommonFunctions.select("SELECT  a.asptbldiagnosismasid,c.patientname,c.lastname,c.age,c.gender,c.tokenno,c.phoneno,date_format(a.modifiedon,'%d-%m-%Y') as modifiedon FROM asptbldiagnosismas a  join asptblregistermas c on c.asptblregistermasid=a.asptblregistermasid  where C.modifiedon='" +dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "'  order by a.asptbldiagnosismasid desc", "asptbldiagnosismas");
                  
                }
                else {
                    CommonFunctions.dt = CommonFunctions.select("SELECT  c.asptblregistermasid,c.patientname,c.lastname, c.age,c.gender,c.tokenno,c.phoneno,c.modifiedon FROM  asptblregistermas c where  C.ACTIVE='F' AND C.modifiedon='" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "'  order by 1", "asptblregistermas");
                 
                }

                if (CommonFunctions.dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in CommonFunctions.dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        if (checkall.Checked == true)
                        {
                            list.SubItems.Add(myRow["asptbldiagnosismasid"].ToString());
                        }
                        else
                        {
                            list.SubItems.Add("");
                        }
                        list.SubItems.Add(myRow["patientname"].ToString());
                        list.SubItems.Add(myRow["lastname"].ToString());                    
                        list.SubItems.Add(myRow["age"].ToString());
                        list.SubItems.Add(myRow["gender"].ToString());
                        list.SubItems.Add(myRow["tokenno"].ToString());
                        list.SubItems.Add(myRow["phoneno"].ToString());
                        list.SubItems.Add(myRow["modifiedon"].ToString());
                        listfilter1.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        i++;
                        listView2.Items.Add(list);
                    }

                    lbltotal1.Text = "Total Count    :" + listView2.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            try
            {
                listView3 .Items.Clear(); listfilter3.Items.Clear();
                CommonFunctions.dt = CommonFunctions.select("select a.asptbllabtestmasid, a.labtest  from   asptbllabtestmas a where a.active='T'  order by 2", "asptbllabtestmas");
              
                if (CommonFunctions.dt.Rows.Count > 0 || CommonFunctions.dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in CommonFunctions.dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbllabtestmasid"].ToString());
                        list.SubItems.Add(myRow["labtest"].ToString());                        
                        list.SubItems.Add("");
                        listfilter3.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        i++;
                        listView3.Items.Add(list);
                    }
                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void Searchs()
        {

            tabControl1.SelectTab(tabPage3);
           
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

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void txtweight_TextChanged(object sender, EventArgs e)
        {

        }

       
        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {

               
                    ListViewItem it2 = new ListViewItem();
                if (e.Item.Checked == true)
                {
                    e.Item.SubItems[6].Text = "✔";
                    it2.SubItems.Add(e.Item.SubItems[1].Text);
                    it2.SubItems.Add(e.Item.SubItems[2].Text);
                    it2.SubItems.Add(e.Item.SubItems[3].Text);
                    it2.SubItems.Add(e.Item.SubItems[4].Text);
                    it2.SubItems.Add(e.Item.SubItems[5].Text);
                    it2.SubItems.Add(e.Item.SubItems[6].Text);
                    it2.SubItems.Add(e.Item.SubItems[7].Text);
                    it2.SubItems.Add(e.Item.SubItems[8].Text);
                    it2.SubItems.Add(e.Item.SubItems[9].Text);
                    it2.SubItems.Add(e.Item.SubItems[10].Text);
                    Class.Users.Staticallip.Items.Add(it2);


                    Cursor = Cursors.Default;
                }
                    if (e.Item.Checked == false && e.Item.SubItems[6].Text == "✔")
                    {
                       

                        e.Item.SubItems[6].Text = "✖";
                        for (int c = 0; c < Class.Users.Staticallip.Items.Count; c++)
                        {
                            if (e.Item.SubItems[2].Text == Class.Users.Staticallip.Items[c].SubItems[2].Text)
                            {

                            Class.Users.Staticallip.Items[c].Remove();
                                c--;
                            }
                        }
                        Cursor = Cursors.Default;
                    }

              


            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

        private void checklabtest_Click(object sender, EventArgs e)
        {
            if (Class.Users.Staticallip.Items.Count > 0)
            {
                Class.Users.Bisconnectclear = true;Class.Users.UserTime = 0;
                timer1.Start();                
                this.Enabled = false;
                Class.Users.Enabled = false;
                Master.Hospital.PrescriptionPopUp popup = new PrescriptionPopUp();
               
                popup.Show();
            }
           
        }
        void GridLoadData(string id)
        {
            //DataTable dt;
            //if (checkall.Checked == true) {
            //    string sel1 = "select a.asptblregistermasid, b.asptblpatientmasid,b.patientname,a.phoneno,b.gender,b.lastname,b.address,  a.majordiagnosis,b.outpatient,b.active, a.pp,a.suger,a.pules,a.age,a.tokenno,a.weight,a.temp,b.patientphoto   from  asptblregistermas a join asptblpatientmas b on a.patientname=b.asptblpatientmasid   join gtcompmast c on a.compcode=c.gtcompmastid where a.asptblregistermasid=" + id + "  order by 1";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblregistermas");
            //    dt = ds.Tables["asptblregistermas"];
            //}
            //else
            //{
            //    string sel1 = "select a.asptblregistermasid, b.asptblpatientmasid,b.patientname,a.phoneno,b.gender,b.lastname,b.address,  a.majordiagnosis,b.outpatient,b.active, a.pp,a.suger,a.pules,a.age,a.tokenno,a.weight,a.temp,b.patientphoto   from  asptblregistermas a join asptblpatientmas b on a.patientname=b.asptblpatientmasid   join gtcompmast c on a.compcode=c.gtcompmastid where a.asptblregistermasid=" + id + "  order by 1";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblregistermas");
            //     dt = ds.Tables["asptblregistermas"];
            //}
            //if (dt.Rows.Count > 0)
            //{
               
            //    combopatientname.ValueMember = "asptblpatientmasid";
            //    combopatientname.DisplayMember = "patientname";
            //    combopatientname.DataSource = dt;

            //    combotokenno.ValueMember = "tokenno";
            //    combotokenno.DisplayMember = "tokenno";
            //    combotokenno.DataSource = dt;
            //    txtregisterid.Text = dt.Rows[0]["asptblregistermasid"].ToString();
            //    txtpatientid.Text = dt.Rows[0]["asptblpatientmasid"].ToString();

            //    if (dt.Rows[0]["gender"].ToString() == "M") { radiomale.Checked = true; } else { radiomale.Checked = true; }
            //    if (dt.Rows[0]["gender"].ToString() == "F") { radiofemale.Checked = true; } else { radiofemale.Checked = true; }
            //    if (dt.Rows[0]["gender"].ToString() == "M") { radioothers.Checked = true; } else { radioothers.Checked = true; }
            //    txtregisterid.Text = dt.Rows[0]["asptblregistermasid"].ToString();
            //    txtpp.Text = dt.Rows[0]["pp"].ToString();
            //    txtsuger.Text = dt.Rows[0]["suger"].ToString();
            //    txtpules.Text = dt.Rows[0]["pules"].ToString();
            //    txtweight.Text = dt.Rows[0]["weight"].ToString();

            //    txtsymptoms.Text = dt.Rows[0]["symptoms"].ToString();
            //    txtmedicine.Text = dt.Rows[0]["medicine"].ToString();
            //    txtdiagnosis.Text = dt.Rows[0]["diagonisis"].ToString();

            //    pictureBoxPatientphoto.Image = null; patientbytes = null;


            //    if (dt.Rows[0]["patientphoto"].ToString() != "")
            //    {

            //        patientbytes = (byte[])dt.Rows[0]["patientphoto"];
            //        Image img = Models.Device.ByteArrayToImage(patientbytes);
            //        pictureBoxPatientphoto.Image = img;

            //    }
                //if (dt.Rows[0]["medicineclob"].ToString() != "")
                //{

                //    //patientbytes = (byte[])dt.Rows[0]["medicineclob"];
                //    //string img = Models.Device.BytesToString(patientbytes);
                //    //txtmedicine.Text = img;
                //}

                //string sel3 = "select distinct b.asptbllabtestmasid,b.labtest from asptbldiatest a join asptbllabtestmas b on a.asptbllabtestmasid = b.asptbllabtestmasid where a.asptbldiagnosismasid='" + txtdiagnosisid.Text + "'  order by 1";
                //DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptbldiatest");
                //DataTable dt3 = ds3.Tables["asptbldiatest"];

                //foreach (DataRow myRow in dt3.Rows)
                //{
                //    ListViewItem list = new ListViewItem();
                //    list.SubItems.Add("");
                //}

                //string sel2 = "select a.asptbldiatestid, b.labtest,a.labtestitem,c.rate, a.asptbldiagnosismasid,a.asptbldiatestid ,a.asptblpatientmasid,a.asptblregistermasid,a.tokenno  from  asptbldiatest a  join asptbllabtestmas b on a.asptbllabtestmasid=b.asptbllabtestmasid  join asptbllabtestitemmas c on c.asptbllabtestitemmasid=a.asptbllabtestitemmasid where  a.asptbldiagnosismasid='" + txtdiagnosisid.Text + "'   and  a.asptblpatientmasid='" + txtpatientid.Text + "'  order by 1";
                //DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldiatest");
                //DataTable dt2 = ds2.Tables["asptbldiatest"];
                //if (dt2.Rows.Count > 0)
                //{
                //    int i = 1; Class.Users.Staticallip.Items.Clear();
                //    foreach (DataRow myRow in dt2.Rows)
                //    {
                //        ListViewItem list = new ListViewItem();
                //        list.SubItems.Add(i.ToString());
                //        list.SubItems.Add(myRow["asptbldiatestid"].ToString());
                //        list.SubItems.Add(myRow["labtest"].ToString());
                //        list.SubItems.Add(myRow["labtestitem"].ToString());
                //        list.SubItems.Add(myRow["rate"].ToString());
                //        list.SubItems.Add("");
                //        list.SubItems.Add(myRow["asptbldiagnosismasid"].ToString());
                //        list.SubItems.Add(myRow["asptbldiatestid"].ToString());
                //        list.SubItems.Add(myRow["asptblpatientmasid"].ToString());
                //        list.SubItems.Add(myRow["asptblregistermasid"].ToString());
                //        list.SubItems.Add(myRow["tokenno"].ToString());
                //        listfilter.Items.Add((ListViewItem)list.Clone());
                //        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                //        Class.Users.Staticallip.Items.Add((ListViewItem)list.Clone());
                //        i++;
                //    }

                //    lbltotal1.Refresh();
                //    lbltotal1.Text = "Total Count    :" + listView2.Items.Count;

                //}

               // SpecifiGridLoad(Convert.ToString(txtregisterid.Text),combotokenno.Text);


                //tabControl1.SelectTab(tabPage1);

                //txtsymptoms.Focus();
                //txtsymptoms.Select();
           // }
        }
       
        private void listView2_ItemActivate(object sender, EventArgs e)
        {
                txtdiagnosis.Text = "";dataGridView1.Rows.Clear();
            if (checkall.Checked == true)
            {
                CommonFunctions.dt = CommonFunctions.select("SELECT  a.asptbldiagnosismasid,c.asptblregistermasid,c.modifiedon,c.patientname,c.lastname,c.age,c.gender,C.phoneno,a.diagonisis,a.medicine, a.symptoms,c.tokenno,c.pp,c.suger,c.pules,c.temp,c.weight, c.patientphoto FROM asptbldiagnosismas a   join asptblregistermas c on c.asptblregistermasid=a.asptblregistermasid join asptbldocmas d on d.active='T'  where a.asptbldiagnosismasid='" + listView2.SelectedItems[0].SubItems[2].Text + "' and  c.TOKENNO='" + listView2.SelectedItems[0].SubItems[7].Text + "'   order by a.asptbldiagnosismasid desc", "asptbldiagnosismas");
                if (CommonFunctions.dt.Rows.Count > 0)
                {
                    txtdiagnosisid.Text = CommonFunctions.dt.Rows[0]["asptbldiagnosismasid"].ToString();
                    txtdiagnosis.Text = CommonFunctions.dt.Rows[0]["diagonisis"].ToString();
                    combopatientname.ValueMember = "asptblregistermasid";
                    combopatientname.DisplayMember = "patientname";
                    combopatientname.DataSource = CommonFunctions.dt;
                    combotokenno.ValueMember = "tokenno";
                    combotokenno.DisplayMember = "tokenno";
                    combotokenno.DataSource = CommonFunctions.dt;
                    txtphoneno.Text = CommonFunctions.dt.Rows[0]["phoneno"].ToString();
                    if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "M") { radiomale.Checked = true; }
                    if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "F") { radiofemale.Checked = true; }
                    if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "O") { radioothers.Checked = true; }
                    txtsymptoms.Text = CommonFunctions.dt.Rows[0]["symptoms"].ToString();
                    txtregisterid.Text = CommonFunctions.dt.Rows[0]["asptblregistermasid"].ToString();
                    txtpp.Text = CommonFunctions.dt.Rows[0]["pp"].ToString();
                    txtsuger.Text = CommonFunctions.dt.Rows[0]["suger"].ToString();
                    txtpules.Text = CommonFunctions.dt.Rows[0]["pules"].ToString();
                    pictureBoxPatientphoto.Image = null; patientbytes = null;
                    if (CommonFunctions.dt.Rows[0]["patientphoto"].ToString() != "")
                    {
                        patientbytes = (byte[])CommonFunctions.dt.Rows[0]["patientphoto"];
                        Image img = Models.Device.ByteArrayToImage(patientbytes);
                        pictureBoxPatientphoto.Image = img;
                    }

                    txtweight.Text = CommonFunctions.dt.Rows[0]["weight"].ToString();
                    txttemp.Text = CommonFunctions.dt.Rows[0]["temp"].ToString();
                }
            }
            else
            {
                CommonFunctions.dt = CommonFunctions.select("SELECT c.asptblregistermasid,c.patientname,c.lastname, c.age,c.gender,c.tokenno,c.phoneno,c.majordiagnosis AS symptoms, c.pp,c.suger,c.pules,c.temp,c.weight,c.modifiedon,c.patientphoto FROM asptblregistermas c join asptbldocmas d on d.active='T' where c.patientname='" + listView2.SelectedItems[0].SubItems[3].Text + "' and  c.phoneno='" + listView2.SelectedItems[0].SubItems[8].Text + "'  and c.tokenno='" + listView2.SelectedItems[0].SubItems[7].Text + "'  AND C.modifiedon='" + dateTimePicker1.Value.Date.ToString("yyyy-MM-dd") + "'  order by 1", "asptblregistermas");
                if (CommonFunctions.dt.Rows.Count > 0)
                {
                    txtphoneno.Text = CommonFunctions.dt.Rows[0]["phoneno"].ToString();
                    combopatientname.ValueMember = "asptblregistermasid";
                    combopatientname.DisplayMember = "patientname";
                    combopatientname.DataSource = CommonFunctions.dt;
                    combotokenno.ValueMember = "tokenno";
                    combotokenno.DisplayMember = "tokenno";
                    combotokenno.DataSource = CommonFunctions.dt;
                    txtphoneno.Text = CommonFunctions.dt.Rows[0]["phoneno"].ToString();
                    if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "M") { radiomale.Checked = true; }
                    if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "F") { radiofemale.Checked = true; }
                    if (CommonFunctions.dt.Rows[0]["gender"].ToString() == "O") { radioothers.Checked = true; }
                    txtsymptoms.Text = CommonFunctions.dt.Rows[0]["symptoms"].ToString();
                    txtregisterid.Text = CommonFunctions.dt.Rows[0]["asptblregistermasid"].ToString();
                    txtpp.Text = CommonFunctions.dt.Rows[0]["pp"].ToString();
                    txtsuger.Text = CommonFunctions.dt.Rows[0]["suger"].ToString();
                    txtpules.Text = CommonFunctions.dt.Rows[0]["pules"].ToString();
                    pictureBoxPatientphoto.Image = null; patientbytes = null;
                    if (CommonFunctions.dt.Rows[0]["patientphoto"].ToString() != "")
                    {
                        patientbytes = (byte[])CommonFunctions.dt.Rows[0]["patientphoto"];
                        Image img = Models.Device.ByteArrayToImage(patientbytes);
                        pictureBoxPatientphoto.Image = img;
                    }

                    txtweight.Text = CommonFunctions.dt.Rows[0]["weight"].ToString();
                    txttemp.Text = CommonFunctions.dt.Rows[0]["temp"].ToString();

                }
            }
            if (CommonFunctions.dt.Rows.Count > 0)
            {
                SpecifiGridLoad(combopatientname.Text.ToString(), txtphoneno.Text);
                CommonFunctions.dt = CommonFunctions.select("SELECT asptblmedmasid,asptbldiagnosismasid,tokenno,medicine,dose,mor,afn,eve,nig,bef,aff from asptblmedmas where asptbldiagnosismasid = '" + txtdiagnosisid.Text + "' AND  patientname = '" + combopatientname.SelectedValue + "'", "asptblmedmas");
            }

            if (CommonFunctions.dt.Rows.Count>0)
            {
                int i = 0, j = 1; dataGridView1.Rows.Clear();
                foreach (DataRow row in CommonFunctions.dt.Rows)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = j.ToString();
                    dataGridView1.Rows[i].Cells[1].Value = row.ItemArray[0].ToString();
                    dataGridView1.Rows[i].Cells[2].Value = row.ItemArray[1].ToString();
                    dataGridView1.Rows[i].Cells[3].Value = row.ItemArray[2].ToString();
                    dataGridView1.Rows[i].Cells[4].Value = row.ItemArray[3].ToString();
                    dataGridView1.Rows[i].Cells[5].Value = row.ItemArray[4].ToString();
                    if (row.ItemArray[5].ToString() == "T") { dataGridView1.Rows[i].Cells[6].Value = true; } else { dataGridView1.Rows[i].Cells[6].Value = false; }
                    if (row.ItemArray[6].ToString() == "T") { dataGridView1.Rows[i].Cells[7].Value = true; } else { dataGridView1.Rows[i].Cells[7].Value = false; }
                    if (row.ItemArray[7].ToString() == "T") { dataGridView1.Rows[i].Cells[8].Value = true; } else { dataGridView1.Rows[i].Cells[8].Value = false; }
                    if (row.ItemArray[8].ToString() == "T") { dataGridView1.Rows[i].Cells[9].Value = true; } else { dataGridView1.Rows[i].Cells[9].Value = false; }
                    if (row.ItemArray[9].ToString() == "T") { dataGridView1.Rows[i].Cells[10].Value = true; } else { dataGridView1.Rows[i].Cells[10].Value = false; }
                    if (row.ItemArray[10].ToString() == "T") { dataGridView1.Rows[i].Cells[11].Value = true; } else { dataGridView1.Rows[i].Cells[11].Value = false; }
                    i++;j++;
                }
            }
            CommonFunctions.dt = CommonFunctions.select("select distinct b.asptbllabtestmasid,b.labtest from asptbldiatest a join asptbllabtestmas b on a.asptbllabtestmasid = b.asptbllabtestmasid where a.asptbldiagnosismasid='" + txtdiagnosisid.Text + "'  order by 1", "asptbldiagnosismas");

            if (CommonFunctions.dt.Rows.Count > 0)
            {
                foreach (DataRow myRow in CommonFunctions.dt.Rows)
                {
                    ListViewItem list = new ListViewItem();
                    list.SubItems.Add("");
                }
            }
            CommonFunctions.dt = CommonFunctions.select("select a.asptbldiatestid, b.labtest,a.labtestitem,c.asptbllabtestitemmasid,a.patientname, a.asptbldiagnosismasid,b.asptbllabtestmasid ,a.asptblregistermasid,a.tokenno,a.phoneno  from  asptbldiatest a  join asptbllabtestmas b on a.asptbllabtestmasid=b.asptbllabtestmasid  join asptbllabtestitemmas c on c.asptbllabtestitemmasid=a.asptbllabtestitemmasid where  a.asptbldiagnosismasid='" + txtdiagnosisid.Text + "'   order by 1", "asptbldiatest");
          
            if (CommonFunctions.dt.Rows.Count > 0)
            {
                int i = 1; Class.Users.Staticallip.Items.Clear();
                foreach (DataRow myRow in CommonFunctions.dt.Rows)
                {
                    ListViewItem list = new ListViewItem();
                    list.SubItems.Add(i.ToString());
                    list.SubItems.Add(myRow["asptbldiatestid"].ToString());
                    list.SubItems.Add(myRow["labtest"].ToString());
                    list.SubItems.Add(myRow["labtestitem"].ToString());
                    list.SubItems.Add(myRow["asptbllabtestitemmasid"].ToString());               
                    list.SubItems.Add(myRow["asptbllabtestmasid"].ToString());
                    list.SubItems.Add(myRow["asptbldiagnosismasid"].ToString());
                    list.SubItems.Add(myRow["patientname"].ToString());
                    list.SubItems.Add(myRow["asptblregistermasid"].ToString());
                    list.SubItems.Add(myRow["tokenno"].ToString());
                    list.SubItems.Add(myRow["phoneno"].ToString());
                    listfilter.Items.Add((ListViewItem)list.Clone());
                    list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                    Class.Users.Staticallip.Items.Add((ListViewItem)list.Clone());
                    i++;
                }

                lbltotal1.Refresh();
                lbltotal1.Text = "Total Count    :" + listView2.Items.Count;
                
            }
            CommonFunctions.dt = CommonFunctions.select("select max(a.asptbldiagnosismasid) as asptbldiagnosismasid from asptbldiagnosismas a  join asptblregistermas b on a.asptblregistermasid=a.asptblregistermasid where b.patientname='" + combopatientname.Text + "' and b.phoneno='" + txtphoneno.Text + "' ", "asptbldiagnosismas");
            CommonFunctions.dt1 = CommonFunctions.select("SELECT asptblmedmasid,asptbldiagnosismasid,tokenno,medicine,mor,afn,eve,nig,bef,aff from asptblmedmas where patientname='" + combopatientname.SelectedValue + "' and phoneno='" + txtphoneno.Text + "' and asptbldiagnosismasid = '" + CommonFunctions.dt.Rows[0]["asptbldiagnosismasid"].ToString() + "' ", "asptblmedmas");//and patientname='" + combopatientname.SelectedValue + "'
            if (CommonFunctions.dt1.Rows.Count > 0)
            {
                int i = 0, j = 1; dataGridView2.Rows.Clear();
                foreach (DataRow row in CommonFunctions.dt1.Rows)
                {
                    dataGridView2.Rows.Add();
                    dataGridView2.Rows[i].Cells[1].Value = j.ToString();
                    dataGridView2.Rows[i].Cells[2].Value = row.ItemArray[0].ToString();
                    dataGridView2.Rows[i].Cells[3].Value = row.ItemArray[1].ToString();
                    dataGridView2.Rows[i].Cells[4].Value = row.ItemArray[2].ToString();
                    dataGridView2.Rows[i].Cells[5].Value = row.ItemArray[3].ToString();                    
                    if (row.ItemArray[4].ToString() == "T") { dataGridView2.Rows[i].Cells[6].Value = true; } else { dataGridView2.Rows[i].Cells[6].Value = false; }
                    if (row.ItemArray[5].ToString() == "T") { dataGridView2.Rows[i].Cells[7].Value = true; } else { dataGridView2.Rows[i].Cells[7].Value = false; }
                    if (row.ItemArray[6].ToString() == "T") { dataGridView2.Rows[i].Cells[8].Value = true; } else { dataGridView2.Rows[i].Cells[8].Value = false; }
                    if (row.ItemArray[7].ToString() == "T") { dataGridView2.Rows[i].Cells[9].Value = true; } else { dataGridView2.Rows[i].Cells[9].Value = false; }
                    if (row.ItemArray[8].ToString() == "T") { dataGridView2.Rows[i].Cells[10].Value = true; } else { dataGridView2.Rows[i].Cells[10].Value = false; }
                    if (row.ItemArray[9].ToString() == "T") { dataGridView2.Rows[i].Cells[11].Value = true; } else { dataGridView2.Rows[i].Cells[11].Value = false; }

                    i++; j++;
                }
            }
            tabControl1.SelectTab(tabPage1);
        }

        private void txtsearch1_TextChanged(object sender, EventArgs e)
        {

            try
            {
                int item0 = 0; i = 0;
                if (txtsearch1.Text.Length > 0)
                {
                    listView2.Items.Clear();
                    foreach (ListViewItem item in listfilter1.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter1.Items[item0].SubItems[3].ToString().Contains(txtsearch1.Text) || listfilter1.Items[item0].SubItems[8].ToString().Contains(txtsearch1.Text) || listfilter1.Items[item0].SubItems[7].ToString().Contains(txtsearch1.Text))
                        {
                            list.Text = listfilter1.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter1.Items[item0].SubItems[8].Text);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView2.Items.Add(list);
                        }
                        item0++;
                        lbltotal1.Refresh();
                        lbltotal1.Text = "Total Count    :" + listView2.Items.Count;
                    }
                }
                else
                {
                    try
                    {
                        listView2.Items.Clear(); item0 = 0;
                        foreach (ListViewItem item in listfilter1.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.SubItems.Add(item.SubItems[7].Text);
                            list.SubItems.Add(item.SubItems[8].Text);
            
                            if (i % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            i++;
                            listView2.Items.Add(list);
                        }
                        lbltotal1.Refresh();
                        lbltotal1.Text = "Total Count    :" + listView2.Items.Count;
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

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            //DialogResult result1 = MessageBox.Show("Do You want to Delete  '" + listView1.SelectedItems[0].SubItems[2].Text + "'  ??\n'    ", "'" + Class.Users.ProjectID + "' - ", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //if (result1.Equals(DialogResult.OK))
            //{
            //    if (listView1.SelectedItems[0].SubItems[6].Text == "✖")
            //    {
            //        string del = "delete from asptbldiatest where asptbldiatestid='" + listView1.SelectedItems[0].SubItems[8].Text + "'";
            //        Utility.ExecuteNonQuery(del);
            //        label7.Refresh();
            //        label7.Text = "Deleted";
            //        foreach (ListViewItem eachItem in listView1.SelectedItems)
            //        {
            //            listView1.Items.Remove(eachItem);
            //        }
            //        listView1.EndUpdate();
            //        lbltotal1.Refresh();
            //        lbltotal1.Text = "Total Count    :" + listView1.Items.Count;
            //    }
            //}
        }
        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtsearch1.Focus();txtsearch1.Select();
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
                panel5.BackColor = Class.Users.BackColors;
                butborder.BackColor = Class.Users.BackColors;
                panel2.BackColor = Class.Users.BackColors; Class.Users.UserTime = 0;
                label19.Text = "Search"; label19.ForeColor = System.Drawing.Color.Black;

                butheader.BackColor = Class.Users.BackColors;
                butheader.Text = "Prescription Entry";
                tabControl1.TabPages.Remove(tabPage2);
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
            {
                panel5.BackColor = Class.Users.BackColors;
                butborder.BackColor = Class.Users.BackColors;
                panel2.BackColor = Class.Users.BackColors; Class.Users.UserTime = 0;
                label19.Text = "Search"; label19.ForeColor = System.Drawing.Color.Black;

                butheader.BackColor = Class.Users.BackColors;
                butheader.Text = "Details";
                txtsearch1.Focus();txtsearch1.Select(); tabControl1.TabPages.Remove(tabPage2);
                GridLoad(); empty();
            }
        }

        private void listView3_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //try
            //{

            
            ListViewItem it2 = new ListViewItem();
            if (e.Item.Checked == true)
            {               
                e.Item.SubItems[4].Text = "✔";
                it2.SubItems.Add(e.Item.SubItems[1].Text);
                it2.SubItems.Add(e.Item.SubItems[2].Text);
                it2.SubItems.Add(e.Item.SubItems[3].Text);
                it2.SubItems.Add(e.Item.SubItems[4].Text);
                if (e.Item.SubItems[3].Text != null)
                {
                    Class.Users.UniqueID = txtdiagnosisid.Text;
                   
                    Class.Users.asptbltestmasid = e.Item.SubItems[2].Text;
                    Class.Users.patientname = combopatientname.SelectedValue.ToString();
                    Class.Users.asptbldiagnosisid = txtdiagnosisid.Text;
                    Class.Users.asptblregisterid = txtregisterid.Text;
                    Class.Users.tokenno = combotokenno.Text;
                    Class.Users.Bisconnectclear = true;
                     listfilter.Items.Clear();
                    Class.Users.Paramlistivew = e.Item.SubItems[3].Text;
                    this.Enabled = false;
                    Class.Users.Enabled = false; Class.Users.UserTime = 0;
                    timer1.Start();
                    Master.Hospital.PrescriptionPopUp popup = new PrescriptionPopUp();
                 
                    popup.ShowDialog();
                    
                }
            }
            if (e.Item.Checked == false && e.Item.SubItems[4].Text == "✔")
            {
                Class.Users.Bisconnectclear = false;
                e.Item.SubItems[4].Text = "✖";
                for (int c = 0; c < Class.Users.Staticallip.Items.Count; c++)
                {
                    if (e.Item.SubItems[3].Text == Class.Users.Staticallip.Items[c].SubItems[3].Text)
                    {

                        Class.Users.Staticallip.Items[c].Remove();
                        c--;
                    }
                }
                Cursor = Cursors.Default;
            }




            //}
            //catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

        private void checkbloodtest_CheckedChanged(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void listView3_ItemActivate(object sender, EventArgs e)
        {
            //if (listView3.SelectedItems[0].SubItems[3].Text != null)
            //{

            //    listView1.Items.Clear(); listfilter.Items.Clear();
            //    tabControl2.Width = 600; 
            //    tabControl2.Height = 270; checklabtest.Visible = false;
            //    string sel1 = " select a.asptbllabtestitemmasID, b.labtest,a.labtestitem,a.rate, a.active  from  asptbllabtestitemmas a  join asptbllabtestmas b on a.labtest=b.asptbllabtestmasid where b.labtest='" + listView3.SelectedItems[0].SubItems[3].Text + "'  order by 1";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
            //    DataTable dt = ds.Tables["asptbllabtestitemmas"];
            //    if (dt.Rows.Count > 0)
            //    {
            //        int i = 1;
            //        foreach (DataRow myRow in dt.Rows)
            //        {
            //            ListViewItem list = new ListViewItem();                       
            //            list.SubItems.Add(i.ToString());
            //            list.SubItems.Add(myRow["asptbllabtestitemmasid"].ToString());
            //            list.SubItems.Add(myRow["labtest"].ToString());
            //            list.SubItems.Add(myRow["labtestitem"].ToString());
            //            list.SubItems.Add(myRow["rate"].ToString());
            //            list.SubItems.Add("");
            //            list.SubItems.Add("0");
            //            list.SubItems.Add("0");
            //            list.SubItems.Add("0");
            //            list.SubItems.Add(listView3.SelectedItems[0].SubItems[2].Text);
            //            listfilter.Items.Add((ListViewItem)list.Clone());
            //            if (i % 2 == 0)
            //            {
            //                list.BackColor = Color.White;
            //            }
            //            else
            //            {
            //                list.BackColor = Color.WhiteSmoke;
            //            }
            //            i++;
            //            listView1.Items.Add(list);
            //        }
            //        lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            //        label7.Refresh();
            //        label7.Text = "Total Count    :" + listView1.Items.Count;
            //    }
            //    else
            //    {
            //        lbltotal.Refresh();
            //        lbltotal.Text = "No data Found.";
            //        label7.Refresh();
            //        label7.Text = "No data Found.";
            //    }
            //}
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void combopatientname_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Class.Users.Enabled == true)
            {
                this.timer1.Stop();
                 this.Enabled = true;
                Class.Users.Enabled = true;
            }
        }

        private void txtsymptoms_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtsymptoms.Lines.Length > 3)
            {
                txtsymptoms.ScrollBars = ScrollBars.Vertical;
            }
            else
            {
                txtsymptoms.ScrollBars = ScrollBars.None;
            }
        }

        private void txtdiagnosis_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtdiagnosis.Lines.Length > 3)
            {
                txtdiagnosis.ScrollBars = ScrollBars.Vertical;
            }
            else
            {
                txtdiagnosis.ScrollBars = ScrollBars.None;
            }
        }

        private void txtmedicine_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

      
 

        private void combopatientname_Enter(object sender, EventArgs e)
        {
            combopatientname.BackColor = Class.Users.Color2;
        }

        private void combopatientname_Leave(object sender, EventArgs e)
        {
            combopatientname.BackColor = Class.Users.Color1;
        }

        private void combotokenno_Enter(object sender, EventArgs e)
        {
            combotokenno.BackColor = Class.Users.Color2;
        }

        private void combotokenno_Leave(object sender, EventArgs e)
        {
            combotokenno.BackColor = Class.Users.Color1;
        }

        private void txtregisterid_Enter(object sender, EventArgs e)
        {
            txtregisterid.BackColor = Class.Users.Color2;
        }

        private void txtregisterid_Leave(object sender, EventArgs e)
        {
            txtregisterid.BackColor = Class.Users.Color1;
        }

        private void txtpp_Leave(object sender, EventArgs e)
        {
            txtpp.BackColor = Class.Users.Color1;
        }

        private void txtpp_Enter(object sender, EventArgs e)
        {
            txtpp.BackColor = Class.Users.Color2;
        }

        private void txtsuger_Leave(object sender, EventArgs e)
        {
            txtsuger.BackColor = Class.Users.Color1;
        }

        private void txtsuger_FontChanged(object sender, EventArgs e)
        {

        }

        private void txtsuger_Enter(object sender, EventArgs e)
        {
            txtsuger.BackColor = Class.Users.Color2;
        }

        private void txtsymptoms_Leave(object sender, EventArgs e)
        {
            txtsymptoms.BackColor = Class.Users.Color1;
        }

        private void txtsymptoms_Enter(object sender, EventArgs e)
        {
            txtsymptoms.BackColor = Class.Users.Color2;
        }

        private void txtdiagnosis_Leave(object sender, EventArgs e)
        {
            txtdiagnosis.BackColor = Class.Users.Color1;
        }

        private void txtdiagnosis_Enter(object sender, EventArgs e)
        {
            txtdiagnosis.BackColor = Class.Users.Color2;
        }

       

        private void txtpules_Leave(object sender, EventArgs e)
        {
            txtpules.BackColor = Class.Users.Color1;
        }

        private void txtpules_Enter(object sender, EventArgs e)
        {
            txtpules.BackColor = Class.Users.Color2;
        }

        private void txtweight_Leave(object sender, EventArgs e)
        {
            txtweight.BackColor = Class.Users.Color1;
        }

        private void txtweight_Enter(object sender, EventArgs e)
        {
            txtweight.BackColor = Class.Users.Color2;
        }

        private void txttemp_Enter(object sender, EventArgs e)
        {
            txttemp.BackColor = Class.Users.Color2;
        }

        private void txttemp_Leave(object sender, EventArgs e)
        {
            txttemp.BackColor = Class.Users.Color1;
        }

        private void listView1_ItemActivate_1(object sender, EventArgs e)
        {
           patientNameload(listView1.SelectedItems[0].SubItems[2].Text,listView1.SelectedItems[0].SubItems[8].Text);
           ////SpecifiGridLoad(listView1.SelectedItems[0].SubItems[2].Text,txtphoneno.Text);
            label19.Text = "Search"; label19.ForeColor = System.Drawing.Color.Black;
        }

        byte[] stdbytes1; byte[] combytes1;
        // Report.Hospital.TestReport rd1 = new Report.Hospital.TestReport();
        Report.Hospital.SummaryReport rd = new Report.Hospital.SummaryReport();
        Report.Hospital.Summary rd1 = new Report.Hospital.Summary();
        private void button2_Click(object sender, EventArgs e)
        {
            if (txtdiagnosisid.Text != "")
            {
                label19.Text = "Search"; label19.ForeColor = System.Drawing.Color.Black;
                tabControl1.TabPages.Add(tabPage2);
                tabControl1.SelectTab(tabPage2); Class.Users.dt = null;
                try
                {
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.Refresh();
                    int cnt = 0;
                   
                    string sel1 = "SELECT a.asptbldiagnosismasid,d.compcode,d.compname,d.address, d.patientname,d.gender,a.diagonisis, a.symptoms,'' medicine,date_format(a.diagnoseddate,'%d-%m-%Y %h:%m:%s') AS diagnoseddate,d.tokenno,d.asptblregistermasid ,d.patientphoto,E.sign,''labtest,'' labtestitem,E.doctorname,E.asptbldocmasid,c.companylogo as complogo,''dose,'' mor,'' afn,'' eve,'' nig,'' bef,'' aff FROM asptbldiagnosismas a join gtcompmast c on c.gtcompmastid=a.compcode    join asptblregistermas d on d.asptblregistermasid=A.asptblregistermasid     join asptbldocmas e    where e.active='T' AND a.asptbldiagnosismasid= '" + txtdiagnosisid.Text + "'   ";//and e.asptbldocmasid='" + Class.Users.DoctorName + "' and d.asptblregistermasid='" + Class.Users.PatientName + "'
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbldiagnosismas");
                    DataTable dt1 = ds1.Tables["asptbldiagnosismas"];

                    try
                    {
                        crystalReportViewer1.ReportSource = null;
                        crystalReportViewer1.Refresh();
                        cnt = 0; int i = 0; int j = 0;            
                        if (dt1.Rows.Count > 0)
                        {
                            if (Class.Users.dt == null)
                            {
                                Class.Users.dt = dt1.Clone();
                            }
                            foreach (DataRow porow in dt1.Rows)
                            {
                                Class.Users.dt.Rows.Add();
                                if (Class.Users.dt.Rows[0]["patientphoto"].ToString() == "")
                                {
                                    if (dt1.Rows[cnt]["patientphoto"].ToString() != "")
                                    {
                                        stdbytes = (byte[])dt1.Rows[cnt]["patientphoto"];
                                        Class.Users.dt.Rows[cnt]["patientphoto"] = stdbytes;
                                    }
                                    if (dt1.Rows[cnt]["sign"].ToString() != "")
                                    {
                                        stdbytes1 = (byte[])dt1.Rows[cnt+i]["sign"];
                                        Class.Users.dt.Rows[cnt+i]["sign"] = stdbytes1;
                                    }

                                    if (dt1.Rows[cnt]["complogo"].ToString() != "")
                                    {
                                        combytes1 = (byte[])dt1.Rows[cnt+i]["complogo"];
                                        Class.Users.dt.Rows[cnt+i]["complogo"] = combytes1;
                                    }
                                }
                                Class.Users.dt.Rows[cnt + i]["asptbldiagnosismasid"] = porow.ItemArray[0].ToString();
                                Class.Users.dt.Rows[cnt + i]["compcode"] = porow.ItemArray[1].ToString();
                                Class.Users.dt.Rows[cnt + i]["compname"] = porow.ItemArray[2].ToString();
                                Class.Users.dt.Rows[cnt + i]["address"] = porow.ItemArray[3].ToString();
                                Class.Users.dt.Rows[cnt + i]["patientname"] = porow.ItemArray[4].ToString();
                                Class.Users.dt.Rows[cnt + i]["gender"] = porow.ItemArray[5].ToString();
                                Class.Users.dt.Rows[cnt + i]["diagonisis"] = porow.ItemArray[6].ToString();
                                Class.Users.dt.Rows[cnt + i]["symptoms"] = porow.ItemArray[7].ToString();
                                Class.Users.dt.Rows[cnt + i]["medicine"] = porow.ItemArray[8].ToString();
                                Class.Users.dt.Rows[cnt + i]["diagnoseddate"] = porow.ItemArray[9].ToString();
                                Class.Users.dt.Rows[cnt + i]["tokenno"] = porow.ItemArray[10].ToString();
                                Class.Users.dt.Rows[cnt + i]["asptblregistermasid"] = porow.ItemArray[11].ToString();
                                Class.Users.dt.Rows[cnt + i]["doctorname"] = porow.ItemArray[16].ToString();
                                Class.Users.dt.Rows[cnt + i]["asptbldocmasid"] = porow.ItemArray[17].ToString();
                                CommonFunctions.dt = CommonFunctions.select("SELECT f.medicine,f.dose,f.mor,f.afn,f.eve,f.nig,f.bef,f.aff,f.asptblmedmasid,a.diagonisis, a.symptoms FROM asptbldiagnosismas a   join gtcompmast c on c.gtcompmastid=a.compcode    join asptblregistermas d on d.asptblregistermasid=A.asptblregistermasid   join asptbldocmas e  join asptblmedmas f on f.asptbldiagnosismasid=a.asptbldiagnosismasid and f.patientname=a.asptblregistermasid  and f.tokenno=a.tokenno  where e.active='T'   and  f.asptbldiagnosismasid= '" + txtdiagnosisid.Text + "' and f.patientname= '" + combopatientname.SelectedValue + "' order by 9", "asptbldiagnosismas");//and e.asptbldocmasid='" + Class.Users.DoctorName + "' and d.asptblregistermasid='" + Class.Users.PatientName + "'
                               
                                if (CommonFunctions.dt != null)
                                {
                                    int k = 0;
                                    foreach (DataRow porow1 in CommonFunctions.dt.Rows)
                                    {
                                        if (k == 0)
                                        {
                                            Class.Users.dt.Rows[cnt + i]["medicine"] = porow1.ItemArray[0].ToString();
                                            Class.Users.dt.Rows[cnt + i]["dose"] = porow1.ItemArray[1].ToString();
                                            if (porow1.ItemArray[2].ToString() == "T") { Class.Users.dt.Rows[cnt + i]["mor"] = "✔"; }
                                            if (porow1.ItemArray[3].ToString() == "T") { Class.Users.dt.Rows[cnt + i]["afn"] = "✔"; }
                                            if (porow1.ItemArray[4].ToString() == "T") { Class.Users.dt.Rows[cnt + i]["eve"] = "✔"; }
                                            if (porow1.ItemArray[5].ToString() == "T") { Class.Users.dt.Rows[cnt + i]["nig"] = "✔"; }
                                            if (porow1.ItemArray[6].ToString() == "T") { Class.Users.dt.Rows[cnt + i]["bef"] = "✔"; }
                                            if (porow1.ItemArray[7].ToString() == "T") { Class.Users.dt.Rows[cnt + i]["aff"] = "✔"; }
                                            Class.Users.dt.Rows[cnt + i]["diagonisis"] = porow1.ItemArray[9].ToString();
                                            Class.Users.dt.Rows[cnt + i]["symptoms"] = porow1.ItemArray[10].ToString();
                                            k++;
                                        }
                                        else
                                        {
                                            Class.Users.dt.Rows.Add();
                                            Class.Users.dt.Rows[j]["medicine"] = porow1.ItemArray[0].ToString();
                                            Class.Users.dt.Rows[j]["dose"] = porow1.ItemArray[1].ToString();
                                            if (porow1.ItemArray[2].ToString() == "T") { Class.Users.dt.Rows[j]["mor"] = "✔"; }
                                            if (porow1.ItemArray[3].ToString() == "T") { Class.Users.dt.Rows[j]["afn"] = "✔"; }
                                            if (porow1.ItemArray[4].ToString() == "T") { Class.Users.dt.Rows[j]["eve"] = "✔"; }
                                            if (porow1.ItemArray[5].ToString() == "T") { Class.Users.dt.Rows[j]["nig"] = "✔"; }
                                            if (porow1.ItemArray[6].ToString() == "T") { Class.Users.dt.Rows[j]["bef"] = "✔"; }
                                            if (porow1.ItemArray[7].ToString() == "T") { Class.Users.dt.Rows[j]["aff"] = "✔"; }
                                            Class.Users.dt.Rows[cnt + i]["diagonisis"] = porow1.ItemArray[9].ToString();
                                            Class.Users.dt.Rows[cnt + i]["symptoms"] = porow1.ItemArray[10].ToString();
                                            
                                        }
                                        i++; j++;
                                    }
                                }

                                CommonFunctions.dt = CommonFunctions.select("SELECT A.modifiedon, g.labtest,h.labtestitem  FROM asptbldiagnosismas a   join gtcompmast c on c.gtcompmastid=a.compcode     join asptblregistermas d on  d.asptblregistermasid = a.asptblregistermasid  join asptbldocmas e on  e.active='T' join asptbldiatest f on f.asptbldiagnosismasid=a.asptbldiagnosismasid   join asptbllabtestmas g on g.asptbllabtestmasid=f.asptbllabtestmasid    join asptbllabtestitemmas h on h.labtest=f.asptbllabtestmasid and h.asptbllabtestitemmasid=f.asptbllabtestitemmasid where a.asptbldiagnosismasid= '" + txtdiagnosisid.Text + "' and d.asptblregistermasid= '" + txtregisterid.Text + "'  ", "asptbldiagnosismas");//and e.asptbldocmasid='" + Class.Users.DoctorName + "' and d.asptblregistermasid='" + Class.Users.PatientName + "'
                             
                                if (CommonFunctions.dt != null)
                                {
                                    int m = 0, n = 0; 
                                    foreach (DataRow porow2 in CommonFunctions.dt.Rows)
                                    {
                                        m++;
                                        Class.Users.dt.Rows[n]["labtestitem"] += "(" + m + ") " + porow2.ItemArray[2].ToString().Trim() + "  \n";
                                   
                                    }
                                }

                            }
                        }
                        rd.Database.Tables["DataTable1"].SetDataSource(Class.Users.dt);
                        crystalReportViewer1.ReportSource = rd;
                        crystalReportViewer1.Refresh();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

               
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else { label19.Text = "Invalid";label19.ForeColor = System.Drawing.Color.Red; }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (txtdiagnosisid.Text != "")
            {
                label19.Text = "Search"; label19.ForeColor = System.Drawing.Color.Black;
                tabControl1.TabPages.Add(tabPage2);
                tabControl1.SelectTab(tabPage2); Class.Users.dt = null;
                try
                {
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.Refresh();
                    int cnt = 0; int i = 0; int j = 0;
                    CommonFunctions.dt = CommonFunctions.select("SELECT a.asptbldiagnosismasid,c.compcode,c.compname,c.address, d.patientname,d.gender,a.diagonisis, a.symptoms,'' medicine,date_format(a.diagnoseddate,'%d-%m-%Y %h:%m:%s') AS diagnoseddate,d.tokenno as test1,d.asptblregistermasid ,d.patientphoto,E.sign,''labtest,'' labtestitem,E.doctorname,E.asptbldocmasid,c.companylogo as complogo,'' mor,'' afn,'' eve,'' nig,'' bef,'' aff FROM asptbldiagnosismas a   join gtcompmast c on c.gtcompmastid=a.compcode  join asptblregistermas d on d.asptblregistermasid=A.asptblregistermasid     join asptbldocmas e  where e.active='T'    AND d.asptblregistermasid= '" + txtregisterid.Text + "'  order by 1,11 ", "asptbldiagnosismas");

                    if (CommonFunctions.dt.Rows.Count > 0)
                    {
                        if (Class.Users.dt == null)
                        {
                            Class.Users.dt = CommonFunctions.dt.Clone();            
                        }
                        foreach (DataRow porow in CommonFunctions.dt.Rows)
                        {
                            Class.Users.dt.Rows.Add();
                            if (porow["patientphoto"].ToString() != "")
                            {
                                stdbytes = (byte[])porow["patientphoto"];
                                Class.Users.dt.Rows[cnt + i]["patientphoto"] = stdbytes;
                            }
                            if (porow["sign"].ToString() != "")
                            {
                                stdbytes1 = (byte[])porow["sign"];
                                Class.Users.dt.Rows[cnt + i]["sign"] = stdbytes1;
                            }
                            if (porow["complogo"].ToString() != "")
                            {
                                combytes1 = (byte[])porow["complogo"];
                                Class.Users.dt.Rows[cnt + i]["complogo"] = combytes1;
                            }

                            Class.Users.dt.Rows[cnt + i]["asptbldiagnosismasid"] = porow.ItemArray[0].ToString();
                            Class.Users.dt.Rows[cnt + i]["compcode"] = porow.ItemArray[1].ToString();
                            Class.Users.dt.Rows[cnt + i]["compname"] = porow.ItemArray[2].ToString();
                            Class.Users.dt.Rows[cnt + i]["address"] = porow.ItemArray[3].ToString();
                            Class.Users.dt.Rows[cnt + i]["patientname"] = porow.ItemArray[4].ToString();
                            Class.Users.dt.Rows[cnt + i]["gender"] = porow.ItemArray[5].ToString();
                            Class.Users.dt.Rows[cnt + i]["diagonisis"] = porow.ItemArray[6].ToString();
                            Class.Users.dt.Rows[cnt + i]["symptoms"] = porow.ItemArray[7].ToString();
                            Class.Users.dt.Rows[cnt + i]["medicine"] = porow.ItemArray[8].ToString();
                            Class.Users.dt.Rows[cnt + i]["diagnoseddate"] = porow.ItemArray[9].ToString();
                            Class.Users.dt.Rows[cnt + i]["test1"] = porow.ItemArray[10].ToString();
                            Class.Users.dt.Rows[cnt + i]["asptblregistermasid"] = porow.ItemArray[11].ToString();
                            Class.Users.dt.Rows[cnt + i]["doctorname"] = porow.ItemArray[16].ToString();
                            Class.Users.dt.Rows[cnt + i]["asptbldocmasid"] = porow.ItemArray[17].ToString();
                            CommonFunctions.dt = CommonFunctions.select("SELECT f.medicine,f.mor,f.afn,f.eve,f.nig,f.bef,f.aff,f.asptblmedmasid FROM asptbldiagnosismas a join gtcompmast c on c.gtcompmastid=a.compcode    join asptblregistermas d on d.asptblregistermasid=A.asptblregistermasid   join asptbldocmas e  join asptblmedmas f on f.asptbldiagnosismasid=a.asptbldiagnosismasid and f.patientname=a.asptblregistermasid  and f.tokenno=a.tokenno  where e.active='T'   and  f.asptbldiagnosismasid= '" + porow.ItemArray[0].ToString() + "' and d.asptblregistermasid= '" + txtregisterid.Text + "' order by 7", "asptbldiagnosismas");//and e.asptbldocmasid='" + Class.Users.DoctorName + "' and d.asptblregistermasid='" + Class.Users.PatientName + "'
                         
                            if (CommonFunctions.dt != null)
                            {
                                int k = 0;string mor="", afn="", eve="", nig="", bef="", aff = "";
                                foreach (DataRow porow1 in CommonFunctions.dt.Rows)
                                {
                                    k++;
                                    if (porow1.ItemArray[1].ToString() == "T") { mor = " Morning "; } else { mor = ""; }
                                    if (porow1.ItemArray[2].ToString() == "T") { afn = " AfterNoon "; } else { afn = ""; }
                                    if (porow1.ItemArray[3].ToString() == "T") { eve = " Evening "; } else { eve = ""; }
                                    if (porow1.ItemArray[4].ToString() == "T") { nig = " Night "; } else { nig = ""; }
                                    if (porow1.ItemArray[5].ToString() == "T") { bef = " BeforeFood "; } else { bef = ""; }
                                    if (porow1.ItemArray[6].ToString() == "T") { aff = " AfterFood "; } else { aff = ""; }
                                    Class.Users.dt.Rows[cnt + i]["medicine"] += "("+ k +") " + porow1.ItemArray[0].ToString() + "-" + mor + "-" + afn + "-" + eve + "-" + nig + "-" + bef + "-" + aff + "\n";
                                    
                                }
                            }

                            CommonFunctions.dt = CommonFunctions.select("SELECT A.modifiedon, g.labtest,h.labtestitem  FROM asptbldiagnosismas a  join gtcompmast c on c.gtcompmastid=a.compcode     join asptblregistermas d on  d.asptblregistermasid = a.asptblregistermasid  join asptbldocmas e on  e.active='T' join asptbldiatest f on f.asptbldiagnosismasid=a.asptbldiagnosismasid join asptbllabtestmas g on g.asptbllabtestmasid=f.asptbllabtestmasid    join asptbllabtestitemmas h on h.labtest=f.asptbllabtestmasid and h.asptbllabtestitemmasid=f.asptbllabtestitemmasid where a.asptbldiagnosismasid= '" + txtdiagnosisid.Text + "' and d.asptblregistermasid= '" + txtregisterid.Text + "'  ", "asptbldiagnosismas");//and e.asptbldocmasid='" + Class.Users.DoctorName + "' and d.asptblregistermasid='" + Class.Users.PatientName + "'
                         
                            if (CommonFunctions.dt != null)
                            {
                                int m = 1, n = 0; int k = 0;
                                foreach (DataRow porow2 in CommonFunctions.dt.Rows)
                                {
                                    k++;
                                        Class.Users.dt.Rows[cnt+i]["labtestitem"] += "(" + k + ") " + porow2.ItemArray[2].ToString().Trim() + "\n";
                                }                            
                            }
                            i++;
                        }
                    }
                    rd1.Database.Tables["DataTable1"].SetDataSource(Class.Users.dt);
                    crystalReportViewer1.ReportSource = rd1;
                    crystalReportViewer1.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else { label19.Text = "Invalid"; label19.ForeColor = System.Drawing.Color.Red; }

        }

        private void butreportback_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabPage2);
            tabControl1.SelectTab(tabPage1);
        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            News();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checklabtest1.Checked == true)
            {
                listView3.Visible = true;
                this.dataGridView1.Size = new System.Drawing.Size(485, 170);
                dataGridView1.Location = new System.Drawing.Point(247, 289);
               
            }
            else
            {
                listView3.Visible = false;              
               this.dataGridView1.Size = new System.Drawing.Size(725, 170);
                dataGridView1.Location = new System.Drawing.Point(10, 289);
                
            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }

        private void txtdiagnosis_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        int rowindex = 0;
        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {
                DialogResult result1 = MessageBox.Show("Do You want to Delete  '" + std3 + "'  ??\n'    ", "'" + Class.Users.ProjectID + "' - ", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result1.Equals(DialogResult.OK))
                {
                    CommonFunctions.query = "delete FROM asptblmedmas a  WHERE a.asptblmedmasid='" + std + "')";
                    Utility.ExecuteNonQuery(CommonFunctions.query);
                    dataGridView1.Rows.RemoveAt(this.rowindex);
                }
            }
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                dataGridView1.Rows[e.RowIndex].Selected = true;
                rowindex = e.RowIndex;
                dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[4];
                std3 = dataGridView1.Rows[e.RowIndex].Cells[4].EditedFormattedValue.ToString();
                std = 0;
                std = Convert.ToInt64(dataGridView1.Rows[e.RowIndex].Cells[1].FormattedValue.ToString());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }
        int rowind = 1;
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            //bool isSelected = Convert.ToBoolean(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected);
            //if (isSelected)
            //{
            //    dataGridView1.Rows.Add(rowind, 0, 0, 0, dataGridView2.Rows[e.RowIndex].Cells[5].Value);
            //    rowind++;
            //}
            

        }

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

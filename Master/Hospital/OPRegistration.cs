using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tulpep.NotificationWindow;

namespace Pinnacle.Master.Hospital
{
    public partial class OPRegistration : Form, ToolStripAccess
    {
        private static OPRegistration _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] patientbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
      
        public static OPRegistration Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new OPRegistration();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public OPRegistration()
        {
            InitializeComponent();

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;            
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;           
            
        }
        public void News()
        {
            GridLoad(); empty();
        }
        private bool check1 = false;
        private bool check()
        {

            if (txtphoneno.Text == "")
            {
                MessageBox.Show("Pls Enter Mobile Number", "Mobile");
                txtphoneno.Select();
                return false;
            }
            if (combopatientname.Text == "")
            {
                MessageBox.Show("Pls Select Patient Name", "Patient Name");
                combopatientname.Select();
                return false;
            }
            if (txtage.Text == "")
            {
                MessageBox.Show("Pls Enter Age", "Age");
                txtage.Select();
                return false;

            }
            if (txtmajordiagnosis.Text == "")
            {
                MessageBox.Show("Pls Enter Patient Input", "PatientInput");
                txtmajordiagnosis.Select();
                txtmajordiagnosis.Focus();
                return false;
            }
           
            
            return true;
        }
      
        public void Saves()
        {
            try
            {
                MySqlCommand cmd;
                if (check() == true)
                {
                    string chk = "F"; string gen = "";
                    if (radiomale.Checked == true) { gen = "M"; radiomale.Checked = false; } else { gen = "F"; radiomale.Checked = false; radiomale.Checked = true; }
                    CommonFunctions.dt = CommonFunctions.select("SELECT asptblregistermasid FROM asptblregistermas where  modifiedon='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' and asptblregistermasid='" + txtregisterid.Text + "' and compcode='" + Class.Users.COMPCODE + "'  and    patientname='" + combopatientname.Text + "' and    lastname='" + txtlastname.Text + "' and    address='" + txtaddress.Text + "' and    gender='" + gen + "'  and age='" + txtage.Text + "' and    phoneno='" + txtphoneno.Text + "'  AND majordiagnosis='" + txtmajordiagnosis.Text + "'  and pp='" + txtpp.Text + "' and suger='" + txtsuger.Text + "' and pules='" + txtpules.Text + "'  and weight='" + txtwt.Text + "' and temp='" + txttemp.Text + "' ", "asptblregistermas");
                    if (CommonFunctions.dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtregisterid.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (CommonFunctions.dt.Rows.Count != 0 && Convert.ToInt32("0" + txtregisterid.Text) == 0 || Convert.ToInt32("0" + txtregisterid.Text) == 0)
                    {                       
                        CommonFunctions.dt = CommonFunctions.select("select  COUNT(a.TOKENNO)+1 CNT  from asptblpatientmas a where  a.modifiedon='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' AND patientname='" + combopatientname.Text + "' AND phoneno='" + txtphoneno.Text + "'", "asptblpatientmas");
                        txttokkenno.Text = "";
                        if (CommonFunctions.dt.Rows.Count == 0) { txttokkenno.Text = "1"; }
                        if (CommonFunctions.dt.Rows.Count > 0)
                        { txttokkenno.Text = CommonFunctions.dt.Rows[0]["CNT"].ToString(); }
                        string ins = "insert into asptblregistermas(TOKENNO,compcode, compname,    patientname,    lastname,    address,    gender,age, phoneno, majordiagnosis,active,    username,    ipaddress,    createdby,    createdon,    modifiedon,pp,suger,pules,weight, temp)  VALUES('" + txttokkenno.Text + "','" + Class.Users.COMPCODE + "','" + Class.Users.COMPCODE + "','" + combopatientname.Text.ToUpper() + "','" + txtlastname.Text.ToUpper() + "','" + txtaddress.Text.ToUpper() + "','" + gen + "','" + Convert.ToString(txtage.Text) + "','" + Convert.ToString(txtphoneno.Text) + "','" + txtmajordiagnosis.Text.ToUpper() + "','" + chk + "', '" + Class.Users.USERID + "','" + Class.Users.IPADDRESS + "','" + Class.Users.USERID + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + txtpp.Text + "','" + txtsuger.Text + "','" + txtpules.Text + "' , '" + Convert.ToInt64("0" + txtwt.Text) + "','" + txttemp.Text + "' )";
                        Utility.ExecuteNonQuery(ins);
                        CommonFunctions.dt = CommonFunctions.select("select MAX(asptblregistermasID) ID    from  asptblregistermas   WHERE patientname='" + combopatientname.Text + "' AND phoneno='" + txtphoneno.Text + "' and GENDER='" + gen + "'", "asptblregistermas");
                        if (CommonFunctions.dt.Rows.Count > 0 && stdbytes != null)
                        {
                            CommonFunctions.query = "UPDATE  asptblregistermas SET patientphoto=@patientphoto where  asptblregistermasID='" + CommonFunctions.dt.Rows[0]["ID"].ToString() + "'";
                            cmd = new MySqlCommand(CommonFunctions.query, Utility.Connect());
                            cmd.Parameters.Add(new MySqlParameter("@patientphoto", stdbytes));
                            cmd.ExecuteNonQuery();   
                        }
                        
                            string ins2 = "insert into asptblpatientmas(asptblregistermasID,TOKENNO,compcode, compname,patientname,lastname,address,gender,age, phoneno, majordiagnosis,active,username,ipaddress,createdby,createdon,modifiedon,pp,suger,pules,weight,temp)VALUES('" + CommonFunctions.dt.Rows[0]["ID"].ToString() + "','" + txttokkenno.Text + "','" + Class.Users.COMPCODE + "','" + Class.Users.COMPCODE + "','" + combopatientname.Text.ToUpper() + "','" + txtlastname.Text.ToUpper() + "','" + txtaddress.Text.ToUpper() + "','" + gen + "','" + Convert.ToString(txtage.Text) + "','" + Convert.ToString(txtphoneno.Text) + "','" + txtmajordiagnosis.Text + "','" + chk + "', '" + Class.Users.USERID + "','" + Class.Users.IPADDRESS + "','" + Class.Users.USERID + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + txtpp.Text + "','" + txtsuger.Text + "','" + txtpules.Text + "' , '" + Convert.ToInt64("0" + txtwt.Text) + "','" + txttemp.Text + "' )";
                            Utility.ExecuteNonQuery(ins2);
                      
                        MessageBox.Show("Record Saved Successfully " + combopatientname.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        mas.pop("Notification", combopatientname.Text, txttokkenno.Text);
                        GridLoad(); empty();
                    }
                    else
                    {

                        CommonFunctions.dt = CommonFunctions.select("SELECT asptblregistermasid FROM asptblregistermas where  modifiedon='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' and asptblregistermasid='" + txtregisterid.Text + "' and compcode='" + Class.Users.COMPCODE + "'  and    patientname='" + combopatientname.Text + "' and    lastname='" + txtlastname.Text + "' and    address='" + txtaddress.Text + "' and    gender='" + gen + "'  and age='" + txtage.Text + "' and    phoneno='" + txtphoneno.Text + "'  AND majordiagnosis='" + txtmajordiagnosis.Text + "'  and pp='" + txtpp.Text + "' and suger='" + txtsuger.Text + "' and pules='" + txtpules.Text + "'  and weight='" + txtwt.Text + "' and temp='" + txttemp.Text + "' ", "asptblregistermas");
                        if (CommonFunctions.dt.Rows.Count != 0 || txttokkenno.Text== "")
                        {
                            DataTable dt1 = CommonFunctions.select("select MAX(asptblregistermasID) asptblregistermasID    from  asptblregistermas", "asptblregistermas");
                            CommonFunctions.dt = CommonFunctions.select("select  COUNT(a.TOKENNO)+1 CNT  from asptblpatientmas a where  a.modifiedon='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' AND patientname='" + combopatientname.Text + "' AND phoneno='" + txtphoneno.Text + "'", "asptblpatientmas");
                            txttokkenno.Text = "";
                            txttokkenno.Text = CommonFunctions.dt.Rows[0]["CNT"].ToString();
                            string qry = "insert into asptblpatientmas(asptblregistermasID,TOKENNO,compcode, compname,patientname,lastname,address,gender,age, phoneno, majordiagnosis,active,username,ipaddress,createdby,createdon,modifiedon,pp,suger,pules,weight,temp)VALUES('" + dt1.Rows[0]["asptblregistermasid"].ToString() + "','" + txttokkenno.Text + "','" + Class.Users.COMPCODE + "','" + Class.Users.COMPCODE + "','" + combopatientname.Text.ToUpper() + "','" + txtlastname.Text.ToUpper() + "','" + txtaddress.Text.ToUpper() + "','" + gen + "','" + Convert.ToString(txtage.Text) + "','" + Convert.ToString(txtphoneno.Text) + "','" + txtmajordiagnosis.Text.ToUpper() + "','" + chk + "', '" + Class.Users.USERID + "','" + Class.Users.IPADDRESS + "','" + Class.Users.USERID + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + txtpp.Text + "','" + txtsuger.Text + "','" + txtpules.Text + "' , '" + Convert.ToInt64("0" + txtwt.Text) + "','" + txttemp.Text + "' )";
                            Utility.ExecuteNonQuery(qry);
                        }
                        else
                        {
                            CommonFunctions.query = "update  asptblpatientmas  set  modifiedon='" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "',pp='" + txtpp.Text + "',suger='" + txtsuger.Text + "',pules='" + txtpules.Text + "' , weight='" + Convert.ToInt64("0" + txtwt.Text) + "',temp='" + txttemp.Text + "' where patientname='" + combopatientname.Text + "' and tokenno='" + txttokkenno.Text + "'";
                            Utility.ExecuteNonQuery(CommonFunctions.query);
                        }
                        CommonFunctions.query = "update  asptblregistermas  set  TOKENNO='" + txttokkenno.Text + "',asptblregistermasID='" + txtregisterid.Text + "',patientname='" + combopatientname.Text.ToUpper() + "', compcode='" + Class.Users.COMPCODE + "' ,   compname='" + Class.Users.COMPCODE + "' ,lastname='" + txtlastname.Text.ToUpper() + "' ,address='" + txtaddress.Text.ToUpper() + "' ,gender='" + gen + "' ,age='" + Convert.ToString(txtage.Text) + "' ,majordiagnosis='" + txtmajordiagnosis.Text.ToUpper() + "' ,active='" + chk + "' ,phoneno='" + txtphoneno.Text + "', username='" + Class.Users.USERID + "',    ipaddress='" + Class.Users.IPADDRESS + "', pp='" + txtpp.Text + "',suger='" + txtsuger.Text + "',pules='" + txtpules.Text + "' , weight='" + Convert.ToInt64("0" + txtwt.Text) + "' , temp='" + txttemp.Text + "',modifiedon='" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "',asptblpatientmasid='0' where asptblregistermasid='" + txtregisterid.Text + "'";
                        Utility.ExecuteNonQuery(CommonFunctions.query);
                        if (txtregisterid.Text != "" && stdbytes != null)
                        {
                            CommonFunctions.query = "UPDATE  asptblregistermas SET patientphoto=@patientphoto where  asptblregistermasID='" + txtregisterid.Text + "'";
                            cmd = new MySqlCommand(CommonFunctions.query, Utility.Connect());
                            cmd.Parameters.Add(new MySqlParameter("@patientphoto", stdbytes));
                            cmd.ExecuteNonQuery();
                        }
                        mas.pop("Notification", combopatientname.Text, txttokkenno.Text);
                        MessageBox.Show("Record Updated Successfully " + combopatientname.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty();
                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
           
        }

       
        public void Deletes()
        {
            if (txtregisterid.Text != "")
            {
                string sel1 = "select a.asptblregistermasid from asptblregistermas a join asptbldiagnosismas b on b.doctorname =a.doctorname and b.asptblregistermasid=a.asptblregistermasid and b.tokenno=a.TOKENNO where a.asptblregistermasid='" + txtregisterid.Text + "' and a.patientname='" + combopatientname.SelectedValue + "' and a.tokenno='" + txttokkenno.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblregistermas");
                DataTable dt = ds.Tables["asptblregistermas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + combopatientname.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblregistermas where asptblregistermasid='" + Convert.ToInt64("0" + txtregisterid.Text) + "' and patientname='" + combopatientname.SelectedValue + "' and tokenno='" + txttokkenno.Text + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + combopatientname.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }


        private void PatientMaster_Load(object sender, EventArgs e)
        {

            deptload();
            News();

        }
        void deptload()
        {
            ////string sel1 = " select a.asptblhosdeptmasID, a.department   from  asptblhosdeptmas a  where a.active='T' order by 1";
            ////DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
            ////DataTable dt = ds.Tables["asptblhosdeptmas"];
           
            ////combodept.ValueMember = "asptblhosdeptmasID";
            ////combodept.DisplayMember = "department";
            ////combodept.DataSource = dt;
        }
        void deptload(string s)
        {
            //if (combodept.SelectedIndex >= 0)
            //{
                
            //    string sel1 = " select b.asptbldocmasid,b.doctorname  from  asptblhosdeptmas a join asptbldocmas b on a.asptblhosdeptmasid=b.department where  a.active='T'  AND a.asptblhosdeptmasid='" + s + "'  order by 1";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
            //    DataTable dt = ds.Tables["asptblhosdeptmas"];
               
            //    combodoctor.ValueMember = "asptbldocmasid";
            //    combodoctor.DisplayMember = "doctorname";
            //    combodoctor.DataSource = dt;
            //}
        }
        private void empty()
        {
           
            txtregisterid.Text = ""; 
            combopatientname.Text = "";
            txtaddress.Text = ""; txtwt.Text = "";
            txtlastname.Text = ""; txtpp.Text = ""; txtsuger.Text = ""; txtpules.Text = "";
            radiomale.Checked = false;
            radiofemale.Checked = false; txttemp.Text = "";
            txtmajordiagnosis.Text = ""; txttokkenno.Text = "";
            txtage.Text = "";
            pictureBoxPatientphoto.Image = null;
            radiofemale.Checked = false; radiomale.Checked = false; txtphoneno.Text = "";
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;           
            tabControl2.SelectTab(tabPage1); 
            txtphoneno.Select(); txtphoneno.Focus();
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                  
                    txtregisterid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    CommonFunctions.dt = CommonFunctions.select("select a.asptblregistermasid, a.patientname,a.phoneno,a.gender,a.lastname,a.address,  a.majordiagnosis,a.outpatient,a.active, a.pp,a.suger,a.pules,a.age,a.tokenno,a.weight,a.temp,a.patientphoto   from  asptblregistermas a   join gtcompmast c on a.compcode=c.gtcompmastid where a.asptblregistermasid=" + txtregisterid.Text + "","asptblregistermas");
                  
                    if (CommonFunctions.dt.Rows.Count > 0)
                    {                    
                        foreach (DataRow myRow in CommonFunctions.dt.Rows)
                        {
                            txtregisterid.Text = Convert.ToString(myRow["asptblregistermasid"].ToString());
                            txtphoneno.Text = Convert.ToString(myRow["phoneno"].ToString());
                            combopatientname.Text = Convert.ToString(myRow["patientname"].ToString());
                            txtlastname.Text = Convert.ToString(myRow["lastname"].ToString());
                            txtaddress.Text = Convert.ToString(myRow["address"].ToString());
                            if (myRow["gender"].ToString() == "M") { radiomale.Checked = true; }
                            else if (myRow["gender"].ToString() == "F") { radiofemale.Checked = true; }
                            else { radioothers.Checked = true; }
                            txtmajordiagnosis.Text= Convert.ToString(myRow["majordiagnosis"].ToString());
                            txttokkenno.Text=Convert.ToString(myRow["tokenno"].ToString());
                            txtpp.Text = Convert.ToString(myRow["pp"].ToString());
                            txtsuger.Text = Convert.ToString(myRow["suger"].ToString());
                            txtpules.Text = Convert.ToString(myRow["pules"].ToString());
                            txtage.Text = Convert.ToString(myRow["age"].ToString());
                            txtwt.Text = Convert.ToString(myRow["weight"].ToString());
                            txttemp.Text = Convert.ToString(myRow["temp"].ToString());
                            pictureBoxPatientphoto.Image = null; patientbytes = null;
                            if (myRow["patientphoto"].ToString() != "")
                            {
                                patientbytes = (byte[])myRow["patientphoto"];
                                Image img = Models.Device.ByteArrayToImage(patientbytes);
                                pictureBoxPatientphoto.Image = img;
                            }
                           
                            txtmajordiagnosis.Select();
                            tabControl2.SelectTab(tabPage1);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }

        //private void combodept_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    deptload(combodept.SelectedValue.ToString());
        //}

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[6].ToString().Contains(txtsearch.Text))
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
                            if (item0 % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {
                    try
                    {
                        listView1.Items.Clear(); item0 = 0;
                        foreach (ListViewItem item in listfilter.Items)
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
           
                            if (item0 % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            item0++;
                            listView1.Items.Add(list);
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

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deptload();
            GridLoad();
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void combodoctor_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
        }
        QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
        void QrcodeGenerate(string s)
        {
            //try
            //{
            //    s = "Patient ID  : " + txtregisterid.Text + "    :   Name  : " + combopatientname.Text + "    Date Timne :  " + System.DateTime.Now.ToString() + "  Token Number: " + Convert.ToInt32("0" + txttokkenno.Text);
            //    Label lblqrheader = new Label();
            //    var mydata = qc.CreateQrCode(s, QRCoder.QRCodeGenerator.ECCLevel.L);
            //    var code = new QRCoder.QRCode(mydata);
            //    pictureBoxqrcode.Image = code.GetGraphic(5, Color.Black, Color.White, true);
            //    MemoryStream stream = new MemoryStream();
            //    pictureBoxqrcode.Image.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);




            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("error" + ex.ToString());
            //}


        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtmajordiagnosis_KeyDown(object sender, KeyEventArgs e)
        {
          
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void tokendate_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkactive_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void butsearch_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void txtphoneno_TextChanged(object sender, EventArgs e)
        {
            if (txtphoneno.Text.Length >= 10)
            {
                string sel1 = "select distinct a.asptblregistermasid, a.patientname,a.lastname,a.address,a.gender,a.phoneno from asptblregistermas a where  a.phoneno='" + txtphoneno.Text + "' order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblregistermas");
                DataTable dt = ds.Tables["asptblregistermas"];
                if (dt.Rows.Count > 0)
                {

                    combopatientname.DisplayMember = "patientname";
                    combopatientname.ValueMember = "asptblregistermasid";
                    combopatientname.DataSource = dt;
                    this.combopatientname.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;

                    txtregisterid.Text = Convert.ToString(dt.Rows[0]["asptblregistermasid"].ToString());
                    txtphoneno.Text = Convert.ToString(dt.Rows[0]["phoneno"].ToString());
                    combopatientname.Text = Convert.ToString(dt.Rows[0]["patientname"].ToString());
                    txtlastname.Text = Convert.ToString(dt.Rows[0]["lastname"].ToString());
                    txtaddress.Text = Convert.ToString(dt.Rows[0]["address"].ToString());
                    if (dt.Rows[0]["gender"].ToString() == "M") { radiomale.Checked = true; }
                    else if (dt.Rows[0]["gender"].ToString() == "F") { radiofemale.Checked = true; }
                    else { radioothers.Checked = true; }

                }
                else
                {
                    combopatientname.DataSource = null;
                    txtlastname.Text = "";
                    txtaddress.Text = ""; this.combopatientname.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple;
                    txtage.Text = "";
                    pictureBoxPatientphoto.Image = null; patientbytes = null;
                }
            }
        }

        private void combopatientname_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (combopatientname.Text != null && txtphoneno.Text != null)
                {

                    CommonFunctions.dt = CommonFunctions.select("select a.asptblregistermasid,b.compcode,a.patientname,A.lastname,a.address,a.gender, a.gender,a.phoneno,a.bloodgroup,a.age, a.outpatient,a.active,a.patientphoto  from asptblregistermas  a   join gtcompmast b on a.compcode=b.gtcompmastid  where a.patientname='" + combopatientname.Text + "'", "asptblregistermas");
                    if (CommonFunctions.dt.Rows.Count > 0)
                    {
                       
                        foreach (DataRow myRow in CommonFunctions.dt.Rows)
                        {
                            txtregisterid.Text = Convert.ToString(myRow["asptblregistermasid"].ToString());
                            combopatientname.Text = Convert.ToString(myRow["patientname"].ToString());
                            txtlastname.Text = Convert.ToString(myRow["lastname"].ToString());
                            txtaddress.Text = Convert.ToString(myRow["address"].ToString());
                            txtage.Text = Convert.ToString(myRow["age"].ToString());
                            if (myRow["gender"].ToString() == "M") { radiomale.Checked = true; }
                            if (myRow["gender"].ToString() == "F") { radiofemale.Checked = true; }
                            if (myRow["gender"].ToString() == "O") { radioothers.Checked = true; }
                            if (myRow["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                            pictureBoxPatientphoto.Image = null; patientbytes = null;
                            if (myRow["patientphoto"].ToString() != "")
                            {

                                patientbytes = (byte[])myRow["patientphoto"];
                                Image img = Models.Device.ByteArrayToImage(patientbytes);
                                pictureBoxPatientphoto.Image = img;
                            }

                            combopatientname.Focus();
                            combopatientname.Select();

                        }
                    }
                    else
                    {

                        txtregisterid.Text = ""; txtlastname.Select();
                        txtaddress.Text = ""; txtwt.Text = "";
                        txtlastname.Text = ""; txtpp.Text = ""; txtsuger.Text = ""; txtpules.Text = "";
                        radiomale.Checked = false;
                        radiofemale.Checked = false; txttemp.Text = "";
                        txtmajordiagnosis.Text = ""; txttokkenno.Text = "";
                        txtage.Text = "";
                        pictureBoxPatientphoto.Image = null;
                        radiofemale.Checked = false; radiomale.Checked = false;

                    }

                }

            }
            catch (Exception ex)
            {

            }

        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad(); empty();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
           
                GridLoad();
           
          
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                DataTable dt = new DataTable();
                if (checkBox1.Checked == true)
                {                   
                    CommonFunctions.dt = CommonFunctions.select("select a.asptblregistermasid,a.patientname,a.lastname,a.gender,a.tokenno ,a.phoneno,a.lastname,a.age,a.weight,date_format(a.modifiedon,'%d-%m-%Y') as modifiedon from asptblregistermas a   join gtcompmast c on a.compcode=c.gtcompmastid    order by a.asptblregistermasid desc", "asptblregistermas");
                }
                else
                {
                    CommonFunctions.dt = CommonFunctions.select("select a.asptblregistermasid,a.patientname,a.lastname,a.gender,a.tokenno ,a.phoneno,a.lastname,a.age,a.weight,date_format(a.modifiedon,'%d-%m-%Y') as modifiedon from asptblregistermas a join gtcompmast c on a.compcode=c.gtcompmastid   where a.modifiedon='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'  order by a.asptblregistermasid desc", "asptblregistermas");
                }
                if (CommonFunctions.dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in CommonFunctions.dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblregistermasid"].ToString());
                        list.SubItems.Add(myRow["patientname"].ToString());
                        list.SubItems.Add(myRow["lastname"].ToString());
                        list.SubItems.Add(myRow["gender"].ToString());
                        list.SubItems.Add(myRow["tokenno"].ToString());
                        list.SubItems.Add(myRow["age"].ToString());
                        list.SubItems.Add(myRow["weight"].ToString());
                        list.SubItems.Add(myRow["phoneno"].ToString());                       
                        list.SubItems.Add(myRow["modifiedon"].ToString());
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                        listView1.Items.Add(list);
                    }
                    lbltotal.Text = "Total Count    :" + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void txtage_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtphoneno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtwt_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtage_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtpules_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtsuger_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtpp_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == '/' || e.KeyChar == (char)Keys.Back);
        }

        private void pictureBoxPatientphoto_Click(object sender, EventArgs e)
        {
            try
            {
                stdbytes = null;
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {

                        p.Image = new Bitmap(open.FileName);
                        stdbytes = Models.Device.ImageToByteArray(p);
                        System.Text.Encoding enc = System.Text.Encoding.ASCII;
                        std = Convert.ToInt64("0" + stdbytes.Length);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label17_Click_1(object sender, EventArgs e)
        {

        }

        private void pictureBoxPatientphoto_Click_1(object sender, EventArgs e)
        {

        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage1"])
            {
              
                butheader.BackColor = Class.Users.BackColors;
                panel3.BackColor = Class.Users.BackColors;
              
                txtphoneno.Select();
            }
            else
            {
               
                butheader.BackColor = Class.Users.BackColors;
                panel3.BackColor = Class.Users.BackColors;
               
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void txtwt_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        }

        private void txtmajordiagnosis_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtmajordiagnosis.Lines.Length > 12)
            {
                txtmajordiagnosis.ScrollBars = ScrollBars.Vertical;
            }
            else
            {
                txtmajordiagnosis.ScrollBars = ScrollBars.None;
            }
        }

        private void txtaddress_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (txtaddress.Lines.Length >3)
            {
                txtaddress.ScrollBars = ScrollBars.Vertical;
            }
            else
            {
                txtaddress.ScrollBars = ScrollBars.None;
            }
        }

        private void txtphoneno_Leave(object sender, EventArgs e)
        {
            txtphoneno.BackColor = Class.Users.Color1;
        }

        private void txtphoneno_Enter(object sender, EventArgs e)
        {
            txtphoneno.BackColor = Class.Users.Color2;
           
        }

        private void combopatientname_Enter(object sender, EventArgs e)
        {          
            txtphoneno.BackColor = Class.Users.Color2;
        }

        private void combopatientname_Leave(object sender, EventArgs e)
        {
            combopatientname.BackColor = Class.Users.Color1;

            
            if (combopatientname.Text != null)
            {
                CommonFunctions.dt = CommonFunctions.select("select a.asptblregistermasid,b.compcode,a.patientname,A.lastname,a.address,a.gender, a.gender,a.phoneno,a.bloodgroup,a.age, a.outpatient,a.active,a.patientphoto  from asptblregistermas  a   join gtcompmast b on a.compcode=b.gtcompmastid  where a.asptblregistermasid='" + combopatientname.SelectedValue + "'", "asptblregistermas");

                if (CommonFunctions.dt.Rows.Count > 0)
                {

                }
                else
                {
                    txtregisterid.Text = ""; txtlastname.Select();
                    txtaddress.Text = ""; txtwt.Text = "";
                    txtlastname.Text = ""; txtpp.Text = ""; txtsuger.Text = ""; txtpules.Text = "";
                    radiomale.Checked = false;
                    radiofemale.Checked = false; txttemp.Text = "";
                    txtmajordiagnosis.Text = ""; txttokkenno.Text = "";
                    txtage.Text = "";
                    pictureBoxPatientphoto.Image = null;
                    radiofemale.Checked = false; radiomale.Checked = false;

                }

            }

        }

        private void txtlastname_Leave(object sender, EventArgs e)
        {
            txtlastname.BackColor = Class.Users.Color1;
        }

        private void txtlastname_Enter(object sender, EventArgs e)
        {
            txtlastname.BackColor = Class.Users.Color2;
           
        }

        private void txtaddress_Leave(object sender, EventArgs e)
        {
            txtaddress.BackColor = Class.Users.Color1;
        }

        private void txtaddress_Enter(object sender, EventArgs e)
        {
            txtaddress.BackColor = Class.Users.Color2;
        }

        private void txtage_Leave(object sender, EventArgs e)
        {
            txtage.BackColor = Class.Users.Color1;
        }

        private void txtage_Enter(object sender, EventArgs e)
        {
            txtage.BackColor = Class.Users.Color2;
        }

        private void txtmajordiagnosis_Leave(object sender, EventArgs e)
        {
            txtmajordiagnosis.BackColor = Class.Users.Color1;
        }

        private void txtmajordiagnosis_Enter(object sender, EventArgs e)
        {
            txtmajordiagnosis.BackColor = Class.Users.Color2;
        }

        private void combopatientname_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtmajordiagnosis_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void combopatientname_TextChanged_1(object sender, EventArgs e)
        {
           
        }

        private void combopatientname_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyData == Keys.Tab)
            {
                if (combopatientname.Text != null)
                {
                    CommonFunctions.dt = CommonFunctions.select("select a.asptblregistermasid,b.compcode,a.patientname,A.lastname,a.address,a.gender, a.gender,a.phoneno,a.bloodgroup,a.age, a.outpatient,a.active,a.patientphoto  from asptblregistermas  a   join gtcompmast b on a.compcode=b.gtcompmastid  where a.asptblregistermasid='" + combopatientname.SelectedValue + "'", "asptblregistermas");

                    if (CommonFunctions.dt.Rows.Count > 0)
                    {

                    }
                    else
                    {
                        txtregisterid.Text = ""; txtlastname.Select();
                        txtaddress.Text = ""; txtwt.Text = "";
                        txtlastname.Text = ""; txtpp.Text = ""; txtsuger.Text = ""; txtpules.Text = "";
                        radiomale.Checked = false;
                        radiofemale.Checked = false; txttemp.Text = "";
                        txtmajordiagnosis.Text = ""; txttokkenno.Text = "";
                        txtage.Text = "";
                        pictureBoxPatientphoto.Image = null;
                        radiofemale.Checked = false; radiomale.Checked = false;

                    }

                }
                e.IsInputKey = true;
            }
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}
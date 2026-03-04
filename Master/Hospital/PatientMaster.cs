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

namespace Pinnacle.Master.Hospital
{
    public partial class PatientMaster : Form,ToolStripAccess
    {


        private static PatientMaster _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
        public static PatientMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PatientMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public PatientMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
          
         
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
           empty();GridLoad();
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
            if (txtpatientname.Text == "")
            {
                MessageBox.Show("Pls Select Patient Name", "Patient Name");

                txtpatientname.Select();
                return false;
            }
            if (txtaddress.Text == "")
            {
                MessageBox.Show("Pls Enter  Address", "Address Name");

                txtaddress.Select();
                return false;
            }
            if (radiomale.Checked==false && radiofemale.Checked == false && radioothers.Checked == false)
            {
                MessageBox.Show("Pls Select Male or FeMale","Radio Button");
               
                return false;
            }
            if (txtage.Text =="")
            {
                MessageBox.Show("Pls Enter Age", "Age");
                txtage.Select();
                return false;
            }
            return true;
        }
        public void Saves()
        {
            try {
                
            
                if (check()==true)
                {

                    MySqlCommand cmd;
                    string gen = "";
                    if (radiomale.Checked == true) { gen = "M"; radiofemale.Checked = false;radioothers.Checked = false; }
                    if (radiofemale.Checked == true) { gen = "F"; radiomale.Checked = false; radioothers.Checked = false; }
                    if (radioothers.Checked == true) { gen = "O"; radioothers.Checked = true; radiofemale.Checked = false; radiomale.Checked = false; }

                    string chk = "",  inpa = "";
                    inpa = comboinout.Text;
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "SELECT asptblpatientmasid FROM asptblpatientmas where compcode='" + Class.Users.COMPCODE + "'  and    patientname='" + txtpatientname.Text + "' and    lastname='" + txtlastname.Text + "' and    address='" + txtaddress.Text + "' and    gender='" + gen + "' and    dateofbirth='" + Convert.ToString(txtage.Text) + "' and    phoneno='" + txtphoneno.Text + "' and     bloodgroup='" + combobroup.Text + "'  and     age='" + txtage.Text + "'  and     outpatient='" + inpa + "'  and     active='" + chk + "'  and patientbytes='" + std + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpatientmas");
                    DataTable dt = ds.Tables["asptblpatientmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtregisterid.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtregisterid.Text) == 0 || Convert.ToInt32("0" + txtregisterid.Text) == 0)
                    {
                        string ins = "insert into asptblpatientmas(compcode, compname,    patientname,    lastname,    address,    gender,    dateofbirth,    phoneno,    bloodgroup,    age,outpatient,    active,    username,    ipaddress,    createdby,    createdon,    modifiedon)  VALUES('" + Class.Users.COMPCODE + "','" + Class.Users.COMPCODE + "','" + txtpatientname.Text + "','" + txtlastname.Text + "','" + txtaddress.Text + "','" + gen + "','" + Convert.ToString(txtage.Text) + "','" + txtphoneno.Text + "','" + combobroup.Text + "','" +txtage.Text + "','" + inpa + "','" + chk + "','" + Class.Users.USERID + "','" + Class.Users.IPADDRESS + "','" + Class.Users.USERID + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "','" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "')";
                        Utility.ExecuteNonQuery(ins);
                        string sel1 = "select MAX(asptblpatientmasID) ID    from  asptblpatientmas   WHERE patientname='" + txtpatientname.Text + "' AND phoneno='"+txtphoneno.Text+"' and GENDER='" + gen + "'";
                        DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblpatientmas");
                        DataTable dt1 = ds1.Tables["asptblpatientmas"];
                        if (dt1.Rows.Count > 0 && stdbytes != null)
                        {
                            string ins1 = "UPDATE  asptblpatientmas SET patientphoto=@patientphoto where  asptblpatientmasID='" + dt1.Rows[0]["ID"].ToString() + "'";
                            cmd = new MySqlCommand(ins1, Utility.Connect());
                            cmd.Parameters.Add(new MySqlParameter("@patientphoto", stdbytes));
                            cmd.ExecuteNonQuery();
                        }

            
                        MessageBox.Show("Record Saved Successfully " + txtpatientname.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblpatientmas  set   compcode='" + Class.Users.COMPCODE + "' ,   compname='" + Class.Users.COMPCODE + "' ,    patientname='" + txtpatientname.Text + "' ,    lastname='" + txtlastname.Text + "' ,address='" + txtaddress.Text + "' ,gender='" + gen + "' ,dateofbirth='" + Convert.ToString(txtdateofadmission.Value) + "' ,phoneno='" + txtphoneno.Text + "' ,bloodgroup='" + combobroup.Text + "' ,age='" + txtage.Text + "' , outpatient='" + inpa + "' ,active='" + chk + "' , username='" + Class.Users.USERID + "',    ipaddress='" + Class.Users.IPADDRESS + "',    createdby='" + Class.Users.USERID + "',    createdon='" + Convert.ToString(txtage.Text) + "',    modifiedon='" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "' ,patientbytes='" + std + "'   where asptblpatientmasid='" + txtregisterid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        if (txtregisterid.Text != "" && stdbytes != null)
                        {
                            string ins1 = "UPDATE  asptblpatientmas SET patientphoto=@patientphoto where  asptblpatientmasID='" + txtregisterid.Text + "'";
                            cmd = new MySqlCommand(ins1, Utility.Connect());
                            cmd.Parameters.Add(new MySqlParameter("@patientphoto", stdbytes));
                            cmd.ExecuteNonQuery();
                        }

                        MessageBox.Show("Record Updated Successfully " + txtpatientname.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        public void Prints()
        {
           
        }

        public void Searchs()
        {
            
        }

        public void Deletes()
        {
            if (txtregisterid.Text != "")
            {
                string del = "delete from asptblpatientmas  where asptblpatientmasid=" + txtregisterid.Text;
                Utility.ExecuteNonQuery(del);
                GridLoad();empty();
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

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }


        private void PatientMaster_Load(object sender, EventArgs e)
        {


            empty();
            combocompcode.SelectedValue = Class.Users.COMPCODE;
            combocompcode.Text = Class.Users.HCompcode;

        }
        
   
        private void empty()
        {
            txtregisterid.Text = ""; std = 0;
            txtpatientname.Text = "";           
            txtaddress.Text = "";
            txtlastname.Text = "";
            radiomale.Checked = false;
            radiofemale.Checked = false; radioothers.Checked = false;
            txtphoneno.Text = "";
            combobroup.Text = "";
           txtage.Text = "";pictureBoxphoto.Image = null;
            comboinout.SelectedIndex = 0;radiofemale.Checked = false;radiomale.Checked = false;
            butheader.Text = Class.Users.ScreenName;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            tabControl1.SelectTab(tabPage1);
            txtphoneno.Select();
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtregisterid.Text = listView1.SelectedItems[0].SubItems[2].Text;

                    string sel1 = " select a.asptblpatientmasid,b.compcode,a.patientname,A.lastname,a.address,a.gender," +
                        " a.gender,a.phoneno,a.bloodgroup,a.age,a.outpatient,a.active,a.patientphoto  from asptblpatientmas  a join gtcompmast b on a.compcode=b.gtcompmastid  where a.asptblpatientmasid=" + txtregisterid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
                    DataTable dt = ds.Tables["asptblhosdeptmas"];
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow myRow in dt.Rows)
                        {
                            txtregisterid.Text = Convert.ToString(myRow["asptblpatientmasid"].ToString());

                            txtpatientname.Text = Convert.ToString(myRow["patientname"].ToString());
                            txtlastname.Text = Convert.ToString(myRow["lastname"].ToString());
                            txtaddress.Text = Convert.ToString(myRow["address"].ToString());

                            if (myRow["gender"].ToString() == "M") { radiomale.Checked = true; } else { radiomale.Checked = false; }
                            if (myRow["gender"].ToString() == "F") { radiofemale.Checked = true; } else { radiofemale.Checked = false; }
                            if (myRow["gender"].ToString() == "O") { radioothers.Checked = true; } else { radioothers.Checked = false; }
                            txtphoneno.Text = Convert.ToString(myRow["phoneno"].ToString());      
                            if (myRow["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                            if (myRow["patientphoto"].ToString() != "")
                            {
                                pictureBoxphoto.Image = null; stdbytes = null;
                                stdbytes = (byte[])myRow["patientphoto"];
                                Image img = Models.Device.ByteArrayToImage(stdbytes);
                                pictureBoxphoto.Image = img;
                            }
                            else
                            {
                                pictureBoxphoto.Image = null;
                            }
                            txtage.Text = Convert.ToString(myRow["age"].ToString());
                            combobroup.Text = Convert.ToString(myRow["bloodgroup"].ToString());
                           
                        }
                        tabControl1.SelectTab(tabPage1);
                        txtpatientname.Select();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void combodept_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

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
                        if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[7].ToString().Contains(txtsearch.Text))
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
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
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
          
            GridLoad();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

       
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtmajordiagnosis_KeyDown(object sender, KeyEventArgs e)
        {
          
        }

       
     
      

        private void button1_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                GridLoad();
            }
            else { GridLoad(); }
        }

        private void pictureBoxphoto_Click(object sender, EventArgs e)
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

        private void txtphoneno_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void txtphoneno_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        private void txtpatientname_TextChanged(object sender, EventArgs e)
        {
            if (txtphoneno.Text != "" && txtregisterid.Text == "")
            {
                string sel1 = " select count(a.asptblpatientmasid) id1  from asptblpatientmas  a where a.patientname='" + txtpatientname.Text + "' and a.phoneno='" + txtphoneno.Text + "'";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblpatientmas");
                DataTable dt1 = ds1.Tables["asptblpatientmas"];
                Int64 cnt = Convert.ToInt64(dt1.Rows[0]["id1"].ToString());
                if (cnt >= 1)
                {
                    MessageBox.Show("This Name :  '" + txtpatientname.Text + "' Already Register", "Register");
                    txtpatientname.Text = ""; txtpatientname.Select();
                }
            }
        }

        private void txtage_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private void txtphoneno_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        public void GridLoad()
        {
            try
            {
                DataTable dt = new DataTable();
                listView1.Items.Clear(); listfilter.Items.Clear();
                if (checkBox1.Checked == true) {
                    string sel1 = " select a.asptblpatientmasid,a.patientname, a.lastname, a.gender,a.age, a.phoneno,a.address, date_format(a.modifiedon,'%d-%m-%Y') as modifiedon  from asptblpatientmas  a join gtcompmast b on a.compcode=b.gtcompmastid  order by a.asptblpatientmasid desc";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpatientmas");
                     dt = ds.Tables["asptblpatientmas"];
                }              
                else
                {
                    string sel1 = "select a.asptblpatientmasid,a.patientname, a.lastname, a.gender,a.age, a.phoneno,a.address,date_format(a.modifiedon,'%d-%m-%Y') as modifiedon  from asptblpatientmas  a join gtcompmast b on a.compcode=b.gtcompmastid  where a.modifiedon='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' order by a.asptblpatientmasid desc";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpatientmas");
                    dt = ds.Tables["asptblpatientmas"];
                }
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = "";
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblpatientmasid"].ToString());
                        list.SubItems.Add(myRow["patientname"].ToString());
                        list.SubItems.Add(myRow["lastname"].ToString());                       
                        list.SubItems.Add(myRow["gender"].ToString());
                        list.SubItems.Add(myRow["age"].ToString());
                        list.SubItems.Add(myRow["phoneno"].ToString());
                        list.SubItems.Add(myRow["address"].ToString());
                        list.SubItems.Add(myRow["modifiedon"].ToString());
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
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

        private void button1_Click_1(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void txtphoneno_Enter(object sender, EventArgs e)
        {
            txtphoneno.BackColor = Class.Users.Color2 ;
        }

        private void txtphoneno_Leave(object sender, EventArgs e)
        {
            txtphoneno.BackColor = Class.Users.Color1;
        }

        private void txtpatientname_Enter(object sender, EventArgs e)
        {
            txtpatientname.BackColor = Class.Users.Color2;
        }

        private void txtpatientname_Leave(object sender, EventArgs e)
        {
            txtpatientname.BackColor = Class.Users.Color1;
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

        private void txtage_Enter(object sender, EventArgs e)
        {
            txtage.BackColor = Class.Users.Color2;
        }

        private void txtage_Leave(object sender, EventArgs e)
        {
            txtage.BackColor = Class.Users.Color1;
        }

        private void combobroup_Enter(object sender, EventArgs e)
        {
            combobroup.BackColor = Class.Users.Color2;
        }

        private void combobroup_Leave(object sender, EventArgs e)
        {
            combobroup.BackColor = Class.Users.Color1;
        }

        private void comboinout_Leave(object sender, EventArgs e)
        {
            comboinout.BackColor = Class.Users.Color1;
        }

        private void comboinout_Enter(object sender, EventArgs e)
        {
            comboinout.BackColor = Class.Users.Color2;
        }

        private void combocompcode_Leave(object sender, EventArgs e)
        {
            combocompcode.BackColor = Class.Users.Color1;
        }

        private void combocompcode_Enter(object sender, EventArgs e)
        {
            combocompcode.BackColor = Class.Users.Color2;
        }

        private void butpatientphoto_Click(object sender, EventArgs e)
        {
            pictureBoxphoto.Image = null;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

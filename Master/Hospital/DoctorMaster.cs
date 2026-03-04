using MySql.Data.MySqlClient;
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
    public partial class DoctorMaster : Form, ToolStripAccess
    {
        private static DoctorMaster _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] digitalbytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
        public static DoctorMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DoctorMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public DoctorMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            GlobalVariables.CurrentForm = this;
     
          
        }
        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    empty();
                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        void deptload()
        {
            string sel1 = " select '0' as asptblhosdeptmasID, '' as department   from dual union  select a.asptblhosdeptmasID, a.department   from  asptblhosdeptmas a  where a.active='T'  order by 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
            DataTable dt = ds.Tables["asptblhosdeptmas"];
           
            combodept.ValueMember = "asptblhosdeptmasID";
            combodept.DisplayMember = "department";
            combodept.DataSource = dt;
        }
        private void empty()
        {
            txtdoctorid.Text = "";
            txtdoctorname.Text = "";
            txtexperience.Text = "";txtmedicalno.Text = "";
            combodept.Text = "";txtfees.Text = "";pictureBox1.Image = null; pictureBox2.Image = null;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;          
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
          
        }
        public void News()
        {
            GridLoad(); deptload(); txtdoctorname.Select();
            empty();
        }
        private bool check()
        {

            if (txtdoctorname.Text == "")
            {
                MessageBox.Show("Pls Enter Doctor Name", "Doctor");
                txtdoctorname.Select();
                return false;
            }
            if (txtmedicalno.Text == "")
            {
                MessageBox.Show("Pls Enter Medical RegisterNo", "Patient Name");

                txtmedicalno.Select();
                return false;
            }
            
            if (combodept.Text == "")
            {
                MessageBox.Show("Pls Select  Department", "Department");

                combodept.Select();
                return false;
            }
            if (txteducation.Text == "")
            {
                MessageBox.Show("Pls Enter Education", "Education");
                txteducation.Select();
                return false;

            }
            if (txtexperience.Text == "")
            {
                MessageBox.Show("Pls Enter Experience ", "Experience");
                txtexperience.Select();
            
                return false;
            }
            if (txtfees.Text == "")
            {
                MessageBox.Show("Pls Enter Field", "Fees");

                txtfees.Select();
                return false;
            }

            return true;
        }
        public void Saves()
        {
            try
            {
                if (check()==true)
                {


                    MySqlCommand cmd;
                    string chk = ""; string gen = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    if (radiomale.Checked == true) { gen = "M"; radiomale.Checked = false; } else { gen = "F"; radiomale.Checked = false; radiomale.Checked = true; }
                    string sel = "SELECT asptbldocmasid FROM asptbldocmas   WHERE doctorname='" + txtdoctorname.Text + "' AND department='" + combodept.SelectedValue + "' AND gender='" + gen + "' AND experience='" + txtexperience.Text + "' AND active='" + chk + "' AND FEES='" + txtfees.Text + "' and education='" + txteducation.Text + "' and doctorbytes='" + std + "' and signbytes='" + std1 + "' and medicalno='" + txtmedicalno.Text + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbldocmas");
                    DataTable dt = ds.Tables["asptbldocmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtdoctorid.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtdoctorid.Text) == 0 || Convert.ToInt32("0" + txtdoctorid.Text) == 0)
                    {
                        string ins = "insert into asptbldocmas(doctorname,department,gender,experience,active,FEES,education,doctorbytes,medicalno,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtdoctorname.Text.ToUpper() + "','" + combodept.SelectedValue + "','" + gen + "','" + txtexperience.Text + "','" + chk + "','" + txtfees.Text + "','" + txteducation.Text + "','" + std + "','" + txtmedicalno.Text + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "')";
                        Utility.ExecuteNonQuery(ins);
                        string sel1 = "select MAX(asptbldocmasid) ID    from  asptbldocmas   WHERE doctorname='" + txtdoctorname+ "' and department='" + combodept.Text + "'";
                        DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbldocmas");
                        DataTable dt1 = ds1.Tables["asptbldocmas"];
                        if (dt1.Rows.Count > 0 && stdbytes != null)
                        {
                            string ins1 = "UPDATE  asptbldocmas SET sign=@sign,doctorphoto=@doctorphoto where  gtcompmastid='" + dt1.Rows[0]["ID"].ToString() + "'";
                            cmd = new MySqlCommand(ins1, Utility.Connect());
                            cmd.Parameters.Add(new MySqlParameter("@doctorphoto", stdbytes));
                            cmd.Parameters.Add(new MySqlParameter("@sign", digitalbytes));
                            cmd.ExecuteNonQuery();
                        }                       
                        MessageBox.Show("Record Saved Successfully " + txtdoctorname.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptbldocmas  set   doctorname='" + txtdoctorname.Text + "' , department='" + combodept.SelectedValue + "' , gender='" + gen + "' , experience='" + txtexperience.Text + "' , active='" + chk + "'  , FEES='" + txtfees.Text + "',education='" + txteducation.Text + "',doctorbytes='" + std + "',medicalno='" + txtmedicalno.Text + "',modifiedby='" + Class.Users.USERID + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbldocmasID='" + txtdoctorid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        if (txtdoctorid.Text != "")
                        {
                            string up1 = "UPDATE  asptbldocmas SET  doctorphoto=@doctorphoto,sign=@sign where  asptbldocmasid='" + txtdoctorid.Text + "'";
                            cmd = new MySqlCommand(up1, Utility.Connect());
                            cmd.Parameters.Add(new MySqlParameter("@doctorphoto", stdbytes));
                            cmd.Parameters.Add(new MySqlParameter("@sign", digitalbytes));
                            cmd.ExecuteNonQuery();

                        }
                        MessageBox.Show("Record Updated Successfully " + txtdoctorname.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            throw new NotImplementedException();
        }

        public void Searchs()
        {
            throw new NotImplementedException();
        }

        public void Deletes()
        {
            if (txtdoctorid.Text != "")
            {
                string sel1 = "select b.asptbldocmasid from asptblpatientmas a join asptbldocmas b on b.asptbldocmasid=a.doctorname where a.asptbldocmasid='" + txtdoctorid.Text + "' and a.department='" + combodept.SelectedValue + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpatientmas");
                DataTable dt = ds.Tables["asptblpatientmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtdoctorname.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptbldocmas where asptbldocmasid='" + Convert.ToInt64("0" + txtdoctorid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtdoctorname.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
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

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtdoctorid.Text = listView1.SelectedItems[0].SubItems[1].Text;

                    string sel1 = " select a.asptbldocmasid,a.doctorname, b.department ,a.experience,a.gender, a.active,A.FEES,a.doctorphoto,a.education,a.medicalno,a.sign  from  asptbldocmas a join asptblhosdeptmas b  on a.department=b.asptblhosdeptmasid where a.asptbldocmasid=" + txtdoctorid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
                    DataTable dt = ds.Tables["asptblhosdeptmas"];
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow myRow in dt.Rows)
                        {
                            txtdoctorid.Text = Convert.ToString(myRow["asptbldocmasid"].ToString());
                            txtdoctorname.Text = Convert.ToString(myRow["doctorname"].ToString());
                            combodept.Text = Convert.ToString(myRow["department"].ToString());

                            if (myRow["gender"].ToString() == "M") { radiomale.Checked = true; } else { radiofemale.Checked = true; checkactive.Checked = false; }
                            txtexperience.Text = Convert.ToString(myRow["experience"].ToString());
                            if (myRow["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                            txtfees.Text = myRow["FEES"].ToString();
                            pictureBox1.Image = null; stdbytes = null;
                            if (myRow["doctorphoto"].ToString() != "")
                            {
                                stdbytes = (byte[])myRow["doctorphoto"];
                                Image img = Models.Device.ByteArrayToImage(stdbytes);
                                pictureBox1.Image = img;
                            }
                            pictureBox2.Image = null; digitalbytes = null;
                            if (myRow["sign"].ToString() != "")
                            {
                                digitalbytes = (byte[])myRow["sign"];
                                Image img = Models.Device.ByteArrayToImage(digitalbytes);
                                pictureBox2.Image = img;
                            }
                            txteducation.Text = myRow["education"].ToString();
                            txtmedicalno.Text = myRow["medicalno"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
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

        private void DoctorMaster_Load(object sender, EventArgs e)
        {
            News();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deptload();GridLoad();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
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

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void txtfees_TextChanged(object sender, EventArgs e)
        {

        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.asptbldocmasid,a.doctorname, b.department ,a.experience,a.gender, a.active  from  asptbldocmas a join asptblhosdeptmas b  on a.department=b.asptblhosdeptmasid order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
                DataTable dt = ds.Tables["asptblhosdeptmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["asptbldocmasid"].ToString());
                        list.SubItems.Add(myRow["doctorname"].ToString());
                        list.SubItems.Add(myRow["department"].ToString());                      
                        list.SubItems.Add(myRow["gender"].ToString());
                        list.SubItems.Add(myRow["experience"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
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

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            try
            {
                digitalbytes = null;
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {

                        p.Image = new Bitmap(open.FileName);
                        digitalbytes = Models.Device.ImageToByteArray(p);
                        System.Text.Encoding enc = System.Text.Encoding.ASCII;
                        std1 = Convert.ToInt64("0" + digitalbytes.Length);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void txtdoctorname_Leave(object sender, EventArgs e)
        {
            txtdoctorname.BackColor = Class.Users.Color1;
        }

        private void txtdoctorname_Enter(object sender, EventArgs e)
        {
            txtdoctorname.BackColor = Class.Users.Color2;
        }

        private void txtmedicalno_Enter(object sender, EventArgs e)
        {
            txtmedicalno.BackColor = Class.Users.Color2;
        }

        private void txtmedicalno_Leave(object sender, EventArgs e)
        {
            txtmedicalno.BackColor = Class.Users.Color1;
        }

        private void combodept_Leave(object sender, EventArgs e)
        {
            combodept.BackColor = Class.Users.Color1;
        }

        private void combodept_Enter(object sender, EventArgs e)
        {
            combodept.BackColor = Class.Users.Color2;
        }

        private void txteducation_Leave(object sender, EventArgs e)
        {
           txteducation.BackColor = Class.Users.Color1;
        }

        private void txteducation_Enter(object sender, EventArgs e)
        {
            txteducation.BackColor = Class.Users.Color2;
        }

        private void txtexperience_Leave(object sender, EventArgs e)
        {
            txtexperience.BackColor = Class.Users.Color1;
        }

        private void txtexperience_Enter(object sender, EventArgs e)
        {
            txtexperience.BackColor = Class.Users.Color2;
        }

        private void txtfees_Leave(object sender, EventArgs e)
        {
            txtfees.BackColor = Class.Users.Color1;
        }

        private void txtfees_Enter(object sender, EventArgs e)
        {
            txtfees.BackColor = Class.Users.Color2;
        }

        private void txtsearch_Enter(object sender, EventArgs e)
        {
            txtsearch.BackColor = Class.Users.Color2;
        }

        private void txtsearch_Leave(object sender, EventArgs e)
        {
            txtsearch.BackColor = Class.Users.Color1;
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

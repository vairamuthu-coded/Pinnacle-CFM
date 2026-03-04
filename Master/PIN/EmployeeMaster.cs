using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master.PIN
{
    public partial class EmployeeMaster : Form,ToolStripAccess
    {
        private static EmployeeMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Employee em = new Models.Employee();
        Models.UserRights sm = new Models.UserRights();
        byte[] bytes; string myString = "";
       
        OpenFileDialog open = new OpenFileDialog(); int i = 0;
         ListView listfilter = new ListView(); Byte[] Signbytes;
        public static EmployeeMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new EmployeeMaster();
                return _instance;
            }
        }
        public EmployeeMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
       
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
        private void EmployeeMaster_Load(object sender, EventArgs e)
        {

           DataTable dt1 = mas.findcomcode(Class.Users.HCompcode, Class.Users.HUserName);
         //   DataTable dt1 = mas.findcomcode(Class.Users.HCompcode);
           
            if (dt1.Rows.Count > 0)
            {


                empcompcode.DisplayMember = "COMPCODE";
                empcompcode.ValueMember = "gtcompmastid";
                empcompcode.DataSource = dt1;

                empcompname.DisplayMember = "COMPNAME";
                empcompname.ValueMember = "gtcompmastid";
                empcompname.DataSource = dt1;
            }
           
           
            DataTable dt2 = mas.dept();
            if (dt2.Rows.Count > 0)
            {
                combodept.DisplayMember = "DEPARTMENT";
                combodept.ValueMember = "gtdeptdesgmastID";
                combodept.DataSource = dt2;
                
            }
            //Utility.Load_Combo(combobatch, "select asptblbacmasid,batch from asptblbacmas  where active='T' order by 2", "asptblbacmasid","batch");
           // Utility.Load_Combo(combocategory, "select asptblcatmasid,category from asptblcatmas  where active='T' order by 2", "asptblcatmasid", "category");
           // Utility.Load_Combo(combodesignation, "select asptbldesigmasid,designation from asptbldesigmas  where active='T' order by 2", "asptbldesigmasid", "designation");
           // Utility.Load_Combo(combograde, "select asptblgramasid,grade from asptblgramas  where active='T' order by 2", "asptblgramasid", "grade");

            GridLoad(); empty();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;

        }
        public void autono()
        {


            //string sel1 = "select max(asptblempid1) as asptblempid from asptblemp;";
            //DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblemp");
            //DataTable dt = ds.Tables["asptblemp"];
            //Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["asptblempid1"].ToString());
           
            //if (cnt == 0)
            //{
            //    int s = 10001;
            //    txtempid.Text = s.ToString();

            //}
            //else
            //{
            //    txtempid1.Text = Convert.ToInt64("0" + dt.Rows[0]["asptblempid"].ToString()).ToString();
            //}

        }

            //DataTable dt2 = mas.dept();
            //if (dt2.Rows.Count > 0)
            //{
            //    combodept.DisplayMember = "DEPARTMENT";
            //    combodept.ValueMember = "asptbldeptmasid";
            //    combodept.DataSource = dt2;

            //}
            //combodept.SelectedIndex = -1;
       

      

        bool ch = false; int rowcount = 0;
        private bool Checks()
        {

            if (PictureBox1.Image == null)
            {
                MessageBox.Show("Image is empty");
                this.PictureBox1.Focus();
                this.PictureBox1.Select(); 
                return false;
            }

            if (empcompcode.Text == "")
            {
                MessageBox.Show("CompCode is empty");
                this.Focus();
                this.empcompcode.Select(); return false;
            }
          
            ch = true;
            return ch;

        }

        private bool Checks1()
        {
           
            return true;
        }
        public void Saves()
        {

            try
            {

                MySqlCommand ascmd = new MySqlCommand();
                if (empcompcode.Text != "" && txtempname.Text != "" && txtidcardno.Text != "")
                {
                    //if (txtempid.Text == "") { autono(); } else { txtempid1.Text = txtempid.Text; }
                    em.ASPTBLEMPID = Convert.ToInt64("0" + txtempid.Text);
                    em.ASPTBLEMPID1 = Convert.ToInt64("0" + txtempid.Text);
                    em.COMPCODE = Convert.ToInt64("0" + empcompcode.SelectedValue);
                    em.COMNAME = Convert.ToInt64("0" + empcompcode.SelectedValue);
                    em.EMPNAME = Convert.ToString(txtempname.Text.ToUpper());
                    em.LASTNAME = Convert.ToString(txtlastname.Text.ToUpper());
                    em.ADDRESS = Convert.ToString(txtaddress.Text.ToUpper());
                    if (radiomale.Checked == true) { em.GENDER = "T"; } else { radiofemale.Checked = false; em.GENDER = "F"; }
                    em.EMPLOYEETYPE = Convert.ToString(combocategory.Text);
                    em.DATEOFBIRTH = Convert.ToString(txtdateofbirth.Text);
                    em.DEPARTMENT = Convert.ToInt64("0" + combodept.SelectedValue);
                    em.DATEOFJOIN = Convert.ToString(txtdateofjoin.Text);
                    em.IDCARDNO = Convert.ToInt64("0" + txtidcardno.Text);
                    em.CONTACT = Convert.ToString(txtcontactno.Text);
                    em.BLOODGROUP = Convert.ToString(combobroup.Text);
                    if (checkactive.Checked) em.ACTIVE = "T"; else em.ACTIVE = "F";
                    em.USERNAME = Convert.ToInt64("0" + Class.Users.USERID);
                    em.IPADDRESS = Class.Users.IPADDRESS;
                    em.CREATEDON = Convert.ToString(Class.Users.CREATED);
                    em.MODIFIEDON = Convert.ToString(Class.Users.CREATED);
                    if (Signbytes==null) { em.imagebyte = 0;  }
                    else
                    {
                        em.imagebyte = Convert.ToInt64("0" + Signbytes.Length);
                    }
                    em.SALARY = Convert.ToInt64("0" + txtsalary.Text);
                    MySqlCommand cmd;
                    DataTable dt = em.select(em.COMPCODE, em.EMPNAME, em.LASTNAME, em.ADDRESS, em.GENDER, em.DATEOFBIRTH, em.DEPARTMENT, em.DATEOFJOIN, em.IDCARDNO, em.CONTACT, em.BLOODGROUP, em.ACTIVE, em.USERNAME, em.EMPLOYEETYPE,em.SALARY,em.imagebyte);
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found", "Exception");
                    }
                    else if (dt.Rows.Count == 0 && em.ASPTBLEMPID == 0)
                    {
                        string ins = "insert into asptblemp(asptblempid1,compcode,comname,empname,lastname,address,gender,dateofbirth,department,dateofjoin,idcardno,contact,bloodgroup,active,username,ipaddress,createdby," +
                            "createdon,modifiedon,employeetype,salary)VALUES('" + em.ASPTBLEMPID1 + "'," + em.COMPCODE + "," + em.COMNAME + ",'" + em.EMPNAME + "','" + em.LASTNAME + "','" + em.ADDRESS + "','" + em.GENDER + "','" + em.DATEOFBIRTH + "'," + em.DEPARTMENT + ",'" + em.DATEOFJOIN + "'," + em.IDCARDNO + ",'" + em.CONTACT + "','" + em.BLOODGROUP + "','" + em.ACTIVE + "'," + em.USERNAME + ",'" + em.IPADDRESS + "','" + em.CREATEDON + "','" + em.CREATEDON + "','" + em.MODIFIEDON + "','" + em.EMPLOYEETYPE + "','" + em.SALARY + "')";
                              Utility.ExecuteNonQuery(ins);
                        string sel1 = "select  max(asptblempid) as asptblempid from asptblemp";
                        DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblemp");
                        DataTable dt1 = ds1.Tables["asptblemp"];
                        if (dt1 != null &&  PictureBox1.Image != null)
                        {
                            string ins1 = "update   asptblemp set EMPIMAGE=@EMPIMAGE,imagebyte='" + em.imagebyte + "' where asptblempid='" + dt1.Rows[0]["asptblempid"].ToString() + "'  and idcardno='" + txtidcardno.Text + "' ";
                            cmd = new MySqlCommand(ins1, Utility.Connect());
                            cmd.Parameters.Add(new MySqlParameter("@EMPIMAGE", Signbytes));
                            cmd.ExecuteNonQuery();
                        }

                        MessageBox.Show("Record Saved Saved Successfully", "Success"); GridLoad(); empty();
                    }
                    else
                    { 
                        string query = "update   asptblemp set  asptblempid1='"+em.ASPTBLEMPID1+"' ,  compcode=" + em.COMPCODE + ",comname=" + em.COMNAME + ",empname='" + em.EMPNAME + "',lastname='" + em.LASTNAME + "' ,address='" + em.ADDRESS + "',gender='" + em.GENDER + "',dateofbirth='" + em.DATEOFBIRTH + "',department=" + em.DEPARTMENT + ",dateofjoin='" + em.DATEOFJOIN + "' ,idcardno=" + em.IDCARDNO + ",contact='" + em.CONTACT + "',bloodgroup='" + em.BLOODGROUP + "',active='" + em.ACTIVE + "' ,username=" + em.USERNAME + ",ipaddress='" + em.IPADDRESS + "',createdby='" + em.IDCARDNO + "',modifiedon='" + em.MODIFIEDON + "',employeetype='" + em.EMPLOYEETYPE + "', salary='" + em.SALARY + "',EMPIMAGE=@EMPIMAGE,imagebyte='" + em.imagebyte+"' where asptblempid='" + em.ASPTBLEMPID + "'";
                        cmd = new MySqlCommand(query, Utility.Connect());
                        cmd.Parameters.Add(new MySqlParameter("@EMPIMAGE", Signbytes));
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Updated  Successfully", "Success"); GridLoad(); empty();
                    }

                }

                else
                {
                    MessageBox.Show("pls Enter Mandatary Fields");

                }
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());
            }
        }

        private void EmployeeMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void News()
        {
            empty();GridLoad();
        }


        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }


        private void Empcompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        private void Lvlogs_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (lvlogs.Items.Count > 0)
                {
                    Cursor = Cursors.WaitCursor;
                    PictureBox1.Image = null; bytes = null;
                    em.ASPTBLEMPID = Convert.ToInt64("0" + lvlogs.SelectedItems[0].SubItems[2].Text);
                    string sel1 = "select a.asptblempid,b.compcode ,b.compname,a.empname,a.lastname ,a.address,a.gender,a.dateofbirth,a.idcardno,a.department,a.dateofjoin ,a.contact,a.bloodgroup,a.active,a.employeetype,a.salary,a.EMPIMAGE from  asptblemp a join   gtcompmast b on b.gtcompmastid = a.compcode  where a.asptblempid=" + em.ASPTBLEMPID;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblemp");
                    DataTable dt = ds.Tables["asptblemp"];
                    foreach (DataRow myRow in dt.Rows)
                    {
                        txtempid.Text = Convert.ToString(myRow["asptblempid"].ToString());    
                        empcompcode.Text = Convert.ToString(myRow["COMPCODE"].ToString());
                        empcompname.Text = Convert.ToString(myRow["COMPNAME"].ToString());
                        txtempname.Text = Convert.ToString(myRow["EMPNAME"].ToString());
                        txtlastname.Text = Convert.ToString(myRow["LASTNAME"].ToString());
                        txtaddress.Text = Convert.ToString(myRow["ADDRESS"].ToString());
                        if (myRow["GENDER"].ToString() == "T") radiomale.Checked = true;
                        else radiofemale.Checked = true;
                        txtdateofbirth.Text = Convert.ToString(myRow["DATEOFBIRTH"].ToString());
                        combodept.SelectedValue = Convert.ToInt64(myRow["DEPARTMENT"].ToString());
                        txtdateofjoin.Text = Convert.ToString(myRow["DATEOFJOIN"].ToString());
                        txtidcardno.Text = Convert.ToString(myRow["IDCARDNO"].ToString());
                        txtcontactno.Text = Convert.ToString(myRow["CONTACT"].ToString());

                        combobroup.Text = Convert.ToString(myRow["BLOODGROUP"].ToString());
                        txtsalary.Text = Convert.ToString(myRow["salary"].ToString());
                        if (myRow["ACTIVE"].ToString() == "T") checkactive.Checked = true; else checkactive.Checked = false;
                        if (myRow["EMPIMAGE"].ToString() != "")
                        {

                            bytes = (byte[])myRow["EMPIMAGE"];
                            Image img = Models.Device.ByteArrayToImage(bytes);
                            PictureBox1.Image = img;


                        }
                       



                        combocategory.Text = Convert.ToString(myRow["EMPLOYEETYPE"].ToString());

                        Cursor = Cursors.Default;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            tabControl1.SelectTab(tabPage1);
        }
        //private void butview_Click(object sender, EventArgs e)
        //{

        //    //try
        //    //{
        //    //    lvlogs.Items.Clear(); int r = 1; listfilter.Items.Clear();
        //    //    string sel1 = "SELECT A.ASPTBLEMPID,B.COMPCODE,A.EMPNAME,A.IDCARDNO ,A.GENDER,C.MNNAME1 AS DEPARTMENT,A.CONTACT,A.BLOODGROUP,A.EMPLOYEETYPE,A.ACTIVE   FROM  ASPTBLEMP A   JOIN   GTCOMPMAST B ON B.gtcompmastid = A.COMPCODE JOIN gtdeptdesgmast C ON C.asptbldeptmasid = A.DEPARTMENT   order by 1;";
        //    //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLEMP");
        //    //    DataTable dt = ds.Tables["ASPTBLEMP"];
        //    //    if (dt.Rows.Count > 0)
        //    //    {
        //    //        foreach (DataRow myRow in dt.Rows)
        //    //        {
        //    //            ListViewItem list = new ListViewItem();
        //    //            list.Text = r.ToString();
        //    //            list.SubItems.Add(myRow["ASPTBLEMPID"].ToString());
        //    //            list.SubItems.Add(myRow["COMPCODE"].ToString());
        //    //            list.SubItems.Add(myRow["EMPNAME"].ToString());
        //    //            list.SubItems.Add(myRow["IDCARDNO"].ToString());
        //    //            if (myRow["GENDER"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");

        //    //            list.SubItems.Add(myRow["DEPARTMENT"].ToString());

        //    //            list.SubItems.Add(myRow["CONTACT"].ToString());
        //    //            list.SubItems.Add(myRow["BLOODGROUP"].ToString());
        //    //            list.SubItems.Add(myRow["EMPLOYEETYPE"].ToString());

        //    //            if (myRow["ACTIVE"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");
        //    //            if (r % 2 == 0)
        //    //            {
        //    //                list.BackColor = Color.White;
        //    //            }
        //    //            else
        //    //            {
        //    //                list.BackColor = Color.WhiteSmoke;
        //    //            }
        //    //            r++;
        //    //            this.listfilter.Items.Add((ListViewItem)list.Clone());
        //    //            lvlogs.Items.Add(list);
        //    //        }

        //    //        lblemptot.Text = "Total Rows    :" + lvlogs.Items.Count;
        //    //    }
        //    //    else
        //    //    {
        //    //        lblemptot.Text = "Total Rows  : 0";
        //    //    }
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    //MessageBox.Show(ex.ToString());
        //    //}

        //}
        public void GridLoad()
        {
            try
            {
                lvlogs.Items.Clear();
                int i = 1;



                string sel1 = "select a.asptblempid,a.empname,a.idcardno , a.active from  asptblemp a join   gtcompmast b on b.gtcompmastid = a.compcode  where b.compcode='" + Class.Users.HCompcode + "' order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblemp");
                DataTable dt = ds.Tables["asptblemp"];
                if (dt != null)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblempid"].ToString());
                        list.SubItems.Add(myRow["EMPNAME"].ToString());
                        list.SubItems.Add(myRow["IDCARDNO"].ToString());
                        if (myRow["ACTIVE"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");

                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;



                        this.listfilter.Items.Add((ListViewItem)list.Clone());

                        lvlogs.Items.Add(list);
                        i++;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void empty()
        {
            txtempid.Text = "";

            empcompcode.Text = "";
            empcompname.Text = "";
            txtempname.Text = "";
            txtlastname.Text = "";
            txtaddress.Text = "";
            radiomale.Checked = true;
            txtdateofbirth.Text = "";
            combodept.Text = "";
            txtdateofjoin.Text = "";
            txtidcardno.Text = "";
            txtcontactno.Text = "";
            combocategory.Text = "";
            checkactive.Checked = true;
            bytes = null;txtsalary.Text = "";
            PictureBox1.Image = null;
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            lvlogs.Font= Class.Users.FontName;
        }

        
        private void Butimageclear_Click(object sender, EventArgs e)
        {
            PictureBox1.Image = null; bytes = null;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Hide();

        }

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EmployeeMaster_Load(sender, e);
        }

        public void Deletes()
        {

            if (txtempid.Text != "")
            {
                string sel1 = "delete  from asptblemp A  WHERE A.asptblempid='" + txtempid.Text + "';";
                Utility.ExecuteNonQuery(sel1);  MessageBox.Show("Record Deleted");
                empty(); GridLoad();
            }
         

        }

     

    
        private void lblemptot_Click(object sender, EventArgs e)
        {

        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            EmployeeMaster_Load(sender, e);
        }

       





     
        //privatevoidDeleteRowFromGrid(int id)
        //{
        //    try
        //    {
        //        conx.Open(); //Open Connection  
        //        SqlCommandcmd = newSqlCommand("DELETE FROM Customers WHERE CustomerID = '" + id + "'", conx); //Here, we specify query with {Id} parameter which allow us to delete rows from Db.  
        //        int Result = cmd.ExecuteNonQuery(); // Execute Query for deleting all rows selected from DataGridView  
        //        conx.Close(); // Close Connection  
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    finally
        //    {
        //        conx.Close();
        //    }
        //}
        //Delete Button  
        private void button1_Click(object sender, EventArgs e)
        {

        }
        

        private void dataGridView1_MouseHover(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {

        }

       



        public void Prints()
        {

        }

        public void Searchs()
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            try
            {
                bytes = null;
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {

                        p.Image = new Bitmap(open.FileName);
                        Signbytes = Models.Device.ImageToByteArray(p);


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        public void DownLoads()
        {
            throw new NotImplementedException();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txtsearch.Text.Length > 0)
                {
                    lvlogs.Items.Clear();

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text))
                        {

                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);

                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;




                            lvlogs.Items.Add(list);
                        }

                        item0++;
                    }
                    lblemptot.Text = "Total Rows    :" + lvlogs.Items.Count;
                }
                else
                {
                    try
                    {
                        lvlogs.Items.Clear();
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            

                                list.SubItems.Add(item.SubItems[1].Text);
                                list.SubItems.Add(item.SubItems[2].Text);
                                list.SubItems.Add(item.SubItems[3].Text);
                                list.SubItems.Add(item.SubItems[4].Text);

                                list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;




                                lvlogs.Items.Add(list);
                     

                            item0++;
                        }
                        lblemptot.Text = "Total Rows    :" + lvlogs.Items.Count;
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

       

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

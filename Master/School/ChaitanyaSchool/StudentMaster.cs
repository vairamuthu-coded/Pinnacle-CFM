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

namespace Pinnacle.Master.School.ChaitanyaSchool
{
    public partial class StudentMaster : Form,ToolStripAccess
    {
        private static StudentMaster _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std,std1=0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
         int i = 0;
        public static StudentMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new StudentMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public StudentMaster()
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

                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        private void StudentMaster_Load(object sender, EventArgs e)
        {
            GlobalVariables.CurrentForm = this;
            string sel11 = "select a.asptblschoolmasid,a.compcode,a.compname from asptblschoolmas a where a.active='T'   ;";
            DataSet ds11 = Utility.ExecuteSelectQuery(sel11, "asptblschoolmas");
            DataTable dt11 = ds11.Tables["asptblschoolmas"];
            if (dt11.Rows.Count > 0)
            {


                comboschoolcode.DisplayMember = "COMPCODE";
                comboschoolcode.ValueMember = "asptblschoolmasid";
                comboschoolcode.DataSource = dt11;

             
            }


            string sel0 = "select a.asptblstandardmasID,a.standard from asptblstandardmas a where a.active='T'   ;";
            DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "asptblstandardmas");
            DataTable dt0 = ds0.Tables["asptblstandardmas"];
            if (dt0.Rows.Count > 0)
            {
                combostandard.DisplayMember = "standard";
                combostandard.ValueMember = "asptblstandardmasID";
                combostandard.DataSource = dt0;

            }
            string sel1 = "select a.asptblsectionmasid,a.section from asptblsectionmas a where a.active='T'   ;";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblsectionmas");
            DataTable dt1 = ds1.Tables["asptblsectionmas"];
            if (dt1.Rows.Count > 0)
            {
                combosection.DisplayMember = "section";
                combosection.ValueMember = "asptblsectionmasid";
                combosection.DataSource = dt1;

            }

            string sel2 = "select a.asptblblockmasid,a.blockno from asptblblockmas a where a.active='T'   ;";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblblockmas");
            DataTable dt2 = ds2.Tables["asptblblockmas"];
            if (dt2.Rows.Count > 0)
            {
                comboblock.DisplayMember = "blockno";
                comboblock.ValueMember = "asptblblockmasid";
                comboblock.DataSource = dt2;

            }

            string sel3 = "select A.asptblelectionmasID,A.POSTNAME from asptblelectionmas a where a.active='T';";
            DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblelectionmas");
            DataTable dt3 = ds3.Tables["asptblelectionmas"];
            if (dt0.Rows.Count > 0)
            {
                comboelectionpost.DisplayMember = "POSTNAME";
                comboelectionpost.ValueMember = "asptblelectionmasID";
                comboelectionpost.DataSource = dt3;

            }
            GridLoad(); empty();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;

        }
        private void comboschoolcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sel11 = "select a.asptblschoolmasid,a.compname from asptblschoolmas a where a.active='T' AND A.COMPCODE='"+comboschoolcode.Text+"'   ;";
            DataSet ds11 = Utility.ExecuteSelectQuery(sel11, "asptblschoolmas");
            DataTable dt11 = ds11.Tables["asptblschoolmas"];
            if (dt11.Rows.Count > 0)
            {
                comboschoolname.DisplayMember = "compname";
                comboschoolname.ValueMember = "asptblschoolmasid";
                comboschoolname.DataSource = dt11;
            }
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
        //    combodept.ValueMember = "gtdeptdesgmastid";
        //    combodept.DataSource = dt2;

        //}
        //combodept.SelectedIndex = -1;




        bool ch = false; int rowcount = 0;
        private bool Checks()
        {

            if (PictureBox1.Image == null)
            {
                MessageBox.Show("Image is empty", "Photo");
                this.PictureBox1.Focus();
                this.PictureBox1.Select(); return false;
            }

            if (comboschoolcode.Text == "")
            {
                MessageBox.Show("CompCode is empty", "CompCode");
                this.Focus();
                this.comboschoolcode.Select(); return false;
            }
            if (txtstudentname.Text == "")
            {
                MessageBox.Show("StudentName is empty", "Student Name");
                this.Focus();
                this.txtstudentname.Select(); return false;
            }
            if (txtrollno.Text == "")
            {
                MessageBox.Show("RollNo is empty", "RollNo");
                this.Focus();
                this.txtrollno.Select(); return false;
            }
            if (combosection.Text == "")
            {
                MessageBox.Show("Section is empty", "Section");
                this.Focus();
                this.combosection.Select(); return false;
            }
            if (combostandard.Text == "" || combostandard.Text == null)
            {
                MessageBox.Show("Grade is empty","Grade/Standard");
                this.Focus();
                this.combostandard.Select(); return false;
            }
            if (comboblock.Text == "" || comboblock.Text == null)
            {
                MessageBox.Show("Block is empty", "Block");
                this.Focus();
                this.comboblock.Select(); return false;
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
                if (Checks())
                {
                    MySqlCommand cmd;

                    em.asptblstudentmasid = Convert.ToInt64("0" + txtstudentid.Text);
                    em.asptblstudentmasid1 = Convert.ToInt64("0" + txtstudentid.Text);
                    em.compcode = Convert.ToInt64("0" + comboschoolcode.SelectedValue);
                    em.compname = Convert.ToInt64("0" + comboschoolcode.SelectedValue);
                    em.studentname = Convert.ToString(txtstudentname.Text.ToUpper());
                    em.lastname = Convert.ToString(txtlastname.Text.ToUpper());
                    em.address = Convert.ToString(txtaddress.Text.ToUpper());
                    if (radiomale.Checked == true) { em.gender = "M"; } else { radiofemale.Checked = false; em.gender = "F"; }
                    em.dateofbirth = Convert.ToString(txtdateofbirth.Value.Date.ToString("yyyy-MM-dd"));
                    em.ROLLNO = Convert.ToString(txtrollno.Text);
                    em.dateofjoin = Convert.ToString(txtdateofjoin.Value.Date.ToString("yyyy-MM-dd"));
                    em.standard = Convert.ToInt64("0" + combostandard.SelectedValue);
                    em.SECTION = Convert.ToInt64(combosection.SelectedValue);
                    em.BLOCKNO = Convert.ToInt64("0" + comboblock.SelectedValue);
                    em.bloodgroup = Convert.ToString(combobroup.Text);
                    if (checkactive.Checked) em.active = "T"; else em.active = "F";
                    em.contact = Convert.ToString(txtcontactno.Text);
                    em.username = Convert.ToInt64("0" + Class.Users.USERID);
                    em.ipaddress = Class.Users.IPADDRESS;
                    em.createdon = Convert.ToString(System.DateTime.Now.ToString());
                    em.createdby = Convert.ToString(Class.Users.CREATED);
                    em.modifiedon = Convert.ToString(System.DateTime.Now.ToString());

                    if (checkelection.Checked == true) { em.election = "T"; } else { checkelection.Checked = false; em.election = "F"; }

                    em.STUDENTIMAGE = stdbytes;
                    em.STUDENTVOTELOG = votebytes;
                    em.STUDENTIMAGEBYTES = std;
                    em.STUDENTVOTEBYTES = std1;
                    em.electionpost = Convert.ToInt64(comboelectionpost.SelectedValue);
                    em.electiondate = Convert.ToString(dateTimePickerelectiondate.Value.Date.ToString("yyyy-MM-dd"));

                    string sel = "select asptblstudentmasid    from  asptblstudentmas    WHERE  compcode='" + em.compcode + "' and compname='" + em.compname + "' and studentname='" + em.studentname + "' and lastname='" + em.lastname + "' and address='" + em.address + "' and gender='" + em.gender + "' and  ROLLNO='" + em.ROLLNO + "' AND dateofbirth=DATE_FORMAT('" + em.dateofbirth.ToString().Substring(0, 10) + "','%y-%m-%d')  and dateofjoin=DATE_FORMAT('" + em.dateofjoin.ToString().Substring(0, 10) + "','%y-%m-%d') and standard='" + em.standard + "' and bloodgroup='" + em.bloodgroup + "' and SECTION='" + em.SECTION + "'  and contact='" + em.contact + "' and BLOCKNO='" + em.BLOCKNO + "' and active='" + em.active + "' and STUDENTIMAGEBYTES='" + em.STUDENTIMAGEBYTES + "' and STUDENTVOTEBYTES='" + em.STUDENTVOTEBYTES + "' and election='" + em.election + "' and  electionpost='" + em.electionpost + "' and electiondate='" + em.electiondate + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblstudentmas");
                    DataTable dt = ds.Tables["asptblstudentmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found", "Exception");
                    }
                    else if (dt.Rows.Count == 0 && em.asptblstudentmasid == 0)
                    {
                        string ins = "insert into asptblstudentmas(compcode,compname,studentname,lastname,address,gender,rollno,dateofbirth,dateofjoin,standard,SECTION,BLOCKNO,bloodgroup,contact,active,STUDENTIMAGEBYTES,STUDENTVOTEBYTES,username,ipaddress,createdby,createdon,modifiedon,election,STUDENTIMAGE,STUDENTVOTELOG,electionpost,electiondate)VALUES('" + em.compcode + "','" + em.compname + "','" + em.studentname + "','" + em.lastname + "','" + em.address + "','" + em.gender + "','" + em.ROLLNO + "', DATE_FORMAT('" + em.dateofbirth.ToString().Substring(0, 10) + "','%y-%m-%d')  , DATE_FORMAT('" + em.dateofjoin.ToString().Substring(0, 10) + "','%y-%m-%d'),'" + em.standard + "',  '" + em.SECTION + "','" + em.BLOCKNO + "','" + em.bloodgroup + "','" + em.contact + "','" + em.active + "','" + em.STUDENTIMAGEBYTES + "','" + em.STUDENTVOTEBYTES + "','" + em.username + "','" + em.ipaddress + "','" + em.username + "','" + em.createdon + "','" + em.modifiedon + "','" + em.election + "',@STUDENTIMAGE,@STUDENTVOTELOG,'" + em.electionpost + "','" + em.electiondate + "')";//@STUDENTVOTELOG,@STUDENTIMAGEBYTES,
                        cmd = new MySqlCommand(ins, Utility.con);
                        cmd.Parameters.Add(new MySqlParameter("@STUDENTIMAGE", stdbytes));
                        cmd.Parameters.Add(new MySqlParameter("@STUDENTVOTELOG", votebytes));
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Saved Saved Successfully", "Success"); GridLoad(); empty();
                    }
                    else
                    {
                        string query = "update   asptblstudentmas set  compcode='" + em.compcode + "',compname='" + em.compname + "',studentname='" + em.studentname + "',lastname='" + em.lastname + "',address='" + em.address + "',gender='" + em.gender + "' ,rollno='" + em.ROLLNO + "',dateofbirth=DATE_FORMAT('" + em.dateofbirth.ToString().Substring(0, 10) + "','%y-%m-%d')  , dateofjoin=DATE_FORMAT('" + em.dateofjoin.ToString().Substring(0, 10) + "','%y-%m-%d'),standard='" + em.standard + "',bloodgroup='" + em.bloodgroup + "',SECTION='" + em.SECTION + "',contact='" + em.contact + "',BLOCKNO='" + em.BLOCKNO + "',STUDENTIMAGE=@STUDENTIMAGE,STUDENTVOTELOG=@STUDENTVOTELOG,active='" + em.active + "',STUDENTIMAGEBYTES='" + em.STUDENTIMAGEBYTES + "',STUDENTVOTEBYTES='" + em.STUDENTVOTEBYTES + "',username='" + em.username + "',ipaddress='" + em.ipaddress + "',modifiedon='" + em.modifiedon + "',election='" + em.election + "'  ,  electionpost='" + em.electionpost + "' , electiondate='" + em.electiondate + "' where asptblstudentmasid=" + em.asptblstudentmasid;
                        cmd = new MySqlCommand(query, Utility.con);
                        cmd.Parameters.Add(new MySqlParameter("@STUDENTIMAGE", stdbytes));
                        cmd.Parameters.Add(new MySqlParameter("@STUDENTVOTELOG", votebytes));
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Record Updated  Successfully", "Success"); GridLoad(); empty();
                    }

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
            empty(); GridLoad();
        }


        private void PictureBox1_Click(object sender, EventArgs e)
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
                    PictureBox1.Image = null; stdbytes = null;
                    PictureBox2.Image = null; votebytes = null;
                    em.asptblstudentmasid = Convert.ToInt64("0" + lvlogs.SelectedItems[0].SubItems[1].Text);
                    string sel1 = "select A.ASPTBLSTUDENTMASID, B.compcode,B.compname,A.studentname,A.lastname,A.address,A.gender,A.rollno,A.dateofbirth,A.dateofjoin, C.standard,D.SECTION,E.BLOCKNO,A.bloodgroup,A.contact,A.active,A.STUDENTIMAGE,A.STUDENTVOTELOG,a.election,F.postname,A.ELECTIONDATE from  ASPTBLSTUDENTMAS a join   asptblschoolmas b on b.asptblschoolmasID = a.compcode join asptblstandardmas C ON C.asptblstandardmasID=A.STANDARD   JOIN asptblsectionmas  D ON D.asptblsectionmasid=A.SECTION JOIN asptblblockmas E ON E.asptblblockmasID=A.BLOCKNO JOIN asptblelectionmas F ON F.ASPTBLELECTIONMASID=A.ELECTIONPOST  where a.ASPTBLSTUDENTMASID=" + em.asptblstudentmasid;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
                    DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
                    foreach (DataRow myRow in dt.Rows)
                    {
                        txtstudentid.Text = Convert.ToString(myRow["ASPTBLSTUDENTMASID"].ToString());                      
                        comboschoolcode.Text = Convert.ToString(myRow["COMPCODE"].ToString());
                        comboschoolname.Text = Convert.ToString(myRow["COMPNAME"].ToString());
                        txtstudentname.Text = Convert.ToString(myRow["studentname"].ToString());
                        txtlastname.Text = Convert.ToString(myRow["LASTNAME"].ToString());
                        txtaddress.Text = Convert.ToString(myRow["ADDRESS"].ToString());
                        if (myRow["GENDER"].ToString() == "M") radiomale.Checked = true;
                        else radiofemale.Checked = true;
                        txtrollno.Text = Convert.ToString(myRow["rollno"].ToString());
                        txtdateofbirth.Text = Convert.ToString(myRow["DATEOFBIRTH"].ToString());
                        txtdateofjoin.Text = Convert.ToString(myRow["DATEOFJOIN"].ToString());
                        combostandard.Text = Convert.ToString(myRow["standard"].ToString());
                        combosection.Text = Convert.ToString(myRow["SECTION"].ToString());
                        comboblock.Text = Convert.ToString(myRow["BLOCKNO"].ToString());
                        combobroup.Text = Convert.ToString(myRow["bloodgroup"].ToString());
                        txtcontactno.Text = Convert.ToString(myRow["contact"].ToString());
                        if (myRow["ELECTION"].ToString() == "T") checkelection.Checked = true;
                        else checkelection.Checked = true;
                        comboelectionpost.Text = Convert.ToString(myRow["postname"].ToString());
                        dateTimePickerelectiondate.Value = Convert.ToDateTime(myRow["ELECTIONDATE"].ToString());
                        if (myRow["ACTIVE"].ToString() == "T") checkactive.Checked = true; else checkactive.Checked = false;
                        if (myRow["STUDENTIMAGE"].ToString() != "")
                        {

                            stdbytes = (byte[])myRow["STUDENTIMAGE"];
                            Image img = Models.Device.ByteArrayToImage(stdbytes);
                            PictureBox1.Image = img;


                        }
                        if (myRow["STUDENTVOTELOG"].ToString() != "")
                        {

                            votebytes = (byte[])myRow["STUDENTVOTELOG"];
                            Image img = Models.Device.ByteArrayToImage(votebytes);
                            PictureBox2.Image = img;


                        }
                       



                        Cursor = Cursors.Default;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            comboschoolcode.Select();
        }
        //private void butview_Click(object sender, EventArgs e)
        //{

        //    //try
        //    //{
        //    //    lvlogs.Items.Clear(); int r = 1; listfilter.Items.Clear();
        //    //    string sel1 = "SELECT A.ASPTBLEMPID,B.COMPCODE,A.EMPNAME,A.IDCARDNO ,A.GENDER,C.MNNAME1 AS DEPARTMENT,A.CONTACT,A.BLOODGROUP,A.EMPLOYEETYPE,A.ACTIVE   FROM  ASPTBLEMP A   JOIN   GTCOMPMAST B ON B.gtcompmastid = A.COMPCODE JOIN gtdeptdesgmast C ON C.gtdeptdesgmastid = A.DEPARTMENT   order by 1;";
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
        private void txtrollno_TextChanged(object sender, EventArgs e)
        {


            

        }
        public void GridLoad()
        {
            try
            {
                lvlogs.Items.Clear();
                int i = 1;


                string sel1 = "select A.ASPTBLSTUDENTMASID, B.compcode,B.compname,A.studentname,A.lastname,A.address,A.gender,A.rollno,A.dateofbirth,A.dateofjoin, C.standard,D.SECTION,E.BLOCKNO,A.bloodgroup,A.contact,A.active,A.STUDENTIMAGE,A.STUDENTVOTELOG,F.POSTNAME from  ASPTBLSTUDENTMAS a join   asptblschoolmas b on b.asptblschoolmasID = a.compcode join asptblstandardmas C ON C.asptblstandardmasID=A.STANDARD   JOIN asptblsectionmas  D ON D.asptblsectionmasid=A.SECTION JOIN asptblblockmas E ON E.asptblblockmasID=A.BLOCKNO   JOIN asptblelectionmas F ON F.ASPTBLELECTIONMASID=A.ELECTIONPOST where a.active='T' order by a.ASPTBLSTUDENTMASID desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
                DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
                foreach (DataRow myRow in dt.Rows)
                {
                    ListViewItem list = new ListViewItem();
                    list.Text = i.ToString();
                    list.SubItems.Add(myRow["ASPTBLSTUDENTMASID"].ToString());
                    list.SubItems.Add(myRow["studentname"].ToString());
                    list.SubItems.Add(myRow["rollno"].ToString());
                    list.SubItems.Add(myRow["POSTNAME"].ToString());
                    if (myRow["ACTIVE"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");
                    list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                    this.listfilter.Items.Add((ListViewItem)list.Clone());                   

                    if (myRow["ACTIVE"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");
                    lvlogs.Items.Add(list);
                    i++;

                }
                lbltotal.Refresh(); lbltotal.Text = "Total :" + lvlogs.Items.Count.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void empty()
        {
            txtstudentid.Text = ""; checkelection.Checked = false;
            txtstudentid1.Text = "";
            // empcompcode.Text = "";
            // empcompname.Text = "";
            txtstudentname.Text = "";
            txtlastname.Text = "";
            txtaddress.Text = "";
            radiomale.Checked = true;
            txtdateofbirth.Text = "";
            combosection.Text = "";
            txtdateofjoin.Text = "";
            txtrollno.Text = "";
            txtcontactno.Text = "";
            comboblock.Text = "";
            checkactive.Checked = true;
            stdbytes = null; votebytes = null;
            PictureBox1.Image = null; PictureBox2.Image = null;
            lblelectionpost.Visible = false;
            comboelectionpost.Visible = false;
            lblelectiondate.Visible = false;
            dateTimePickerelectiondate.Visible = false;
            lblsearch.ForeColor = Class.Users.Color1;
            lbltotal.ForeColor = Class.Users.Color1;
            butheader.BackColor = Class.Users.BackColors;
            panel1.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            lvlogs.Font = Class.Users.FontName;
            comboschoolcode.Focus();
        }


       

  

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        public void Deletes()
        {

            if (txtstudentid.Text != "")
            {
                string sel1 = "delete  from ASPTBLSTUDENTMAS A  WHERE A.ASPTBLSTUDENTMASID='" + txtstudentid.Text + "';";
                Utility.ExecuteNonQuery(sel1); MessageBox.Show("Record Deleted");
                empty(); GridLoad();
            }


        }




        private void lblemptot_Click(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            
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
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

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
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;


                            lvlogs.Items.Add(list);
                        }

                        item0++;
                       
                    }
                    lbltotal.Refresh();
                    lbltotal.Text = "Total Rows    :" + lvlogs.Items.Count;
                }
                else
                {
                    try
                    {
                        GridLoad();
                        lbltotal.Refresh();
                        lbltotal.Text = "Total Rows    :" + lvlogs.Items.Count.ToString();
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

        private void butimageclear_Click_1(object sender, EventArgs e)
        {

        }

        private void checkelection_CheckedChanged(object sender, EventArgs e)
        {
            if (checkelection.Checked == true)
            {
                lblelectionpost.Visible = true;
                comboelectionpost.Visible = true;
                lblelectiondate.Visible = true;
                dateTimePickerelectiondate.Visible = true;
            }
            else
            {
                lblelectionpost.Visible = false;
                comboelectionpost.Visible = false;
                lblelectiondate.Visible = false;
                dateTimePickerelectiondate.Visible = false;
            }
        }

        private void txtrollno_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void txtrollno_KeyDown(object sender, KeyEventArgs e)
        {
           
        }

        private void txtrollno_MouseLeave(object sender, EventArgs e)
        {

        }

        private void txtrollno_KeyUp(object sender, KeyEventArgs e)
        {
            string sel1 = "select rollno from  ASPTBLSTUDENTMAS where rollno='" + txtrollno.Text + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
            DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
            if (dt.Rows.Count >= 1)
            {
                txtrollno.BackColor = Color.Red;
                txtrollno.ForeColor = Color.White;
                txtrollno.Select(); 
            }
            else
            {
                txtrollno.BackColor = Color.White;
                txtrollno.ForeColor = Color.Black;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            try
            {
                votebytes = null;
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {

                        p.Image = new Bitmap(open.FileName);
                        votebytes = Models.Device.ImageToByteArray(p);
                        System.Text.Encoding enc = System.Text.Encoding.ASCII;
                        std1 = Convert.ToInt64("0" + votebytes.Length);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            em.asptblstudentmasid = Convert.ToInt64("0" + lvlogs.SelectedItems[0].SubItems[1].Text);

            string sel1 = "select A.ASPTBLSTUDENTMASID, B.compcode,B.compname,A.studentname,A.lastname,A.address,A.gender,A.rollno,A.dateofbirth,A.dateofjoin, C.standard,D.SECTION,E.BLOCKNO,A.bloodgroup,A.contact,A.active,A.STUDENTIMAGE,A.STUDENTVOTELOG,a.election,F.postname,A.ELECTIONDATE from  ASPTBLSTUDENTMAS a join   asptblschoolmas b on b.asptblschoolmasID = a.compcode join asptblstandardmas C ON C.asptblstandardmasID=A.STANDARD   JOIN asptblsectionmas  D ON D.asptblsectionmasid=A.SECTION JOIN asptblblockmas E ON E.asptblblockmasID=A.BLOCKNO JOIN asptblelectionmas F ON F.ASPTBLELECTIONMASID=A.ELECTIONPOST  where a.ASPTBLSTUDENTMASID=" + em.asptblstudentmasid;
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
            DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
            foreach (DataRow myRow in dt.Rows)
            {
                txtstudentid.Text = "";
                comboschoolcode.Text = Convert.ToString(myRow["COMPCODE"].ToString());
                comboschoolname.Text = Convert.ToString(myRow["COMPNAME"].ToString());
                txtstudentname.Text = Convert.ToString(myRow["studentname"].ToString());
                txtlastname.Text = Convert.ToString(myRow["LASTNAME"].ToString());
                txtaddress.Text = Convert.ToString(myRow["ADDRESS"].ToString());
                if (myRow["GENDER"].ToString() == "M") radiomale.Checked = true;
                else radiofemale.Checked = true;
                txtrollno.Text = Convert.ToString(myRow["rollno"].ToString());
                txtdateofbirth.Text = Convert.ToString(myRow["DATEOFBIRTH"].ToString());
                txtdateofjoin.Text = Convert.ToString(myRow["DATEOFJOIN"].ToString());
                combostandard.Text = Convert.ToString(myRow["standard"].ToString());
                combosection.Text = Convert.ToString(myRow["SECTION"].ToString());
                comboblock.Text = Convert.ToString(myRow["BLOCKNO"].ToString());
                combobroup.Text = Convert.ToString(myRow["bloodgroup"].ToString());
                txtcontactno.Text = Convert.ToString(myRow["contact"].ToString());
                if (myRow["ELECTION"].ToString() == "T") checkelection.Checked = true;
                else checkelection.Checked = true;
                comboelectionpost.Text = Convert.ToString(myRow["postname"].ToString());
                dateTimePickerelectiondate.Value = Convert.ToDateTime(myRow["ELECTIONDATE"].ToString());
                if (myRow["ACTIVE"].ToString() == "T") checkactive.Checked = true; else checkactive.Checked = false;
                if (myRow["STUDENTIMAGE"].ToString() != "")
                {

                    stdbytes = (byte[])myRow["STUDENTIMAGE"];
                    Image img = Models.Device.ByteArrayToImage(stdbytes);
                    PictureBox1.Image = img;


                }
                if (myRow["STUDENTVOTELOG"].ToString() != "")
                {

                    votebytes = (byte[])myRow["STUDENTVOTELOG"];
                    Image img = Models.Device.ByteArrayToImage(votebytes);
                    PictureBox2.Image = img;


                }




                Cursor = Cursors.Default;

            }
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

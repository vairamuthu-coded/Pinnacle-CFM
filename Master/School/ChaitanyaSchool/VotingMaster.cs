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
    public partial class VotingMaster : Form,ToolStripAccess
    {
        private static VotingMaster _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();

        Models.UserRights sm = new Models.UserRights(); string[] s;
        byte[] bytes; string myString = ""; int i = 0; byte[] stdbytes; byte[] votebytes;
        public static VotingMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new VotingMaster();
                return _instance;
            }
        }
        public VotingMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            GlobalVariables.CurrentForm = this;
        }

        private void VotingMaster_Load(object sender, EventArgs e)
        {
           
        }
        void electionpost()
        {
            string sel1 = "select 0  as asptblelectionmasid,'' as postname from asptblelectionmas union  select distinct a.asptblelectionmasid,a.postname from asptblelectionmas a join ASPTBLSTUDENTMAS b on a.asptblelectionmasid=b.electionpost where b.election='T' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblelectionmas");
            DataTable dt = ds.Tables["asptblelectionmas"];
  
            comboselectelection.DataSource = dt;
            comboselectelection.DisplayMember = "postname";
            comboselectelection.ValueMember = "asptblelectionmasid";

            comboelecationpost.DataSource = dt;
            comboelecationpost.DisplayMember = "postname";
            comboelecationpost.ValueMember = "asptblelectionmasid";
        }
        private void comboselectelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                lblvoteing.Text = "";

                pop(comboselectelection.Text);
                posttotalcount(comboselectelection.Text);
            }
        }
        private void pop(string s)
        {
          
            string sel1 = "select A.ASPTBLSTUDENTMASID, B.compcode,B.compname,A.studentname,A.lastname,A.address,A.gender,A.rollno,A.dateofbirth,A.dateofjoin, C.standard ,D.SECTION , E.BLOCKNO,A.bloodgroup,A.contact,A.active,A.STUDENTIMAGE,A.STUDENTVOTELOG,F.POSTNAME from  ASPTBLSTUDENTMAS a join   asptblschoolmas b on b.asptblschoolmasID = a.compcode join asptblstandardmas C ON C.asptblstandardmasID=A.STANDARD   JOIN asptblsectionmas  D ON D.asptblsectionmasid=A.SECTION JOIN asptblblockmas E ON E.asptblblockmasID=A.BLOCKNO  JOIN asptblelectionmas   F ON F.asptblelectionmasid=A.ELECTIONPOST where a.election='T' AND A.ACTIVE='T' and F.POSTNAME='" + s + "' ORDER BY F.POSTNAME ASC ";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
            DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
            Pinnacle.UserControls.CTS.Voting[] items = new UserControls.CTS.Voting[dt.Rows.Count];
            flowLayoutPanel1.Controls.Clear();
            foreach (DataRow myRow in dt.Rows)
            {
                items[i] = new UserControls.CTS.Voting();


                items[i].studentid.Text = Convert.ToString(myRow["ASPTBLSTUDENTMASID"].ToString());

                items[i].ElectionPost.Text = Convert.ToString(myRow["POSTNAME"].ToString());
                items[i].studentname.Text = "Name : " + Convert.ToString(myRow["studentname"].ToString() + "      - " + myRow["rollno"].ToString() + " ");
                items[i].standard.Text = Convert.ToString(myRow["standard"].ToString() + "  " + myRow["section"].ToString());
                flowLayoutPanel1.Controls.Add(items[i]);
                if (myRow["STUDENTIMAGE"].ToString() != "")
                {
                    stdbytes = (byte[])myRow["STUDENTIMAGE"];
                    Image img = Models.Device.ByteArrayToImage(stdbytes);
                    items[i].studentimage = img;
                }

                if (myRow["STUDENTVOTELOG"].ToString() != "")
                {
                    votebytes = (byte[])myRow["STUDENTVOTELOG"];
                    Image img = Models.Device.ByteArrayToImage(votebytes);
                    items[i].studentLogoimage = img;
                }

                items[i].VottingOkButton.Text = items[i].studentid.Text;
                items[i].VottingOkButton.Click += VottingOkButton_Click1;
               
            }

        }

        private void VottingOkButton_Click1(object sender, EventArgs e)
        {
            
            if (checkBox1.Checked == false)
            {
                s = sender.ToString().Split(','); Cursor = Cursors.WaitCursor;
                em.asptblstudentmasid = Convert.ToInt64(s[1].Substring(7).TrimEnd());
                string sel1 = "select A.ASPTBLSTUDENTMASID, A.compcode,A.studentname,A.lastname,A.address,F.electiontime," +
                            "A.gender,A.rollno,A.dateofbirth,A.dateofjoin, a.standard,a.SECTION,A.BLOCKNO,A.bloodgroup,A.contact,A.active," +
                            "A.STUDENTIMAGE,A.STUDENTVOTELOG,F.ASPTBLELECTIONMASID,A.ELECTIONDATE from  ASPTBLSTUDENTMAS a join   asptblschoolmas b on b.asptblschoolmasID = a.compcode join asptblstandardmas C ON C.asptblstandardmasID=A.STANDARD   JOIN asptblsectionmas  D ON D.asptblsectionmasid=A.SECTION JOIN asptblblockmas E ON E.asptblblockmasID=A.BLOCKNO JOIN ASPTBLELECTIONMAS  F ON F.ASPTBLELECTIONMASID=A.ELECTIONPOST  where a.ASPTBLSTUDENTMASID=" + em.asptblstudentmasid;
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
                DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
                Class.Users.ElectionTime = Convert.ToInt32("0" + dt.Rows[0]["electiontime"].ToString());
                timer3.Enabled = true; 
                this.Enabled = false;
                //lblvoteing.Refresh();
               // lblvoteing.Text = "Student Name :" + Convert.ToString(dt.Rows[0]["studentname"].ToString() + "Session Time : " + Class.Users.UserTime;
                foreach (DataRow myRow in dt.Rows)
                {
                    em.asptblstudentmasid = Convert.ToInt64(myRow["ASPTBLSTUDENTMASID"].ToString());
                    em.compcode = Convert.ToInt64(myRow["compcode"].ToString());
                    em.compname = Convert.ToInt64(myRow["compcode"].ToString());
                    em.studentname = Convert.ToString(myRow["studentname"].ToString());
                    em.lastname = Convert.ToString(myRow["LASTNAME"].ToString());
                    em.address = Convert.ToString(myRow["ADDRESS"].ToString());
                    if (myRow["GENDER"].ToString() == "T") em.gender = "M";
                    else em.gender = "F";
                    em.ROLLNO = Convert.ToString(myRow["rollno"].ToString());
                    em.dateofbirth = Convert.ToString(Convert.ToDateTime(myRow["dateofbirth"].ToString()).ToString("yyyy-MM-dd"));
                    em.dateofjoin = Convert.ToString(Convert.ToDateTime(myRow["DATEOFJOIN"].ToString()).ToString("yyyy-MM-dd"));
                    em.standard = Convert.ToInt64(myRow["standard"].ToString());
                    em.SECTION = Convert.ToInt64(myRow["SECTION"].ToString());
                    em.BLOCKNO = Convert.ToInt64(myRow["BLOCKNO"].ToString());
                    em.bloodgroup = Convert.ToString(myRow["bloodgroup"].ToString());
                    em.contact = Convert.ToString(myRow["contact"].ToString());
                    if (myRow["ACTIVE"].ToString() == "T") em.active = "T"; else em.active = "F";
                    em.username = Convert.ToInt64("0" + Class.Users.USERID);
                    em.ipaddress = Class.Users.IPADDRESS;
                    em.createdon = Convert.ToString(System.DateTime.Now.ToString());
                    em.createdby = Convert.ToString(Class.Users.CREATED);
                    em.modifiedon = Convert.ToString(System.DateTime.Now.ToString());
                    em.election = "T";
                    em.votedate = Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd"));
                    em.electionpost = Convert.ToInt64(myRow["ASPTBLELECTIONMASID"].ToString());
                    em.electiondate = Convert.ToString(Convert.ToDateTime(myRow["electiondate"].ToString()).ToString("yyyy-MM-dd"));
                    string ins = "insert into asptblvotemas(asptblstudentmasid,compcode,compname,studentname,lastname,address,gender,rollno,dateofbirth,dateofjoin,standard,SECTION,BLOCKNO,bloodgroup,contact,active,username,ipaddress,createdby,createdon,modifiedon,election,votedate,votecount,postname,electiondate)VALUES('" + em.asptblstudentmasid + "','" + em.compcode + "','" + em.compname + "','" + em.studentname + "','" + em.lastname + "','" + em.address + "','" + em.gender + "','" + em.ROLLNO + "', DATE_FORMAT('" + em.dateofbirth.ToString() + "','%y-%m-%d')  , DATE_FORMAT('" + em.dateofjoin.ToString() + "','%y-%m-%d'),'" + em.standard + "',  '" + em.SECTION + "','" + em.BLOCKNO + "','" + em.bloodgroup + "','" + em.contact + "','" + em.active + "','" + em.username + "','" + em.ipaddress + "','" + em.username + "','" + em.createdon + "','" + em.modifiedon + "','" + em.election + "','" + em.votedate + "','1','" + em.electionpost + "','" + em.electiondate + "')";//@STUDENTVOTELOG,@STUDENTIMAGEBYTES,
                    Utility.ExecuteNonQuery(ins);
                }
            }
            Cursor = Cursors.Default;
        }

        private void pop()
        {


            string sel1 = "select A.ASPTBLSTUDENTMASID, B.compcode,B.compname,A.studentname,A.lastname,A.address,A.gender,A.rollno,A.dateofbirth,A.dateofjoin, C.standard ,D.SECTION , E.BLOCKNO,A.bloodgroup,A.contact,A.active,A.STUDENTIMAGE,A.STUDENTVOTELOG,F.POSTNAME from  ASPTBLSTUDENTMAS a join   asptblschoolmas b on b.asptblschoolmasID = a.compcode join asptblstandardmas C ON C.asptblstandardmasID=A.STANDARD   JOIN asptblsectionmas  D ON D.asptblsectionmasid=A.SECTION JOIN asptblblockmas E ON E.asptblblockmasID=A.BLOCKNO  JOIN asptblelectionmas   F ON F.asptblelectionmasid=A.ELECTIONPOST where a.election='T' ORDER BY F.POSTNAME ASC ";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
            DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
            Pinnacle.UserControls.CTS.Voting[] items = new UserControls.CTS.Voting[dt.Rows.Count];
            flowLayoutPanel1.Controls.Clear();
            foreach (DataRow myRow in dt.Rows)
            {
                items[i] = new UserControls.CTS.Voting();


                items[i].studentid.Text = Convert.ToString(myRow["ASPTBLSTUDENTMASID"].ToString());

                items[i].ElectionPost.Text = Convert.ToString(myRow["POSTNAME"].ToString());
                items[i].studentname.Text = "Name :" + Convert.ToString(myRow["studentname"].ToString() + "-" + myRow["rollno"].ToString() + " ");
                items[i].standard.Text = Convert.ToString(myRow["standard"].ToString() + "-" + myRow["section"].ToString());
                flowLayoutPanel1.Controls.Add(items[i]);
                if (myRow["STUDENTIMAGE"].ToString() != "")
                {
                    stdbytes = (byte[])myRow["STUDENTIMAGE"];
                    Image img = Models.Device.ByteArrayToImage(stdbytes);
                    items[i].studentimage = img;
                }
                if (myRow["STUDENTVOTELOG"].ToString() != "")
                {
                    votebytes = (byte[])myRow["STUDENTVOTELOG"];
                    Image img = Models.Device.ByteArrayToImage(votebytes);
                    items[i].studentLogoimage = img;
                }

                items[i].VottingOkButton.Text = items[i].studentid.Text;
                items[i].VottingOkButton.Click += VottingOkButton_Click;
                
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
                Class.Users.UserTime += 1;
           
            if (Class.Users.UserTime > Class.Users.ElectionTime)
            {
                timer3.Stop();
                this.Enabled = true; Cursor = Cursors.Default;
                Class.Users.UserTime = 0;
                lblvoteing.Refresh();                
                lblvoteing.Text = "Session Time  : " + Class.Users.UserTime;                
                totalcount();
            }
            else
            {
                this.Enabled = false;
                lblvoteing.Refresh(); 
                lblvoteing.Text = "Poling to  RollNo :" + em.ROLLNO  +"  : Name : "+ em.studentname +  " : Session Time : " + Class.Users.UserTime;

            }

        }
        private void VottingOkButton_Click(object sender, EventArgs e)
        {
          
            if (checkBox1.Checked == true)
            {
                Class.Users.UserTime = 0;
                s = sender.ToString().Split(','); Cursor = Cursors.WaitCursor;
                em.asptblstudentmasid = Convert.ToInt64(s[1].Substring(7).TrimEnd());
                string sel1 = "select A.ASPTBLSTUDENTMASID, A.compcode,A.studentname,A.lastname,A.address,F.electiontime," +
                            "A.gender,A.rollno,A.dateofbirth,A.dateofjoin, a.standard,a.SECTION,A.BLOCKNO,A.bloodgroup,A.contact,A.active," +
                            "A.STUDENTIMAGE,A.STUDENTVOTELOG,F.ASPTBLELECTIONMASID,A.ELECTIONDATE from  ASPTBLSTUDENTMAS a join   asptblschoolmas b on b.asptblschoolmasID = a.compcode join asptblstandardmas C ON C.asptblstandardmasID=A.STANDARD   JOIN asptblsectionmas  D ON D.asptblsectionmasid=A.SECTION JOIN asptblblockmas E ON E.asptblblockmasID=A.BLOCKNO JOIN ASPTBLELECTIONMAS  F ON F.ASPTBLELECTIONMASID=A.ELECTIONPOST  where a.ASPTBLSTUDENTMASID=" + em.asptblstudentmasid;
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTUDENTMAS");
                DataTable dt = ds.Tables["ASPTBLSTUDENTMAS"];
                Class.Users.ElectionTime = Convert.ToInt32("0" + dt.Rows[0]["electiontime"].ToString());
                timer3.Enabled = true; 
                this.Enabled = false; 
                //lblvoteing.Text = "Session Time  : " + Class.Users.UserTime;

                foreach (DataRow myRow in dt.Rows)
                {
                    em.asptblstudentmasid = Convert.ToInt64(myRow["ASPTBLSTUDENTMASID"].ToString());
                    em.compcode = Convert.ToInt64(myRow["compcode"].ToString());
                    em.compname = Convert.ToInt64(myRow["compcode"].ToString());
                    em.studentname = Convert.ToString(myRow["studentname"].ToString());
                    em.lastname = Convert.ToString(myRow["LASTNAME"].ToString());
                    em.address = Convert.ToString(myRow["ADDRESS"].ToString());
                    if (myRow["GENDER"].ToString() == "T") em.gender = "M";
                    else em.gender = "F";
                    em.ROLLNO = Convert.ToString(myRow["rollno"].ToString());
                    em.dateofbirth = Convert.ToString(Convert.ToDateTime(myRow["dateofbirth"].ToString()).ToString("yyyy-MM-dd"));
                    em.dateofjoin = Convert.ToString(Convert.ToDateTime(myRow["DATEOFJOIN"].ToString()).ToString("yyyy-MM-dd"));
                    em.standard = Convert.ToInt64(myRow["standard"].ToString());
                    em.SECTION = Convert.ToInt64(myRow["SECTION"].ToString());
                    em.BLOCKNO = Convert.ToInt64(myRow["BLOCKNO"].ToString());
                    em.bloodgroup = Convert.ToString(myRow["bloodgroup"].ToString());
                    em.contact = Convert.ToString(myRow["contact"].ToString());
                    if (myRow["ACTIVE"].ToString() == "T") em.active = "T"; else em.active = "F";
                    em.username = Convert.ToInt64("0" + Class.Users.USERID);
                    em.ipaddress = Class.Users.IPADDRESS;
                    em.createdon = Convert.ToString(System.DateTime.Now.ToString());
                    em.createdby = Convert.ToString(Class.Users.CREATED);
                    em.modifiedon = Convert.ToString(System.DateTime.Now.ToString());
                    em.election = "T";
                    em.votedate = Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd"));
                    em.electionpost = Convert.ToInt64(myRow["ASPTBLELECTIONMASID"].ToString());
                    em.electiondate = Convert.ToString(Convert.ToDateTime(myRow["electiondate"].ToString()).ToString("yyyy-MM-dd"));
                    string ins = "insert into asptblvotemas(asptblstudentmasid,compcode,compname,studentname,lastname,address,gender,rollno,dateofbirth,dateofjoin,standard,SECTION,BLOCKNO,bloodgroup,contact,active,username,ipaddress,createdby,createdon,modifiedon,election,votedate,votecount,postname,electiondate)VALUES('" + em.asptblstudentmasid + "','" + em.compcode + "','" + em.compname + "','" + em.studentname + "','" + em.lastname + "','" + em.address + "','" + em.gender + "','" + em.ROLLNO + "', DATE_FORMAT('" + em.dateofbirth.ToString() + "','%y-%m-%d')  , DATE_FORMAT('" + em.dateofjoin.ToString() + "','%y-%m-%d'),'" + em.standard + "',  '" + em.SECTION + "','" + em.BLOCKNO + "','" + em.bloodgroup + "','" + em.contact + "','" + em.active + "','" + em.username + "','" + em.ipaddress + "','" + em.username + "','" + em.createdon + "','" + em.modifiedon + "','" + em.election + "','" + em.votedate + "','1','" + em.electionpost + "','" + em.electiondate + "')";//@STUDENTVOTELOG,@STUDENTIMAGEBYTES,
                    Utility.ExecuteNonQuery(ins);
                    //lblvoteing.Refresh();
                    //lblvoteing.Text = "Session Time : " + Class.Users.UserTime;

                }
            }
            Cursor.Current = Cursors.Default;
        }
        void totalcount()
        {
            string sel2 = "select count(asptblvotemasid) as asptblvotemasid from asptblvotemas;";// where a.votedate='" + Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd")) + "'";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblvotemas");
            DataTable dt2 = ds2.Tables["asptblvotemas"];
            lblvoteing.Text = ""; lblvoteing.Refresh();
            lblvoteing.Text = "Election Poling Count  :" + dt2.Rows[0]["asptblvotemasid"].ToString();
        }
        void posttotalcount(string post)
        {
            string sel2 = "select sum(a.votecount)  as votecount from asptblvotemas a join ASPTBLSTUDENTMAS  b on a.ASPTBLSTUDENTMASid=b.ASPTBLSTUDENTMASid join asptblschoolmas c on c.asptblschoolmasid=a.compcode and c.asptblschoolmasid=b.compcode join asptblstandardmas d on d.asptblstandardmasid=a.standard join asptblsectionmas e on e.asptblsectionmasid=a.section JOIN asptblelectionmas   F ON F.asptblelectionmasid=B.ELECTIONPOST where b.election='T' AND F.POSTNAME='" + post + "'";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblvotemas");
            DataTable dt2 = ds2.Tables["asptblvotemas"];
            lblvoteing.Refresh(); lblvoteing.Text = "";
            lblvoteing.Text = "Election Poling Count  :" + dt2.Rows[0]["votecount"].ToString();
        }
        public void News()
        {
            empty();
            GridLoad();
            pop(); lblvoteing.Text = ""; em.votecount = 0;
            RollNo();
          ////  string sel1 = "select distinct a.asptblelectionmasid,a.postname,date_format(b.ELECTIONDATE,'%d-%m-%Y') as ELECTIONDATE from asptblelectionmas a join ASPTBLSTUDENTMAS b on a.asptblelectionmasid=b.electionpost where b.election='T' ;";
            string sel1 = "select distinct date_format(b.ELECTIONDATE,'%d-%m-%Y') as ELECTIONDATE from asptblelectionmas a join ASPTBLSTUDENTMAS b on a.asptblelectionmasid=b.electionpost where b.election='T' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblelectionmas");
            DataTable dt = ds.Tables["asptblelectionmas"];
            comboelectiondate1.DataSource = dt;
            comboelectiondate1.DisplayMember = "ELECTIONDATE";
            comboelectiondate1.ValueMember = "ELECTIONDATE";
            electionpost(); totalcount();
        }
        private void empty() {
            panel1.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            lblvoteing.ForeColor= Class.Users.Color1;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;Class.Users.UserTime = 0;
            lblpost.ForeColor = Class.Users.Color1;
            lbldate1.ForeColor = Class.Users.Color1;
            lblstudent.ForeColor= Class.Users.Color1;
            lbldate.ForeColor = Class.Users.Color1;
            lblelectionpost.ForeColor = Class.Users.Color1;
        }

        public void Saves()
        {

        }

        public void Prints()
        {

        }

        public void Searchs()
        {

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

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        public void GridLoad()
        {

        }
        Report.CTS.VotingResultReport rd3 = new Report.CTS.VotingResultReport();
        Report.CTS.VoteDetailsReport rd4 = new Report.CTS.VoteDetailsReport();
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
               // combostudentvotecount_SelectedIndexChanged(sender, e);
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                //DataTable dt2 = new DataTable(); ;

                //string sel2 = "select A.ASPTBLSTUDENTMASID,A.ROLLNO from asptblstudentmas a  where a.election='T'";
                //DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblstudentmas");
                //dt2 = ds2.Tables["asptblstudentmas"];
                //if (dt2.Rows.Count > 0)
                //{
                //    comborollno.DataSource = dt2;
                //    comborollno.DisplayMember = "ROLLNO";
                //    comborollno.ValueMember = "ASPTBLSTUDENTMASID";


                //}
            }
        }
        void RollNo()
        {
            DataTable dt2 = new DataTable(); 

            string sel2 = "select A.ASPTBLSTUDENTMASID,A.ROLLNO from asptblstudentmas a  where a.election='T'";
            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblstudentmas");
            dt2 = ds2.Tables["asptblstudentmas"];
            if (dt2.Rows.Count > 0)
            {
                comborollno.DataSource = dt2;
                comborollno.DisplayMember = "ROLLNO";
                comborollno.ValueMember = "ASPTBLSTUDENTMASID";


            }
        }
        private void combostudentvotecount_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboelecationpost.Text != "System.Data.DataRowView" && comboelecationpost.Text != "")
            {
                DataTable dt2 = new DataTable(); ;

                string sel2 = "select a.asptblvotemasid,b.rollno,b.studentname,d.standard, e.section,c.compname,b.STUDENTIMAGE,b.STUDENTVOTELOG,sum(a.votecount) as votecount,F.POSTNAME AS TEST1,B.ELECTIONDATE AS votedate from asptblvotemas a join ASPTBLSTUDENTMAS  b on a.ASPTBLSTUDENTMASid=b.ASPTBLSTUDENTMASid join asptblschoolmas c on c.asptblschoolmasid=a.compcode and c.asptblschoolmasid=b.compcode join asptblstandardmas d on d.asptblstandardmasid=a.standard join asptblsectionmas e on e.asptblsectionmasid=a.section JOIN asptblelectionmas   F ON F.asptblelectionmasid=B.ELECTIONPOST where b.election='T' AND B.ACTIVE='T' AND F.POSTNAME='" + comboelecationpost.Text + "' group by a.asptblvotemasid,b.rollno,b.studentname,d.standard, e.section,c.compname,b.STUDENTIMAGE,b.STUDENTVOTELOG,F.POSTNAME,B.ELECTIONDATE";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblvotemas");
                dt2 = ds2.Tables["asptblvotemas"];                
                if (dt2.Rows.Count > 0)
                {
                    rd3.SetDataSource(dt2);
                    crystalReportViewer1.ReportSource = null;
                    crystalReportViewer1.ReportSource = rd3;
                    crystalReportViewer1.Refresh();
                }
                else
                {
                    crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();
                }
            }
        }
        void electioncandidateload()
        {
            //string sel2 = "select 0 AS ASPTBLSTUDENTMASid, 'ALL' AS  studentname from ASPTBLSTUDENTMAS UNION select b.ASPTBLSTUDENTMASid, b.studentname from ASPTBLSTUDENTMAS  b where b.election='T'";
            //DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLSTUDENTMAS");
            //DataTable dt2 = ds2.Tables["ASPTBLSTUDENTMAS"];
            //if (dt2 !=null)
            //{
            //    combostudentvotecount.DataSource = dt2;
            //    combostudentvotecount.DisplayMember = "studentname";
            //    combostudentvotecount.ValueMember = "ASPTBLSTUDENTMASid";
            //}
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pop();
        }

        private void comborollno_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comborollno.SelectedItem.ToString() != "")
            {
                string sel3 = "select A.ASPTBLSTUDENTMASID,A.STUDENTNAME from asptblstudentmas a  where a.election='T' AND A.ROLLNO='" + comborollno.Text + "'";
                DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblstudentmas");
                DataTable dt3 = ds3.Tables["asptblstudentmas"];
                if (dt3.Rows.Count > 0)
                {
                    combostudent.DataSource = dt3;
                    combostudent.DisplayMember = "STUDENTNAME";
                    combostudent.ValueMember = "ASPTBLSTUDENTMASID";



                }
               
            }
           
        }

        private void combostudent_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combostudent.SelectedItem.ToString() != "")
            {
                string sel4 = "select distinct A.ELECTIONDATE from asptblstudentmas a  where a.election='T' AND A.ROLLNO='" + comborollno.Text + "'";
                DataSet ds4 = Utility.ExecuteSelectQuery(sel4, "asptblstudentmas");
                DataTable dt4 = ds4.Tables["asptblstudentmas"];
                if (dt4.Rows.Count > 0)
                {
                    comboElectiondate.DataSource = dt4;
                    comboElectiondate.DisplayMember = "ELECTIONDATE";
                    comboElectiondate.ValueMember = "ELECTIONDATE";
                }
            }
        }

        private void comboElectiondate_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt2 = new DataTable(); ;
            if (comborollno.Text != "System.Data.DataRowView" && comboElectiondate.Text !="System.Data.DataRowView")
            {
                string sel2 = "select a.asptblvotemasid,b.rollno,b.studentname,d.standard, e.section,c.compname,b.STUDENTIMAGE,b.STUDENTVOTELOG,a.votecount,F.POSTNAME AS TEST1,date_format(b.ELECTIONDATE,'%d-%m-%Y') as votedate  ,a.createdon test2 from asptblvotemas a join ASPTBLSTUDENTMAS  b on a.ASPTBLSTUDENTMASid=b.ASPTBLSTUDENTMASid join asptblschoolmas c on c.asptblschoolmasid=a.compcode and c.asptblschoolmasid=b.compcode join asptblstandardmas d on d.asptblstandardmasid=a.standard join asptblsectionmas e on e.asptblsectionmasid=a.section JOIN asptblelectionmas   F ON F.asptblelectionmasid=B.ELECTIONPOST where b.election='T' AND B.ACTIVE='T' and b.rollno='" + comborollno.Text + "'  and b.studentname='" + combostudent.Text + "' and b.ELECTIONDATE='" +Convert.ToString(Convert.ToDateTime(comboElectiondate.Text).ToString("yyyy-MM-dd")) + "' ";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblvotemas");
                dt2 = ds2.Tables["asptblvotemas"];

                if (dt2.Rows.Count > 0)
                {
                    rd4.SetDataSource(dt2);
                    crystalReportViewer2.ReportSource = null;
                    crystalReportViewer2.ReportSource = rd4;
                    crystalReportViewer2.Refresh();
                }
                else
                {
                    crystalReportViewer2.ReportSource = null; crystalReportViewer2.Refresh();
                 
                }
            }
            
        }


        private void comboelecationpost_SelectedIndexChanged(object sender, EventArgs e)
        {
            combostudentvotecount_SelectedIndexChanged(sender, e);
        }

        private void LBLHEADING_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                lblvoteing.Text = "";
                comboselectelection.Enabled = true; GlobalVariables.CurrentForm = this;
                pop(comboselectelection.Text);
            }
            else
            {
                comboelecationpost.Text = "";
                lblvoteing.Text = "";totalcount();
                comboselectelection.Enabled = false; GlobalVariables.CurrentForm = this;
                pop();
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            electionpost(); RollNo();
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

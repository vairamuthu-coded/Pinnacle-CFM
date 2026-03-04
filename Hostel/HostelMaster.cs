using System;
using System.Data;
using System.Windows.Forms;

namespace Pinnacle.Hostel
{
    public partial class HostelMaster : Form, ToolStripAccess
    {
        public HostelMaster()
        {
            InitializeComponent();

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }

        public void ReadOnlys()
        {

        }

        private static HostelMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Employee em = new Models.Employee();
        Models.UserRights sm = new Models.UserRights();
       
        public static HostelMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new HostelMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        public void News()
        {

            txthostelblock.Text = "";
            txthostelmasid.Text = ""; combo_compcode.SelectedIndex = -1;
            txtidcardno.Text = "";
            comboempname.Text = "";
            comboempname.SelectedIndex = -1;
            txthostelroom.Text = "";

            checkactive.Checked = true;

        }

        public void Saves()
        {
            try
            {
                if (combo_compcode.Text != "" && comboempname.Text != "" && txtidcardno.Text != "" && txthostelblock.Text != "" && txthostelroom.Text != "")
                {
                    string chk = "";
                    string hostel = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }
                    if (radioBoysHostel.Checked == true) { hostel = "BOYS HOSTEL"; } else { hostel = "GIRLS HOSTEL"; radioBoysHostel.Checked = false; }
                    string sel = "select A.ASPTBLHOSTELMASID FROM ASPTBLHOSTELMAS A  WHERE A.COMPCODE=" + combo_compcode.SelectedValue + " AND A.EMPNAME=" + comboempname.SelectedValue + " AND A.IDCARDNO=" + txtidcardno.Text + " AND A.HOSTELNAME='" + hostel + "' AND A.HOSTELBLOCK='" + txthostelblock.Text.ToUpper() + "' AND A.HOSTELROOM='" + txthostelroom.Text.ToUpper() + "'  AND A.ACTIVE='" + chk + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLHOSTELMAS");
                    DataTable dt = ds.Tables["ASPTBLHOSTELMAS"];
                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("Child Record Found    :" + hostel, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else if (dt.Rows.Count == 0 && Convert.ToInt32("0" + txthostelmasid.Text) == 0 || Convert.ToInt32("0" + txthostelmasid.Text) == 0)
                    {
                        string ins = "INSERT INTO ASPTBLHOSTELMAS(COMPCODE,EMPNAME,IDCARDNO, HOSTELNAME,HOSTELBLOCK,HOSTELROOM,ACTIVE, USERNAME,  MODIFIED,  CREATEDON,  IPADDRESS)VALUES(" + combo_compcode.SelectedValue + "," + comboempname.SelectedValue + "," + txtidcardno.Text + ",'" + hostel + "','" + txthostelblock.Text.ToUpper() + "','" + txthostelroom.Text.ToUpper() + "','" + chk + "'," + Class.Users.USERID + ",to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:ss'),to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),'" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully   :" + hostel, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    else
                    {
                        string up = "UPDATE ASPTBLHOSTELMAS SET COMPCODE=" + combo_compcode.SelectedValue + ", EMPNAME=" + comboempname.SelectedValue + ",IDCARDNO=" + txtidcardno.Text + ", HOSTELNAME='" + hostel + "',HOSTELBLOCK='" + txthostelblock.Text.ToUpper() + "', HOSTELROOM='" + txthostelroom.Text.ToUpper() + "',ACTIVE='" + chk + "', USERNAME=" + Class.Users.USERID + ",  MODIFIED=to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),IPADDRESS='" + Class.Users.IPADDRESS + "' WHERE  ASPTBLHOSTELMASID=" + txthostelmasid.Text;
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated     :" + hostel, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    GridLoad();
                    News();
                }
                else
                {
                    MessageBox.Show("PLS Enter Mandatary Fields", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GridLoad()
        {
            try
            {
                listmaster.Items.Clear();
                string sel1 = "SELECT A.ASPTBLHOSTELMASID, B.COMPCODE ,C.FNAME AS EMPNAME,D.MIDCARD AS  IDCARDNO,  A.HOSTELNAME,A.HOSTELBLOCK,A.HOSTELROOM,A.ACTIVE  FROM  ASPTBLHOSTELMAS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN HREMPLOYMAST C ON C.COMPCODE = B.GTCOMPMASTID AND C.HREMPLOYMASTID = A.EMPNAME JOIN HREMPLOYDETAILS D ON C.HREMPLOYMASTID = D.HREMPLOYMASTID   WHERE B.COMPCODE='" + Class.Users.HCompcode + "' ORDER BY 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELMAS");
                DataTable dt = ds.Tables["ASPTBLHOSTELMAS"];
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(myRow["ASPTBLHOSTELMASID"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["EMPNAME"].ToString());
                        list.SubItems.Add(myRow["IDCARDNO"].ToString());
                        list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                        list.SubItems.Add(myRow["HOSTELBLOCK"].ToString());
                        list.SubItems.Add(myRow["HOSTELROOM"].ToString());
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        listmaster.Items.Add(list);
                    }
                    lbltotcount.Text = "Total Count    :" + listmaster.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void HostelMaster_Load(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = mas.findcomcode(Class.Users.HCompcode, Class.Users.HUserName);
                if (dt.Rows.Count > 0)
                {
                    combo_compcode.DisplayMember = "COMPCODE";
                    combo_compcode.ValueMember = "GTCOMPMASTID";
                    combo_compcode.DataSource = dt;


                }

                //string sel1 = "SELECT A.ASPTBLEMPID,A.EMPNAME FROM ASPTBLEMP A  ORDER BY 1";
                //DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLEMP");
                //DataTable dt1 = ds1.Tables["ASPTBLEMP"];
                //if (dt1.Rows.Count > 0)
                //{
                //    comboempname.DisplayMember = "EMPNAME";
                //    comboempname.ValueMember = "ASPTBLEMPID";
                //    comboempname.DataSource = dt1;


                //}
                combo_compcode.SelectedIndex = -1;
                comboempname.SelectedIndex = -1;
                txtidcardno.Text = ""; combo_compcode.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            GridLoad();
            this.combo_compcode.Focus();
        }

        private void Listmaster_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listmaster.Items.Count > 0)
                {

                    txthostelmasid.Text = listmaster.SelectedItems[0].SubItems[1].Text;
                    string sel1 = "SELECT A.ASPTBLHOSTELMASID, B.COMPCODE ,C.FNAME AS EMPNAME,D.MIDCARD AS  IDCARDNO,  A.HOSTELNAME,A.HOSTELBLOCK,A.HOSTELROOM,A.ACTIVE  FROM  ASPTBLHOSTELMAS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN HREMPLOYMAST C ON C.COMPCODE = B.GTCOMPMASTID AND C.HREMPLOYMASTID = A.EMPNAME JOIN HREMPLOYDETAILS D ON C.HREMPLOYMASTID = D.HREMPLOYMASTID WHERE A.ASPTBLHOSTELMASID=" + txthostelmasid.Text;
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELMAS");//C.EMPNAME,C.IDCARDNO,
                    DataTable dt = ds1.Tables["ASPTBLHOSTELMAS"];
                    if (dt.Rows.Count > 0)
                    {
                        txthostelmasid.Text = Convert.ToString(dt.Rows[0]["ASPTBLHOSTELMASID"].ToString());
                        combo_compcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                        comboempname.Text = Convert.ToString(dt.Rows[0]["EMPNAME"].ToString());
                        txtidcardno.Text = Convert.ToString(dt.Rows[0]["IDCARDNO"].ToString());
                        if (dt.Rows[0]["HOSTELNAME"].ToString() == "BOYS HOSTEL") { radioBoysHostel.Checked = true; } else { radioGirlsHostel.Checked = true; }
                        // if (dt.Rows[0]["HOSTELNAME"].ToString() == "GIRLS HOSTEL") { radioGirlsHostel.Checked = true; } else { radioBoysHostel.Checked = false; }
                        txthostelblock.Text = Convert.ToString(dt.Rows[0]["HOSTELBLOCK"].ToString());
                        txthostelroom.Text = Convert.ToString(dt.Rows[0]["HOSTELROOM"].ToString());

                        if (dt.Rows[0]["ACTIVE"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        //  txtcompcode.Text = Convert.ToString(dt.Rows[0]["GTCOMPMASTID"].ToString());

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Comboempname_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                if (comboempname.SelectedIndex >= 0)
                {
                    string sel1 = "SELECT B.MIDCARD AS IDCARDNO FROM HREMPLOYMAST A JOIN HREMPLOYDETAILS B ON A.HREMPLOYMASTID = B.HREMPLOYMASTID JOIN GTCOMPMAST C ON A.COMPCODE = C.GTCOMPMASTID JOIN GTDEPTDESGMAST D ON D.GTDEPTDESGMASTID = B.DEPTNAME WHERE  A.HREMPLOYMASTID=" + comboempname.SelectedValue;
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "HREMPLOYMAST");
                    DataTable dt1 = ds1.Tables["HREMPLOYMAST"];
                    if (dt1.Rows.Count > 0)
                    {
                        txtidcardno.DisplayMember = "IDCARDNO";
                        txtidcardno.ValueMember = "IDCARDNO";
                        txtidcardno.DataSource = dt1;

                        // txtidcardno.Text = dt1.Rows[0]["IDCARDNO"].ToString();


                    }
                }
            }
            catch (Exception EX)
            {
                // MessageBox.Show(EX.Message);
            }
        }

        private void Txthostelsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txthostelsearch.Text != "")
                {
                    listmaster.Items.Clear(); int iGLCount = 1;
                    string sel1 = "SELECT   A.ASPTBLHOSTELMASID,D.COMPCODE,B.FNAME AS EMPNAME,B.IDCARDNO, A.HOSTELNAME,A.HOSTELBLOCK,A.HOSTELROOM,A.ACTIVE  FROM ASPTBLHOSTELMAS A JOIN HREMPLOYMAST B ON A.EMPNAME=B.HREMPLOYMASTID JOIN HREMPLOYDETAILS C ON C.HREMPLOYMASTID=B.HREMPLOYMASTID AND C.MIDCARD=A.IDCARDNO AND B.IDCARDNO=C.IDCARD  JOIN GTCOMPMAST D ON D.GTCOMPMASTID = A.COMPCODE AND D.GTCOMPMASTID = B.COMPCODE  WHERE C.MIDCARD LIKE'%" + txthostelsearch.Text + "%' OR B.FNAME LIKE'%" + txthostelsearch.Text + "%' OR A.HOSTELNAME LIKE'%" + txthostelsearch.Text + "%' OR  A.HOSTELBLOCK LIKE'%" + txthostelsearch.Text + "%' OR A.HOSTELROOM LIKE'%" + txthostelsearch.Text + "%'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLHOSTELMAS");
                    DataTable dt = ds.Tables["ASPTBLHOSTELMAS"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["ASPTBLHOSTELMASID"].ToString());
                            list.SubItems.Add(myRow["COMPCODE"].ToString());
                            list.SubItems.Add(myRow["EMPNAME"].ToString());
                            list.SubItems.Add(myRow["IDCARDNO"].ToString());
                            list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                            list.SubItems.Add(myRow["HOSTELBLOCK"].ToString());
                            list.SubItems.Add(myRow["HOSTELROOM"].ToString());
                            list.SubItems.Add(myRow["ACTIVE"].ToString());
                            listmaster.Items.Add(list);
                            iGLCount++;
                        }
                        lbltotcount.Text = "Total Count    :" + listmaster.Items.Count;
                    }
                    else
                    {
                        listmaster.Items.Clear();
                    }
                }
                else
                {

                    listmaster.Items.Clear();
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void EmpNameRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Combo_compcode_SelectedIndexChanged(sender, e);
        }


        private void Combo_compcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt3 = mas.EmpName(Class.Users.HCompcode);
            if (dt3.Rows.Count > 0)
            {
                comboempname.DisplayMember = "EMPNAME";
                comboempname.ValueMember = "ASPTBLEMPID";
                comboempname.DataSource = dt3;
            }
        }

        private void Deletes_Click(object sender, EventArgs e)
        {
            string sel1 = "DELETE  FROM ASPTBLHOSTELMAS A  WHERE A.ASPTBLHOSTELMASID=" + txthostelmasid.Text;
            Utility.ExecuteNonQuery(sel1); GridLoad(); News();
            MessageBox.Show("Record Deleted");
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
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
            this.Hide();
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

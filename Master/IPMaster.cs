using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class IPMaster : Form,ToolStripAccess
    {
        public IPMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
      
        }
        ListView listfilter = new ListView();

        private static IPMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Employee em = new Models.Employee();
        Models.UserRights sm = new Models.UserRights();

     
        public static IPMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new IPMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void ReadOnlys()
        { }
        public void News()
        {
            txtmachineid.Text = "";
            combo_compcode.Text = ""; txtsessiontime.Text = ""; comboMTYPE2.Text = ""; comboMTYPE2.SelectedIndex = -1;
            comboipaddress.Text = "";
            combousername.Text = "";
            checkactive.Checked = true;
          
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            panel3.BackColor = Class.Users.BackColors; GridLoad();

        }
        private void userload()
        {


            string sel1 = "SELECT distinct  D.USERNAME   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'  JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME  ORDER BY 4";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
            DataTable dt = ds.Tables["ASPTBLMACHINEMAS"];
            if (dt != null)
            {
                comboUser.DataSource = dt;
                comboUser.DisplayMember = "USERNAME";
                comboUser.ValueMember = "USERNAME";
            }
        }
        private void combo_compcode1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = mas.username(combo_compcode1.Text); Class.Users.UserTime = 0;
            if (dt!=null)
            {
                comboUser.DataSource = dt;
                comboUser.DisplayMember = "USERNAME";
                comboUser.ValueMember = "USERNAME";
            }
            else
            {
                comboUser.DataSource = null;
                comboUser.Text = ""; comboUser.SelectedIndex = -1;
            }
        }
        //private void hostelload()
        //{
        //    try
        //    {
        //        string sel3 = " select DISTINCT  A.HOSTELNAME from HOSTELLIVEDATA A  ORDER BY 1";
        //        DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "HOSTELLIVEDATA");
        //        DataTable dt3 = ds3.Tables["HOSTELLIVEDATA"];
        //        if (dt3.Rows.Count > 0)
        //        {
        //            int i = 1;
        //            foreach (DataRow myRow in dt3.Rows)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                list.SubItems.Add(i.ToString());
        //                list.SubItems.Add(myRow["HOSTELNAME"].ToString());

        //                listView1.Items.Add(list);

        //                i++;
        //            }

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Data Source Not Connected" + ex.Message);
        //    }
        //}
        public void Saves()
        {
            try
            {
                if (combo_compcode.Text != "" && combousername.Text != "" && comboipaddress.Text != "" && comboMTYPE2.Text != "")
                {
                    string chk = "";

                    string hostel = "", agfchk = "", flfchk = "", flfdchk = "", agfmchk = "", agfmgiichk = "", agfcchk = "", agfpchk = "", agfkchk = "";

                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }

                  string sel = "select A.ASPTBLMACHINEMASID,a.username FROM ASPTBLMACHINEMAS A  WHERE A.COMPCODE=" + combo_compcode.SelectedValue + " AND A.USERNAME=" + combousername.SelectedValue + " AND A.IPADDRESS=" + comboipaddress.SelectedValue + " AND A.HOSTELNAME='" + hostel + "'  AND A.ACTIVE='" + chk + "' AND A.MTYPE2='" + comboMTYPE2.Text + "' and A.SESSIONTIME='" + txtsessiontime.Text + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLMACHINEMAS");
                    DataTable dt = ds.Tables["ASPTBLMACHINEMAS"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found     :" + hostel, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (dt.Rows.Count == 0 && Convert.ToInt32("0" + txtmachineid.Text) == 0 || Convert.ToInt32("0" + txtmachineid.Text) == 0)
                    {
                        string ins = "INSERT INTO ASPTBLMACHINEMAS(COMPCODE,USERNAME,IPADDRESS,HOSTELNAME,ACTIVE,USERNAME1,  MODIFIED,  CREATEDON,  IPADDRESS1,MACIP,MTYPE2,SESSIONTIME)VALUES(" + combo_compcode.SelectedValue + "," + combousername.SelectedValue + "," + comboipaddress.SelectedValue + ",'" + hostel + "','" + chk + "'," + Class.Users.USERID + ",'" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "','" + Convert.ToDateTime(Class.Users.CREATED).ToString()+"','" + Class.Users.IPADDRESS + "' ,'" + comboipaddress.Text + "','" + comboMTYPE2.Text + "','" + txtsessiontime.Text + "' )";
                        Utility.ExecuteNonQuery(ins);

                        MessageBox.Show("Record Saved Successfully    :" + hostel, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    else
                    {
                        string up = "UPDATE ASPTBLMACHINEMAS SET COMPCODE=" + combo_compcode.SelectedValue + ", USERNAME=" + combousername.SelectedValue + ",IPADDRESS=" + comboipaddress.SelectedValue + ", HOSTELNAME='" + hostel + "',ACTIVE='" + chk + "', USERNAME1=" + Class.Users.USERID + ",  MODIFIED='" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "',IPADDRESS1='" + Class.Users.IPADDRESS + "' , MACIP='" + comboipaddress.Text + "',MTYPE2='" + comboMTYPE2.Text + "'  , SESSIONTIME='" + txtsessiontime.Text + "' WHERE  ASPTBLMACHINEMASID=" + txtmachineid.Text;
                        Utility.ExecuteNonQuery(up);
                        string up1 = "update asptblusermas set SESSIONTIME='" + txtsessiontime.Text + "' where   userid='" + combousername.SelectedValue + "'";
                        Utility.ExecuteNonQuery(up1);
                        MessageBox.Show("Record Updated     :" + hostel, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    GridLoad();

                    News();
                }
                else
                {
                    MessageBox.Show("PLS Enter Mandatary Fields");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        public void GridLoad()
        {
            try
            {
                int i = 1;
                listmachine.Items.Clear(); DataTable dt = new DataTable(); listfilter.Items.Clear();
                if (Class.Users.HUserName == "VAIRAM")
                {
                    string sel1 = "SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP  AS IPADDRESS ,A.HOSTELNAME,A.MTYPE2,D.SESSIONTIME,A.ACTIVE   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'  JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME  ORDER BY 2,4";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                    dt = ds.Tables["ASPTBLMACHINEMAS"];
                }
                else
                {
                    string sel1 = "SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP  AS IPADDRESS ,A.HOSTELNAME,A.MTYPE2,D.SESSIONTIME, A.ACTIVE   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS  AND C.ACTIVE='T'   JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME  AND D.COMPCODE=B.GTCOMPMASTID  WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND D.USERNAME='" + Class.Users.HUserName + "' ORDER BY 4";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                    dt = ds.Tables["ASPTBLMACHINEMAS"];
                }

                if (dt.Rows.Count > 0)
                {
                    i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLMACHINEMASID"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["USERNAME"].ToString());
                        list.SubItems.Add(myRow["IPADDRESS"].ToString());
                        list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                        list.SubItems.Add(myRow["MTYPE2"].ToString());
                        list.SubItems.Add(myRow["SESSIONTIME"].ToString());
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        if (myRow["MTYPE2"].ToString() == "TRANSER IP")
                        {
                            list.BackColor = this.BackColor;
                            list.ForeColor = Color.White;
                        }
                        if (myRow["MTYPE2"].ToString() == "OUTPASS")
                        {
                            list.BackColor = Color.Navy;
                            list.ForeColor = Color.White;
                        }
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listmachine.Items.Add(list); i++;
                    }
                    lbltotal.Text = "Total Count    :" + listmachine.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void comboUser_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                listmachine.Items.Clear(); DataTable dt = new DataTable(); listfilter.Items.Clear();
                //if (Class.Users.HUserName == "VAIRAM")
                //{
                //    string sel1 = "SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP  AS IPADDRESS ,A.HOSTELNAME,C.MTYPE2,  A.ACTIVE   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'  JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME ORDER BY 4";
                //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                //    dt = ds.Tables["ASPTBLMACHINEMAS"];
                //}
                //else
                //{
                string sel1 = "SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP  AS IPADDRESS ,A.HOSTELNAME,A.ACTIVE,A.MTYPE2,D.SESSIONTIME   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS  AND C.ACTIVE='T'   JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME  AND D.COMPCODE=B.GTCOMPMASTID  WHERE B.COMPCODE='" + combo_compcode1.Text + "' AND D.USERNAME='" + comboUser.Text + "' ORDER BY 4";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                dt = ds.Tables["ASPTBLMACHINEMAS"];
                // }
                int i = 1;
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLMACHINEMASID"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["USERNAME"].ToString());
                        list.SubItems.Add(myRow["IPADDRESS"].ToString());
                        list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                        list.SubItems.Add(myRow["MTYPE2"].ToString());
                        list.SubItems.Add(myRow["SESSIONTIME"].ToString());
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        if (myRow["MTYPE2"].ToString() == "TRANSER IP")
                        {
                            list.BackColor = Color.Navy;
                            list.ForeColor = Color.White;
                        }
                        if (myRow["MTYPE2"].ToString() == "OUTPASS")
                        {
                            list.BackColor = Color.Blue;
                            list.ForeColor = Color.White;
                        }
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listmachine.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count    :" + listmachine.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void LoadUser()
        {
            DataTable dt;
            if (Class.Users.HUserName == "VAIRAM")
            {
                string sel = "SELECT 0 GTCOMPMASTID, 'ALL' COMPCODE FROM DUAL UNION select  DISTINCT A.GTCOMPMASTID, a.compcode  from  gtcompmast a join asptblusermas b on a.gtcompmastid = b.compcode JOIN ASPTBLMACIP C ON C.COMPCODE=A.GTCOMPMASTID AND C.COMPCODE=B.COMPCODE   order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                dt = ds.Tables["gtcompmast"];
            }
            else
            {
                string sel = "select  DISTINCT A.GTCOMPMASTID, a.compcode  from  gtcompmast a join asptblusermas b on a.gtcompmastid = b.compcode JOIN ASPTBLMACIP C ON C.COMPCODE=A.GTCOMPMASTID AND C.COMPCODE=B.COMPCODE WHERE A.COMPCODE='" + Class.Users.HCompcode + "'  AND B.USERNAME='" + Class.Users.HUserName + "'   order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                dt = ds.Tables["gtcompmast"];
            }
            if (dt.Rows.Count > 0)
            {
                combo_compcode1.DisplayMember = "COMPCODE";
                combo_compcode1.ValueMember = "GTCOMPMASTID";
                combo_compcode1.DataSource = dt;
            }
        }
        private void IPMaster_Load(object sender, EventArgs e)
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


                combousername.SelectedIndex = -1;
                combo_compcode.SelectedIndex = -1;
                comboipaddress.SelectedIndex = -1;
                LoadUser();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
            GridLoad(); userload(); News();
            this.combo_compcode.Select();
        }
        void ipload(string s)
        {
            string sel1 = "SELECT DISTINCT  D.ASPTBLMACIPID ,D.MACIP FROM ASPTBLMACIP D  JOIN GTCOMPMAST B ON D.COMPCODE=B.GTCOMPMASTID   WHERE D.ACTIVE='T' AND B.COMPCODE='" + combo_compcode.Text + "'  ORDER BY 2 ";//AND  D.COMPCODE='" + Class.Users.HCompcode + "' AND  C.USERNAME='" + Class.Users.HUserName + "'
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACIP");
            DataTable dt1 = ds1.Tables["ASPTBLMACIP"];

            if (dt1.Rows.Count > 0)
            {


                comboipaddress.DisplayMember = "MACIP";
                comboipaddress.ValueMember = "ASPTBLMACIPID";
                comboipaddress.DataSource = dt1;


            }

        }
        private void Listmachine_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listmachine.Items.Count > 0)
                {
                    Class.Users.UserTime = 0;
                    txtmachineid.Text = listmachine.SelectedItems[0].SubItems[2].Text;

                    string sel1 = " SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP  as IPADDRESS ,A.HOSTELNAME,A.MTYPE2,D.SESSIONTIME,A.ACTIVE   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   JOIN  asptblusermas D ON D.userid=A.USERNAME   WHERE A.ASPTBLMACHINEMASID=" + txtmachineid.Text;
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");//C.EMPNAME,C.IDCARDNO,
                    DataTable dt = ds1.Tables["ASPTBLMACHINEMAS"];
                    if (dt.Rows.Count > 0)
                    {
                        txtmachineid.Text = Convert.ToString(dt.Rows[0]["ASPTBLMACHINEMASID"].ToString());
                        combo_compcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                        combousername.Text = Convert.ToString(dt.Rows[0]["USERNAME"].ToString());
                        comboipaddress.Text = Convert.ToString(dt.Rows[0]["IPADDRESS"].ToString());                      

                        comboMTYPE2.Text = Convert.ToString(dt.Rows[0]["MTYPE2"].ToString());
                        combousername_SelectedIndexChanged(sender, e);

                        if (dt.Rows[0]["ACTIVE"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Combo_compcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (combo_compcode.SelectedIndex >= 0)
                {
                    ipload(combo_compcode.Text);
                    Int64 s = Convert.ToInt64(combo_compcode.SelectedValue);
                    DataTable dt1 = mas.comcode1(s);
                    if (dt1.Rows.Count > 0)
                    {
                        combousername.DisplayMember = "USERNAME";
                        combousername.ValueMember = "userid";
                        combousername.DataSource = dt1;

                    }
                    combousername.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void WardenRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (combo_compcode.SelectedIndex >= 0)
                {

                    Int64 s = Convert.ToInt64(combo_compcode.SelectedValue);
                    DataTable dt1 = mas.comcode1(s);
                    if (dt1.Rows.Count > 0)
                    {
                        combousername.DisplayMember = "USERNAME";
                        combousername.ValueMember = "userid";
                        combousername.DataSource = dt1;


                    }
                    combousername.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }



        }

        private void IPRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt1;
            try
            {
                if (Class.Users.HUserName == "VAIRAM")
                {
                    string sel1 = "SELECT DISTINCT  D.ASPTBLMACIPID ,D.MACIP FROM ASPTBLMACIP D  JOIN GTCOMPMAST B ON D.COMPCODE=B.GTCOMPMASTID   WHERE D.ACTIVE='T'  ORDER BY 2 ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                    dt1 = ds.Tables["ASPTBLMACHINEMAS"];
                }
                else
                {
                    string sel2 = "SELECT DISTINCT  D.ASPTBLMACIPID ,D.MACIP FROM ASPTBLMACIP D  JOIN GTCOMPMAST B ON D.COMPCODE=B.GTCOMPMASTID   WHERE D.ACTIVE='T' AND B.COMPCODE='" + combo_compcode.Text + "'  ORDER BY 2 ";//AND  D.COMPCODE='" + Class.Users.HCompcode + "' AND  C.USERNAME='" + Class.Users.HUserName + "'
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLMACIP");
                    dt1 = ds2.Tables["ASPTBLMACIP"];
                }
                //string sel1 = " SELECT A.ASPTBLMACIPID , A.MACIP     FROM  ASPTBLMACIP   A  WHERE A.ACTIVE='T' ORDER BY 2";
                //DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACIP");
                //DataTable dt1 = ds1.Tables["ASPTBLMACIP"];

                if (dt1.Rows.Count >= 0)
                {


                    comboipaddress.DisplayMember = "MACIP";
                    comboipaddress.ValueMember = "ASPTBLMACIPID";
                    comboipaddress.DataSource = dt1;


                }
                else
                {
                    comboipaddress.DataSource = null;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void CompCodeRefreshToolStripMenuItem_Click(object sender, EventArgs e)
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
            }
            catch (Exception ex)
            { }
        }

        private void Txtmachinesearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    if (txtmachinesearch.Text != "")
            //    {
            //        listmachine.Items.Clear(); int iGLCount = 1;
            //        string sel1 = " SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP as IPADDRESS,A.HOSTELNAME,A.ACTIVE    FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN HRMACIPENTRYDET C ON C.HRMACIPENTRYDETID=A.IPADDRESS  JOIN  asptblusermas D ON D.userid=A.USERNAME  AND D.COMPCODE=B.GTCOMPMASTID WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND  D.USERNAME LIKE'%" + txtmachinesearch.Text + "%'  OR A.HOSTELNAME LIKE'%" + txtmachinesearch.Text + "%'  OR C.MACIP LIKE'%" + txtmachinesearch.Text + "%'  OR A.ACTIVE LIKE'%" + txtmachinesearch.Text + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
            //        DataTable dt = ds.Tables["ASPTBLMACHINEMAS"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["ASPTBLMACHINEMASID"].ToString());
            //                list.SubItems.Add(myRow["COMPCODE"].ToString());
            //                list.SubItems.Add(myRow["USERNAME"].ToString());
            //                list.SubItems.Add(myRow["HOSTELNAME"].ToString());
            //                list.SubItems.Add(myRow["IPADDRESS"].ToString());
            //                list.SubItems.Add(myRow["ACTIVE"].ToString());
            //                listmachine.Items.Add(list);
            //                iGLCount++;
            //            }
            //            lbltotal.Text = "Total Count    :" + listmachine.Items.Count;
            //        }
            //        else
            //        {
            //            listmachine.Items.Clear();
            //        }
            //    }
            //    else
            //    {
            //        listmachine.Items.Clear();
            //        GridLoad();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            try
            {
                listmachine.Items.Clear(); DataTable dt = new DataTable(); listfilter.Items.Clear();
                if (Class.Users.HUserName == "VAIRAM")
                {
                    string sel1 = "SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP  AS IPADDRESS ,A.HOSTELNAME,C.MTYPE2,  A.ACTIVE   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'  JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME ORDER BY 4";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                    dt = ds.Tables["ASPTBLMACHINEMAS"];
                }
                else
                {
                    string sel1 = "SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS USERNAME,C.MACIP  AS IPADDRESS ,A.HOSTELNAME,A.ACTIVE,A.MTYPE2,D.SESSIONTIME   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS  AND C.ACTIVE='T'   JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME  AND D.COMPCODE=B.GTCOMPMASTID  WHERE B.COMPCODE='" + combo_compcode1.Text + "' AND  D.USERNAME='" + comboUser.Text + "'  OR A.HOSTELNAME LIKE'%" + txtmachinesearch.Text + "%'  OR C.MACIP LIKE'%" + txtmachinesearch.Text + "%'  OR A.ACTIVE LIKE'%" + txtmachinesearch.Text + "%' ORDER BY 4";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                    dt = ds.Tables["ASPTBLMACHINEMAS"];
                }
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(myRow["ASPTBLMACHINEMASID"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["USERNAME"].ToString());
                        list.SubItems.Add(myRow["IPADDRESS"].ToString());
                        list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                        list.SubItems.Add(myRow["MTYPE2"].ToString());
                        list.SubItems.Add(myRow["SESSIONTIME"].ToString());
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        // this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listmachine.Items.Add(list);
                    }
                    lbltotal.Text = "Total Count    :" + listmachine.Items.Count;
                }
                else
                {
                    listmachine.Items.Clear();
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        public void Deletes()
        {

            string sel1 = "DELETE  FROM ASPTBLMACHINEMAS WHERE ASPTBLMACHINEMASID=" + txtmachineid.Text;
            Utility.ExecuteNonQuery(sel1); GridLoad(); MessageBox.Show("Record Deleted"); News();
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

        private void butok_Click(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
        }

        private void combousername_SelectedIndexChanged(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
            string sel = "select  b.sessiontime from  gtcompmast  a join asptblusermas b on  a.gtcompmastid=b.compcode  where  a.compcode='" + combo_compcode.Text + "' and b.username='" + combousername.Text + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            if (dt.Rows.Count > 0)
            {
                txtsessiontime.Text = dt.Rows[0]["sessiontime"].ToString();
            }
            else
            {
                txtsessiontime.Text = "";
            }
        }

        private void sessionEnableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            txtsessiontime.ReadOnly = false;
        }

        private void comboipaddress_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void Searchs(int id)
        {
            throw new NotImplementedException();
        }

        public void Deletes(int id)
        {
            throw new NotImplementedException();
        }

        private void IPMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            _instance = null;
        }
    }
}

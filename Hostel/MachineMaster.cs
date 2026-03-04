using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Hostel
{
    public partial class MachineMaster : Form,ToolStripAccess
    {
        public MachineMaster()
        {
            InitializeComponent();
      
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }

    

        private static MachineMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Employee em = new Models.Employee();
        Models.UserRights sm = new Models.UserRights();
     
       
        public static MachineMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new MachineMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void ReadOnlys()
        { }
        public void News()
        {
            txtmachineid.Text = "";
            combo_compcode.SelectedIndex = -1;
            comboipaddress.SelectedIndex = -1;
            combowardenname.SelectedIndex = -1;
            checkactive.Checked = true;
            radioBoysHostel.Checked = true;
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
                if (combo_compcode.Text != "" && combowardenname.Text != "" && comboipaddress.Text != "")
                {
                    string chk = "";
                    string hostel, agfchk, flfchk, flfdchk, agfmchk = ""; ;

                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }

                    if (radioAGF.Checked == true) { agfchk = "AGF"; hostel = "AGF"; } else { agfchk = ""; hostel = ""; }
                    if (radioFLF.Checked == true) { flfchk = "FLF"; hostel = "FLF"; } else { flfchk = ""; hostel = ""; }
                    if (radioFLFD.Checked == true) { flfdchk = "FLFD"; hostel = "FLFD"; } else { flfdchk = ""; hostel = ""; }
                    if (radioAGFM.Checked == true) { agfmchk = "AGFM"; hostel = "AGFM"; } else { agfmchk = ""; hostel = ""; }
                    if (radioBoysHostel.Checked == true) { hostel = "WORKING GENTS HOSTEL"; } else {   }
                    if (radioGirlsHostel.Checked == true) { hostel = "WOMENS HOSTEL"; } else { }
                    if (radiosecurity.Checked == true) { hostel = "SECURITY"; } else { }  
                        string sel = "select A.ASPTBLMACHINEMASID FROM ASPTBLMACHINEMAS A  WHERE A.COMPCODE=" + combo_compcode.SelectedValue + " AND A.WARDENNAME=" + combowardenname.SelectedValue + " AND A.IPADDRESS=" + comboipaddress.SelectedValue + " AND A.HOSTELNAME='" + hostel + "'  AND A.ACTIVE='" + chk + "' AND A.AGF='" + agfchk + "' AND A.FLF='" + flfchk + "' AND A.FLFD='" + flfdchk + "' AND A.AGFM='" + agfmchk + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLMACHINEMAS");
                        DataTable dt = ds.Tables["ASPTBLMACHINEMAS"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found     :" + hostel, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (dt.Rows.Count == 0 && Convert.ToInt32("0" + txtmachineid.Text) == 0 || Convert.ToInt32("0" + txtmachineid.Text) == 0)
                        {
                            string ins = "INSERT INTO ASPTBLMACHINEMAS(COMPCODE,WARDENNAME,IPADDRESS,HOSTELNAME,ACTIVE,AGF,FLF,FLFD,AGFM, USERNAME,  MODIFIED,  CREATEDON,  IPADDRESS1)VALUES(" + combo_compcode.SelectedValue + "," + combowardenname.SelectedValue + "," + comboipaddress.SelectedValue + ",'" + hostel + "','" + chk + "','" + agfchk + "','" + flfchk + "','" + flfdchk + "','" + agfmchk + "'," + Class.Users.USERID + ",to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),'" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully    :" + hostel, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        }
                        else
                        {
                            string up = "UPDATE ASPTBLMACHINEMAS SET COMPCODE=" + combo_compcode.SelectedValue + ", WARDENNAME=" + combowardenname.SelectedValue + ",IPADDRESS=" + comboipaddress.SelectedValue + ", HOSTELNAME='" + hostel + "',ACTIVE='" + chk + "', AGF='" + agfchk + "' , FLF='" + flfchk + "' , FLFD='" + flfdchk + "' , AGFM='" + agfmchk + "', USERNAME=" + Class.Users.USERID + ",  MODIFIED=to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),IPADDRESS1='" + Class.Users.IPADDRESS + "' WHERE  ASPTBLMACHINEMASID=" + txtmachineid.Text;
                            Utility.ExecuteNonQuery(up);
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

      public  void GridLoad()
        {
            try
            {
                listmachine.Items.Clear();
                string sel1 = "SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS WARDENNAME,C.MACIP  as IPADDRESS ,A.HOSTELNAME,A.ACTIVE   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN HRMACIPENTRYDET C ON C.HRMACIPENTRYDETID=A.IPADDRESS  JOIN  asptblusermas D ON D.userid=A.WARDENNAME  AND D.COMPCODE=B.GTCOMPMASTID WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND D.USERNAME='"+Class.Users.HUserName+"' ORDER BY 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                DataTable dt = ds.Tables["ASPTBLMACHINEMAS"];
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(myRow["ASPTBLMACHINEMASID"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["WARDENNAME"].ToString());
                        list.SubItems.Add(myRow["IPADDRESS"].ToString());
                        list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        listmachine.Items.Add(list);
                    }
                    lbltotal.Text = "Total Count    :" + listmachine.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void MachineMaster_Load(object sender, EventArgs e)
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

                string sel1 = "SELECT B.HRMACIPENTRYDETID,B.MACIP FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID  JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  JOIN ASPTBLUSERMAS  D ON D.COMPCODE=C.GTCOMPMASTID   WHERE B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES' AND C.COMPCODE='" + Class.Users.HCompcode + "' AND  D.USERNAME='"+ Class.Users.HUserName + "' ORDER BY 2";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "HRMACIPENTRYDET");
                DataTable dt1 = ds1.Tables["HRMACIPENTRYDET"];
              
                if (dt1.Rows.Count >= 0)
                {


                    comboipaddress.DisplayMember = "MACIP";
                    comboipaddress.ValueMember = "HRMACIPENTRYDETID";
                    comboipaddress.DataSource = dt1;


                }
                else
                {
                    comboipaddress.DataSource = null;
                }
                combowardenname.SelectedIndex = -1;
                combo_compcode.SelectedIndex = -1;
                comboipaddress.SelectedIndex = -1;
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
            GridLoad(); 
            this.combo_compcode.Select();
        }

        private void Listmachine_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listmachine.Items.Count > 0)
                {

                    txtmachineid.Text = listmachine.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS WARDENNAME,C.MACIP  as IPADDRESS ,A.HOSTELNAME,A.AGF,A.FLF,A.FLFD,A.AGFM, A.ACTIVE   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN HRMACIPENTRYDET C ON C.HRMACIPENTRYDETID=A.IPADDRESS  JOIN  asptblusermas D ON D.userid=A.WARDENNAME   WHERE A.ASPTBLMACHINEMASID=" + txtmachineid.Text;
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");//C.EMPNAME,C.IDCARDNO,
                    DataTable dt = ds1.Tables["ASPTBLMACHINEMAS"];
                    if (dt.Rows.Count > 0) 
                    {
                        txtmachineid.Text = Convert.ToString(dt.Rows[0]["ASPTBLMACHINEMASID"].ToString());
                        combo_compcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                        combowardenname.Text = Convert.ToString(dt.Rows[0]["WARDENNAME"].ToString());
                        comboipaddress.Text = Convert.ToString(dt.Rows[0]["IPADDRESS"].ToString());
                        if (dt.Rows[0]["HOSTELNAME"].ToString() == "WORKING GENTS HOSTEL") { radioBoysHostel.Checked = true; }

                        if (dt.Rows[0]["AGF"].ToString() == "AGF") { radioAGF.Checked = true; } else { radioAGF.Checked = false; }
                        if (dt.Rows[0]["FLF"].ToString() == "FLF") { radioFLF.Checked = true; } else { radioFLF.Checked = false; }
                        if (dt.Rows[0]["FLFD"].ToString() == "FLFD") { radioFLFD.Checked = true; } else { radioFLFD.Checked = false; }
                        if (dt.Rows[0]["AGFM"].ToString() == "AGFM") { radioAGFM.Checked = true; } else { radioAGFM.Checked = false; }
                        if (dt.Rows[0]["HOSTELNAME"].ToString() == "WOMENS HOSTEL") { radioGirlsHostel.Checked = true; }
                  

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

                    Int64 s = Convert.ToInt64(combo_compcode.SelectedValue);
                    DataTable dt1 = mas.comcode1(s);
                    if (dt1.Rows.Count > 0)
                    {
                        combowardenname.DisplayMember = "USERNAME";
                        combowardenname.ValueMember = "userid";
                        combowardenname.DataSource = dt1;

                    }
                    combowardenname.SelectedIndex = -1;
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
                        combowardenname.DisplayMember = "USERNAME";
                        combowardenname.ValueMember = "userid";
                        combowardenname.DataSource = dt1;

                    }
                    combowardenname.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }


           
        }

        private void IPRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string sel1 = "SELECT B.HRMACIPENTRYDETID,B.MACIP FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID  JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES' AND C.COMPCODE='" + Class.Users.HCompcode + "'  ORDER BY 2";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "HRMACIPENTRYDET");
                DataTable dt1 = ds1.Tables["HRMACIPENTRYDET"];

                if (dt1.Rows.Count >= 0)
                {


                    comboipaddress.DisplayMember = "MACIP";
                    comboipaddress.ValueMember = "HRMACIPENTRYDETID";
                    comboipaddress.DataSource = dt1;


                }
                else
                {
                    comboipaddress.DataSource = null;
                }
            }
            catch(Exception ex)
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
            try
            {
                if (txtmachinesearch.Text != "")
                {
                    listmachine.Items.Clear(); int iGLCount = 1;
                    string sel1 = " SELECT A.ASPTBLMACHINEMASID, B.COMPCODE , D.USERNAME AS WARDENNAME,C.MACIP as IPADDRESS,A.HOSTELNAME,A.ACTIVE    FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN HRMACIPENTRYDET C ON C.HRMACIPENTRYDETID=A.IPADDRESS  JOIN  asptblusermas D ON D.userid=A.WARDENNAME  AND D.COMPCODE=B.GTCOMPMASTID WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND  D.USERNAME LIKE'%" + txtmachinesearch.Text + "%'  OR A.HOSTELNAME LIKE'%" + txtmachinesearch.Text + "%'  OR C.MACIP LIKE'%" + txtmachinesearch.Text + "%'  OR A.ACTIVE LIKE'%" + txtmachinesearch.Text + "%'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                    DataTable dt = ds.Tables["ASPTBLMACHINEMAS"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["ASPTBLMACHINEMASID"].ToString());
                            list.SubItems.Add(myRow["COMPCODE"].ToString());
                            list.SubItems.Add(myRow["WARDENNAME"].ToString());
                            list.SubItems.Add(myRow["HOSTELNAME"].ToString());
                            list.SubItems.Add(myRow["IPADDRESS"].ToString());
                            list.SubItems.Add(myRow["ACTIVE"].ToString());
                            listmachine.Items.Add(list);
                            iGLCount++;
                        }
                        lbltotal.Text = "Total Count    :" + listmachine.Items.Count;
                    }
                    else
                    {
                        listmachine.Items.Clear();
                    }
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
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
            this.Hide();
        }

        public void Deletes()
        {

            string sel1 = "DELETE  FROM ASPTBLMACHINEMAS WHERE ASPTBLMACHINEMASID=" + txtmachineid.Text;
            Utility.ExecuteNonQuery(sel1); GridLoad();MessageBox.Show("Record Deleted"); News();
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

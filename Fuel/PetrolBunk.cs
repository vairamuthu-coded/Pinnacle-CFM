using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Fuel
{
    public partial class PetrolBunk : Form,ToolStripAccess
    {
        private static PetrolBunk _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();


        public PetrolBunk()
        {
            InitializeComponent();      
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
          //  usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }


        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {
                        if (dt1.Rows[r]["READONLY"].ToString() == "T") { } else { this.contextMenuStrip1.Enabled = true; }
                       


                    }
                }

            }
            else
            {
                GlobalVariables.Toolstrip1.Enabled = false;

            }

        }
        public static PetrolBunk Instance
        {
            get { if (_instance == null) _instance = new PetrolBunk(); GlobalVariables.CurrentForm = _instance; return _instance; }
        }
        public void ReadOnlys()
        {

        }
        void empty()
        {
            txtbunkaddress.Text = "";
            txtbunkname.Text = "";
            combocompcode.SelectedIndex = -1;
            txtcontactno.Text = "";
            txtpetrolbunkid.Text = "";
            txtsearch.Text = "";
        }
        public void News()
        {
            empty();
        }
        void CompCodeLoad()
        {
            try
            {
                DataTable dt = mas.findcomcode(Class.Users.HCompcode);
                combocompcode.ValueMember = "GTCOMPMASTID";
                combocompcode.DisplayMember = "COMPCODE";
                combocompcode.DataSource = dt;
            }
            catch (Exception EX) { }

        }
        private void PetrolBunk_Load(object sender, EventArgs e)
        {
            CompCodeLoad(); GridLoad();txtbunkname.Select();
        }
        bool ch = false;
        private bool Checks()
        {
            try
            {



                if (combocompcode.Text=="")
                {
                    MessageBox.Show("CompCode Number is empty");

                    this.combocompcode.Focus(); return false;
                }
                if (txtbunkname.Text == "")
                {
                    MessageBox.Show("Petrol Bunk Name is empty");
                    this.txtbunkname.Focus();
                    return false;
                }
               

                return true;
            }
            catch (Exception EX)
            {
                MessageBox.Show("Error" + EX.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            return ch;

        }
       public void GridLoad()
        {
            try
            {
                listviewpetrol.Items.Clear();
                

                string sel = "SELECT A.ASPTBLPETMASID,B.COMPCODE,A.BUNKNAME,A.BUNKADDRESS,A.CONTACTNO,A.ACTIVE FROM ASPTBLPETMAS A JOIN GTCOMPMAST B ON A.COMPCODE = B.GTCOMPMASTID   WHERE  A.COMPCODE='" + Class.Users.COMPCODE + "' order by 1";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "ASPTBLPETMAS");
                DataTable dt1 = ds1.Tables["ASPTBLPETMAS"];

                if (dt1.Rows.Count > 0)
                {

                    int i = 1;
                    foreach (DataRow myRow in dt1.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLPETMASID"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["BUNKNAME"].ToString());
                        list.SubItems.Add(myRow["BUNKADDRESS"].ToString());
                        list.SubItems.Add(myRow["CONTACTNO"].ToString());
                        if(myRow["ACTIVE"].ToString()=="T")
                        list.SubItems.Add("T");
                        else
                            list.SubItems.Add("F");
                        listviewpetrol.Items.Add(list);
                        i++;
                    }


                }
                else
                {
                   
                    lblcount.Text = "Total Rows Count:  " + listviewpetrol.Items.Count.ToString();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        public void Saves()
        {
            if (Checks() == true)
            {
                string chk = "";
                if (checActive.Checked == true) chk = "T"; else chk = "F";
                string sel = "select ASPTBLPETMASID FROM ASPTBLPETMAS WHERE COMPCODE='" + combocompcode.SelectedValue + "' AND BUNKNAME='" + txtbunkname.Text + "'  AND BUNKADDRESS='" + txtbunkaddress.Text + "' AND CONTACTNO='" + txtcontactno.Text + "' AND ACTIVE='" + chk + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLPETMAS");
                DataTable dt = ds.Tables["ASPTBLPETMAS"];
                if (dt.Rows.Count != 0)
                {
                    MessageBox.Show("Child Record Found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); empty();
                    return;
                }
                else if (dt.Rows.Count != 0 && Convert.ToInt64("0" + txtpetrolbunkid.Text) == 0 || Convert.ToInt64("0" + txtpetrolbunkid.Text) == 0)
                {
                    string ins = "insert into ASPTBLPETMAS(COMPCODE,BUNKNAME,BUNKADDRESS,CONTACTNO,ACTIVE)values('" + combocompcode.SelectedValue + "','" + txtbunkname.Text + "','" + txtbunkaddress.Text + "','" + txtcontactno.Text + "','" + chk + "')";
                    Utility.ExecuteNonQuery(ins);
                    MessageBox.Show("Record Saved Successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad();empty();
                }
                else
                {
                    string up = "update   ASPTBLPETMAS set COMPCODE='" + combocompcode.SelectedValue + "',BUNKNAME='" + txtbunkname.Text + "',BUNKADDRESS='" + txtbunkaddress.Text + "',CONTACTNO='" + txtcontactno.Text + "', ACTIVE='" + chk + "' where ASPTBLPETMASID='" + txtpetrolbunkid.Text + "'";
                    Utility.ExecuteNonQuery(up);
                    MessageBox.Show("Record Updated Successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void Listviewpetrol_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                empty();
                if (listviewpetrol.Items.Count > 0)
                {
                    Models.HRVehicle v1 = new Models.HRVehicle();
                    txtpetrolbunkid.Text = Convert.ToString(listviewpetrol.SelectedItems[0].SubItems[2].Text);
                    string sel1 = "SELECT A.ASPTBLPETMASID,A.COMPCODE,A.BUNKNAME,A.BUNKADDRESS,A.CONTACTNO,A.ACTIVE FROM ASPTBLPETMAS A JOIN GTCOMPMAST B ON A.COMPCODE = B.GTCOMPMASTID   WHERE  A.ASPTBLPETMASID = '" + txtpetrolbunkid.Text + "'  ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLPETMAS");
                    DataTable dt = ds.Tables["ASPTBLPETMAS"];
                    txtpetrolbunkid.Text = Convert.ToString(dt.Rows[0]["ASPTBLPETMASID"].ToString());
                    combocompcode.SelectedValue = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                    txtbunkname.Text = Convert.ToString(dt.Rows[0]["BUNKNAME"].ToString());
                    txtbunkaddress.Text = dt.Rows[0]["BUNKADDRESS"].ToString();
                    txtcontactno.Text = dt.Rows[0]["CONTACTNO"].ToString();
                    if (dt.Rows[0]["ACTIVE"].ToString() == "T") checActive.Checked=true; else checActive.Checked = false;
                   
                   
                    lblcount.Text = "Total Rows ount:  " + listviewpetrol.Items.Count.ToString();
                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Deletes()
        {
            try
            {
                if (txtpetrolbunkid.Text != "")
                {
                    string sel = "SELECT DISTINCT C.ASPTBLPETMASID FROM ASPTBLVEHTOKEN A JOIN GTCOMPMAST b ON b.GTCOMPMASTID=A.COMPCODE  JOIN ASPTBLPETMAS c ON C.COMPCODE=A.COMPCODE and C.COMPCODE=B.GTCOMPMASTID AND C.ASPTBLPETMASID=A.BUNKNAME WHERE C.ASPTBLPETMASID=" + txtpetrolbunkid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLPETMAS");
                    DataTable dt = ds.Tables["ASPTBLPETMAS"];
                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("Child Record Found.Can not Delete.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string del = "delete from   ASPTBLPETMAS where COMPCODE='" + combocompcode.SelectedValue + "' and  ASPTBLPETMASID='" + txtpetrolbunkid.Text + "'";
                        Utility.ExecuteNonQuery(del);
                        MessageBox.Show("Record Deleted Successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                }
                else
                {
                    MessageBox.Show("Invalid", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Question);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Prints_Click(object sender, EventArgs e)
        {

        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
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

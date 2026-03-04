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
    public partial class ReasonMaster : Form,ToolStripAccess
    {
        public ReasonMaster()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
        }

     

        private static ReasonMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Employee em = new Models.Employee();
        Models.UserRights sm = new Models.UserRights();
      
        public static ReasonMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ReasonMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void ReadOnlys()
        {

        }
        public void News()
        {
            txtreason.Text = "";
            txtreasonid.Text="";
            combocompcode.SelectedIndex = -1;
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
        }

        public void Saves()
        {
            try
            {
                string chk = "";

                if (combocompcode.Text != "" && txtreason.Text != "")
                {
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }
                    string sel = "select a.asptblreasonmasid from asptblreasonmas a  where a.compcode=" + combocompcode.SelectedValue + " and a.reason='" + txtreason.Text + "'  and a.active='" + chk + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblreasonmas");
                    DataTable dt = ds.Tables["asptblreasonmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found    :" + txtreason.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    else if (dt.Rows.Count == 0 && Convert.ToInt32("0" + txtreasonid.Text) == 0 || Convert.ToInt32("0" + txtreasonid.Text) == 0)
                    {
                        string ins = "insert into asptblreasonmas(compcode,  reason,active, username,  modified,  createdon,  ipaddress)values(" + combocompcode.SelectedValue + ",'" + txtreason.Text.ToUpper() + "','" + chk + "'," + Class.Users.USERID + ",to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),'" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully     :" + txtreason.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    else
                    {
                        string up = "update asptblreasonmas set compcode=" + combocompcode.SelectedValue + ",  reason='" + txtreason.Text.ToUpper() + "',active='" + chk + "', username=" + Class.Users.USERID + ",  modified=to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),ipaddress='" + Class.Users.IPADDRESS + "' where  asptblreasonmasid=" + txtreasonid.Text;
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated      :" + txtreason.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    GridLoad();
                    News();
                }
                else
                {
                    MessageBox.Show("Pls Enter Data", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void ReasonMaster_Load(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = mas.findcomcode(Class.Users.HCompcode, Class.Users.HUserName);
                if (dt.Rows.Count > 0)
                {
                    combocompcode.DisplayMember = "COMPCODE";
                    combocompcode.ValueMember = "GTCOMPMASTID";
                    combocompcode.DataSource = dt;


                }
                combocompcode.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data Source Not Connected" + ex.Message);
            }
            GridLoad();
            combocompcode.Select();

        }
      public  void GridLoad()
        {
            try
            {
                listreason.Items.Clear();
                string sel1 = "  select a.asptblreasonmasid, b.compcode , a.reason,a.active  from  asptblreasonmas a  join   gtcompmast b on b.gtcompmastid = a.compcode where b.compcode='" + Class.Users.HCompcode + "' order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblreasonmas");
                DataTable dt = ds.Tables["asptblreasonmas"];
                if (dt.Rows.Count > 0)
                {
                    int mycount = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = mycount.ToString();
                        list.SubItems.Add(myRow["asptblreasonmasid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["reason"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        if (mycount % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        listreason.Items.Add(list);
                        mycount++;
                    }
                    lblreasontotal.Text = "Total Rows    :" + listreason.Items.Count;
;                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Listreason_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listreason.Items.Count > 0)
                {
                    txtreasonid.Text = listreason.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.asptblreasonmasid, b.compcode , a.reason,a.active  from  asptblreasonmas a  join   gtcompmast b on b.gtcompmastid = a.compcode  where a.asptblreasonmasid=" + txtreasonid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblreasonmas");
                    DataTable dt = ds.Tables["asptblreasonmas"];
                    if (dt.Rows.Count > 0)
                    {
                        txtreasonid.Text = Convert.ToString(dt.Rows[0]["asptblreasonmasid"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        txtreason.Text = Convert.ToString(dt.Rows[0]["reason"].ToString());

                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        //combocompcode.Text = "COMPCODE";
                        //combocompcode.ValueMember = "GTCOMPMASTID";
                        //combocompcode.DataSource = dt;
                         //combocompcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TxtreasonSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtreasonSearch.Text != "")
                {
                    listreason.Items.Clear(); int iGLCount = 1;
                    string sel1 = "  select  a.asptblreasonmasid,b.compcode,a.reason,a.active from   asptblreasonmas a join   gtcompmast b on b.gtcompmastid = a.compcode  where a.reason like'%" + txtreasonSearch.Text + "%'  or a.active like'%" + txtreasonSearch.Text + "%'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblreasonmas");
                    DataTable dt = ds.Tables["asptblreasonmas"];
                    if (dt.Rows.Count > 0)
                    {
                        
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["asptblreasonmasid"].ToString());
                            list.SubItems.Add(myRow["COMPCODE"].ToString());
                            list.SubItems.Add(myRow["REASON"].ToString());                   
                            list.SubItems.Add(myRow["active"].ToString());
                            listreason.Items.Add(list);
                            iGLCount++;
                        }
                        lblreasontotal.Text = "Total Count    :" + listreason.Items.Count;
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        public void Deletes()
        {
            string sel1 = "DELETE  FROM asptblreasonmas A  WHERE A.asptblreasonmasid=" + txtreasonid.Text;
            Utility.ExecuteNonQuery(sel1); GridLoad(); News();MessageBox.Show("Record Deleted Successfully");
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
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

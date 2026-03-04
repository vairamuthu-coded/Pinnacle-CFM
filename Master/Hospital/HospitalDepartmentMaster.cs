using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master.Hospital
{
    public partial class HospitalDepartmentMaster : Form, ToolStripAccess
    {
        private static HospitalDepartmentMaster _instance;
        Models.Master mas = new Models.Master();       
        Models.UserRights sm = new Models.UserRights();
       ListView listfilter = new ListView();

        int i = 0;
       
        public static HospitalDepartmentMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new HospitalDepartmentMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public HospitalDepartmentMaster()
        {
            InitializeComponent();

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
         
            GlobalVariables.CurrentForm = this;
           
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;          
            this.BackColor = Class.Users.BackColors;
        }

      

        private void HopitalDepartmentMaster_Load(object sender, EventArgs e)
        {
           empty(); 
           
        }

      

        private bool ValidateContact(TextBox s)
        {
            if (!Regex.Match(s.Text, "^[A-Z][a-zA-Z]*$").Success)
            {

                s.Text.Remove(s.Text.Length - 1);

            }
            return true;
        }

        public void Saves()
        {
            try
            {
                if (txthospitaldept.Text != "")
                {
                   

                       
                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "SELECT a.asptblhosdeptmasid FROM asptblhosdeptmas a    WHERE department='" + txthospitaldept.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblhosdeptmas");
                        DataTable dt = ds.Tables["asptblhosdeptmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txthospitaldept.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txthospitaldeptid.Text) == 0 || Convert.ToInt32("0" + txthospitaldeptid.Text) == 0)
                        {
                            string ins = "insert into asptblhosdeptmas(department,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txthospitaldept.Text + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txthospitaldept.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblhosdeptmas  set   department='" + txthospitaldept.Text + "' , active='" + chk + "' , modifiedby='" + Class.Users.USERID + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblhosdeptmasID='" + txthospitaldeptid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txthospitaldept.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void HopitalDepartmentMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        public void News()
        {
            GridLoad();
            empty();
        }
        private void empty()
        {
            txthospitaldeptid.Text = "";
            txthospitaldept.Text = "";
           
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            txthospitaldept.Select();


        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptblhosdeptmasID, a.department , a.active  from  asptblhosdeptmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
                DataTable dt = ds.Tables["asptblhosdeptmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                  

                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["asptblhosdeptmasID"].ToString());
                        list.SubItems.Add(myRow["department"].ToString());
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

       
        public void Searchs(int EditID)
        {
            txthospitaldeptid.Text = EditID.ToString();
            string sel1 = " select a.asptblhosdeptmasID, a.department , a.active    from  asptblhosdeptmas a    where a.asptblhosdeptmasID=" + txthospitaldeptid.Text;
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
            DataTable dt = ds.Tables["asptblhosdeptmas"];

            if (dt.Rows.Count > 0)
            {
                txthospitaldeptid.Text = Convert.ToString(dt.Rows[0]["asptblhosdeptmasID"].ToString());
                txthospitaldept.Text = Convert.ToString(dt.Rows[0]["department"].ToString());
                if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                
            }

        }

            private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txthospitaldeptid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.asptblhosdeptmasID, a.department , a.active    from  asptblhosdeptmas a    where a.asptblhosdeptmasID=" + txthospitaldeptid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
                    DataTable dt = ds.Tables["asptblhosdeptmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txthospitaldeptid.Text = Convert.ToString(dt.Rows[0]["asptblhosdeptmasID"].ToString());
                        txthospitaldept.Text = Convert.ToString(dt.Rows[0]["department"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Search()
        {
            try
            {
                if (txtsearch.Text.ToUpper() != "")
                {
                    listView1.Items.Clear(); int iGLCount = 1;
                    string sel1 = "  SELECT  a.asptblhosdeptmasID,a.department,a.active from asptblhosdeptmas a  where a.department LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
                    DataTable dt = ds.Tables["asptblhosdeptmas"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["asptblhosdeptmasID"].ToString());
                            list.SubItems.Add(myRow["department"].ToString());
                            list.SubItems.Add(myRow["active"].ToString());
                            listView1.Items.Add(list);
                            iGLCount++;
                        }
                        lbltotal.Text = "Total Count    :" + listView1.Items.Count;
                    }
                    else
                    {
                        listView1.Items.Clear();
                    }
                }
                else
                {

                    listView1.Items.Clear();
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Deletes()
        {
            if (txthospitaldeptid.Text != "")
            {
                string sel1 = "select a.asptblhosdeptmasID from asptblhosdeptmas a join asptbldocmas b on a.asptblhosdeptmasID=b.department where a.asptblhosdeptmasID='" + txthospitaldeptid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblhosdeptmas");
                DataTable dt = ds.Tables["asptblhosdeptmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txthospitaldept.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblhosdeptmas where asptblhosdeptmasID='" + Convert.ToInt64("0" + txthospitaldeptid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txthospitaldept.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        public void Prints()
        {
           
        }

        public void Searchs()
        {
           
        }



        public void ReadOnlys()
        {
          
        }

        public void Imports()
        {
            throw new NotImplementedException();
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
            empty();
            GlobalVariables.HeaderName.Text = "";
            HospitalDepartmentMaster dtp = new HospitalDepartmentMaster();
            dtp.Close();
        
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
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
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
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
                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            if (item0 % 2 == 0)
                            {
                                item.BackColor = Color.White;
                            }
                            else
                            {
                                item.BackColor = Color.WhiteSmoke;
                            }
                            listView1.Items.Add(list);
                            item0++;
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

        private void txthospitaldept_Leave(object sender, EventArgs e)
        {
            txthospitaldept.BackColor = Class.Users.Color1;
        }

        private void txthospitaldept_Enter(object sender, EventArgs e)
        {
            txthospitaldept.BackColor = Class.Users.Color2;
        }
    }
}


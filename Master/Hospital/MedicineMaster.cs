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
    public partial class MedicineMaster : Form,ToolStripAccess
    {
        private static MedicineMaster _instance;
        Models.Master mas = new Models.Master();
        Models.CTS.StudentMasterModel em = new Models.CTS.StudentMasterModel();
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
        public static MedicineMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new MedicineMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public MedicineMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
         
          
        }
        private void MedicineMaster_Load(object sender, EventArgs e)
        {
           
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
                if (txtmedicine.Text != "")
                {



                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "SELECT a.asptblmedicalmasid FROM asptblmedicalmas a    WHERE medicine='" + txtmedicine.Text + "' and Description='" + txtdescriptions.Text.ToUpper() + "' and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblmedicalmas");
                    DataTable dt = ds.Tables["asptblmedicalmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtmedicine.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtmedicineid.Text) == 0 || Convert.ToInt32("0" + txtmedicineid.Text) == 0)
                    {
                        string ins = "insert into asptblmedicalmas(medicine,Description,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtmedicine.Text.ToUpper() + "','" + txtdescriptions.Text.ToUpper() + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtmedicine.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblmedicalmas  set   medicine='" + txtmedicine.Text.ToUpper() + "' , Description='" + txtdescriptions.Text.ToUpper() + "', active='" + chk + "' , modifiedby='" + Class.Users.USERID + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblmedicalmasID='" + txtmedicineid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtmedicine.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void MedicineMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            txtmedicineid.Text = "";
            txtmedicine.Text = "";
            butheader.Text = Class.Users.ScreenName;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            this.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            txtmedicine.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.asptblmedicalmasID, a.medicine ,a.Description, a.active  from  asptblmedicalmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmedicalmas");
                DataTable dt = ds.Tables["asptblmedicalmas"];
                if (dt.Rows.Count > 0 || dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblmedicalmasID"].ToString());
                        list.SubItems.Add(myRow["medicine"].ToString());
                        list.SubItems.Add(myRow["Description"].ToString());
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
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtmedicineid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.asptblmedicalmasID, a.medicine ,a.Description, a.active    from  asptblmedicalmas a    where a.asptblmedicalmasID=" + txtmedicineid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmedicalmas");
                    DataTable dt = ds.Tables["asptblmedicalmas"];
                    if (dt.Rows.Count > 0)
                    {
                        txtmedicineid.Text = Convert.ToString(dt.Rows[0]["asptblmedicalmasID"].ToString());
                        txtmedicine.Text = Convert.ToString(dt.Rows[0]["medicine"].ToString());
                        txtdescriptions.Text = Convert.ToString(dt.Rows[0]["Description"].ToString());
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
                    string sel1 = "  SELECT  a.asptblmedicalmasID,a.medicine,a.Description,a.active from asptblmedicalmas a  where a.medicine LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmedicalmas");
                    DataTable dt = ds.Tables["asptblmedicalmas"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["asptblmedicalmasID"].ToString());
                            list.SubItems.Add(myRow["medicine"].ToString());
                            list.SubItems.Add(myRow["Description"].ToString());
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
            if (txtmedicineid.Text != "")
            {
                string sel1 = "select a.asptblmedicalmasID from asptblmedicalmas a join asptbldiagnosismas b on a.asptblmedicalmasID=b.medicine where a.asptblmedicalmasID='" + txtmedicineid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmedicalmas");
                DataTable dt = ds.Tables["asptblmedicalmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtmedicine.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblmedicalmas where asptblmedicalmasID='" + Convert.ToInt64("0" + txtmedicineid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtmedicine.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
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
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        listView1.Items.Add(list);
                    }

                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("---" + ex.ToString());
            }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();empty();
        }

        private void txtmedicine_Leave(object sender, EventArgs e)
        {
            txtmedicine.BackColor = Class.Users.Color1;
        }

        private void txtmedicine_Enter(object sender, EventArgs e)
        {
            txtmedicine.BackColor = Class.Users.Color2;
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}


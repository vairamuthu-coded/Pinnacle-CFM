using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Thulliam.Master
{
    public partial class VarietyItemMaster : Form
    {
        public VarietyItemMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            lbl_Header.Text = Class.Users.ScreenName;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static VarietyItemMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        ThulliamMdi mdi = new ThulliamMdi();
        public static VarietyItemMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new VarietyItemMaster();
                return _instance;
            }
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

                        if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                        if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                        if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                        if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                        if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                        if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                        if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                        if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                        if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                        if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                        if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                        if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                        if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                        if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }

                    }
                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        private void VarietyItemMaster_Load(object sender, EventArgs e)
        {
            News_Click(sender, e);
        }

        private void Saves_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtitem.Text == "")
                {
                    MessageBox.Show("Godown Name is empty " + " Alert " + txtitem.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtitem.Text != "")
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsStringNumbericslace(txtitem.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select asptblitemmastid    from  asptblitemmast    WHERE itemname='" + txtitem.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblitemmast");
                        DataTable dt = ds.Tables["asptblitemmast"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtitem.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtitemid.Text) == 0 || Convert.ToInt32("0" + txtitemid.Text) == 0)
                        {
                            string ins = "insert into asptblitemmast(itemname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtitem.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtitem.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            gridload(); empty();
                        }
                        else
                        {
                            string up = "update  asptblitemmast  set   itemname='" + txtitem.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblitemmastid='" + txtitemid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtitem.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            gridload();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("itemname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void VarietyItemMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void News_Click(object sender, EventArgs e)
        {

            empty(); gridload();
        }
        private void empty()
        {
            txtitemid.Text = "";
            txtitem.Text = "";
            checkactive.Checked = false;
        }
        private void gridload()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "   SELECT A.asptblitemmastid, A.itemname , a.active  FROM  asptblitemmast a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblitemmast");
                DataTable dt = ds.Tables["asptblitemmast"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["asptblitemmastid"].ToString());
                        list.SubItems.Add(myRow["itemname"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listView1.Items.Add(list);
                        i++;
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

                    txtitemid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.asptblitemmastid, a.itemname , a.active    from  asptblitemmast a    where a.asptblitemmastid=" + txtitemid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblitemmast");
                    DataTable dt = ds.Tables["asptblitemmast"];

                    if (dt.Rows.Count > 0)
                    {
                        txtitemid.Text = Convert.ToString(dt.Rows[0]["asptblitemmastid"].ToString());
                        txtitem.Text = Convert.ToString(dt.Rows[0]["itemname"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }


                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtsearch.Text.ToUpper() != "")
                {
                    listView1.Items.Clear(); int iGLCount = 1;
                    string sel1 = "  SELECT  a.asptblitemmastid,a.itemname,a.active from asptblitemmast a  where a.itemname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblitemmast");
                    DataTable dt = ds.Tables["asptblitemmast"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["asptblitemmastid"].ToString());
                            list.SubItems.Add(myRow["itemname"].ToString());
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
                    gridload();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

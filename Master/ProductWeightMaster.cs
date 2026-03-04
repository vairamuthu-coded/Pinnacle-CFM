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
    public partial class ProductWeightMaster : Form
    {
        public ProductWeightMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            lbl_Header.Text = Class.Users.ScreenName;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToShortTimeString().ToString());
        }

        private static ProductWeightMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        ThulliamMdi mdi = new ThulliamMdi();
        public static ProductWeightMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ProductWeightMaster();
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
        private void ProductWeightMaster_Load(object sender, EventArgs e)
        {
            News_Click(sender, e);
        }

        public DataTable loadproduct()
        {
            string sel = "SELECT  a.asptblproductmasid,a.productname from asptblproductmas a where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproductmas");
            DataTable dt = ds.Tables["asptblproductmas"];
            comboproductname.DataSource = dt;
            comboproductname.DisplayMember = "productname";
            comboproductname.ValueMember = "asptblproductmasid";
            return dt;
        }
       
        private void Saves_Click(object sender, EventArgs e)
        {
            try
            {
                Models.Validate va = new Models.Validate();
                if (comboproductname.SelectedValue == null)
                {
                    MessageBox.Show("'Product Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.comboproductname.Focus();
                    return;
                }
                if (txtnetweight.Text == "")
                {
                    MessageBox.Show("'Netweight Weight  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtnetweight.Focus();
                    return;
                }
                if (txtgrossweight.Text == "")
                {
                    MessageBox.Show("'Gross Weight  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtgrossweight.Focus(); return;
                }
                else
                {

                    if (va.IsInteger(comboproductname.SelectedValue.ToString()) == true && va.IsDecimal(txtnetweight.Text) == true && va.IsDecimal(txtgrossweight.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.asptblproductweightmasid    from  asptblproductweightmas a join asptblproductmas b on a.productname=b.asptblproductmasid   WHERE a.productname='" + comboproductname.SelectedValue + "' and a.netweight='" + txtnetweight.Text + "'and a.grossweight='" + txtgrossweight.Text + "' and a.active='" + chk + "' and  a.productname1='" + txtproductname1.Text + "' and a.asptblproductweightmasid='" + txtproductweightid.Text + "';";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproductweightmas");
                        DataTable dt = ds.Tables["asptblproductweightmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + comboproductname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtproductweightid.Text) == 0 || Convert.ToInt32("0" + txtproductweightid.Text) == 0)
                        {
                            string ins = "insert into asptblproductweightmas(productname, netweight,grossweight,active,productname1,createdby,modifiedby,ipaddress)  VALUES('" + comboproductname.SelectedValue + "','" + txtnetweight.Text + "','" + txtgrossweight.Text + "','" + chk + "','" + txtproductname1.Text + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + comboproductname.SelectedValue, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            gridload(); empty();
                        }
                        else
                        {
                            string up = "update  asptblproductweightmas  set productname='" + comboproductname.SelectedValue + "', netweight='" + txtnetweight.Text + "',grossweight='" + txtgrossweight.Text + "' , active='" + chk + "' , productname1='"+txtproductname1.Text+"',modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblproductweightmasid='" + txtproductweightid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + comboproductname.SelectedValue, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            gridload();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid Data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;

                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("grossweight " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void ProductWeightMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void News_Click(object sender, EventArgs e)
        {

            gridload(); loadproduct(); empty();
        }
        private void empty()
        {
            txtproductweightid.Text = "";
            comboproductname.SelectedIndex = -1;;
            txtnetweight.Text = ""; ; txtgrossweight.Text = "";
            txtgrossweight.Text = ""; txtsearch.Text = "";txtproductname1.Text = "";
            checkactive.Checked = false;
        }
        private void gridload()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = " select a.asptblproductweightmasid,b.productname, a.netweight, a.grossweight , a.active,a.productname1    from  asptblproductweightmas a  join asptblproductmas b on a.productname=b.asptblproductmasid    order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproductweightmas");
                DataTable dt = ds.Tables["asptblproductweightmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblproductweightmasid"].ToString());
                        list.SubItems.Add(myRow["productname"].ToString());
                        list.SubItems.Add(myRow["netweight"].ToString());
                        list.SubItems.Add(myRow["grossweight"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        list.SubItems.Add(myRow["productname1"].ToString());
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

                    txtproductweightid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = "select a.asptblproductweightmasid,b.productname, a.netweight, a.grossweight , a.active,a.productname1    from  asptblproductweightmas a  join asptblproductmas b on a.productname=b.asptblproductmasid where a.asptblproductweightmasid=" + txtproductweightid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproductweightmas");
                    DataTable dt = ds.Tables["asptblproductweightmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtproductweightid.Text = Convert.ToString(dt.Rows[0]["asptblproductweightmasid"].ToString());
                        comboproductname.Text = Convert.ToString(dt.Rows[0]["productname"].ToString());
                        txtnetweight.Text = Convert.ToString(dt.Rows[0]["netweight"].ToString());
                        txtgrossweight.Text = Convert.ToString(dt.Rows[0]["grossweight"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                      //  txtproductname1.Text = Convert.ToString(dt.Rows[0]["productname1"].ToString())+;

                        txtproductname1.Text = comboproductname.Text + "-" + txtnetweight.Text;
                    }
                    loadproduct();
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
                    string sel1 = "select a.asptblproductweightmasid,b.productname, a.netweight, a.grossweight , a.active,a.productname1    from  asptblproductweightmas a  join asptblproductmas b on a.productname=b.asptblproductmasid  where b.productname LIKE'%" + txtsearch.Text.ToUpper() + "%'  || a.netweight LIKE'%" + txtsearch.Text.ToUpper() + "%'  || a.grossweight LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproductweightmas");
                    DataTable dt = ds.Tables["asptblproductweightmas"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(iGLCount.ToString());
                            list.SubItems.Add(myRow["asptblproductweightmasid"].ToString());
                            list.SubItems.Add(myRow["productname"].ToString());
                            list.SubItems.Add(myRow["netweight"].ToString());
                            list.SubItems.Add(myRow["grossweight"].ToString());
                            list.SubItems.Add(myRow["active"].ToString());
                            list.SubItems.Add(myRow["productname1"].ToString());
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

      
        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            loadproduct();
        }

        private void Txtgrossweight_TextChanged(object sender, EventArgs e)
        {
            txtproductname1.Text = "";
            txtproductname1.Text = comboproductname.Text + "-" + txtnetweight.Text; checkactive.Checked = true;
        }

        private void Txtnetweight_TextChanged(object sender, EventArgs e)
        {
            txtproductname1.Text = "";
            txtproductname1.Text = comboproductname.Text + "-" + txtnetweight.Text; checkactive.Checked = true;
        }

        private void Deletes_Click(object sender, EventArgs e)
        {
            if (txtproductweightid.Text != "")
            {
                string sel1 = "select a.asptblproductweightmasid from asptblproductweightmas a join asptbldeliverydet b on a.asptblproductweightmasid=b.productname where a.asptblproductweightmasid='" + txtproductweightid.Text+"';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproductweightmas");
                DataTable dt = ds.Tables["asptblproductweightmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + comboproductname.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblproductweightmas where asptblproductweightmasid='" + Convert.ToInt64("0" + txtproductweightid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + comboproductname.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    gridload(); empty();
                }
            }
        }
    }
}

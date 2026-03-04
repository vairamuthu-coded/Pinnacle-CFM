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
    public partial class CategoryMaster : Form,ToolStripAccess
    {
        public CategoryMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
          
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
 
        }

        private static CategoryMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();

        public static CategoryMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new CategoryMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                   
                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        private void CategoryMaster_Load(object sender, EventArgs e)
        {
            News();GridLoad();
        }

        public void Saves()
        {
            try
            {
                if (txtcategory.Text == "")
                {
                    MessageBox.Show("category Name is empty " + " Alert " + txtcategory.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtcategory.Text != "")
                {


                    string chk = "",sew="",qc="";
                    if (checksewing.Checked == true) { sew = "T"; } else { sew = "F"; checksewing.Checked = false; }
                    if (checkqc.Checked == true) { qc = "T"; } else { qc = "F"; checkqc.Checked = false; }
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblcatmasid    from  asptblcatmas    WHERE category='" + txtcategory.Text + "' and sewingcategory='" + sew + "' and qccategory='" + qc + "' and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblcatmas");
                    DataTable dt = ds.Tables["asptblcatmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtcategory.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtcategoryid.Text) == 0 || Convert.ToInt32("0" + txtcategoryid.Text) == 0)
                    {
                        string ins = "insert into asptblcatmas(category,sewingcategory,qccategory,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtcategory.Text.ToUpper() + "','" + sew + "','" + qc + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtcategory.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblcatmas  set   category='" + txtcategory.Text.ToUpper() + "' , sewingcategory='" + sew + "',qccategory='" + qc + "',active='" + chk + "' ,compcode='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblcatmasid='" + txtcategoryid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtcategory.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty();
                    }

                }
                else
                {
                    MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("category " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void categoryMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       

        public void News()
        {

            empty(); GridLoad();
        }
        private void empty()
        {
            txtcategoryid.Text = "";
            txtcategory.Text = "";
            checksewing.Checked = false;
            checkqc.Checked = false;
            checkactive.Checked = true;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "   SELECT asptblcatmasid, category ,sewingcategory as sew,qccategory as qc, active  FROM  asptblcatmas    order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcatmas");
                DataTable dt = ds.Tables["asptblcatmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblcatmasid"].ToString());
                        list.SubItems.Add(myRow["category"].ToString());
                        list.SubItems.Add(myRow["sew"].ToString());
                        list.SubItems.Add(myRow["qc"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
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

                    txtcategoryid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select asptblcatmasid, category ,sewingcategory,qccategory, active    from  asptblcatmas a    where a.asptblcatmasid=" + txtcategoryid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcatmas");
                    DataTable dt = ds.Tables["asptblcatmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtcategoryid.Text = Convert.ToString(dt.Rows[0]["asptblcatmasid"].ToString());
                        txtcategory.Text = Convert.ToString(dt.Rows[0]["category"].ToString());
                        if (dt.Rows[0]["sewingcategory"].ToString() == "T") { checksewing.Checked = true; } else {  checksewing.Checked = false; }

                        if (dt.Rows[0]["qccategory"].ToString() == "T") {checkqc.Checked = true; } else {  checkqc.Checked = false; }

                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }


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
                int i = 1;

                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length > 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text))
                        {
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                        }
                        i++;

                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {
                    i = 1;
                    ListView ll = new ListView();

                    listView1.Items.Clear(); item0 = listfilter.Items.Count;
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.SubItems.Add(item.SubItems[5].Text);
                        list.SubItems.Add(item.SubItems[6].Text);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listView1.Items.Add(list);
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }

        }

        public void Delete()
        {
            if (txtcategoryid.Text != "")
            {
                string sel1 = "select a.asptblcatmasid from asptblcatmas a join gtstatemast b on a.asptblcatmasid=b.country where a.asptblcatmasid='" + txtcategoryid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcatmas");
                DataTable dt = ds.Tables["asptblcatmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcategory.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblcatmas where asptblcatmasid='" + Convert.ToInt64("0" + txtcategoryid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtcategory.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void Deletes()
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

       
    }
}

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
    public partial class DefectMaster : Form,ToolStripAccess
    {
        public DefectMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());

        }

        private static DefectMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();

        public static DefectMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DefectMaster();
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
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                    }
                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        private void DefectMaster_Load(object sender, EventArgs e)
        {
            News();GridLoad(); qcdefectload();


        }

       

        private void DefectMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

     
        public void qcdefectload()
        {
            try
            {
                string sel = "select asptblcatmasid,category from  asptblcatmas  where category='T' order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblcatmas");
                DataTable dt = ds.Tables["asptblcatmas"];
              
                comboqccategory.DisplayMember = "category";
                comboqccategory.ValueMember = "asptblcatmasid";
                comboqccategory.DataSource = dt;

            }
            catch (Exception EX)
            { }
        }
      
        private void empty()
        {
            txtdefectid.Text = "";
            txtdefectname.Text = "";
            comboqccategory.SelectedIndex = -1; comboqccategory.Text = "";
            txtdefectno.Text = "";
            checkactive.Checked = false;
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font= Class.Users.FontName;
        }

       
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtdefectid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptbldefmasid, a.defectname ,b.category,a.defectno, a.active  from  asptbldefmas a join asptblcatmas b on a.category=b.asptblcatmasid   where a.asptbldefmasid=" + txtdefectid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldefmas");
                    DataTable dt = ds.Tables["asptbldefmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtdefectid.Text = Convert.ToString(dt.Rows[0]["asptbldefmasid"].ToString());
                        txtdefectname.Text = Convert.ToString(dt.Rows[0]["defectname"].ToString());
                        comboqccategory.Text = Convert.ToString(dt.Rows[0]["category"].ToString());
                        txtdefectno.Text = Convert.ToString(dt.Rows[0]["defectno"].ToString());
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


                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length > 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text) || item.SubItems[4].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {

                    ListView ll = new ListView();

                    listView1.Items.Clear(); listView1.BackColor = System.Drawing.Color.LightGray;
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();


                        list.Text = item.SubItems[0].Text;
                        list.SubItems.Add(item.SubItems[1].Text);
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.SubItems.Add(item.SubItems[5].Text);
                        list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);



                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
            //try
            //{
            //    if (txtsearch.Text.ToUpper() != "")
            //    {
            //        listView1.Items.Clear(); int iGLCount = 1;
            //        string sel1 = "  SELECT  a.asptbldefmasid,a.defectname,a.active from asptbldefmas a  where a.defectname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldefmas");
            //        DataTable dt = ds.Tables["asptbldefmas"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptbldefmasid"].ToString());
            //                list.SubItems.Add(myRow["defectname"].ToString());
            //                list.SubItems.Add(myRow["active"].ToString());
            //                listView1.Items.Add(list);
            //                iGLCount++;
            //            }
            //            lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            //        }
            //        else
            //        {
            //            listView1.Items.Clear();
            //        }
            //    }
            //    else
            //    {

            //        listView1.Items.Clear();
            //        GridLoad();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

       

        private void refeshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void refeshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            qcdefectload();
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
            if (txtdefectid.Text != "")
            {
                string sel1 = "select a.asptbldefmasid from asptbldefmas a join gtstatemast b on a.asptbldefmasid=b.country where a.asptbldefmasid='" + txtdefectid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldefmas");
                DataTable dt = ds.Tables["asptbldefmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtdefectname.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptbldefmas where asptbldefmasid='" + Convert.ToInt64("0" + txtdefectid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtdefectname.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
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

        public void News()
        {
            empty();GridLoad();
        }

        public void Saves()
        {
            try
            {
                if (txtdefectname.Text == "")
                {
                    MessageBox.Show("defectname Name is empty " + " Alert " + txtdefectname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtdefectname.Text != "")
                {
                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptbldefmasid    from  asptbldefmas    WHERE defectname='" + txtdefectname.Text + "' and category='" + comboqccategory.SelectedValue + "' and defectno='" + txtdefectno.Text + "' and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbldefmas");
                    DataTable dt = ds.Tables["asptbldefmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtdefectname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtdefectid.Text) == 0 || Convert.ToInt32("0" + txtdefectid.Text) == 0)
                    {
                        string ins = "insert into asptbldefmas(defectname,category,defectno,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtdefectname.Text.ToUpper() + "','" + comboqccategory.SelectedValue + "','" + txtdefectno.Text + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtdefectname.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptbldefmas  set   defectname='" + txtdefectname.Text.ToUpper() + "' , category='" + comboqccategory.SelectedValue + "',defectno='" + txtdefectno.Text + "',active='" + chk + "' ,compcode='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbldefmasid='" + txtdefectid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtdefectname.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("defectname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "SELECT A.asptbldefmasid, A.defectname ,b.category, a.active  FROM  asptbldefmas a join asptblcatmas b on a.category=asptblcatmasid  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldefmas");
                DataTable dt = ds.Tables["asptbldefmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldefmasid"].ToString());
                        list.SubItems.Add(myRow["defectname"].ToString());
                        list.SubItems.Add(myRow["category"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

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
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master.Hospital
{
    public partial class PrescriptionPopUp : Form,ToolStripAccess
    {
       
        Models.Master mas = new Models.Master();
      
        Models.UserRights sm = new Models.UserRights();
         ListView listfilter = new ListView();       
        int i = 0;        
        public PrescriptionPopUp()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
           
            panel3.BackColor = SystemColors.Highlight;
            listView1.Font = Class.Users.FontName;
        }

        public void News()
        {
            GridLoad();
        }

        public void Saves()
        {
           
        }

        public void Prints()
        {
            
        }

        public void Searchs()
        {
            
        }

        public void Deletes()
        {
            
        }

        public void ReadOnlys()
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
            GlobalVariables.MdiPanel.Show(); News();
            this.Hide(); GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
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
                        if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[10].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[11].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[12].Text);
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

                    listView1.Items.Clear(); item0 = 0;i = 1;
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = "";
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.SubItems.Add(item.SubItems[5].Text);
                        list.SubItems.Add(item.SubItems[6].Text);
                        list.SubItems.Add(item.SubItems[7].Text);
                        list.SubItems.Add(item.SubItems[8].Text);
                        list.SubItems.Add(item.SubItems[9].Text);
                        list.SubItems.Add(item.SubItems[10].Text);
                        list.SubItems.Add(item.SubItems[11].Text);
                        list.SubItems.Add(item.SubItems[12].Text);
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        i++;
                        listView1.Items.Add(list);
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("---" + ex.ToString());
            }
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    Class.Users.Paramid = Convert.ToInt64(listView1.SelectedItems[0].SubItems[2].Text);
                    this.Hide();
                 //   GlobalVariables.CurrentForm = "{PSSRobot.Master.Hospital.DiagnosisMaster, Text: DiagnosisMaster}";
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }

        private void ListViewMaster_Load(object sender, EventArgs e)
        {

            News();
            timer1.Start();
            txtsearch.Focus(); txtsearch.Select();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            Class.Users.Enabled = true;
            this.Close(); Class.Users.Bisconnectclear = false; Class.Users.Paramlistivew = "";
           
        }

        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
               

               ListViewItem it2 = new ListViewItem();
                if (e.Item.Checked == true)
                {


                    e.Item.SubItems[12].Text = "✔";
                    it2.SubItems.Add(e.Item.SubItems[1].Text);
                    it2.SubItems.Add(e.Item.SubItems[2].Text);
                    it2.SubItems.Add(e.Item.SubItems[3].Text);
                    it2.SubItems.Add(e.Item.SubItems[4].Text);
                    it2.SubItems.Add(e.Item.SubItems[5].Text);
                    it2.SubItems.Add(e.Item.SubItems[6].Text);
                    it2.SubItems.Add(e.Item.SubItems[7].Text);
                    it2.SubItems.Add(e.Item.SubItems[8].Text);
                    it2.SubItems.Add(e.Item.SubItems[9].Text);
                    it2.SubItems.Add(e.Item.SubItems[10].Text);
                    it2.SubItems.Add(e.Item.SubItems[11].Text);
                    it2.SubItems.Add(e.Item.SubItems[12].Text);
                    Class.Users.Staticallip.Items.Add(it2);

                    Cursor = Cursors.Default;
                }
                if (e.Item.Checked == false && e.Item.SubItems[12].Text == "✔")
                {


                    e.Item.SubItems[12].Text = "✖";
                    for (int c = 0; c < Class.Users.Staticallip.Items.Count; c++)
                    {
                        if (e.Item.SubItems[2].Text == Class.Users.Staticallip.Items[c].SubItems[2].Text)
                        {
                            Class.Users.Staticallip.Items[c].Remove();
                            c--;
                        }
                    }
                    Cursor = Cursors.Default;
                }




            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {

                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select '' as asptbldiatestid, b.labtest,a.labtestitem,a.asptbllabtestitemmasid,b.asptbllabtestmasid a.active  from  asptbllabtestitemmas a  join asptbllabtestmas b on a.labtest=b.asptbllabtestmasid  order by 2";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
                DataTable dt = ds.Tables["asptbllabtestitemmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();

                        list.Text = "";
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldiatestid"].ToString());
                        list.SubItems.Add(myRow["labtest"].ToString());
                        list.SubItems.Add(myRow["labtestitem"].ToString());
                        list.SubItems.Add(myRow["asptbllabtestitemmasid"].ToString());
                        list.SubItems.Add("asptbllabtestmasid");
                        list.SubItems.Add("0");
                        list.SubItems.Add("0");
                        list.SubItems.Add("0");
                        list.SubItems.Add("");
                        list.SubItems.Add(""); list.SubItems.Add("");
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        i++;
                        listView1.Items.Add(list);
                    }
                    lbltotal.Text = "Test Count    :" + listView1.Items.Count;
                   
                }
                else
                {
                    lbltotal.Refresh();
                    lbltotal.Text = "No data Found.";
                   
                }
            }
            else
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select '' as asptbldiatestid, b.labtest,a.labtestitem,a.asptbllabtestitemmasid,b.asptbllabtestmasid a.active  from  asptbllabtestitemmas a  join asptbllabtestmas b on a.labtest=b.asptbllabtestmasid where b.labtest='" + Class.Users.Paramlistivew + "'  order by 2";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
                DataTable dt = ds.Tables["asptbllabtestitemmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();

                        list.Text = "";
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldiatestid"].ToString());
                        list.SubItems.Add(myRow["labtest"].ToString());
                        list.SubItems.Add(myRow["labtestitem"].ToString());
                        list.SubItems.Add(myRow["asptbllabtestitemmasid"].ToString());
                        list.SubItems.Add("asptbllabtestmasid");
                        list.SubItems.Add("0");
                        list.SubItems.Add("0");
                        list.SubItems.Add("0");
                        list.SubItems.Add("");
                        list.SubItems.Add(""); list.SubItems.Add("");
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        i++;
                        listView1.Items.Add(list);
                    }
                    lbltotal.Text = "Test Count    :" + listView1.Items.Count;

                }
                else
                {
                    lbltotal.Refresh();
                    lbltotal.Text = "No data Found.";

                }
            }
            txtsearch.Focus();txtsearch.Select();
        }

        private void listView1_ItemActivate_1(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Do You want to Delete  '" + listView1.SelectedItems[0].SubItems[2].Text + "'  ??\n'    ", "'" + Class.Users.ProjectID + "' - ", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result1.Equals(DialogResult.OK))
            {
                if (listView1.SelectedItems[0].SubItems[6].Text == "✖")
                {
                    string del = "delete from asptbldiatest where  asptbldiatestid='" + listView1.SelectedItems[0].SubItems[2].Text + "'";
                    Utility.ExecuteNonQuery(del);
                    label7.Refresh();
                    label7.Text = "Deleted";
                    foreach (ListViewItem eachItem in listView1.SelectedItems)
                    {
                        listView1.Items.Remove(eachItem);
                    }
                    listView1.EndUpdate();
                    label7.Refresh();
                    label7.Text = "Total Count    :" + listView1.Items.Count;
                }
            }
        }

        public void GridLoad()
        {

            if (Class.Users.Bisconnectclear==true && Class.Users.Paramlistivew != null)
            {
                listView1.Items.Clear();
                string sel1 = "select '' as asptbldiatestid, b.labtest,a.labtestitem,a.asptbllabtestitemmasid,b.asptbllabtestmasid  from  asptbllabtestitemmas a  join asptbllabtestmas b on a.labtest=b.asptbllabtestmasid where b.labtest='" + Class.Users.Paramlistivew + "'  order by 3";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllabtestitemmas");
                DataTable dt = ds.Tables["asptbllabtestitemmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = "";
                        list.SubItems.Add(i.ToString());
                        Class.Users.asptbltestitemmasid = myRow["asptbllabtestitemmasid"].ToString();
                        list.SubItems.Add(myRow["asptbldiatestid"].ToString());
                        list.SubItems.Add(myRow["labtest"].ToString());
                        list.SubItems.Add(myRow["labtestitem"].ToString());
                        list.SubItems.Add(myRow["asptbllabtestitemmasid"].ToString());
                        list.SubItems.Add(myRow["asptbllabtestmasid"].ToString());
                        list.SubItems.Add("");
                        list.SubItems.Add("");
                        list.SubItems.Add("");
                        list.SubItems.Add("");
                        list.SubItems.Add(""); 
                        list.SubItems.Add("");
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor=i % 2 ==0 ? Class.Users.Color1:Class.Users.Color2;
                        i++;
                        listView1.Items.Add(list);
                    }
                    lbltotal.Text = "Test Count    :" + listView1.Items.Count;
                   
                }
                else
                {
                    if (Class.Users.Staticallip.Items.Count >= 0)
                    {
                        int item0 = 0;
                        int i = 1;
                        listView1.Items.Clear(); listfilter.Items.Clear();
                        foreach (ListViewItem item in Class.Users.Staticallip.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = "";
                            list.SubItems.Add(i.ToString());
                            Class.Users.asptbltestitemmasid = item.SubItems[5].Text;
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.SubItems.Add(item.SubItems[7].Text);
                            list.SubItems.Add(item.SubItems[8].Text);
                            list.SubItems.Add(item.SubItems[9].Text);
                            list.SubItems.Add(item.SubItems[10].Text);
                            list.SubItems.Add(item.SubItems[11].Text);
                            list.SubItems.Add("");
                            listView1.Items.Add(list);          
                            listfilter.Items.Add((ListViewItem)list.Clone());
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            item0++; i++;
                        }

                        lbltotal.Text = "Test Count    :" + listView1.Items.Count;
                    }
                }
            }
            else
            {
                int item0 = 0;
                int i = 1;
                listView1.Items.Clear(); listfilter.Items.Clear();
                foreach (ListViewItem item in Class.Users.Staticallip.Items)
                {
                    ListViewItem list = new ListViewItem();
                    list.Text = "";
                    list.SubItems.Add(i.ToString());
                    Class.Users.asptbltestitemmasid = item.SubItems[5].Text;
                    list.SubItems.Add(item.SubItems[2].Text);
                    list.SubItems.Add(item.SubItems[3].Text);
                    list.SubItems.Add(item.SubItems[4].Text);
                    list.SubItems.Add(item.SubItems[5].Text);
                    list.SubItems.Add(item.SubItems[6].Text);
                    list.SubItems.Add(item.SubItems[7].Text);
                    list.SubItems.Add(item.SubItems[8].Text);
                    list.SubItems.Add(item.SubItems[9].Text);
                    list.SubItems.Add(item.SubItems[10].Text);
                    list.SubItems.Add(item.SubItems[11].Text);
                    list.SubItems.Add("");
                    listView1.Items.Add(list);
                    listfilter.Items.Add((ListViewItem)list.Clone());
                    list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                    item0++; i++;
                }

                lbltotal.Text = "Test Count    :" + listView1.Items.Count;

            }

            
        }

        private void PrescriptionPopUp_FormClosed(object sender, FormClosedEventArgs e)
        {
           Class.Users.Bisconnectclear = false; Class.Users.Paramlistivew = "";
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Class.Users.Enabled == true)
            {
               
                timer1.Stop();
            }
        }
    }
}

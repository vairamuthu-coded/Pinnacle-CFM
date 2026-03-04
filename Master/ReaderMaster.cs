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
    public partial class ReaderMaster : Form, ToolStripAccess
    {
        Models.Master mas = new Models.Master();
        Models.Sewing sew = new Models.Sewing();
        Models.UserRights sm = new Models.UserRights();
        private static ReaderMaster _instance; ListView listfilter = new ListView();
        public static ReaderMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ReaderMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public ReaderMaster()
        {
            InitializeComponent();
            Class.Users.UserTime = 0;

           
        }
       
        private void ReaderMaster_Load(object sender, EventArgs e)
        {
            News(); compload();sectionload();
        }
        public void compload()
        {
            try
            {

                DataTable dt = mas.comcode();
                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;
                combocompcode.Text = ""; combocompcode.SelectedIndex = -1;

                combocompcodesearch.DisplayMember = "compcode";
                combocompcodesearch.ValueMember = "gtcompmastid";
                combocompcodesearch.DataSource = dt;
                combocompcodesearch.Text = ""; combocompcodesearch.SelectedIndex = -1;

            }
            catch (Exception EX)
            { }
        }
        public void sectionload()
        {
            try
            {

                DataTable dt = sew.section();
                combosection.DisplayMember = "section";
                combosection.ValueMember = "asptblsecmasid";
                combosection.DataSource = dt;
                combosection.Text = ""; combosection.SelectedIndex = -1;

            }
            catch (Exception EX)
            { }
        }
        public void Saves()
        {
            try
            {
                if (txtreader.Text == "")
                {
                    MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtreader.Focus();
                    return;
                }

                else
                {
                    if (txtreader.Text != "")
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.ASPTBLREAMASID    from  ASPTBLREAMAS a     WHERE a.READER='" + txtreader.Text + "' and a.compcode='"+combocompcode.SelectedValue+"' and a.active='" + chk + "' and a.ASPTBLREAMASID='" + txtreaderid.Text + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLREAMAS");
                        DataTable dt = ds.Tables["ASPTBLREAMAS"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtreader.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtreaderid.Text) == 0 || Convert.ToInt32("0" + txtreaderid.Text) == 0)
                        {
                            string ins = "insert into ASPTBLREAMAS(READER,compcode,active,createdby,modifiedby,ipaddress)  VALUES('" + txtreader.Text.ToUpper() + "','"+combocompcode.SelectedValue+"','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtreader.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  ASPTBLREAMAS  set READER='" + txtreader.Text.ToUpper() + "',compcode='"+combocompcode.SelectedValue+"' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where ASPTBLREAMASID='" + txtreaderid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtreader.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
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

                MessageBox.Show("READER " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void ReaderMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }

        public void News()
        {

            GridLoad(); empty();
        }
        private void empty()
        {
            txtreaderid.Text = "";
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;

            Class.Users.UserTime = 0;
            txtreader.Text = ""; txtreader.Select();
            combocompcode.Text = "";combosection.Text = "";
            txtsearch.Text = "";
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.ASPTBLREAMASID,B.COMPCODE,c.section, a.READER , a.active    from  ASPTBLREAMAS a   join gtcompmast b on a.compcode=b.gtcompmastid join asptblsecmas c on c.COMPCODE=b.gtcompmastid and c.COMPCODE=a.compcode and c.asptblsecmasid=a.section order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLREAMAS");
                DataTable dt = ds.Tables["ASPTBLREAMAS"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLREAMASID"].ToString());                      
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["SECTION"].ToString());
                        list.SubItems.Add(myRow["READER"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listView1.Items.Add(list);
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
                Class.Users.UserTime = 0;
                if (listView1.Items.Count > 0)
                {

                    txtreaderid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = "select a.ASPTBLREAMASID,B.COMPCODE,c.section, a.READER , a.active    from  ASPTBLREAMAS a   join gtcompmast b on a.compcode=b.gtcompmastid join asptblsecmas c on c.COMPCODE=b.gtcompmastid and c.COMPCODE=a.compcode and c.asptblsecmasid=a.section   where a.ASPTBLREAMASID=" + txtreaderid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLREAMAS");
                    DataTable dt = ds.Tables["ASPTBLREAMAS"];

                    if (dt.Rows.Count > 0)
                    {
                        txtreaderid.Text = Convert.ToString(dt.Rows[0]["ASPTBLREAMASID"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combosection.Text = Convert.ToString(dt.Rows[0]["SECTION"].ToString());

                        txtreader.Text = Convert.ToString(dt.Rows[0]["READER"].ToString());
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
                int item0 = 0;int i = 1;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text) || item.SubItems[3].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = i.ToString();                                
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
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
                            list.Text = i.ToString();
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
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

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        public void Deletes()
        {
            if (txtreaderid.Text != "")
            {
                string sel1 = "select a.ASPTBLREAMASID from ASPTBLREAMAS a  join asptblbuysaminw b on a.ASPTBLREAMASID=b.READER  where a.ASPTBLREAMASID='" + txtreaderid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLREAMAS");
                DataTable dt = ds.Tables["ASPTBLREAMAS"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtreader.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
                else
                {
                    string del = "delete from ASPTBLREAMAS where ASPTBLREAMASID='" + Convert.ToInt64("0" + txtreaderid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtreader.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty(); return;
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

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad(); 

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void Searchs(int id)
        {
            throw new NotImplementedException();
        }

        public void Deletes(int id)
        {
            throw new NotImplementedException();
        }
    }
}

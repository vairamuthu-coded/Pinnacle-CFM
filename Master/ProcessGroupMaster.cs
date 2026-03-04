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
    public partial class ProcessGroupMaster : Form,ToolStripAccess
    {
        public ProcessGroupMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
        }

        private static ProcessGroupMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights(); ListView listfilter = new ListView();
        public static ProcessGroupMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ProcessGroupMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


       

        private void ProcessGroupMaster_Load(object sender, EventArgs e)
        {
            
        }
        public void Saves()
        {
            Models.ProcessGroupDetModel c = new Models.ProcessGroupDetModel();
            try
            {
               
                Int64 maxid = 0; Int64 j = 0;
                if (txtprocess.Text == "")
                {
                    MessageBox.Show("ProcessGroup Name is empty " + " Alert " + txtprocess.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtshortname.Text == "")
                {
                    MessageBox.Show("ShoftName is empty " + " Alert " + txtshortname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtprocess.Text != "" && txtshortname.Text != "")
                {
                   
                    c.asptblprocessgroupid =Convert.ToInt64("0"+txtprogrpid.Text);
                    c.processgroup = Convert.ToString(txtprocess.Text);
                    string chk = "",sew = "",qc = "";
                   
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblprogrpmasid    from  asptblprogrpmas    WHERE ProcessGroup='" + txtprocess.Text + "' and shortname='" + txtshortname.Text + "'  and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblprogrpmas");
                    DataTable dt = ds.Tables["asptblprogrpmas"];
                    if (dt.Rows.Count != 0)
                    {
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtprogrpid.Text) == 0 || Convert.ToInt32("0" + txtprogrpid.Text) == 0)
                    {
                        string ins = "insert into asptblprogrpmas(ProcessGroup,shortname,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtprocess.Text.ToUpper() + "','" + txtshortname.Text.ToUpper() + "', '" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptblprogrpmasid) as asptblprogrpmasid   from  asptblprogrpmas ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblprogrpmas");
                        DataTable dt2 = ds2.Tables["asptblprogrpmas"]; maxid = 0;
                        maxid = Convert.ToInt64(dt2.Rows[0]["asptblprogrpmasid"].ToString());
                    }
                    else
                    {
                        string up = "update  asptblprogrpmas  set   ProcessGroup='" + txtprocess.Text.ToUpper() + "' , shortname='" + txtshortname.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblprogrpmasid='" + txtprogrpid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        maxid = 0;
                        maxid = Convert.ToInt64(txtprogrpid.Text);
                        
                    }
                    int i = 0;
                    

                    int cc = dataGridView1.Rows.Count - 1;
                    if (cc >= 0)
                    {
                        for (i = 0; i < cc; i++)
                        {
                            if (Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value) >= 1)
                            {
                                if (dataGridView1.Rows[i].Cells[0].Value == null && txtprogrpid.Text == "")
                                {
                                    c.asptblprocessgroupdetid = 0;
                                    c.asptblprocessgroupid = maxid;
                                }
                                if (dataGridView1.Rows[i].Cells[0].Value == null && txtprogrpid.Text != "")
                                {
                                    c.asptblprocessgroupdetid = 0;
                                    c.asptblprocessgroupid = Convert.ToInt64("0" + txtprogrpid.Text);
                                }
                                else
                                {
                                    c.asptblprocessgroupid = Convert.ToInt64("0" + txtprogrpid.Text);
                                    c.asptblprocessgroupdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].EditedFormattedValue.ToString());
                                }
                                c.processgroup = Convert.ToString(txtprocess.Text);
                                c.processname = Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value);
                                c.rate = Convert.ToDecimal(dataGridView1.Rows[i].Cells[4].Value);
                                c.notes = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                                string sel1 = "select asptblprogrpdetid    from  asptblprogrpdet   where  asptblprogrpmasid='" + c.asptblprocessgroupid + "' and  processname='" + c.processname + "'  and processgroup='" + c.processgroup + "' and rate='" + c.rate + "' and notes='" + c.notes + "';";
                                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblprogrpdet");
                                DataTable dt1 = ds1.Tables["asptblprogrpdet"];
                                if (dt1.Rows.Count != 0)   {}
                                else if (dt1.Rows.Count != 0 && Convert.ToInt64("0" + c.asptblprocessgroupdetid) == 0 || Convert.ToInt64("0" + c.asptblprocessgroupdetid) == 0)
                                {
                                    string ins1 = "insert into asptblprogrpdet(asptblprogrpmasid,processgroup,processname,rate,notes) values('" + c.asptblprocessgroupid + "' ,'" + c.processgroup + "','" + c.processname + "' ,'" + c.rate + "','" + c.notes + "' );";
                                    Utility.ExecuteNonQuery(ins1);
                                }
                                else
                                {
                                    string up1 = "update  asptblprogrpdet  set asptblprogrpmasid='" + c.asptblprocessgroupid + "', processname='" + c.processname + "',processgroup='" + c.processgroup + "',rate='" + c.rate + "',notes='" + c.notes + "'  where asptblprogrpdetid='" + c.asptblprocessgroupdetid + "';";
                                    Utility.ExecuteNonQuery(up1);
                                }
                            }
                        }

                    }

                    if (txtprogrpid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtprogrpid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        News();

                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtprogrpid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        News();

                    }

                }
                else
                {
                    MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("ProcessGroup " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void ProcessGroupMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

      

        public void News()
        {

            empty(); GridLoad();
           
            Utility.Load_GridCombo(processname, "select asptblpromasid, Processname    from  asptblpromas  where active='T'   order by 2", "asptblpromasid", "processname");
        }
        private void empty()
        {
            txtprogrpid.Text = "";
            txtprocess.Text = "";
            txtshortname.Text = "";

            checkactive.Checked = true;
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            GlobalVariables.HideCols = new string[] { "asptblprogrpdetid","asptblprogrpmasid", "processgroup" };
            string[] col = { "processname" };
            GlobalVariables.WidthCols = new Int32[] { 400 };
            CommonFunctions.RemoveColumn(dataGridView1, GlobalVariables.HideCols);
            CommonFunctions.SetRowNumber(dataGridView1);
            do
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView1.Rows.Count > 1);

            txtprocess.Select();
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "   SELECT a.asptblprogrpmasid, a.ProcessGroup ,a.shortname,a.processgroup ,a.active  FROM  asptblprogrpmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprogrpmas");
                DataTable dt = ds.Tables["asptblprogrpmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblprogrpmasid"].ToString());
                        list.SubItems.Add(myRow["ProcessGroup"].ToString());
                        list.SubItems.Add(myRow["shortname"].ToString());      
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

                    txtprogrpid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblprogrpmasid, a.ProcessGroup ,a.shortname, a.active    from  asptblprogrpmas a    where a.asptblprogrpmasid=" + txtprogrpid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprogrpmas");
                    DataTable dt = ds.Tables["asptblprogrpmas"];
                     
                    if (dt.Rows.Count > 0)
                    {
                        txtprogrpid.Text = Convert.ToString(dt.Rows[0]["asptblprogrpmasid"].ToString());
                        txtprocess.Text = Convert.ToString(dt.Rows[0]["ProcessGroup"].ToString());
                        txtshortname.Text = Convert.ToString(dt.Rows[0]["shortname"].ToString());

                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        string sel2 = "select b.asptblprogrpdetid,b.asptblprogrpmasid,b.processgroup,c.asptblpromasid,b.rate,b.notes from asptblprogrpmas a join asptblprogrpdet b on a.asptblprogrpmasid=b.asptblprogrpmasid join asptblpromas c on c.asptblpromasid = b.processname   where a.asptblprogrpmasid='" + txtprogrpid.Text + "';";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblprogrpdet");
                        DataTable dt1 = ds2.Tables["asptblprogrpdet"];
                        if (dt1.Rows.Count > 0)
                        {

                            dataGridView1.Rows.Clear();
                            int i = 0;
                            for (i = 0; i < dt1.Rows.Count; i++)
                            {
                                if (Convert.ToInt64(dt1.Rows[i]["asptblprogrpdetid"].ToString()) > 0)
                                {
                                    dataGridView1.Rows.Add();
                                    dataGridView1.Rows[i].Cells[0].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblprogrpdetid"].ToString());
                                    dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblprogrpmasid"].ToString());
                                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToString(dt1.Rows[i]["processgroup"].ToString());
                                    dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64(dt1.Rows[i]["asptblpromasid"].ToString());
                                    dataGridView1.Rows[i].Cells[4].Value = dt1.Rows[i]["rate"].ToString();
                                    dataGridView1.Rows[i].Cells[5].Value = dt1.Rows[i]["notes"].ToString();
                                }
                            }

                            CommonFunctions.SetRowNumber(dataGridView1);
                        }
                        else
                        {
                            dataGridView1.Rows.Clear();
                        }
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

                    int item0 = 0; listView1.Items.Clear();
                    if (txtsearch.Text.Length > 1)
                    {

                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text))
                            {
                                list.Text = listfilter.Items[item0].SubItems[0].Text;
                                list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);               
                                list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);                               
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

                        listView1.Items.Clear(); item0 = listfilter.Items.Count;
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            if (item0 % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }
                            listView1.Items.Add(list);
                            item0++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
        }

        public void Deletes()
        {
            if (txtprogrpid.Text != "")
            {
                string sel1 = "select a.asptblprogrpmasid from asptblprogrpmas a join asptbllay b on a.asptblprogrpmasid=b.processname where a.asptblprogrpmasid='" + txtprogrpid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprogrpmas");
                DataTable dt = ds.Tables["asptblprogrpmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtprocess.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    string del = "delete from asptblprogrpmas where asptblprogrpmasid='" + Convert.ToInt64("0" + txtprogrpid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtprocess.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void refeshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();
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

      

        private void ProcessGroupMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            _instance = null;
        }

        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //CommonFunctions.SetRowNumber(dataGridView1);
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }

        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3 && dataGridView1.CurrentCell.Value != null)
            {
                mas.checkduplicate(e.ColumnIndex, dataGridView1);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
namespace Pinnacle.Master
{
    public partial class SewingLineMaster : Form,ToolStripAccess
    {
        public SewingLineMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

            //this.BackColor = Class.Users.BackColors;
            //butheader.BackColor = Class.Users.BackColors;
            //panel2.BackColor = Class.Users.BackColors;
            //panel3.BackColor = Class.Users.BackColors;
            
           // ucDataGridView1.ColumnHeadersDefaultCellStyle.BackColor= Class.Users.BackColors;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static SewingLineMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();

        public static SewingLineMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new SewingLineMaster();
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
        private void SewingLineMaster_Load(object sender, EventArgs e)
        {
          

            string[] ports = SerialPort.GetPortNames();
            string[] ports1 = SerialPort.GetPortNames();
            combosewport.Items.AddRange(ports);
            combosewport.Text = ports[0];
            comboqcport.Items.AddRange(ports1);
            comboqcport.Text = ports1[0];
            locationload();
            combocodeload(); combocodeSewingload(); sectionload();
            News();
        }
        public void combocodeload()
        {
            try
            {
                string sel = "select gtcompmastid,compcode from  gtcompmast  order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                combocompcodesewingline.DisplayMember = "compcode";
                combocompcodesewingline.ValueMember = "gtcompmastid";
                combocompcodesewingline.DataSource = dt;
            

            }
            catch (Exception EX)
            { }
        }

        public void combocodeSewingload()
        {
            try
            {
               
                    string sel = "select distinct a.compcode from  gtcompmast a join asptblsewlinmas b on a.gtcompmastid=b.compcode  order by 1 ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                    DataTable dt = ds.Tables["gtcompmast"];

                    combocompcodesewingwork.DisplayMember = "compcode";
                    combocompcodesewingwork.ValueMember = "gtcompmastid";
                    combocompcodesewingwork.DataSource = dt;
                    combocompcodeqc.DisplayMember = "compcode";
                    combocompcodeqc.ValueMember = "gtcompmastid";
                    combocompcodeqc.DataSource = dt;
               
       
                
            }
            catch (Exception EX)
            { }
        }
        void sectionload()
        {

   
            Utility.Load_GridCombo(Section, "select a.asptblsewlinmasid,a.linename from asptblsewlinmas a join gtcompmast b on a.compcode=b.gtcompmastid order by 2", "asptblsewlinmasid", "linename");
            Utility.Load_GridCombo(ReaderType, "select a.asptblsewlinmasid,a.linename from asptblsewlinmas a join gtcompmast b on a.compcode=b.gtcompmastid order by 2", "asptblsewlinmasid", "linename");
            Utility.Load_GridCombo(SectionQC, "select a.asptblsewlinmasid,a.linename from asptblsewlinmas a join gtcompmast b on a.compcode=b.gtcompmastid order by 2", "asptblsewlinmasid", "linename");

            
        }
        public void LineSewing(string s)
        {
            try
            {
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
                {
                    string sel = "select a.asptblsewlinmasid,a.linename from asptblsewlinmas a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + s + "'order by 2";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsewlinmas");
                    DataTable dt = ds.Tables["asptblsewlinmas"];
                    if (dt.Rows.Count > 0)
                    {
                        combolinenamesewing.DisplayMember = "linename";
                        combolinenamesewing.ValueMember = "asptblsewlinmasid";
                        combolinenamesewing.DataSource = dt;
               

                    }
                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])//your specific tabname
                {
                    string sel = "select a.asptblsewlinmasid,a.linename from asptblsewlinmas a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + s + "'order by 2";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsewlinmas");
                    DataTable dt = ds.Tables["asptblsewlinmas"];
                    if (dt.Rows.Count > 0)
                    {                       
                        combolinenameqc.DisplayMember = "linename";
                        combolinenameqc.ValueMember = "asptblsewlinmasid";
                        combolinenameqc.DataSource = dt;

                    }
                }
            }
            catch (Exception EX)
            { }
        }
        public void LineSewing(string s,string ss)
        {
            try
            {

                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
                {
                    string sel = "select a.workstartno,a.workendno from asptblsewlinmas a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + s + "' and a.linename='" + ss + "' order by 2";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsewlinmas");
                    DataTable dt = ds.Tables["asptblsewlinmas"];
                    txtsewingwssno.Text = dt.Rows[0]["workstartno"].ToString();
                    txtsewingwseno.Text = dt.Rows[0]["workendno"].ToString();
                }
                if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])//your specific tabname
                {
                    string sel3 = "select a.workstartno,a.workendno from asptblsewlinmas a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + s + "' and a.linename='" + ss + "' order by 2";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblsewlinmas");
                    DataTable dt3 = ds3.Tables["asptblsewlinmas"];
                    txtwssnoqc.Text = dt3.Rows[0]["workstartno"].ToString();
                    txtwsenoqc.Text = dt3.Rows[0]["workendno"].ToString();
                }
            }
            catch (Exception EX)
            { }
        }
        public void locationload()
        {
            try
            {
                string sel = "select asptbllocmasid,location from  asptbllocmas  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllocmas");
                DataTable dt = ds.Tables["asptbllocmas"];

                combolocationseqloc.DisplayMember = "location";
                combolocationseqloc.ValueMember = "asptbllocmasid";
                combolocationseqloc.DataSource = dt;

            }
            catch (Exception EX)
            { }
        }
        public void Saves()
        {
            try
            {
                if (txtlinenamesewline.Text == "")
                {
                    MessageBox.Show("lineName Name is empty " + " Alert " + txtlinenamesewline.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtlinenosewline.Text == "")
                {
                    MessageBox.Show("lineno Name is empty " + " Alert " + txtlinenamesewline.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (combocompcodesewingline.Text == "")
                {
                    MessageBox.Show("CompCode Name is empty " + " Alert " + combocompcodesewingline.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (combolocationseqloc.Text == "")
                {
                    MessageBox.Show("Location  empty " + " Alert " + combolocationseqloc.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtpollingip.Text == "")
                {
                    MessageBox.Show("Polling IpAddress empty " + " Alert " + txtpollingip.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (combosewport.Text == "")
                {
                    MessageBox.Show("Sewing Port empty " + " Alert " + combosewport.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (comboqcport.Text == "")
                {
                    MessageBox.Show("QC Port empty " + " Alert " + comboqcport.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtlinenamesewline.Text != "" && txtlinenosewline.Text != "" && combocompcodesewingline.Text != "" && combolocationseqloc.Text != "" && comboqcport.Text != "" && combosewport.Text != "" )
                {
                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblsewlinmasid    from  asptblsewlinmas    WHERE unit='" + combocompcodesewingline.SelectedValue + "' and lineno='" + txtlinenosewline.Text + "' and linename='" + txtlinenamesewline.Text + "' and location='" + combolocationseqloc.SelectedValue + "' and pollingipaddress='" + txtpollingip.Text + "' and sewingcomport='" + combosewport.Text + "' and qccomport='" + comboqcport.Text + "' and workstartno='" + txtwssno.Text + "' and workendno='" + txtwseno.Text + "' and   active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsewlinmas");
                    DataTable dt = ds.Tables["asptblsewlinmas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtlinenamesewline.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtseineid.Text) == 0 || Convert.ToInt32("0" + txtseineid.Text) == 0)
                    {
                        string ins = "insert into asptblsewlinmas(unit,lineno,linename,location,pollingipaddress,sewingcomport,qccomport,workstartno,workendno,active,compcode,compcode1,username,createdby,modifiedby,ipaddress)  VALUES('" + combocompcodesewingline.SelectedValue + "','" + txtlinenosewline.Text + "','" + txtlinenamesewline.Text.ToUpper() + "','" + combolocationseqloc.SelectedValue + "','" + txtpollingip.Text + "','" + combosewport.Text.ToUpper() + "','" + comboqcport.Text.ToUpper() + "','" + txtwssno.Text + "' , '" + txtwseno.Text + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtlinenamesewline.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblsewlinmas  set   unit='" + combocompcodesewingline.SelectedValue + "',lineno='" + txtlinenosewline.Text + "',linename='" + txtlinenamesewline.Text + "',location='" + combolocationseqloc.SelectedValue + "',pollingipaddress='" + txtpollingip.Text + "',sewingcomport='" + combosewport.Text + "',qccomport='" + comboqcport.Text + "', workstartno='" + txtwssno.Text + "' , workendno='" + txtwseno.Text + "', active='" + chk + "' ,compcode='" + Class.Users.COMPCODE + "',compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblsewlinmasid='" + txtseineid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtlinenamesewline.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("lineno " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void SewingLineMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       

        public void News()
        {
            locationload(); combocodeload();

            GridLoad(); empty();
        }
        private void empty()
        {
            txtseineid.Text = "";
            txtlinenamesewline.Text = "";
            txtlinenosewline.Text = ""; txtwssno.Text = ""; txtwseno.Text = "";
            combocompcodeqc.Text = "";combolinenameqc.Text = "";txtwssnoqc.Text = "";txtwsenoqc.Text = "";
            combocompcodesewingline.Text = ""; combocompcodesewingline.SelectedIndex = -1;
            combolocationseqloc.Text = ""; combolocationseqloc.SelectedIndex = -1;
            txtpollingip.Text = "";
            comboqcport.Text = ""; comboqcport.SelectedIndex = -1;
            combosewport.Text = ""; combosewport.SelectedIndex = -1;
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.EnableHeadersVisualStyles = false;

            dataGridView2.AllowUserToAddRows = true;
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            dataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView2.EnableHeadersVisualStyles = false;
            checkactive.Checked = false;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
        }
       
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "   SELECT A.asptblsewlinmasid,a.lineno,a.linename,a.pollingipaddress,a.sewingcomport,a.qccomport, a.active  FROM  asptblsewlinmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsewlinmas");
                DataTable dt = ds.Tables["asptblsewlinmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblsewlinmasid"].ToString());
                        list.SubItems.Add(myRow["lineno"].ToString());
                        list.SubItems.Add(myRow["linename"].ToString());
                        list.SubItems.Add(myRow["pollingipaddress"].ToString());
                        list.SubItems.Add(myRow["sewingcomport"].ToString());
                        list.SubItems.Add(myRow["qccomport"].ToString());
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

                    txtseineid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select A.asptblsewlinmasid,b.compcode, a.lineno,a.linename,c.location, a.pollingipaddress,a.sewingcomport,a.qccomport, a.workstartno,a.workendno, a.active  from  asptblsewlinmas a  join gtcompmast b on b.gtcompmastid=a.compcode join asptbllocmas c on c.asptbllocmasid=a.location where  a.asptblsewlinmasid=" + txtseineid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsewlinmas");
                    DataTable dt = ds.Tables["asptblsewlinmas"];
                   
                    if (dt.Rows.Count > 0)
                    {
                        txtseineid.Text = Convert.ToString(dt.Rows[0]["asptblsewlinmasid"].ToString());
                        combocompcodesewingline.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        txtlinenosewline.Text = Convert.ToString(dt.Rows[0]["lineno"].ToString());
                        txtlinenamesewline.Text = Convert.ToString(dt.Rows[0]["linename"].ToString());
                        combolocationseqloc.Text = Convert.ToString(dt.Rows[0]["location"].ToString());
                        txtpollingip.Text = Convert.ToString(dt.Rows[0]["pollingipaddress"].ToString());
                        combosewport.Text = Convert.ToString(dt.Rows[0]["sewingcomport"].ToString());
                        comboqcport.Text = Convert.ToString(dt.Rows[0]["qccomport"].ToString());
                        txtwssno.Text = Convert.ToString(dt.Rows[0]["workstartno"].ToString());
                        txtwseno.Text = Convert.ToString(dt.Rows[0]["workendno"].ToString());
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
                        if (listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[5].ToString().Contains(txtsearch.Text))
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
                            listView1.Items.Add(list);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                       

                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {

                    ListView ll = new ListView();

                    listView1.Items.Clear(); int i = listfilter.Items.Count;
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        this.listView1.Items.Add((ListViewItem)item.Clone());
                        if (i % 2 == 0) { item.BackColor = Color.WhiteSmoke; } else { item.BackColor = Color.White; }

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
            //        string sel1 = "  SELECT  a.asptblsewlinmasid,a.lineno,a.active from asptblsewlinmas a  where a.lineno LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsewlinmas");
            //        DataTable dt = ds.Tables["asptblsewlinmas"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptblsewlinmasid"].ToString());
            //                list.SubItems.Add(myRow["lineno"].ToString());
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

        public void Deletes()
        {
            if (txtseineid.Text != "")
            {
                string sel1 = "select a.asptblsewlinmasid from asptblsewlinmas a join gtstatemast b on a.asptblsewlinmasid=b.country where a.asptblsewlinmasid='" + txtseineid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsewlinmas");
                DataTable dt = ds.Tables["asptblsewlinmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtlinenamesewline.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {

                    string del = "delete from asptblsewlinmas where asptblsewlinmasid='" + Convert.ToInt64("0" + txtseineid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtlinenamesewline.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

    

        private void combolinenamesewing_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (combolinenamesewing.SelectedIndex != -1 && combolinenamesewing.SelectedIndex != -1)
            {
                LineSewing(combocompcodesewingwork.Text, combolinenamesewing.Text);
                if (txtsewingwssno.Text != "")
                {
                    dataGridView1.Rows.Clear();
                    int stno = Convert.ToInt32(txtsewingwssno.Text);
                    int enno = Convert.ToInt32(txtsewingwseno.Text);
                    for (int i = 0; i < enno; i++)
                    {
                        dataGridView1.Rows.Add();

                        dataGridView1.Rows[i].Cells[0].Value = stno.ToString();
                        stno += 1;
                    }
                }
            }
        }

        private void combocompcodesewingwork_SelectedIndexChanged(object sender, EventArgs e)
        {
            LineSewing(combocompcodesewingwork.Text);
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }

        private void combocompcodeqc_SelectedIndexChanged(object sender, EventArgs e)
        {
            LineSewing(combocompcodeqc.Text);
        }

        private void combolinenameqc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combocompcodeqc.SelectedIndex != -1 && combolinenameqc.Text != "")
            {
                LineSewing(combocompcodeqc.Text, combolinenameqc.Text);
                if (txtwssnoqc.Text != "")
                {
                    dataGridView2.Rows.Clear();
                    int stno = Convert.ToInt32(txtwssnoqc.Text);
                    int enno = Convert.ToInt32(txtwsenoqc.Text);
                    for (int i = 0; i < enno; i++)
                    {
                        dataGridView2.Rows.Add();

                        dataGridView2.Rows[i].Cells[0].Value = stno.ToString();
                        stno += 1;
                    }
                }
            }
        }
    }
}

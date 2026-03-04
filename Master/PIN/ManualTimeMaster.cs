using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Pinnacle.Master.PIN
{
    public partial class ManualTimeMaster : Form, ToolStripAccess
    {
        private static ManualTimeMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView(); ListView listfilter1 = new ListView();
        public static ManualTimeMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ManualTimeMaster();
                GlobalVariables.CurrentForm = _instance;
                _instance.Font = Class.Users.FontName;
                return _instance;
            }
        }
        protected override bool ProcessCmdKey(ref Message message, Keys keys)
        {
            switch (keys)
            {
                case Keys.N | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    News();
                    return true; // signal that we've processed this key
                case Keys.S | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Saves();
                    return true; // signal that we've processed this key
                //case Keys.E | Keys.Control:
                //    // ... Process Shift+Ctrl+Alt+B ...
                //    updating = true;
                //    adding = false;
                //    EnableText();
                //    return true; // signal that we've processed this key
                case Keys.D | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Deletes();
                    return true; // signal that we've processed this key
            }
            // run base implementation
            return base.ProcessCmdKey(ref message, keys);
        }
        public ManualTimeMaster()
        {
            InitializeComponent();



        }
        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    //for (int r = 0; r < dt1.Rows.Count; r++)
                    //{

                    //    if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                    //    if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                    //    if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                    //    if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                    //    if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                    //    if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                    //    if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                    //    if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                    //    if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                    //    if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                    //    if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                    //    if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }
                    //}
                }


            }
            else
            {

            }

        }

        private void ManualTimeMaster_Load(object sender, EventArgs e)
        {
            GridLoad(); state();
            compcodeLoad();

        }

        public void state()
        {
            string sel = "SELECT  a.asptblempid,a.empname from asptblemp a  where a.active='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblemp");
            DataTable dt = ds.Tables["asptblemp"];
            comboempname.DataSource = dt;
            comboempname.DisplayMember = "empname";
            comboempname.ValueMember = "asptblempid";



        }

        public void compcodeLoad()
        {
            string sel = "SELECT  a.gtcompmastid,a.compcode from gtcompmast a  where a.active='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            combocompcode.DisplayMember = "compcode";
            combocompcode.ValueMember = "gtcompmastid";
            combocompcode.DataSource = dt;

        }
        public DataTable country(string s)
        {
            string sel = "SELECT  a.idcardno from asptblemp a  where a.active='T' and a.asptblempid='" + s + "' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblemp");
            DataTable dt = ds.Tables["asptblemp"];

            return dt;
        }


        public void Saves()
        {
            try
            {


                if (comboempname.Text == "")
                {
                    MessageBox.Show("'EmpName Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.comboempname.Focus();
                    return;
                }
                if (comboidcardno.Text == "")
                {
                    MessageBox.Show("'IDcardno Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.comboidcardno.Focus(); return;
                }
                if (txtintime.Text == "")
                {
                    MessageBox.Show("'InTime  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtintime.Focus(); return;
                }


                else
                {
                    DateTime statetime = new DateTime(); DateTime endtime = new DateTime();
                    TimeSpan differ = new TimeSpan(); TimeSpan differ1 = new TimeSpan();
                    if (comboempname.Text != "" && comboidcardno.Text != "")
                    {

                        string chk = "", chk1 = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.asptblmanualID    from  asptblmanual a    WHERE a.manualdate='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'   and a.empname='" + comboempname.SelectedValue + "' and a.idcardno='" + comboidcardno.SelectedValue + "' and a.active='" + chk + "'";//and a.intime='" + txtintime.Text + "' and a.outtime='" + txtouttime.Text + "' and a.manualtime='" + txtmanualtime.Text + "' and a.permission='" + txtpermission.Text + "'
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblmanual");
                        DataTable dt = ds.Tables["asptblmanual"];
                        if (dt.Rows.Count <= 0)
                        {
                        //    MessageBox.Show("Child Record Found " + " Alert " + txtpermission.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        //}
                        //else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtmanualtimeid.Text) == 0 || Convert.ToInt32("0" + txtmanualtimeid.Text) == 0)
                        //{
                            string ins = "insert into asptblmanual(manualdate,manualtime,permission, empname,idcardno,intime,outtime, active,createdby,modifiedby,ipaddress)  VALUES('" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "','" + txtmanualtime.Text.ToUpper() + "','" + txtpermission.Text + "','" + comboempname.SelectedValue + "','" + comboidcardno.SelectedValue + "','" + txtintime.Text + "','" + txtouttime.Text + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);

                            if (txtmanualtime.Text=="0")
                            {
                                string ins10 = "INSERT INTO " + Class.Users.HCompcode + "TRS_ATTLOG(IPADDRESS, ENROLLNO, DATETIMERECORD,indate,intime,outtime,compcode)VALUES('" + Class.Users.IPADDRESS + "'," + comboidcardno.Text + ",'" + txtouttime.Text + "','" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "','" + txtintime.Text + "','" + txtouttime.Text + "','" + Class.Users.HCompcode + "')";
                                Utility.ExecuteNonQuery(ins10);
                            }
                            else
                            {
                                string ins10 = "INSERT INTO " + Class.Users.HCompcode + "TRS_ATTLOG(IPADDRESS, ENROLLNO, DATETIMERECORD,indate,intime,outtime,compcode)VALUES('" + Class.Users.IPADDRESS + "'," + comboidcardno.Text + ",'" + txtmanualtime.Text + "','" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "','" + txtintime.Text + "','" + txtmanualtime.Text + "','" + Class.Users.HCompcode + "')";
                                Utility.ExecuteNonQuery(ins10);
                            }


                            MessageBox.Show("Record Saved Successfully " + txtpermission.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblmanual  set manualdate='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' ,manualtime='" + txtmanualtime.Text + "' , permission='" + txtpermission.Text.ToUpper() + "', empname='" + comboempname.SelectedValue + "',idcardno='" + comboidcardno.SelectedValue + "' ,intime='" + txtintime.Text + "',outtime='" + txtouttime.Text + "', active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblmanualID='" + txtmanualtimeid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            //if (txtintime.Text != "" && txtouttime.Text == "")
                            //{                                
                            //    string up1 = "update  pintrs_attlog set  outtime='" + txtmanualtime.Text + "' where ENROLLNO='" + comboidcardno.Text + "' and indate='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                            //    Utility.ExecuteNonQuery(up1);
                            //}
                            //if (txtintime.Text != "" && txtouttime.Text != "")
                            //{
                            string up1 = "update  pintrs_attlog set intime='" + txtintime.Text + "', outtime='" + txtmanualtime.Text + "' where ENROLLNO='" + comboidcardno.Text + "' and indate='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'";
                            Utility.ExecuteNonQuery(up1);
                            // }

                            MessageBox.Show("Record Updated Successfully " + txtpermission.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("permission " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void ManualTimeMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
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
            empty();
            GridLoad();
        }
        private void empty()
        {
            txtmanualtimeid.Text = ""; txtmanualtime.Text = "";
            txtpermission.Text = ""; txtouttime.Enabled = true;
            comboempname.Text = ""; comboempname.SelectedIndex = -1;
            comboidcardno.Text = ""; comboidcardno.SelectedIndex = -1; txtintime.Text = ""; txtouttime.Text = "";
            // txtsearch.Text = "";
            //panel2.BackColor = Class.Users.BackColors;
            //panel3.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
          
            listView2.Font = Class.Users.FontName;
            txtpermission.Select();
        }

        public void GridLoad()
        {
            try
            {
                listView2.Items.Clear(); listfilter1.Items.Clear();
                string sel1 = "select  distinct a.asptblmanualid, b.empname, a.idcardno  , a.manualdate,a.manualtime,a.permission,a.intime,a.outtime from  asptblmanual a join asptblemp b on a.idcardno=b.idcardno where a.manualdate='" + frmdate.Value.ToString("yyyy-MM-dd") + "'   order by 2";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmanual");
                DataTable dt = ds.Tables["asptblmanual"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblmanualid"].ToString());
                        list.SubItems.Add(myRow["empname"].ToString());
                        list.SubItems.Add(myRow["idcardno"].ToString());
                        list.SubItems.Add(myRow["manualdate"].ToString().Substring(0, 10));
                        list.SubItems.Add(myRow["manualtime"].ToString());
                        list.SubItems.Add(myRow["permission"].ToString());
                       
                        list.SubItems.Add(myRow["intime"].ToString());
                        list.SubItems.Add(myRow["outtime"].ToString());
                                             
                        listfilter1.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                        listView2.Items.Add(list);
                    }
                    label7.Text = "Total Count: " + listView2.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        

        private void Combostate_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = country(Convert.ToString(comboempname.SelectedValue));
                comboidcardno.DataSource = dt;
                comboidcardno.DisplayMember = "idcardno";
                comboidcardno.ValueMember = "idcardno";
                txtintime.Text = ""; txtouttime.Text = "";
                DataTable dt1 = CommonFunctions.select("select  min(b.intime) as intime,max(b.outtime) as outtime ,b.indate,b.enrollno from pinTRS_ATTLOG b join asptblemp d on  d.idcardno=b.enrollno where d.active='T' and b.enrollno='" + comboidcardno.Text + "' and date_format(b.indate, '%Y-%m-%d') = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'  group by  b.indate,b.enrollno   order by 1,2", "pinTRS_ATTLOG");
                if (dt1.Rows.Count > 0)
                {
                    txtintime.Text = dt1.Rows[0]["intime"].ToString();

                    if (dt1.Rows[0]["outtime"].ToString() != "")
                    {
                        txtouttime.Text = dt1.Rows[0]["outtime"].ToString();
                        txtmanualtime.Enabled = true;
                    }
                    else
                    {

                        
                    }

                    txtmanualtime.Text = ""; txtmanualtime.Select();
                   
                }
                else
                {

                    txtintime.Select();
                    txtmanualtime.Text = "0";
                    txtmanualtime.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }



        }

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            state(); GridLoad();
        }

        public void Deletes()
        {

            if (txtmanualtimeid.Text != "")
            {
                string sel1 = "select a.asptblmanualID,a.permission, b.empname , b.idcardno  , a.active    from  asptblmanual a  join asptblemp b on a.empname=b.asptblempid where a.asptblmanualID='" + txtmanualtimeid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmanual");
                DataTable dt = ds.Tables["asptblmanual"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtpermission.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblmanual where asptblmanualID='" + Convert.ToInt64("0" + txtmanualtimeid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtpermission.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
            else
            {
                string sel1 = "select  a.ATTLOGID from  pintrs_attlog a join asptblemp b on a.ENROLLNO=b.idcardno  join gtcompmast c on c.gtcompmastid=b.compcode where c.compcode='" + Class.Users.HCompcode + "' AND  a.ENROLLNO='" + comboidcardno.Text + "' and  a.indate='" + frmdate.Value.ToString("yyyy-MM-dd") + "'  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "pintrs_attlog");
                DataTable dt = ds.Tables["pintrs_attlog"];
                if (dt.Rows.Count > 0)
                {
                    for (int J = 0; J < dt.Rows.Count; J++)
                    {
                        string del = "delete from pintrs_attlog where ATTLOGID='" + dt.Rows[J]["ATTLOGID"].ToString() + "'";
                        Utility.ExecuteNonQuery(del);
                    }
                    MessageBox.Show("Record Deleted Successfully ", " Deleted ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad();
                }

                empty();
            }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

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

        private void txtcity_TextChanged(object sender, EventArgs e)
        {

        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void txtpermission_KeyDown(object sender, KeyEventArgs e)
        {

        }
        private void txtmanualtime_KeyDown(object sender, KeyEventArgs e)
        {

        }
        private void txtpermission_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == 'A' || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);
        }


        private void butView_Click(object sender, EventArgs e)
        {

            GridLoad();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; int i = 1;
                if (textBox1.Text.Length > 0)
                {
                    listView2.Items.Clear();
                    foreach (ListViewItem item in listfilter1.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[2].ToString().Contains(textBox1.Text.ToUpper()) || item.SubItems[3].ToString().Contains(textBox1.Text.ToUpper()))
                        {
                            list.Text = i.ToString();
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView2.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {

                    listView2.Items.Clear(); item0 = 0;
                    foreach (ListViewItem item in listfilter1.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(item.SubItems[1].Text);
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);

                        list.SubItems.Add(item.SubItems[5].Text);
                        list.SubItems.Add(item.SubItems[6].Text);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView2.Items.Add(list);

                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("---" + ex.ToString());
            }
        }

        private void listView2_ItemActivate(object sender, EventArgs e)
        {

            empty();
            if (listView2.Items.Count > 0)
            {
                string sel1 = "select distinct a.asptblmanualid,c.asptblempid, c.empname,c.idcardno,a.manualdate,a.intime, a.manualtime,a.permission from  asptblmanual a join asptblemp c on c.idcardno=a.idcardno  where c.idcardno='" + listView2.SelectedItems[0].SubItems[4].Text + "' AND  a.manualdate='" + frmdate.Value.ToString("yyyy-MM-dd") + "'  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblmanual");
                DataTable dt = ds.Tables["asptblmanual"];
                if (dt.Rows.Count > 0)
                {
                    txtouttime.Enabled = false;
                    txtmanualtimeid.Text = dt.Rows[0]["asptblmanualid"].ToString();
                    comboempname.SelectedValue = Convert.ToString(dt.Rows[0]["asptblempid"].ToString());
                    dateTimePicker1.Text = Convert.ToString(dt.Rows[0]["manualdate"].ToString());
                    Combostate_SelectedIndexChanged(sender, e);
                    txtpermission.Text = dt.Rows[0]["permission"].ToString();
                    txtmanualtime.Text = dt.Rows[0]["manualtime"].ToString();
         
                    DataTable dt1 = CommonFunctions.select("select  min(b.intime) as intime,max(b.outtime) as outtime ,b.indate,b.enrollno from pinTRS_ATTLOG b join asptblemp d on  d.idcardno=b.enrollno where d.active='T' and b.enrollno='" + dt.Rows[0]["idcardno"].ToString() + "' and date_format(b.indate, '%Y-%m-%d') = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'  group by  b.indate,b.enrollno   order by 1,2", "pinTRS_ATTLOG");
                    if (dt1 != null)
                    {
                        txtintime.Text = dt1.Rows[0]["intime"].ToString();
                        if (dt1.Rows[0]["outtime"].ToString() == "0")
                        {
                            txtouttime.Text = "";
                        }
                        else
                        {
                            txtmanualtime.Text = dt1.Rows[0]["outtime"].ToString();
                            txtouttime.Text = dt1.Rows[0]["outtime"].ToString();
                        }
                    }
                }
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {

            GridLoad();
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void txtmanualtime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) ||  e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);
            
        }

        private void txtintime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);

        }

        private void txtouttime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsNumber(e.KeyChar)  || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back);

        }

        private void txtmanualtime_TextChanged(object sender, EventArgs e)
        {
            if (txtmanualtime.Text.Length >= 5)
            {
                string[] ss = txtmanualtime.Text.Split(':');
                if (ss.Length > 1)
                {
                    int s1 = Convert.ToInt32("0" + ss[0].ToString());
                    int s2 = Convert.ToInt32("0" + ss[1].ToString());
                    int s3 = Convert.ToInt32("0" + ss[2].ToString());
                    if (s1 >= 24 || s2 >= 59 || s3 >= 59) { MessageBox.Show("Invalid  ." + txtmanualtime.Text); }
                }
                else
                {
                    MessageBox.Show("Invalid  ." + txtmanualtime.Text);
                }
            }
        }
       private  void checkformate (string[] s1,int s2,int s3,int s4)
        {
           
        }
        private void txtpermission_TextChanged(object sender, EventArgs e)
        {
            if (txtpermission.Text.Length >= 5)
            {
                string[] ss = txtmanualtime.Text.Split(':');
                if (ss.Length > 1)
                {                   
                    int s1 = Convert.ToInt32("0" + ss[0].ToString());
                    int s2 = Convert.ToInt32("0" + ss[1].ToString());
                    int s3 = Convert.ToInt32("0" + ss[2].ToString());
                    if (s1 >= 24 || s2 >= 59 || s3 >= 59) { MessageBox.Show("Invalid  ." + txtpermission.Text); }
                }
                else
                {
                    MessageBox.Show("Invalid  ." + txtpermission.Text);
                }
            }
        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }
    }
}

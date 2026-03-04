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
    public partial class LineMaster : Form,ToolStripAccess
    {
        public LineMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            ports = SerialPort.GetPortNames();

           
        }
        private static LineMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
    string[] ports;
        public static LineMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new LineMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        public void usercheck(string s, string ss, string sss)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show("usercheck: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LineMaster_Load(object sender, EventArgs e)
        {
            News();
            comboqcport.Items.AddRange(ports);
            combosewingport.Items.AddRange(ports);
            txtsearch.Select();
        }
        public void News()
        {
            empty();
            compcodeload(); locationload(); linename();
            linecompcode(); GridLoad();


        }
        public void linecompcode()
        {
            try
            {

                string sel = " select distinct a.gtcompmastid,a.compcode, a.compname from  gtcompmast  a join asptbllinemas b on a.gtcompmastid=b.compcode order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];
                combosearchline.DisplayMember = "compcode";
                combosearchline.ValueMember = "gtcompmastid";
                combosearchline.DataSource = dt;
                combosearchline.Text = ""; combosearchline.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void linename()
        {
            try
            {

                string sel = "select A.asptbllinemasid,A.LINENO from asptbllinemas a where a.active='T' ORDER BY 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllinemas");
                DataTable dt = ds.Tables["asptbllinemas"];
                comboline.DisplayMember = "LINENO";
                comboline.ValueMember = "asptbllinemasid";
                comboline.DataSource = dt;
                section.DisplayMember = "LINENO";
                section.ValueMember = "asptbllinemasid";
                section.DataSource = dt;
                comboline.Text = ""; comboline.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void compcodeload()
        {
            try
            {

                DataTable dt = mas.comcode();
                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;
                combocompcode.Text = ""; combocompcode.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }

        public void locationload()
        {
            try
            {

                DataTable dt = mas.locaion();
                comboloc.DisplayMember = "location";
                comboloc.ValueMember = "asptbllocmasid";
                comboloc.DataSource = dt;
                comboloc.Text = ""; comboloc.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }






        public void Saves()
        {
            

            try
            {

                if (combocompcode.Text != "")
                {
                    Models.LineModel c1 = new Models.LineModel();
                    c1.active = "";
                    Models.Validate va = new Models.Validate();
                    c1.asptbllinemasid = Convert.ToInt64("0" + txtlineid.Text);
                    c1.compcode = Convert.ToInt64("0"+combocompcode.SelectedValue);
                    c1.lineno = Convert.ToString(txtlinenumber.Text);
                    c1.pollingip = Convert.ToString(txtpolipaddress.Text);
                    c1.noofmachine = Convert.ToString(txtnoofmachine.Text);
                    c1.location = Convert.ToInt64("0" + comboloc.SelectedValue);
                    c1.sewingport = Convert.ToString(combosewingport.Text);
                    c1.startfrom = Convert.ToInt64("0" + txtworkstartform.Text);
                    c1.endto = Convert.ToInt64("0" + txtworkendto.Text);
                    c1.qcport = Convert.ToString(comboqcport.Text);
                    if (checkactive.Checked == true)
                        c1.active = "T";
                    else
                        c1.active = "F";

                    c1.compcode1 = Convert.ToInt64(Class.Users.COMPCODE);
                    c1.username = Convert.ToInt64(Class.Users.USERID);
                    c1.createdby = Convert.ToString(Class.Users.HUserName);
                    c1.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString());
                    c1.modified = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString());
                    c1.modifiedby = Class.Users.HUserName;
                    c1.ipAddress = Class.Users.IPADDRESS;
                    //string sel = "select asptbllinemasid    from  asptbllinemas   WHERE   compcode='" + c1.compcode + "' and lineno='" + c1.lineno + "' and pollingip='" + c1.pollingip + "' and noofmachine='" + c1.noofmachine + "' and location='" + c1.location + "' and startfrom='" + c1.startfrom + "' and endto='" + c1.endto + "' and sewingport='" + c1.sewingport + "' and qcport='" + c1.qcport + "' and active='" + c1.active + "';";
                    //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllinemas");
                    //DataTable dt = ds.Tables["asptbllinemas"];
                    //if (dt.Rows.Count != 0)
                    //{

                    //}
                    if (GlobalVariables.New_Flg == true)//dt.Rows.Count != 0 && Convert.ToInt32("0" + txtlineid.Text) == 0 || Convert.ToInt32("0" + txtlineid.Text) == 0)
                    {
                        string ins = "insert into asptbllinemas(compcode,lineno,pollingip,noofmachine,location,sewingport,qcport,startfrom,endto,active,compcode1,username,createdby,createdon,modified,ipaddress) values('" + c1.compcode + "','" + c1.lineno + "','" + c1.pollingip+ "','" + c1.noofmachine + "','" + c1.location+ "','" + c1.sewingport+ "','" + c1.qcport + "','" + c1.startfrom + "','" + c1.endto + "','" + c1.active + "','" + c1.compcode1 + "','" + c1.username + "','" + c1.createdby + "','" + Convert.ToDateTime(c1.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Class.Users.IPADDRESS + "');";
                        Utility.ExecuteNonQuery(ins);
                        //string sel2 = "select max(asptbllinemasid) as asptbllinemasid   from  asptbllinemas ;";
                        //DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbllinemas");
                        //DataTable dt2 = ds2.Tables["asptbllinemas"]; maxid = 0;
                        //maxid = Convert.ToInt64(dt2.Rows[0]["asptbllinemasid"].ToString());
                    }
                    if (GlobalVariables.New_Flg == false)
                    {
                        string up = "update  asptbllinemas  set   compcode='" + c1.compcode + "',lineno='" + c1.lineno + "',pollingip='" + c1.pollingip + "',noofmachine='" + c1.noofmachine + "',location='" + c1.location + "', startfrom='" + c1.startfrom + "' , endto='" + c1.endto + "',sewingport='" + c1.sewingport + "',qcport='" + c1.qcport + "',active='" + c1.active + "', compcode1='" + Class.Users.COMPCODE + "',  username='" + Class.Users.USERID + "',modifiedby='" + Class.Users.HUserName + "', modified='" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbllinemasid='" + c1.asptbllinemasid + "';";
                        Utility.ExecuteNonQuery(up);
                        //maxid = 0;
                        //maxid = Convert.ToInt64(txtlineid.Text);

                    }
                   // int i = 0;
                    //Models.compcodeModeldet c = new Models.compcodeModeldet();

                    //int cc = dataGridView1.Rows.Count;
                    //if (cc >= 1)
                    //{
                    //    for (i = 0; i < cc; i++)
                    //    {
                    //        if (Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value) >= 1)
                    //        {
                    //            if (txtcompcodeid.Text == "") { c.asptbllinemasid = Convert.ToInt64("0" + maxid); }
                    //            else { c.asptbllinemasid = Convert.ToInt64("0" + txtcompcodeid.Text); }
                    //            c.stylename = Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value);
                    //            c.compcode = Convert.ToString(combocompcode.SelectedValue);
                    //            c.notes = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);

                    //            string sel1 = "select asptblcompcodedetid    from  asptblcompcodedet   where  asptbllinemasid='" + c.asptbllinemasid + "' and  stylename='" + c.stylename + "'  and compcode='" + c.compcode + "'and notes='" + c.notes + "';";
                    //            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblcompcodedet");
                    //            DataTable dt1 = ds1.Tables["asptblcompcodedet"];
                    //            if (dt1.Rows.Count != 0)
                    //            {

                    //            }
                    //            else if (dt1.Rows.Count != 0 && Convert.ToInt64("0" + c.asptblcompcodedetid) == 0 || Convert.ToInt64("0" + c.asptblcompcodedetid) == 0)
                    //            {

                    //                string ins1 = "insert into asptblcompcodedet(asptbllinemasid,stylename,compcode,notes) values('" + c.asptbllinemasid + "' ,'" + c.stylename + "' ,'" + c.compcode + "','" + c.notes + "' );";
                    //                Utility.ExecuteNonQuery(ins1);
                    //            }
                    //            else
                    //            {
                    //                string up1 = "update  asptblcompcodedet  set asptbllinemasid='" + c.asptbllinemasid + "', stylename='" + c.stylename + "',compcode='" + c.compcode + "',notes='" + c.notes + "'  where asptblcompcodedetid='" + c.asptblcompcodedetid + "';";
                    //                Utility.ExecuteNonQuery(up1);
                    //            }
                    //        }
                    //    }

                    //}


                    if (txtlineid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtlinenumber.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();empty();
                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtlinenumber.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Saves_Click " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }





        
        private void empty()
        {
            txtlineid.Text = ""; GlobalVariables.New_Flg = true;
            combocompcode.Text = "";combocompcode.SelectedIndex = -1;
            combosearchline.Text = ""; combosearchline.SelectedIndex = -1;
            comboloc.Text = ""; comboloc.SelectedIndex = -1;
            combosewingport.Text = ""; combosewingport.SelectedIndex = -1;
            comboqcport.Text = ""; comboqcport.SelectedIndex = -1;
            txtpolipaddress.Text = "";
            txtlinenumber.Text = "";
            txtnoofmachine.Text = "";
            txtworkstartform.Text = "";
            txtworkendto.Text = "";
            txtnoofmachine.Text = "";
            GlobalVariables.HideCols = new string[] { "asptblstygrpDetid", "asptblstygrpmasid" };          
            CommonFunctions.RemoveColumn(dataGridView1, GlobalVariables.HideCols);
            CommonFunctions.SetRowNumber(dataGridView1);
            checkactive.Checked = true;
            if (listView1.Items.Count > 0)
            {
                listView1.Items[0].Selected = true;
            }
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;          
            panel4.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel6.BackColor = Class.Users.BackColors;
            panel11.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;          
            listView4.Font = Class.Users.FontName;
            this.Font = Class.Users.FontName;
        }
        public void GridLoad()
        {
            try
            {
                listView2.Items.Clear(); listfilter2.Items.Clear(); //date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')
                string sel1 = "select a.asptbllinemasid,b.compcode,a.lineno,a.pollingip,a.noofmachine,c.location,a.sewingport,a.qcport, a.active from asptbllinemas a join gtcompmast b on a.compcode=b.gtcompmastid join asptbllocmas c on c.asptbllocmasid=a.location  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllinemas");
                DataTable dt = ds.Tables["asptbllinemas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbllinemasid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["lineno"].ToString());
                        list.SubItems.Add(myRow["pollingip"].ToString());
                        list.SubItems.Add(myRow["noofmachine"].ToString());
                        list.SubItems.Add(myRow["location"].ToString());
                        list.SubItems.Add(myRow["sewingport"].ToString());                       
                        list.SubItems.Add(myRow["qcport"].ToString());   
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter2.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView2.Items.Add(list);
                        i++;
                    }
                    lbltotalline.Text = "Total Count: " + listView2.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                // empty();
                if (listView1.Items.Count > 0)
                {
                    GlobalVariables.New_Flg = false;
                    txtlineid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select a.asptbllinemasid,b.compcode,a.lineno,a.pollingip,a.noofmachine,c.location,a.sewingport,a.qcport, a.active from asptbllinemas a join gtcompmast b on a.compcode=b.gtcompmastid join asptbllocmas c on c.asptbllocmasid=a.location    where asptbllinemasid='" + txtlineid.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllinemas");
                    DataTable dt = ds.Tables["asptbllinemas"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtlineid.Text = Convert.ToString(dt.Rows[0]["asptbllinemasid"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        txtlinenumber.Text = Convert.ToString(dt.Rows[0]["lineno"].ToString());
                        txtpolipaddress.Text = Convert.ToString(dt.Rows[0]["pollingip"].ToString());
                        txtnoofmachine.Text = Convert.ToString(dt.Rows[0]["noofmachine"].ToString());
                        comboloc.Text = Convert.ToString(dt.Rows[0]["location"].ToString());
                        combosewingport.Text = Convert.ToString(dt.Rows[0]["sewingport"].ToString());
                        comboqcport.Text = Convert.ToString(dt.Rows[0]["qcport"].ToString());                     
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("ListView1_ItemActivate: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();


            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }

        public void Deletes()
        {
            try
            {
                if (txtlineid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {


                        string del1 = "delete from asptbllinemas where   asptbllinemasid='" + txtlineid.Text + "';";
                        Utility.ExecuteNonQuery(del1);
                        string del = "delete from asptblcompcodedet where  asptblcompcodedetid='" + txtlineid.Text + "';";
                        Utility.ExecuteNonQuery(del);
                        MessageBox.Show("Record Deleted Successfully " + txtlineid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletes_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


      
        private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //try
            //{

            //    foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
            //    {
            //        if (oneCell.Selected)
            //        {

            //            if (txtcompcodeid.Text != "")
            //            {
            //                var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            //                if (confirmation == DialogResult.Yes)
            //                {
            //                    if (griddelrow > 0)
            //                    {
            //                        string del1 = "delete from  asptblcompcodedet     Where  asptblcompcodedetid='" + griddelrow + "';";
            //                        Utility.ExecuteNonQuery(del1);

            //                        griddelrow = 0;
            //                    }
            //                    else
            //                    {
            //                        dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
            //                    }
            //                }
            //            }
            //        }
            //    }

            //}
            //catch (Exception EX)
            //{
            //    // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            //}

        }








        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //try
            //{

            //    foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
            //    {
            //        if (oneCell.Selected)
            //        {

            //            if (txtcompcodeid.Text != "")
            //            {
            //                var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            //                if (confirmation == DialogResult.Yes)
            //                {
            //                    if (griddelrow > 0)
            //                    {
            //                        string del1 = "delete from  asptblcompcodedet     Where  asptblcompcodedetid='" + griddelrow + "';";
            //                        Utility.ExecuteNonQuery(del1);
            //                        dataGridView1.Rows.RemoveAt(Convert.ToInt32(griddelrow));
            //                        griddelrow = 0;
            //                    }
            //                    else
            //                    {
            //                        dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
            //                    }
            //                }
            //            }
            //            else
            //            {
            //                dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
            //            }
            //        }
            //    }

            //}
            //catch (Exception EX)
            //{
            //    // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            //}
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            //    if (txtcompcodedetid.Text != "")
            //    {
            //        griddelrow = 0;
            //        griddelrow = Convert.ToInt64("0" + dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
            //    }
            //   // dataGridView1.BeginEdit(true);
            //}
            //catch(Exception ex)
            //{

            //}
        }



        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            lblsearchtotal.Text = "Total  :" +listView1.Items.Count.ToString();
        }



        private void refreshToolStripMenuItem_Click_3(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {

        }

        private void Searchs_Click(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            compcodeload(); locationload();
            
            comboqcport.Items.AddRange(ports);
           
            combosewingport.Items.AddRange(ports);

        }

        private void txtsearchline_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; listView2.Items.Clear();
                if (txtsearchline.Text.Length >= 1)
                {

                    foreach (ListViewItem item in listfilter2.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter2.Items[item0].SubItems[3].ToString().Contains(txtsearchline.Text) || listfilter2.Items[item0].SubItems[4].ToString().Contains(txtsearchline.Text) || listfilter2.Items[item0].SubItems[5].ToString().Contains(txtsearchline.Text) || listfilter2.Items[item0].SubItems[7].ToString().Contains(txtsearchline.Text) || listfilter2.Items[item0].SubItems[8].ToString().Contains(txtsearchline.Text))
                        {


                            list.Text = listfilter2.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter2.Items[item0].SubItems[9].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView2.Items.Add(list);
                        }
                        item0++;
                    }
                    lbltotalline.Text = "Total Count: " + listView2.Items.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    listView2.Items.Clear();

                    foreach (ListViewItem item in listfilter2.Items)
                    {

                        this.listView2.Items.Add((ListViewItem)item.Clone());

                        item0++;
                    }
                    lbltotalline.Text = "Total Count: " + listView2.Items.Count;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }
        }

        private void combosearchline_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                listView2.Items.Clear(); listfilter2.Items.Clear(); //date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')
                string sel1 = "select a.asptbllinemasid,b.compcode,a.lineno,a.pollingip,a.noofmachine,c.location,a.sewingport,a.qcport, a.active from asptbllinemas a join gtcompmast b on a.compcode=b.gtcompmastid join asptbllocmas c on c.asptbllocmasid=a.location where b.compcode='" + combosearchline.Text + "'  order by 3";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllinemas");
                DataTable dt = ds.Tables["asptbllinemas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbllinemasid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["lineno"].ToString());
                        list.SubItems.Add(myRow["pollingip"].ToString());
                        list.SubItems.Add(myRow["noofmachine"].ToString());
                        list.SubItems.Add(myRow["location"].ToString());
                        list.SubItems.Add(myRow["sewingport"].ToString());
                        list.SubItems.Add(myRow["qcport"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter2.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView2.Items.Add(list);
                        i++;
                    }
                    lbltotalline.Text = "Total Count: " + listView2.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listView2_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                // empty();
                if (listView2.Items.Count > 0)
                {
                    GlobalVariables.New_Flg = false;
                    txtlineid.Text = listView2.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select a.asptbllinemasid,b.compcode,a.lineno,a.pollingip,a.noofmachine,c.location,a.sewingport,a.qcport,a.startfrom,a.endto, a.active from asptbllinemas a join gtcompmast b on a.compcode=b.gtcompmastid join asptbllocmas c on c.asptbllocmasid=a.location    where asptbllinemasid='" + txtlineid.Text + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllinemas");
                    DataTable dt = ds.Tables["asptbllinemas"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtlineid.Text = Convert.ToString(dt.Rows[0]["asptbllinemasid"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        txtlinenumber.Text = Convert.ToString(dt.Rows[0]["lineno"].ToString());
                        txtpolipaddress.Text = Convert.ToString(dt.Rows[0]["pollingip"].ToString());
                        txtnoofmachine.Text = Convert.ToString(dt.Rows[0]["noofmachine"].ToString());
                        comboloc.Text = Convert.ToString(dt.Rows[0]["location"].ToString());
                        combosewingport.Text = Convert.ToString(dt.Rows[0]["sewingport"].ToString());
                        comboqcport.Text = Convert.ToString(dt.Rows[0]["qcport"].ToString());
                        txtworkstartform.Text = Convert.ToString(dt.Rows[0]["startfrom"].ToString());
                        txtworkendto.Text = Convert.ToString(dt.Rows[0]["endto"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }


                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("ListView1_ItemActivate: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtlinenumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar >= 'a' && e.KeyChar <= 'z' || e.KeyChar >= 'A' && e.KeyChar <= 'Z' ||e.KeyChar == '-' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtpolipaddress_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtnoofmachine_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

   

        private void tabPagedel1_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboline_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CommonFunctions.SetRowNumber(dataGridView1);
        }
    }
}

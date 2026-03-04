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
    public partial class HolidayMaster : Form,ToolStripAccess
    {
        public HolidayMaster()
        {
            InitializeComponent();
        
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            combofinyear.Text = Class.Users.Finyear;
            dataGridView1.Controls.Add(dtp);
            dtp.Visible = false;
            dtp.Format = DateTimePickerFormat.Short;
            dtp.TextChanged += Dtp_TextChanged;
        }

        private static HolidayMaster _instance;
        public static HolidayMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new HolidayMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

        DateTimePicker dtp = new DateTimePicker();
        Rectangle rectangle;

        private string readserialvalue;
        decimal listview2totalweight = 0;
        Models.Validate va = new Models.Validate();
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        string dat = "";       
        bool validprint = false;
        ListView listfilter = new ListView();

       
        private void HolidayMaster_Load(object sender, EventArgs e)
        {
            GridLoad();
            companyload(); holidaycateload(); combofinyear.Text = Class.Users.Finyear; empty();
            this.KeyPreview = true;
            tabControl1.SelectTab(tabPagedel2);
            txtsearch.Select(); dateTimePicker1.Value = System.DateTime.Now.AddDays(0);
        }
        public void News()
        {

            holidaycateload(); companyload(); GridLoad();
            tabControl1.SelectTab(tabPagedel1); autonumberload(); combocompcode.Select();
            empty();

        }


        public void companyload()
        {
            try
            {
                string sel = "select a.gtcompmastid,a.compcode, a.compname from  gtcompmast a  where  a.active='T'  order by 2 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;

                //string sel1 = "select distinct a.gtcompmastid,a.compcode from  gtcompmast a join asptblholmas b on a.gtcompmastid=b.compcode  where a.ptransaction ='COMPANY'  order by 2;";
                //DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                //DataTable dt1 = ds1.Tables["gtcompmast"];
                combounit.DisplayMember = "compcode";
                combounit.ValueMember = "gtcompmastid";
                combounit.DataSource = dt;


                combocompname.DisplayMember = "compname";
                combocompname.ValueMember = "gtcompmastid";
                combocompname.DataSource = dt;







            }
            catch (Exception ex)
            {
                MessageBox.Show("companyload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        //string sel = "select b.gtcompmastid,b.compname from  gtcompmast b where b.ptransaction = 'COMPANY' and b.compcode ='" + combocompcode.Text + "'; ";
        //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblholmas");
        //DataTable dt = ds.Tables["asptblholmas"];
        //int cnt = dt.Rows.Count;

        //combocompname.DisplayMember = "compname";
        //        combocompname.ValueMember = "gtcompmastid";
        //        combocompname.DataSource = dt;

        //        try
        //        {
        //            string sel1 = "select distinct  a.month from asptblholdet a order by 1;";
        //DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblholdet");
        //DataTable dt1 = ds1.Tables["asptblholdet"];

        //combomonth.DisplayMember = "month";
        //            combomonth.ValueMember = "month";
        //            combomonth.DataSource = dt1;

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("combocompcode_SelectedIndexChanged: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }

        public void autonumberload()
        {
            try
            {
                //txtholiid.Text = ""; txtdocid.Text = "";

                string sel = "select max(asptblholmasid1)+1 as id,b.gtcompmastid,b.compname from asptblholmas a join gtcompmast b on a.compcode = b.gtcompmastid where a.finyear='" + combofinyear.Text + "' and b.compcode ='" + combocompcode.Text + "'; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblholmas");
                DataTable dt = ds.Tables["asptblholmas"];
                int cnt = dt.Rows.Count;
                if (dt.Rows[0]["id"].ToString() == "")
                {
                    string sel1 = "select b.gtcompmastid, b.compname from  gtcompmast b  where  b.compcode='" + combocompcode.Text + "'; ";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                    DataTable dt1 = ds1.Tables["gtcompmast"];
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt1;

                    txtdocid.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + 1;
                    txtholiid1.Text = "1";
                }
                else
                {

                    txtdocid.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                    txtholiid1.Text = dt.Rows[0]["id"].ToString();
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void holidaycateload()
        {
            try
            {
                string sel = "select a.asptblhrholidaycatmasid,a.holidaycategory from asptblhrholidaycatmas a where a.active='T';";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblhrholidaycatmas");
                DataTable dt = ds.Tables["asptblhrholidaycatmas"];

                HOLIDAYCATEGORY.DisplayMember = "holidaycategory";
                HOLIDAYCATEGORY.ValueMember = "asptblhrholidaycatmasid";
                HOLIDAYCATEGORY.DataSource = dt;

            }
            catch (Exception ex)
            {
                MessageBox.Show("holidaycateload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        bool isvalid = false;

        private bool Checks()
        {
            Models.Validate va = new Models.Validate();
            int rowcount = 0;
            rowcount = dataGridView1.Rows.Count - 1;
            if (combocompcode.Text == "")
            {
                MessageBox.Show("CompCode is Empty." + combocompcode.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                combocompcode.Select();
                return false;

            }
            if (combofinyear.Text == "")
            {
                MessageBox.Show("'combofinyear Field is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.combofinyear.Select();

                return false;
            }



            Models.HolidayModel.HolidayModeldetails c1 = new Models.HolidayModel.HolidayModeldetails();
            for (int i = 0; i < rowcount; i++)
            {


                if (Convert.ToString(dataGridView1.Rows[i].Cells[5].Value.ToString()) == null)
                {

                    MessageBox.Show("nhdate Field is empty");

                    return false;
                }
                if (Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[10].Value) <= 0)
                {
                    MessageBox.Show("holidaycategory Field is empty");
                    return false;
                }

                return true;
            }
            return isvalid;
        }
        public void Saves()
        {
            Int64 maxid = 0; Int64 j = 0;

            try
            {
                if (Checks() == true)
                {
                    Models.HolidayModel c1 = new Models.HolidayModel();
                    c1.active = "";
                    Models.Validate va = new Models.Validate();

    
                 
                    if (txtholiid.Text == "") { autonumberload(); c1.asptblholmasid1 = Convert.ToInt64("0" + txtholiid1.Text); txtholiid.Text = ""; }
                    else { c1.asptblholmasid = Convert.ToInt64("0" + txtholiid.Text); c1.asptblholmasid1 = Convert.ToInt64("0" + txtholiid1.Text); }
                    c1.finyear = Convert.ToString(combofinyear.Text);
                    c1.month = Convert.ToString(combomonth.Text);
                    c1.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    c1.compname = Convert.ToInt64("0" + combocompname.SelectedValue);
                    c1.docid = Convert.ToString(txtdocid.Text);
                    c1.date = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    if (checkactive.Checked == true)
                        c1.active = "T";
                    else
                        c1.active = "F";

                    c1.compcode1 = Convert.ToInt64(Class.Users.COMPCODE);
                    c1.username = Convert.ToInt64(Class.Users.USERID);
                    c1.createdby = Convert.ToString(Class.Users.HUserName);
                    c1.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString());
                    c1.modifiedby = Class.Users.HUserName;
                    c1.ipAddress = Class.Users.IPADDRESS;

                    if (combofinyear.Text == "")
                    {
                        MessageBox.Show("'combofinyear Field is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.combofinyear.Select();

                        return;
                    }
                    if (combocompcode.Text == "")
                    {
                        MessageBox.Show("'CompCode Field is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.combocompcode.Select();

                        return;
                    }
                    else
                    {

                        string sel = "select asptblholmasid    from  asptblholmas   WHERE  asptblholmasid1='" + c1.asptblholmasid1 + "' and finyear='" + c1.finyear + "' and compcode='" + c1.compcode + "' and compname='" + c1.compname + "' and month='" + c1.month + "'   and active='" + c1.active + "';";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblholmas");
                        DataTable dt = ds.Tables["asptblholmas"];
                        if (dt.Rows.Count != 0)
                        {

                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtholiid.Text) == 0 || Convert.ToInt32("0" + txtholiid.Text) == 0)
                        {
                            string ins = "insert into asptblholmas(asptblholmasid1,finyear,compcode,compname,month,docid,date,active,compcode1,username,createdby,createdon,modifiedby,ipaddress) values('" + c1.asptblholmasid1 + "','" + c1.finyear + "','" + c1.compcode + "','" + c1.compname + "','" + c1.month + "','" + c1.docid + "','" + c1.date + "','" + c1.active + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + dateTimePicker1.Value.ToString() + "','" + Class.Users.CREATED + "','" + Class.Users.IPADDRESS + "');";
                            Utility.ExecuteNonQuery(ins);
                            string sel2 = "select max(asptblholmasid) as asptblholmasid   from  asptblholmas   WHERE  finyear='" + c1.finyear + "' and compname='" + c1.compcode + "' ;";
                            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblholmas");
                            DataTable dt2 = ds2.Tables["asptblholmas"]; maxid = 0;
                            maxid = Convert.ToInt64(dt2.Rows[0]["asptblholmasid"].ToString());
                        }
                        else
                        {
                            string up = "update  asptblholmas  set asptblholmasid1='" + c1.asptblholmasid1 + "' , finyear='" + c1.finyear + "' , compcode='" + c1.compcode + "' , compname='" + c1.compname + "' , docid='" + c1.docid + "'  , month='" + c1.month + "' , date='" + c1.date + "',compcode1='" + Class.Users.COMPCODE + "',  username='" + Class.Users.USERID + "',createdby='" + Class.Users.HUserName + "', modifiedby='" + Class.Users.CREATED + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblholmasid='" + c1.asptblholmasid + "';";
                            Utility.ExecuteNonQuery(up);
                            maxid = 0;
                            maxid = Convert.ToInt64(txtholiid.Text);

                        }
                        int i = 0;
                        Models.HolidayModel.HolidayModeldetails c = new Models.HolidayModel.HolidayModeldetails();
                        int cc = dataGridView1.Rows.Count - 1;
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (Convert.ToDateTime(row.Cells["nhdate"].Value).ToString("dd-MM-yyyy") != "" && Convert.ToString(row.Cells["holidaycategory"].Value) != "")
                            {
                                if (txtholiid.Text == "") { c.asptblholmasid = Convert.ToInt64("0" + maxid); c.asptblholmasid1 = Convert.ToInt64("0" + txtholiid1.Text); }
                                else { c.asptblholmasid = Convert.ToInt64("0" + txtholiid.Text); c.asptblholmasid1 = Convert.ToInt64("0" + txtholiid.Text); }
                                c.asptblholdetid = Convert.ToInt64("0" + row.Cells["asptblholdetid"].Value);
                                c.compcode = Class.Users.COMPCODE;
                                c.nhdate = Convert.ToDateTime(row.Cells["nhdate"].Value).ToString("yyyy-MM-dd");
                                c.month = Convert.ToString(row.Cells["month"].Value);
                                c.day = Convert.ToString(row.Cells["day"].Value);
                                c.year = Convert.ToString(row.Cells["year"].Value);
                                c.Reason = Convert.ToString(row.Cells["Reason"].Value);
                                c.holidaycategory = Convert.ToInt64("0" + row.Cells["holidaycategory"].Value);
                                c.WorkingDay = Convert.ToString(row.Cells["WorkingDay"].Value);
                                c.notes = Convert.ToString(row.Cells["notes"].Value);
                                string sel1 = "select asptblholdetid    from  asptblholdet   where asptblholmasid='" + c.asptblholmasid + "'and asptblholmasid1='" + c.asptblholmasid1 + "'and compcode='" + c.compcode + "' and  nhdate='" + c.nhdate + "' and month='" + c.month + "' and day='" + c.day + "' and year='" + c.year + "' and Reason='" + c.Reason + "' and holidaycategory='" + c.holidaycategory + "' and WorkingDay='" + c.WorkingDay + "'and notes='" + c.notes + "';";
                                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblholdet");
                                DataTable dt1 = ds1.Tables["asptblholdet"];
                                if (dt1.Rows.Count != 0)
                                {
                                    tabControl1.SelectTab(tabPagedel2);
                                }
                                else if (dt1.Rows.Count != 0 && Convert.ToInt64("0" + c.asptblholdetid) == 0 || Convert.ToInt64("0" + c.asptblholdetid) == 0)
                                {

                                    string ins1 = "insert into asptblholdet(asptblholmasid,asptblholmasid1,compcode,nhdate,month,day,year,Reason,holidaycategory,WorkingDay,notes) values('" + c.asptblholmasid + "' ,'" + c.asptblholmasid1 + "' ,'" + c.compcode + "' ,'" + c.nhdate + "','" + c.month + "','" + c.day + "','" + c.year + "','" + c.Reason + "','" + c.holidaycategory + "','" + c.WorkingDay + "','" + c.notes + "' );";
                                    Utility.ExecuteNonQuery(ins1);
                                }
                                else
                                {
                                    string up1 = "update  asptblholdet  set asptblholmasid='" + c.asptblholmasid + "',asptblholmasid1='" + c.asptblholmasid1 + "',compcode='" + c.compcode + "',nhdate='" + c.nhdate + "',month='" + c.month + "',day='" + c.day + "',year='" + c.year + "',Reason='" + c.Reason + "',holidaycategory='" + c.holidaycategory + "',WorkingDay='" + c.WorkingDay + "',notes='" + c.notes + "'  where asptblholdetid='" + c.asptblholdetid + "';";
                                    Utility.ExecuteNonQuery(up1);
                                }
                            }
                        }
                        if (txtholiid.Text == "")
                        {
                            MessageBox.Show("Record Saved Successfully " + txtdocid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                            tabControl1.SelectTab(tabPagedel2);
                        }
                        else
                        {
                            MessageBox.Show("Record Updated Successfully " + txtdocid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                            tabControl1.SelectTab(tabPagedel2);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("mandatory  Field is Empty");
                    return;
                }
            }
            catch (Exception ex)
            {
                //  MessageBox.Show("Saves_Click " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //string sel1 = "ROLLBACK";
                //Utility.ExecuteNonQuery(sel1);
                // MessageBox.Show("mandatory  Field is Empty");
            }

        }

        private void HolidayMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

  

        private void empty()
        {
            txtholiid.Text = ""; combocompcode.SelectedIndex = -1; combocompcode.Text = ""; combounit.SelectedIndex = -1; combounit.Text = "";
            combocompname.SelectedIndex = -1; combocompname.Text = ""; combocompcode.Enabled = true; combofinyear.Text = Class.Users.Finyear; dateTimePicker1.Value = System.DateTime.Now;
            combounit.SelectedIndex = -1; combounit.Text = "";

            listfilter.Items.Clear();
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            txtdocid.Text = ""; dtp.Visible = false;
            dataGridView1.AllowUserToAddRows = true;
            listView1.Font = Class.Users.FontName;
            this.Font = Class.Users.FontName;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.EnableHeadersVisualStyles = false;
            do
            {
                int i = 0;
                for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }

            }
            while (dataGridView1.Rows.Count > 1);


            if (listView1.Items.Count >= 1)
            {
                listView1.Items[0].Selected = true;
            }
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear(); //date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')
                string sel1 = "select a.asptblholmasid,a.finyear,a.docid,c.compcode,date_format(a.date,'%d-%m-%Y') as date,a.active from asptblholmas a join gtcompmast c on a.compcode=c.gtcompmastid  order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblholmas");
                DataTable dt = ds.Tables["asptblholmas"];
                if (dt.Rows.Count >= 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblholmasid"].ToString());
                        list.SubItems.Add(myRow["finyear"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["date"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
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

                    txtholiid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select a.asptblholmasid, a.asptblholmasid1,a.finyear,c.gtcompmastid as  compcode,c.compname,a.month,a.docid,a.date,a.active from asptblholmas a join gtcompmast c on a.compcode=c.gtcompmastid  where a.asptblholmasid='" + txtholiid.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblholmas");
                    DataTable dt = ds.Tables["asptblholmas"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtholiid.Text = Convert.ToString(dt.Rows[0]["asptblholmasid"].ToString());
                        txtholiid1.Text = Convert.ToString(dt.Rows[0]["asptblholmasid1"].ToString());
                        combofinyear.Text = Convert.ToString(dt.Rows[0]["finyear"].ToString());
                        combocompcode.SelectedValue = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        combomonth.Text = Convert.ToString(dt.Rows[0]["month"].ToString());
                        txtdocid.Text = Convert.ToString(dt.Rows[0]["docid"].ToString());
                        dateTimePicker1.Text = Convert.ToString(dt.Rows[0]["date"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        string sel2 = "select a.asptblholdetid, a.asptblholmasid,a.asptblholmasid1,a.compcode,a.nhdate,a.day,a.month,a.year,a.Reason,e.asptblhrholidaycatmasid as holidaycategory,a.WorkingDay,a.notes from asptblholdet a join asptblholmas b on a.asptblholmasid = b.asptblholmasid join gtcompmast d on a.compcode = d.gtcompmastid join asptblhrholidaycatmas e on e.asptblhrholidaycatmasid=a.holidaycategory where a.asptblholmasid='" + txtholiid.Text + "';";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblholdet");
                        DataTable dt1 = ds2.Tables["asptblholdet"];
                        if (dt1.Rows.Count > 0)
                        {
                            dataGridView1.DataSource = dt1;

                            for (i = 0; i < dt1.Rows.Count; i++)
                            {

                                if (Convert.ToInt64(dataGridView1.Rows[i].Cells[1].Value) > 0)
                                {
                                    dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblholdetid"].ToString());
                                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblholmasid"].ToString());
                                    dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblholmasid1"].ToString());
                                    dataGridView1.Rows[i].Cells[4].Value = Convert.ToInt64("0" + dt1.Rows[i]["compcode"].ToString());
                                    dataGridView1.Rows[i].Cells[5].Value = Convert.ToString(dt1.Rows[i]["nhdate"].ToString());
                                    dataGridView1.Rows[i].Cells[6].Value = Convert.ToString(dt1.Rows[i]["day"].ToString());
                                    dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(dt1.Rows[i]["month"].ToString());
                                    dataGridView1.Rows[i].Cells[8].Value = dt1.Rows[i]["year"].ToString();
                                    dataGridView1.Rows[i].Cells[9].Value = dt1.Rows[i]["Reason"].ToString();
                                    dataGridView1.Rows[i].Cells[10].Value = dt1.Rows[i]["holidaycategory"].ToString();
                                    dataGridView1.Rows[i].Cells[11].Value = dt1.Rows[i]["WorkingDay"].ToString();
                                    dataGridView1.Rows[i].Cells[12].Value = dt1.Rows[i]["notes"].ToString();

                                }
                            }


                        }
                        combocompcode.Enabled = false;
                        tabControl1.SelectTab(tabPagedel1);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ListView1_ItemActivate: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {

            try
            {
                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length >= 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[5].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[6].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    listView1.Items.Clear();

                    foreach (ListViewItem item in listfilter.Items)
                    {

                        this.listView1.Items.Add((ListViewItem)item.Clone());

                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }


        }





        //private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        //{
        //    GridLoad(); holidaycateload();
        //    companyload(); combocompcode.Select();

        //    usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName); companyload();
        //}



        private void Searchs_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPagedel2);

        }

        private void Deletes_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtholiid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {


                        string del1 = "delete from asptblholdet where compcode='" + combocompcode.SelectedValue + "' and  asptblholmasid='" + txtholiid.Text + "';";
                        Utility.ExecuteNonQuery(del1);
                        string del = "delete from asptblholmas where compcode='" + combocompcode.SelectedValue + "' and  asptblholmasid='" + txtholiid.Text + "';";
                        Utility.ExecuteNonQuery(del);
                        MessageBox.Show("Record Deleted Successfully " + txtholiid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletes_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void Prints_Click(object sender, EventArgs e)
        {

        }


        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtholiid.Text == "")
                {

                    autonumberload();

                }
             


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        bool valid = false;
        //private void txtvechileno_TextChanged(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        Models.Validate va = new Models.Validate();
        //        valid = va.IsStringNumberic(txtvechileno.Text);
        //        if (valid == false)
        //        {
        //            MessageBox.Show("Special Charecters not allowed", "Informmation", MessageBoxButtons.OK, MessageBoxIcon.Error); return;
        //            txtvechileno.Select();
        //            return;
        //        }


        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("txtvechileno_TextChanged" + ex.Message);
        //    }
        //}

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPagedel1"])//your specific tabname
            {

                combocompcode_SelectedIndexChanged(sender, e);
                combocompcode.Select();

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPagedel2"])//your specific tabname
            {
                txtsearch.Select();

            }
        }








        //private void txtbags_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        //}

        //private void txtsecondweight_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
        //    {
        //        e.Handled = false; //Do not reject the input
        //    }
        //    else
        //    {
        //        e.Handled = true; //Reject the input
        //    }
        //}


        //private void txtvechileno_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        //}





        private void combocompcode_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            autonumberload();
        }
        string day1 = "";
        string month1 = "";
        string year1 = "";
        int ind = 0; Int64 griddelrow = 0;
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (e.ColumnIndex == 5)
                {
                    //case 5: // Column index of needed dateTimePicker cell

                    rectangle = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtp.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtp.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtp.Visible = true;
                    ind = e.RowIndex;
                }
                if (txtholiid.Text != "")
                {
                    griddelrow = 0;
                    griddelrow = Convert.ToInt64("0"+dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());

                }
                dataGridView1.BeginEdit(true);
            }
            catch (Exception ex) { }

        }

        private void Dtp_TextChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.CurrentCell.Value = dtp.Value.ToString("dd-MM-yyyy");
                day1 = dtp.Value.ToString("dddd");
                month1 = dtp.Value.ToString("MMMM");
                year1 = dtp.Value.ToString("yyyy");
                dataGridView1.Rows[ind].Cells[6].Value = "";
                dataGridView1.Rows[ind].Cells[7].Value = "";
                dataGridView1.Rows[ind].Cells[8].Value = "";
                dataGridView1.Rows[ind].Cells[6].Value = day1;
                dataGridView1.Rows[ind].Cells[7].Value = month1;
                dataGridView1.Rows[ind].Cells[8].Value = year1;
            }
            catch (Exception ex) { }



            //MessageBox.Show(string.Format("Date changed to {0}", dateTimePicker1.Text.ToString() + "mm" + month1 + "year" + year1)); 

        }

        // on text change of dtp, assign back to cell


        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            try
            {
                // dtp.Visible = false;
            }
            catch (Exception ex) { }
        }

        private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        if (txtholiid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow >=1)
                                {
                                    string del1 = "delete from  asptblholdet     Where  asptblholdetid='" + griddelrow + "';";
                                    Utility.ExecuteNonQuery(del1);
                                    dataGridView1.Rows.RemoveAt(Convert.ToInt32(griddelrow));
                                    griddelrow = 0;
                                }
                                else
                                {
                                    dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                                }
                            }
                        }
                        else
                        {
                            dataGridView1.Rows.RemoveAt(oneCell.RowIndex);

                        }
                    }
                }

            }
            catch (Exception EX)
            {
                // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            }
        }


        private void HolidayMaster_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control == true && e.KeyCode == Keys.S)
            {
                Saves();
            }
        }

        private void combomonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear(); //date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')
                string sel1 = "select a.asptblholmasid,a.finyear,a.docid,c.compcode,date_format(a.date,'%d-%m-%Y') as date,a.active from asptblholmas a join gtcompmast c on a.compname=c.gtcompmastid  order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblholmas");
                DataTable dt = ds.Tables["asptblholmas"];
                if (dt.Rows.Count >= 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblholmasid"].ToString());
                        list.SubItems.Add(myRow["finyear"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["date"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());

                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name.ToString() == "nhdate")
            {
                dataGridView1.Columns[e.ColumnIndex].Name.FirstOrDefault();
                // dataGridView1.CurrentCell = dataGridView1[dataGridView1.CurrentCell.ColumnIndex + 1, dataGridView1.CurrentCell.RowIndex];
            }
        }

        private void combounit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combounit.Text != "")
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptblholmasid, a.asptblholmasid1,a.finyear,c.gtcompmastid as  compcode,c.compname,a.docid,a.date,a.active from asptblholmas a join gtcompmast c on a.compcode=c.gtcompmastid where c.compcode='" + combounit.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblholmas");
                DataTable dt = ds.Tables["asptblholmas"];
                if (dt.Rows.Count >= 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblholmasid"].ToString());
                        list.SubItems.Add(myRow["finyear"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["date"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());

                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad(); holidaycateload();
            companyload(); combocompcode.Select();



        }

        private void refreshToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            holidaycateload();companyload();
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
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }





        //private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (listView2.Items.Count > 0)
        //        {
        //            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        //            if (confirmation == DialogResult.Yes)
        //            {
        //                int i = 0; decimal d1 = Convert.ToDecimal("0" + txtproductweight.Text);
        //                decimal d2 = 0, d3 = 0;
        //                for (i = 0; i < listView2.Items.Count; i++)
        //                {

        //                    if (listView2.Items[i].Selected)
        //                    {
        //                        if (Convert.ToInt64("0" + listView2.Items[i].SubItems[2].Text) > 1)
        //                        {
        //                            string del1 = "delete from ASPTBLNHDAYDET where asptblholdetid='" + listView2.Items[i].SubItems[2].Text + "';";
        //                            Utility.ExecuteNonQuery(del1);
        //                        }
        //                        MessageBox.Show("Index of Row:   " + listView2.Items[i].SubItems[1].Text + "      Name:  " + listView2.Items[i].SubItems[5].Text + "Deleted Successfully", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                        d2 = Convert.ToDecimal("0" + listView2.Items[i].SubItems[7].Text);
        //                        d3 = d1 - d2;

        //                        listView2.Items[i].Remove();
        //                        i--;
        //                    }
        //                    txtproductweight.Text = d3.ToString();
        //                    lbltotalfooterkgs.Text = d3.ToString();
        //                    listview2totalweight = Convert.ToDecimal("0" + d3.ToString()); comboproduct.SelectedIndex = -1;
        //                }
        //                if (listView2.Items.Count == 0)
        //                {
        //                    comboproduct.SelectedIndex = -1;
        //                    txtproductweight.Text = ""; txtdifference.Text = ""; lbltotalfooterkgs.Text = ""; lbltotal1.Text = "Total Count: ";
        //                }


        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("If you want to Remove,Double Click a Specific Row in ListView.", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //    //try
        //    //{
        //    //    if (listView2.Items.Count >= 0)
        //    //    {
        //    //        var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        //    //        if (confirmation == DialogResult.Yes)
        //    //        {
        //    //            int i = 0;
        //    //            for (i = 0; i < listView2.Items.Count; i++)
        //    //            {

        //    //                if (listView2.Items[i].Selected)
        //    //                {
        //    //                    if (Convert.ToInt64("0" + listView2.Items[i].SubItems[2].Text) > 1)
        //    //                    {
        //    //                        string del1 = "delete from ASPTBLNHDAYDET where asptblholdetid='" + listView2.Items[i].SubItems[2].Text + "';";
        //    //                        Utility.ExecuteNonQuery(del1);
        //    //                    }
        //    //                    MessageBox.Show("S.No:   " + listView2.Items[i].SubItems[1].Text + "      Row ID:  " + listView2.Items[i].SubItems[2].Text, "Delete this Record");
        //    //                    listView2.Items[i].Remove();
        //    //                    i--;
        //    //                }
        //    //            }
        //    //        }
        //    //    }
        //    //    else
        //    //    {
        //    //        MessageBox.Show("If you want to Remove,Double Click a Specific Row in ListView.", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //    //    }
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    MessageBox.Show(ex.ToString());
        //    //}
        //}
    }
}

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
    public partial class ShiftMaster : Form, ToolStripAccess
    {
        public ShiftMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            tabControl1.SelectTab(tabPage8);
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            panel6.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            panel9.BackColor = Class.Users.BackColors;
            panel10.BackColor = Class.Users.BackColors;
            panel11.BackColor = Class.Users.BackColors;
            panel12.BackColor = Class.Users.BackColors;
        }
  

    private static ShiftMaster _instance; string coid = "", siid = "";
    Models.Master mas = new Models.Master();
    Models.UserRights sm = new Models.UserRights();
    ListView listfilterbreak = new ListView();
    ListView listfilter = new ListView();
    ListView listfilter2 = new ListView();
    ListView listfilter4 = new ListView();
    public static ShiftMaster Instance
    {
        get
        {
            if (_instance == null)
                _instance = new ShiftMaster();
            GlobalVariables.CurrentForm = _instance; return _instance;
        }
    }


    public void autonumberload()
    {
        try
        {

            string sel = "select max(a.asptblshimasid)+1 as id  from asptblshimas a join gtcompmast b on a.compcode = b.gtcompmastid where b.compcode ='" + combocompcode.Text + "'; ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblshimas");
            DataTable dt = ds.Tables["asptblshimas"];
            int cnt = dt.Rows.Count;
            if (dt.Rows[0]["id"].ToString() == "")
            {



                txtshiftid1.Text = "1";
            }
            else
            {


                txtshiftid1.Text = dt.Rows[0]["id"].ToString();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
        private void ShiftMaster_Load(object sender, EventArgs e)
        {
            News();
        }

    public void Saves()
    {
        try
        {
            if (tabControl1.SelectedTab.Text == "Break Master")
            {
                try
                {
                    if (txtbreakname.Text == "")
                    {
                        MessageBox.Show("breakname Name is empty " + " Alert " + txtbreakname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (txtbreakname.Text != "")
                    {


                        string chk = "";
                        if (checkbreak.Checked == true) { chk = "T"; } else { chk = "F"; checkbreak.Checked = false; }
                        string sel = "select asptblbreakmasid    from  asptblbreakmas    WHERE compcode='" + combocompcodebreak.SelectedValue + "' and breakname='" + txtbreakname.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbreakmas");
                        DataTable dt = ds.Tables["asptblbreakmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtbreakname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtbreakid.Text) == 0 || Convert.ToInt32("0" + txtbreakid.Text) == 0)
                        {
                            string ins = "insert into asptblbreakmas(breakname,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtbreakname.Text.ToUpper() + "','" + chk + "','" + combocompcodebreak.SelectedValue + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtbreakname.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoadbreak(); empty();
                        }
                        else
                        {
                            string up = "update  asptblbreakmas  set   breakname='" + txtbreakname.Text.ToUpper() + "' , active='" + chk + "' ,compcode='" + combocompcodebreak.SelectedValue + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblbreakmasid='" + txtbreakid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtbreakname.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoadbreak();
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

                    MessageBox.Show("breakname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            if (tabControl1.SelectedTab.Text == "     Shift Line Assignment     ")
            {
                if (comboshift.Text == "" && listView4.Items.Count < 0)
                {
                    MessageBox.Show("colorname Name is empty " + " Alert " + comboshift.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (comboshift.Text != "" && listView4.Items.Count > 0)
                {
                    string maxid = "";
                    Models.ShiftModel p = new Models.ShiftModel();
                    p.asptblshimasid = Convert.ToInt64("0" + txtshiftid.Text);
                    p.asptblshimasid1 = Convert.ToString(txtshiftid.Text);
                    p.shiftdate = Convert.ToDateTime(dateTimePicker1.Value).ToString("yyyy-MM-dd");
                        p.compcode = Convert.ToInt64(combocompcode.SelectedValue);
                    p.finyear = Class.Users.Finyear;
                    p.linegroup = Convert.ToInt64("0" + comboline.SelectedValue);
                    p.shiftname = Convert.ToInt64(comboshift.SelectedValue);
                    p.shiftno = Convert.ToInt64(comboshiftno.SelectedValue);
                        p.breaktime = Convert.ToInt64(combobreak.SelectedValue);
                        p.shiftstart =  Convert.ToDateTime(dateshiftstart.Value).ToString("HH:mm:ss");
                        p.shiftend =  Convert.ToDateTime(dateshiftend.Value).ToString("HH:mm:ss");
                        p.breakstart = Convert.ToDateTime(datebreakstart.Value).ToString("HH:mm:ss");
                        p.breakend = Convert.ToDateTime(datebreakend.Value).ToString("HH:mm:ss");
                        p.otminutes = Convert.ToDateTime(dateottime.Value).ToString("HH:mm:ss");
                        p.breakminuts = Convert.ToDateTime(dateottime1.Value).ToString("HH:mm:ss");

                        p.compcode1 = Class.Users.COMPCODE;
                    p.username = Class.Users.USERID;
                    p.createdby = Convert.ToString(Class.Users.HUserName);
                    p.createdon = Convert.ToDateTime(System.DateTime.Now.ToString());
                    p.modified = Convert.ToString(System.DateTime.Now.ToString());
                    p.modifiedby = Class.Users.HUserName;
                    p.ipaddress = Class.Users.IPADDRESS;


                    if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; checkactive.Checked = false; }
                    if (checkshiftcancel.Checked == true) { p.shiftcancel = "T"; } else { p.shiftcancel = "F"; checkshiftcancel.Checked = false; }
                        string sel = "select asptblshimasid    from  asptblshimas    WHERE finyear='"+p.finyear+"' and  shiftdate='" + p.shiftdate + "' and compcode='" + p.compcode + "' and linegroup='" + p.linegroup + "' and shiftname='" + p.shiftname + "' and shiftno='" + p.shiftno + "' and shiftstart='" + p.shiftstart + "'  and shiftend='" + p.shiftend + "' and breaktime='" + p.breaktime + "' and breakstart='" + p.breakstart + "' and breakend='" + p.breakend + "' and otminutes='" + p.otminutes + "' and breakminuts='" + p.breakminuts + "' and active='" + p.active + "' and shiftcancel='" + p.shiftcancel + "' ";

                        // string sel = "select asptblshimasid    from  asptblshimas    WHERE  shiftdate=date_format('" + p.shiftdate + "') and compcode='" + p.compcode + "' and linegroup='" + p.linegroup + "' and shiftname='" + p.shiftname + "' and shiftno='" + p.shiftno + "' and shiftstart=TIME_FORMAT('"+p.shiftstart+ "', ' % h:% i % p')  and shiftend=TIME_FORMAT('" + p.shiftend + "', ' % h:% i % p') and breaktime=TIME_FORMAT('" + p.breaktime + "', ' % h:% i % p') and breakstart=TIME_FORMAT('" + p.breakstart + "', ' % h:% i % p') and breakend=TIME_FORMAT('" + p.breakend + "', ' % h:% i % p') and otminutes=TIME_FORMAT('" + p.otminutes + "', ' % h:% i % p') and breakminuts=TIME_FORMAT('" + p.breakminuts + "', ' % h:% i % p') and active='" + p.active + "' and shiftcancel='" + p.shiftcancel + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblshimas");
                    DataTable dt = ds.Tables["asptblshimas"];
                    if (dt.Rows.Count != 0)
                    {
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtshiftid.Text) == 0 || Convert.ToInt32("0" + txtshiftid.Text) == 0)
                    {
                        string ins = "insert into asptblshimas(finyear,shiftdate,compcode,linegroup,shiftname,shiftno,shiftstart,shiftend,breaktime,breakstart,breakend,otminutes,breakminuts,active,shiftcancel,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.finyear + "','" + p.shiftdate + "','" + p.compcode + "','" + p.linegroup + "','" + p.shiftname + "','" + p.shiftno + "','" + p.shiftstart + "','" + p.shiftend + "','" + p.breaktime + "','" + p.breakstart + "','" + p.breakend + "','" + p.otminutes + "','" + p.breakminuts + "','" + p.active + "','" + p.shiftcancel + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + p.modified + "','" + p.modifiedby + "','" + p.ipaddress + "');";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptblshimasid) id    from  asptblshimas   where  compcode='" + combocompcode.SelectedValue + "' ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblshimas");
                        DataTable dt2 = ds2.Tables["asptblshimas"];
                        maxid = dt2.Rows[0]["id"].ToString();
                    }
                    else
                    {
                            string up = "update  asptblshimas  set   asptblshimasid1='" + p.asptblshimasid + "' ,finyear='" + p.finyear + "',shiftdate='" + p.shiftdate + "' , compcode='" + p.compcode + "' , linegroup='" + p.linegroup + "' , shiftname='" + p.shiftname + "' , shiftno='" + p.shiftno + "' , shiftstart='" + p.shiftstart + "'  , shiftend='" + p.shiftend + "' , breaktime='" + p.breaktime + "' , breakstart='" + p.breakstart + "' , breakend='" + p.breakend + "' , otminutes='" + p.otminutes + "' , breakminuts='" + p.breakminuts + "' , active='" + p.active + "' , shiftcancel='" + p.shiftcancel + "',compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblshimasid='" + txtshiftid.Text + "';";

                            Utility.ExecuteNonQuery(up);
                        maxid = txtshiftid.Text;
                    }
                    Models.ShiftModel.ShiftModeldet pp = new Models.ShiftModel.ShiftModeldet();
                    int i = 0;
                    if (listView4.Items.Count >= 0)
                    {
                        for (i = 0; i < listView4.Items.Count; i++)
                        {
                            pp.asptblshidetid = Convert.ToInt64("0" + listView4.Items[i].SubItems[2].Text);
                            pp.asptblshimasid = Convert.ToInt64("0" + maxid.ToString());
                            pp.asptblshimasid1 = Convert.ToString(maxid.ToString());
                            pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                            pp.shiftname = Convert.ToInt64("0" + comboshift.SelectedValue);
                            pp.weakday = Convert.ToString(listView4.Items[i].SubItems[3].Text);
                            lineid(listView4.Items[i].SubItems[4].Text);
                            pp.linename = siid;
                            string sel1 = "select asptblshimasid    from  asptblshidet   where   asptblshimasid='" + pp.asptblshimasid + "' and asptblshimasid1='" + pp.asptblshimasid1 + "' and compcode='" + pp.compcode + "' and shiftname='" + pp.shiftname + "' and weakday='" + pp.weakday + "' and linename='" + pp.linename + "' ;";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblshidet");
                            DataTable dt1 = ds1.Tables["asptblshidet"];
                            if (dt1.Rows.Count != 0)
                            {

                                
                            }
                            else if (dt1.Rows.Count != 0 && pp.asptblshidetid == 0 || pp.asptblshidetid == 0)
                            {
                                string ins1 = "insert into asptblshidet(asptblshimasid,asptblshimasid1,compcode,shiftname,weakday,linename) values('" + pp.asptblshimasid + "' ,'" + pp.asptblshimasid1 + "' , '" + pp.compcode + "' ,'" + pp.shiftname + "' , '" + pp.weakday + "' ,'" + pp.linename + "');";
                                Utility.ExecuteNonQuery(ins1);
                            }
                            else
                            {
                                string up1 = "update  asptblshidet  set asptblshimasid='" + pp.asptblshimasid + "' ,asptblshimasid1='" + pp.asptblshimasid1 + "', compcode='" + pp.compcode + "' , shiftname='" + pp.shiftname + "' , weakday='" + pp.weakday + "' ,linename='" + pp.linename + "' where asptblshidetid='" + pp.asptblshidetid + "';";
                                Utility.ExecuteNonQuery(up1);
                            }


                        }
                    }
                    if (txtshiftid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtshiftid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        autonumberload();
                        tabControl1.SelectTab(tabPage8);
                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtshiftid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        autonumberload();
                        tabControl1.SelectTab(tabPage8);
                    }
                }
                else
                {
                    MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }
            }
        }
        catch (Exception ex)
        {

            MessageBox.Show("colorname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }

    }

    private void ShiftMaster_FormClosed(object sender, FormClosedEventArgs e)
    {
        _instance = null;
    }



    public void News()
    {

        empty();
        GridLoad(); GridLoadbreak(); linegroupload(); compload(); shiftload(); breakload(); DayListview();
    }
        private void empty()
        {
            txtshiftid1.Text = "";
            txtshiftid.Text = ""; txtbreakid.Text = ""; txtbreakname.Text = ""; txtsearchbreak.Text = "";
            comboshift.Text = ""; listfilterbreak.Items.Clear(); listfilter.Items.Clear();
            comboline.Text = ""; comboline.SelectedIndex = -1;
            combocompcode.Text = ""; combocompcode.SelectedIndex = -1;
            comboshift.Text = ""; comboshift.SelectedIndex = -1;

            txtshiftid1.Text = "";
            checkactive.Checked = true;
            checkshiftcancel.Checked = false; listView4.Items.Clear();
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            panel6.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            panel9.BackColor = Class.Users.BackColors;
            panel10.BackColor = Class.Users.BackColors;
            panel11.BackColor = Class.Users.BackColors;
            panel12.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            listView3.Font = Class.Users.FontName;
            listView4.Font = Class.Users.FontName;
            listViewbreak.Font = Class.Users.FontName;
        }
    public void linegroupload()
    {
        try
        {
            string sel = "select asptbllingrpmasid,linegroup from  asptbllingrpmas  order by 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllingrpmas");
            DataTable dt = ds.Tables["asptbllingrpmas"];

            comboline.DisplayMember = "linegroup";
            comboline.ValueMember = "asptbllingrpmasid";
            comboline.DataSource = dt;
        }
        catch (Exception EX)
        { }
    }
    public void shiftload()
    {
        try
        {
            string sel = "select a.asptblshitypeid,a.shifttype from  asptblshitype a  order by 1 ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblshitype");
            DataTable dt = ds.Tables["asptblshitype"];

            comboshift.DisplayMember = "shifttype";
            comboshift.ValueMember = "asptblshitypeid";
            comboshift.DataSource = dt;

            comboshift.Text = ""; comboshift.SelectedIndex = -1;
            comboshiftno.Text = ""; comboshiftno.SelectedIndex = -1;
        }
        catch (Exception EX)
        { }
    }
    public void breakload()
    {
        try
        {
            string sel = "select asptblbreakmasid,breakname from  asptblbreakmas  order by 1 ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbreakmas");
            DataTable dt = ds.Tables["asptblbreakmas"];

            combobreak.DisplayMember = "breakname";
            combobreak.ValueMember = "asptblbreakmasid";
            combobreak.DataSource = dt;
            combobreak.Text = ""; combobreak.SelectedIndex = -1;
        }
        catch (Exception EX)
        { }
    }


    public void lineid(string s)
    {
        try
        {
            string sel = "select asptbllinemasid from  asptbllinemas where lineno='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllinemas");
            DataTable dt = ds.Tables["asptbllinemas"];
            siid = "";
            siid = Convert.ToString(dt.Rows[0]["asptbllinemasid"].ToString());

        }
        catch (Exception EX)
        { }
    }
    //public void lineid(string s)
    //{
    //    try
    //    {
    //        string sel = "select asptbllinemasid from  asptbllinemas where linename='" + s + "' ;";
    //        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllinemas");
    //        DataTable dt = ds.Tables["asptbllinemas"];
    //        siid = "";
    //        siid = Convert.ToString(dt.Rows[0]["asptbllinemasid"].ToString());

    //    }
    //    catch (Exception EX)
    //    { }
    //}


    public void compload()
    {
        try
        {

            DataTable dt = mas.comcode();

            combocompcode.DisplayMember = "compcode";
            combocompcode.ValueMember = "gtcompmastid";
            combocompcode.DataSource = dt;
            combocompcode.Text = ""; combocompcode.SelectedIndex = -1;
                combocompcodebreak.DisplayMember = "compcode";
                combocompcodebreak.ValueMember = "gtcompmastid";
                combocompcodebreak.DataSource = dt;
                combocompcodebreak.Text = ""; combocompcodebreak.SelectedIndex = -1;
            }
        catch (Exception EX)
        { }
    }
    private void comboshift_SelectedIndexChanged(object sender, EventArgs e)
    {
        string sel = "select a.asptblshitypeid,a.shiftno from  asptblshitype a where a.shifttype='" + comboshift.Text + "' ;";
        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblshifttype");
        DataTable dt = ds.Tables["asptblshifttype"];
        comboshiftno.DisplayMember = "shiftno";
        comboshiftno.ValueMember = "asptblshitypeid";
        comboshiftno.DataSource = dt;
    }
    private void lineload(string lingroup)
    {
        try
        {
            listView3.Items.Clear();
            string sel1 = "SELECT c.asptbllinemasid, c.lineno   FROM  asptbllingrpdet a join asptbllingrpmas b on a.asptbllingrpmasid=b.asptbllingrpmasid join asptbllinemas c on c.asptbllinemasid=a.linename  where b.linegroup='" + lingroup + "'  order by 1;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllinemas");
            DataTable dt = ds.Tables["asptbllinemas"];
            if (dt.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataRow myRow in dt.Rows)
                {
                    ListViewItem list = new ListViewItem();
                    list.SubItems.Add(i.ToString());
                    list.SubItems.Add(myRow["lineno"].ToString());
                    list.SubItems.Add("");
                    listView3.Items.Add(list);
             
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                }
                lbltotalline.Text = "Total Count    :" + listView3.Items.Count;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
    private void GridLoadbreak()
    {
        try
        {
            listViewbreak.Items.Clear();
            string sel1 = "   SELECT A.asptblbreakmasid, A.breakname ,b.compcode, a.active  FROM  asptblbreakmas a join gtcompmast b on a.compcode=b.gtcompmastid where a.active='T'  order by 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbreakmas");
            DataTable dt = ds.Tables["asptblbreakmas"];
            if (dt.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataRow myRow in dt.Rows)
                {
                    ListViewItem list = new ListViewItem();
                    list.SubItems.Add(i.ToString());
                    list.SubItems.Add(myRow["asptblbreakmasid"].ToString());
                    list.SubItems.Add(myRow["breakname"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                    this.listfilterbreak.Items.Add((ListViewItem)list.Clone());
                    listViewbreak.Items.Add(list);
                     
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2; i++;
                }
                lbltotalbreak.Text = "Total Count    :" + listViewbreak.Items.Count;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
    public void GridLoad()
    {
        try
        {
            listView1.Items.Clear();
            string sel1 = "select a.asptblshimasid ,c.compcode,d.shifttype ,d.shiftno,  e.linegroup,a.shiftstart,a.shiftend, f.breakname,a.breakstart,a.breakend,a.active from  asptblshimas a join gtcompmast c on c.gtcompmastid=a.compcode  join asptblshitype d on d.asptblshitypeid=a.shiftname join asptbllingrpmas e on e.asptbllingrpmasid=a.linegroup join asptblbreakmas f on f.asptblbreakmasid=a.breaktime   order by a.asptblshimasid desc;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblshimas");
            DataTable dt = ds.Tables["asptblshimas"];
            if (dt.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataRow myRow in dt.Rows)
                {
                    ListViewItem list = new ListViewItem();
                    list.SubItems.Add(i.ToString());
                    list.SubItems.Add(myRow["asptblshimasid"].ToString());
                    list.SubItems.Add(myRow["compcode"].ToString());
                    list.SubItems.Add(myRow["shifttype"].ToString());
                    list.SubItems.Add(myRow["shiftno"].ToString());
                    list.SubItems.Add(myRow["linegroup"].ToString());
                    list.SubItems.Add(myRow["shiftstart"].ToString());
                    list.SubItems.Add(myRow["shiftend"].ToString());
                    list.SubItems.Add(myRow["breakname"].ToString());
                    list.SubItems.Add(myRow["breakstart"].ToString());
                    list.SubItems.Add(myRow["breakend"].ToString());
                    list.SubItems.Add(myRow["active"].ToString());
                    this.listfilter.Items.Add((ListViewItem)list.Clone());
                    listView1.Items.Add(list);
                     
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2; i++;
                }
                lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
        public void DayListview()
        {
            try
            {                

                int i = 0;
                foreach (ListViewItem myRow in listView2.Items)
                {
                    myRow.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2; i++;
                }
                i = 0;
                foreach (ListViewItem myRow in listView7.Items)
                {
                    myRow.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2; i++;
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

                txtshiftid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                string sel1 = "select a.asptblshimasid,a.asptblshimasid1,a.shiftdate,b.compcode,c.shifttype,c.shiftno,d.linegroup,a.shiftstart, a.shiftend,f.breakname,a.breakstart, a.breakend,a.otminutes, a.breakminuts, a.active,a.shiftcancel from  asptblshimas a join gtcompmast b on b.gtcompmastid=a.compcode    join asptblshitype c on c.asptblshitypeid=a.shiftname join asptbllingrpmas d on d.asptbllingrpmasid=a.linegroup  join asptblbreakmas f on f.asptblbreakmasid=a.breaktime  where a.asptblshimasid=" + txtshiftid.Text;
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblshimas");
                DataTable dt = ds.Tables["asptblshimas"];
                    if (dt.Rows.Count > 0)
                    {
                        txtshiftid.Text = Convert.ToString(dt.Rows[0]["asptblshimasid"].ToString());
                        txtshiftid1.Text = Convert.ToString(dt.Rows[0]["asptblshimasid1"].ToString());
                        dateTimePicker1.Value = Convert.ToDateTime(dt.Rows[0]["shiftdate"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        comboline.Text = Convert.ToString(dt.Rows[0]["linegroup"].ToString());
                        comboshift.Text = Convert.ToString(dt.Rows[0]["shifttype"].ToString());
                        comboshiftno.Text = Convert.ToString(dt.Rows[0]["shiftno"].ToString());
                        comboline.Text = Convert.ToString(dt.Rows[0]["linegroup"].ToString());
                        dateshiftstart.Text = Convert.ToString(dt.Rows[0]["shiftstart"].ToString());
                        dateshiftstart.Text = Convert.ToString(dt.Rows[0]["shiftend"].ToString());
                        combobreak.Text = Convert.ToString(dt.Rows[0]["breakname"].ToString());
                        datebreakstart.Text = Convert.ToString(dt.Rows[0]["breakstart"].ToString());
                        datebreakend.Text = Convert.ToString(dt.Rows[0]["breakend"].ToString());
                        dateottime.Text = Convert.ToString(dt.Rows[0]["otminutes"].ToString());
                        dateottime1.Text = Convert.ToString(dt.Rows[0]["breakminuts"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        if (dt.Rows[0]["shiftcancel"].ToString() == "T") { checkshiftcancel.Checked = true; } else { checkshiftcancel.Checked = false; }
                    }
                string sel2 = "select a.asptblshimasid ,a.weakday, e.lineno from asptblshidet a join asptblshimas b on a.asptblshimasid=b.asptblshimasid join gtcompmast c on c.gtcompmastid=b.compcode  join asptblshitype d on d.asptblshitypeid=a.shiftname join asptbllinemas e on e.asptbllinemasid=a.linename  where a.asptblshimasid=" + txtshiftid.Text;
                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblshimas");
                DataTable dt2 = ds2.Tables["asptblshimas"];
                if (dt2.Rows.Count > 0)
                {
                    int i = 1; listView4.Items.Clear(); listfilter4.Items.Clear();
                    foreach (DataRow myRow in dt2.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblshimasid"].ToString());
                        list.SubItems.Add(myRow["weakday"].ToString());
                        list.SubItems.Add(myRow["lineno"].ToString());
                        this.listfilter4.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }
                        listView4.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count    :" + listView4.Items.Count;
                }
            }

        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }
        tabControl1.SelectTab(tabPage1);
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
                        if (listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[6].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[9].ToString().Contains(txtsearch.Text))
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

                listView1.Items.Clear(); int i = listfilter.Items.Count;
                foreach (ListViewItem item in listfilter.Items)
                {

                    this.listView1.Items.Add((ListViewItem)item.Clone());

                    if (i % 2 == 0) { item.BackColor = Color.WhiteSmoke; } else { item.BackColor = Color.White; }
                    item0++; i++;
                }
                lbltotal.Text = "Total Count: " + listView1.Items.Count;
            }



        }
        catch (Exception ex)
        {
            MessageBox.Show("---" + ex.ToString());
        }
    }

    public void Deletes()
    {
        if (tabControl1.SelectedTab.Text == "Break Master")
        {
            string del = "delete from asptblbreakmas where asptblbreakmasid=" + txtbreakid.Text;
            Utility.ExecuteNonQuery(del);
            GridLoadbreak(); empty();
        }
        if (tabControl1.SelectedTab.Text == "Shift Master")
        {
            if (txtshiftid.Text != "")
            {
                string sel1 = "select a.asptblshimasid from asptblshimas a join gtstatemast b on a.asptblshimasid=b.country where a.asptblshimasid='" + txtshiftid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblshimas");
                DataTable dt = ds.Tables["asptblshimas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + comboshift.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtshiftid.Text != "")
                    {
                        string del = "delete from asptblshimas where asptblshimasid=" + txtshiftid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptblshidet where asptblshimasid=" + txtshiftid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + comboshift.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + comboshift.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                }
            }
        }
    }

    private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
    {

    }

    private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
    {
        GridLoad();
    }

    private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    private void comboline_SelectedIndexChanged(object sender, EventArgs e)
    {

        lineload(comboline.Text);
    }

    private void butlistinsert_Click(object sender, EventArgs e)
    {
        if (allip1.Items.Count > 0 && allip2.Items.Count > 0)
        {
            string[] col;
            string[] siz;
            List<string> list1 = new List<string>();
            for (int j = 0; j < allip1.Items.Count; j++)
            {
                if (allip1.Items[j].SubItems[2].Text == "True")
                {
                    list1.Add(allip1.Items[j].SubItems[1].Text);
                }
            }
            col = list1.ToArray<string>();
            List<string> list2 = new List<string>();
            for (int j = 0; j < allip2.Items.Count; j++)
            {
                if (allip2.Items[j].SubItems[2].Text == "True")
                {
                    list2.Add(allip2.Items[j].SubItems[1].Text);
                }
            }
            siz = list2.ToArray<string>();
            int item0 = 1;
            for (int i = 0; i < col.Length; i++)
            {

                for (int j = 0; j < siz.Length; j++)
                {

                    ListViewItem list = new ListViewItem();

                    list.SubItems.Add(item0.ToString());
                    list.SubItems.Add("");
                    list.SubItems.Add(col[i].ToString());
                    list.SubItems.Add(siz[j].ToString());
                    this.listfilter4.Items.Add((ListViewItem)list.Clone());
                    
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2; listView4.Items.Add(list);
                    item0++;

                }

            }
            allip1.Items.Clear(); allip2.Items.Clear(); checksizeall.Checked = false;
            foreach (ListViewItem item in listView3.Items)
            {
                item.Checked = false;
            }
            foreach (ListViewItem item in listView2.Items)
            {
                item.Checked = false;
            }
        }

        lbltotalall.Text = "Total Count: " + listView4.Items.Count;
    }

    private void checkall_CheckedChanged(object sender, EventArgs e)
    {
        try
        {
            //foreach (ListViewItem item in listView4.Items)
            //{
            //    item.Checked = true;
            //}
            //foreach (ListViewItem item in listView3.Items)
            //{
            //    item.Checked = false;
            //}
            if (checkall.Checked == true)
            {

                var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmation == DialogResult.Yes)
                {
                    int i = 0;
                    allip1.Items.Clear(); allip2.Items.Clear();
                    listView4.Items.Clear();
                    checkall.Checked = false;
                    checksizeall.Checked = false;
                    //listView2.Items.Clear();
                    //listView3.Items.Clear();
                    //colorload();
                    //lineload(comboline.Text);
                }
            }
        }

        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }
    }
    ListView allip = new ListView();
    ListView allip1 = new ListView();
    ListView allip2 = new ListView();
    private void listView2_ItemChecked(object sender, ItemCheckedEventArgs e)
    {
        try
        {


            ListViewItem it = new ListViewItem();
            if (e.Item.Checked == true)
            {

                e.Item.SubItems[3].Text = "Connected";
                it.SubItems.Add(e.Item.SubItems[2].Text);
                it.SubItems.Add(e.Item.Checked.ToString());
                allip1.Items.Add(it);


            }
            if (e.Item.Checked == false && e.Item.SubItems[3].Text == "Connected")
            {
                e.Item.SubItems[3].Text = "DisConnected";
                e.Item.Checked = false;
                for (int c = 0; c < allip1.Items.Count; c++)
                {
                    if (allip1.Items[c].SubItems[2].Text == "True")
                    {
                        allip1.Items[c].Remove();
                        c--;
                    }
                }
                Cursor = Cursors.Default;
            }

        }
        catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
    }

    private void listView3_ItemChecked(object sender, ItemCheckedEventArgs e)
    {
        try
        {
            ListViewItem it = new ListViewItem();
            if (e.Item.Checked == true)
            {

                e.Item.SubItems[3].Text = "Connected";
                it.SubItems.Add(e.Item.SubItems[2].Text);
                it.SubItems.Add(e.Item.Checked.ToString());
                allip2.Items.Add(it);


            }
            if (e.Item.Checked == false && e.Item.SubItems[3].Text == "Connected")
            {
                e.Item.SubItems[3].Text = "DisConnected";
                for (int c = 0; c < allip2.Items.Count; c++)
                {
                    if (allip2.Items[c].SubItems[2].Text == "true")
                    {
                        allip2.Items[c].Remove();
                        c--;
                    }
                }
                Cursor = Cursors.Default;
            }

        }
        catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
    }

    private void butlistdelete_Click(object sender, EventArgs e)
    {




    }

    private void listView4_ItemChecked(object sender, ItemCheckedEventArgs e)
    {
        try
        {
            ListViewItem it = new ListViewItem();
            if (e.Item.Checked == true)
            {
                var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmation == DialogResult.Yes)
                {
                    foreach (ListViewItem eachItem in listView4.SelectedItems)
                    {
                        string ID = eachItem.SubItems[2].Text;
                        if (Convert.ToInt64("0" + ID) > 1)
                        {
                            string del = "delete from asptblshidet where asptblshimasid=" + ID;
                            Utility.ExecuteNonQuery(del);
                        }
                        listView4.Items.Remove(eachItem);
                    }
                }
            }

        }
        catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
    }

    private void checksizeall_CheckedChanged(object sender, EventArgs e)
    {
        try
        {
            allip2.Items.Clear();
            //for (int c = 0; c < listView3.Items.Count; c++)
            //{
            //    ListViewItem it = new ListViewItem();
            //    it.SubItems.Add(listView3.SelectedItems[0].SubItems[2].Text);
            //      allip2.Items.Add(it);
            //}
            if (checksizeall.Checked == true)
            {
                foreach (ListViewItem item in listView3.Items)
                {
                    item.Checked = true;
                }
            }
            else
            {
                foreach (ListViewItem item in listView3.Items)
                {
                    item.Checked = false;
                }
            }

            //if (e.Item.Checked == false && e.Item.SubItems[3].Text == "Connected")
            //{
            //    e.Item.SubItems[3].Text = "DisConnected";
            //    for (int c = 0; c < allip2.Items.Count; c++)
            //    {
            //        if (listView3.SelectedItems[0].SubItems[2].Text == allip2.Items[c].SubItems[2].Text)
            //        {
            //            allip2.Items[c].Remove();
            //            c--;
            //        }
            //    }
            //    Cursor = Cursors.Default;
            //}

        }
        catch (Exception ex) { }

    }

    private void txtpursearch_TextChanged(object sender, EventArgs e)
    {
        try
        {
            int item0 = 0; listView4.Items.Clear();
            if (txtpursearch.Text.Length >= 1)
            {

                foreach (ListViewItem item in listfilter4.Items)
                {
                    ListViewItem list = new ListViewItem();
                    if (listfilter4.Items[item0].SubItems[2].ToString().Contains(txtpursearch.Text) || listfilter4.Items[item0].SubItems[3].ToString().Contains(txtpursearch.Text))
                    {


                        list.Text = listfilter4.Items[item0].SubItems[0].Text;
                        list.SubItems.Add(listfilter4.Items[item0].SubItems[1].Text);
                        list.SubItems.Add(listfilter4.Items[item0].SubItems[2].Text);
                        list.SubItems.Add(listfilter4.Items[item0].SubItems[3].Text);
                        list.SubItems.Add(listfilter4.Items[item0].SubItems[4].Text);

                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2; listView4.Items.Add(list);
                    }
                    item0++;
                }
                lbltotalall.Text = "Total Count: " + listView4.Items.Count;
            }
            else
            {
                ListView ll = new ListView();

                listView4.Items.Clear();
                foreach (ListViewItem item in listfilter4.Items)
                {

                    this.listView4.Items.Add((ListViewItem)item.Clone());

                    item0++;
                }

                lbltotalall.Text = "Total Count: " + listView4.Items.Count;
            }



        }
        catch (Exception ex)
        {
            //MessageBox.Show("---" + ex.ToString());
        }
    }

    private void dateshiftstart_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
        {
            e.Handled = false; //Do not reject the input
        }
        else
        {
            e.Handled = true; //Reject the input
        }
    }

    private void dateshiftend_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
        {
            e.Handled = false; //Do not reject the input
        }
        else
        {
            e.Handled = true; //Reject the input
        }
    }

    private void datebreakstart_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
        {
            e.Handled = false; //Do not reject the input
        }
        else
        {
            e.Handled = true; //Reject the input
        }
    }

    private void datebreakend_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == ':' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
        {
            e.Handled = false; //Do not reject the input
        }
        else
        {
            e.Handled = true; //Reject the input
        }
    }

    private void txtsearchbreak_TextChanged(object sender, EventArgs e)
    {
        try
        {


            int item0 = 0; listViewbreak.Items.Clear();
            if (txtsearchbreak.Text.Length > 1)
            {

                foreach (ListViewItem item in listfilterbreak.Items)
                {
                    ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearchbreak.Text))
                        {


                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            this.listfilterbreak.Items.Add((ListViewItem)list.Clone());
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listViewbreak.Items.Add(list);


                        }
                    item0++;
                }
                lbltotalbreak.Text = "Total Count: " + listViewbreak.Items.Count;
            }
            else
            {

                ListView ll = new ListView();
                item0 = listfilter.Items.Count;
                listViewbreak.Items.Clear(); listViewbreak.BackColor = System.Drawing.Color.LightGray;
                    foreach (ListViewItem item in listfilterbreak.Items)
                    {
                        ListViewItem list = new ListViewItem();



                        list.Text = item.SubItems[0].Text;
                        list.SubItems.Add(item.SubItems[1].Text);
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listViewbreak.Items.Add(list);



                        item0++;
                    }
                    lbltotalbreak.Text = "Total Count: " + listViewbreak.Items.Count;
            }



        }
        catch (Exception ex)
        {
            MessageBox.Show("---" + ex.ToString());
        }
    }

    private void listViewbreak_ItemActivate(object sender, EventArgs e)
    {
        try
        {
            if (listViewbreak.Items.Count > 0)
            {

                txtbreakid.Text = listViewbreak.SelectedItems[0].SubItems[2].Text;
                string sel1 = " select a.asptblbreakmasid, a.breakname ,b.compcode, a.active    from  asptblbreakmas a  join gtcompmast b on a.compcode=b.gtcompmastid where a.active='T'   and a.asptblbreakmasid=" + txtbreakid.Text;
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbreakmas");
                DataTable dt = ds.Tables["asptblbreakmas"];

                if (dt.Rows.Count > 0)
                {
                    txtbreakid.Text = Convert.ToString(dt.Rows[0]["asptblbreakmasid"].ToString());
                    txtbreakname.Text = Convert.ToString(dt.Rows[0]["breakname"].ToString());
                        combocompcodebreak.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());

                        if (dt.Rows[0]["active"].ToString() == "T") { checkbreak.Checked = true; } else { checkbreak.Checked = false; }
                }
            }

        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }
    }

    private void combocompcode_SelectedIndexChanged_1(object sender, EventArgs e)
    {
        autonumberload();
    }

    private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
    {

        autonumberload();
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

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

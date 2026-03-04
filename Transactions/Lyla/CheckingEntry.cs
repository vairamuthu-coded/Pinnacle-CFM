using Pinnacle.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace Pinnacle.Transactions.Lyla
{
    
    public partial class CheckingEntry : Form,ToolStripAccess
    {
       
        public CheckingEntry()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
           
        }
        private string screen = "CheckingEntry";
        Models.CommonClass CC = new Models.CommonClass();
        private static CheckingEntry _instance; string coid = "", siid = "", fabid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
       ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();

        ListView allip = new ListView();
        ListView allip1 = new ListView();
        ListView allip2 = new ListView();
        ListView allip3delete = new ListView();

        public static CheckingEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new CheckingEntry();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

       

        
        private void CheckingEntry_Load(object sender, EventArgs e)
        {
           
            combocompcode.Select();
            
        }
        string maxid = "";
        Models.LYLA.CheckingEntry p = new Models.LYLA.CheckingEntry();
       
        public void colorid(string s)
        {
            try
            {

                CC.ds = Utility.ExecuteSelectQuery("select asptblcolmasid,colorname from  asptblcolmas where active='T' and colorname='" + s + "'  order by 2", "asptblcolmas");
                coid = Convert.ToString(CC.ds.Tables[0].Rows[0]["asptblcolmasid"].ToString());


            }
            catch (Exception EX)
            { }
        }
        public void sizeid(string s)
        {
            try
            {
               
                CC.ds = Utility.ExecuteSelectQuery("select asptblsizmasid from  asptblsizmas where active='T' AND sizename='" + s + "' order by 1", "asptblsizmas");

                if (CC.ds.Tables[0].Rows.Count > 0)
                {
                    siid = "";
                    siid = Convert.ToString(CC.ds.Tables[0].Rows[0]["asptblsizmasid"].ToString());
                }
            }
            catch (Exception EX)
            { }
        }
        string date1 = "", date2 = "";
        public void Saves()
        {
            try
            {
                
                if (combonotes.Text == "")
                {
                    MessageBox.Show("Pls Select Notes Field.");
                    combonotes.Focus();
                    return;
                }

                if (combonotes.Text != "" && combopono.Text != "" && comboprocesstype.Text != "" && dataGridView1.Rows.Count > 0)
                {
                   
                    Cursor.Current = Cursors.WaitCursor;
                    p.asptblpurid = Convert.ToInt64("0" + txtasptblpurid.Text);
                    p.asptblchkid = Convert.ToInt64("0" + txtcheckingid.Text);
                    p.asptblchk1id = Convert.ToInt64("0" + txtcheckid1.Text);
                    p.finyear = Class.Users.Finyear;
                    p.shortcode = Convert.ToString(txtshortcode.Text);
                    p.date = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd").Substring(0, 10));
                    p.checkno = Convert.ToString(txtcheckno.Text);
                    p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                    p.pono = Convert.ToString(combopono.Text);
                    p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                    p.lotno = Convert.ToString(txtlotno.Text.ToUpper());
                    p.bundle = Convert.ToString(txtbundle.Text);
                    p.processname = Convert.ToInt64("0" + comboprocess.SelectedValue);
                    p.processtype = Convert.ToString(comboprocesstype.Text);
                    p.issuetype = Convert.ToString(comboprocesstype.Text);
                    p.notes = Convert.ToString(combonotes.Text);
                    if (comboprocesstype.Text.Trim() == "INWARD") { p.inward = "T"; p.checking = "T"; }
                    else
                    {
                        p.inward = "F";
                    }
                    if (comboprocesstype.Text.Trim() == "DELIVERY") { p.delivery = "T"; p.checking = "T"; }
                    else
                    {
                        p.delivery = "F";
                    }
                    if (comboprocesstype.Text.Trim() == "REWORK") { p.rechecking = "T"; p.checking = "T"; p.delivery = "T"; }
                    else
                    {
                        p.rechecking = "F";
                    }

                    p.sizename = Convert.ToString(txtsize.Text.ToUpper());
                    p.compcode1 = Class.Users.COMPCODE;
                    p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
                    p.username = Class.Users.USERID;
                    p.createdby = Convert.ToString(Class.Users.HUserName);
                    p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                    p.modified = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                    p.modifiedby = Class.Users.HUserName;
                    p.ipaddress = Class.Users.IPADDRESS;
                    if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; }
                    if (checkcancel.Checked == true) { p.pocancel = "T"; } else { p.pocancel = "F"; }
                    Class.Users.Query = "select asptblchkid   from  asptblchk   where  asptblchk1id='" + p.asptblchk1id + "' and issuetype='" + p.issuetype + "' and  restitching='" + p.restitching + "' and rechecking='" + p.rechecking + "' and inward='" + p.inward + "' and delivery='" + p.delivery + "' ";
                    Class.Users.dt1 = CC.select(Class.Users.Query, "asptblchk");
                    if (Class.Users.dt1.Rows.Count != 0)
                    {
                    }
                    else if (Class.Users.dt1.Rows.Count != 0 && p.asptblchkid == 0 || p.asptblchkid == 0)
                    {
                        auto();
                        string ins = "insert into asptblchk(asptblchk1id,finyear,shortcode,wdate,checkno,compcode,buyer,pono,orderqty,stylename,lotno,bundle,size,processname,processtype,issuetype, productioncancel,active,restitching,rechecking,inward,delivery, compcode1,username,createdby,createdon,modified,modifiedby,ipaddress,notes,asptblpurid)  VALUES('" + p.asptblchk1id + "','" + p.finyear + "','" + p.shortcode + "',date_format('" + p.date.ToString().Substring(0, 10) + "','%Y-%m-%d'),'" + p.checkno + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.lotno + "','" + p.bundle + "','" + p.sizename + "','" + p.processname + "','" + p.processtype + "','" + p.issuetype + "','" + p.pocancel + "','" + p.active + "','" + p.restitching + "','" + p.rechecking + "','" + p.inward + "','" + p.delivery + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "',date_format('" + p.modified.ToString().Substring(0, 10) + "','%Y-%m-%d'),'" + p.modifiedby + "','" + p.ipaddress + "','" + p.notes + "','" + p.asptblpurid + "');";
                        Utility.ExecuteNonQuery(ins);
                    }
                    else
                    {
                        //string up = "update  asptblchk   set  modified=date_format('" + p.modified.ToString().Substring(0, 10) + "','%Y-%m-%d'),notes='" + p.notes + "' where asptblchkid='" + p.asptblchkid + "'";
                        //Utility.ExecuteNonQuery(up);
                    }
                    Models.LYLA.CheckingEntrydetail pp = new Models.LYLA.CheckingEntrydetail();
                    int i = 0, j = 1; int cnt = dataGridView1.Rows.Count;
                    if (cnt >= 0)
                    {
                        string sel2 = "select max(asptblchkid) id    from  asptblchk   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and PONO='" + combopono.Text + "' ";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblchk");
                        DataTable dt2 = ds2.Tables["asptblchk"];
                        maxid = dt2.Rows[0]["id"].ToString();
                        for (i = 0; i < cnt; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value == null)
                            {
                                dataGridView1.Rows[i].Cells[0].Value = 0;
                            }
                            else
                            {
                                pp.asptblchkdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                            }
                            pp.barcode = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[5].Value.ToString());
                            pp.asptblpurdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[6].Value.ToString());
                            pp.asptblpurid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[7].Value.ToString());
                            p.asptblchkid = Convert.ToInt64("0" + maxid);
                            p.asptblchk1id = Convert.ToInt64("0" + txtcheckid1.Text);
                            p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                            p.pono = Convert.ToString(combopono.Text);
                            colorid(dataGridView1.Rows[i].Cells[8].Value.ToString());
                            pp.colorname = coid;
                            sizeid(dataGridView1.Rows[i].Cells[9].Value.ToString());
                            pp.sizename = siid;
                            pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[10].Value.ToString());
                            pp.process = Convert.ToString(comboprocess.SelectedValue);
                            pp.processcheck = "T";
                            string sel1 = "select asptblchkdetid   from  asptblchkdet   where  barcode='" + pp.barcode + "'  and asptblchkid='" + p.asptblchkid + "' and asptblchk1id='" + p.asptblchk1id + "' and compcode='" + p.compcode + "' and  pono='" + p.pono + "'  and  colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "' and processname='" + pp.process + "' and  processcheck='" + pp.processcheck + "' and issuetype='" + p.issuetype + "' and processtype='" + p.processtype + "' and restitching='" + p.restitching + "' and rechecking='" + p.rechecking + "' and inward='" + p.inward + "' and delivery='" + p.delivery + "' ";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblchkdet");
                            DataTable dt1 = ds1.Tables["asptblchkdet"];
                            if (dt1.Rows.Count != 0 && pp.asptblchkdetid == 0 || pp.asptblchkdetid == 0 || p.asptblchkid==0)
                            {
                                string ins1 = "insert into asptblchkdet(asptblpurdet1id,barcode,asptblpurdetid,asptblpurid,asptblchkid,asptblchk1id,compcode,pono,colorname,sizename,orderqty,processname,processcheck,processtype,issuetype,restitching,rechecking,inward,delivery,notes,modified) values('" + pp.barcode + "','" + pp.barcode + "','" + pp.asptblpurdetid + "','" + p.asptblpurid + "','" + p.asptblchkid + "','" + p.asptblchk1id + "' , '" + p.compcode + "' ,'" + p.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.orderqty + "','" + pp.process + "','" + pp.processcheck + "','" + p.processtype + "','" + p.issuetype + "','" + p.restitching + "','" + p.rechecking + "','" + p.inward + "','" + p.delivery + "','" + p.notes + "',date_format('" + p.modified + "','%Y-%m-%d'))";
                                Utility.ExecuteNonQuery(ins1);
                                if (comboprocesstype.Text.Trim() == "INWARD")
                                {
                                    string up1 = "update asptblpurdet1 set inward='" + p.inward + "' ,checking='" + p.checking + "',remarks='" + p.issuetype + "' ,issuetype='" + p.issuetype + "', panelmistake='CHECKING INWARD'  where barcode='" + pp.barcode + "' AND PONO='" + p.pono + "' AND COMPCODE='" + p.compcode + "' AND FINYEAR='" + p.finyear + "'";
                                    Utility.ExecuteNonQuery(up1);
                                }
                                if (comboprocesstype.Text.Trim() == "DELIVERY")
                                {
                                    string up1 = "update asptblpurdet1 set  delivery='" + p.delivery + "' ,remarks='" + p.issuetype + "' ,issuetype='" + p.issuetype + "' ,panelmistake='CHECKING DELIVERY'  where barcode='" + pp.barcode + "' AND PONO='" + p.pono + "' AND COMPCODE='" + p.compcode + "' AND FINYEAR='" + p.finyear + "'";
                                    Utility.ExecuteNonQuery(up1);
                                }
                                if (comboprocesstype.Text.Trim() == "REWORK")
                                {
                                    string up1 = "update asptblpurdet1 set  rechecking='" + p.rechecking + "' ,remarks='" + p.issuetype + "' ,issuetype='" + p.issuetype + "' , delivery='" + p.delivery + "',panelmistake='CHECKING DELIVERY' where barcode='" + pp.barcode + "' AND PONO='" + p.pono + "' AND COMPCODE='" + p.compcode + "' AND FINYEAR='" + p.finyear + "'";
                                    Utility.ExecuteNonQuery(up1);
                                }
                                lblcount.Refresh();
                                lblcount.Text = "Inserted " + cnt + " of " + j.ToString();
                                j++;
                            }
                            else
                            {
                                //string up1 = "update  asptblchkdet  set notes='" + combonotes.Text + "' , modified=date_format('" + p.modified + "','%Y-%m-%d') where asptblchkdetid='" + pp.asptblchkdetid + "'";
                                //Utility.ExecuteNonQuery(up1);
                                lblcount.Refresh();
                                lblcount.Text = "Updated " + cnt + " of " + i.ToString();
                            }
                        }
                    }

                    if (txtcheckingid.Text == "")
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Record Saved Successfully " + txtcheckingid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        tabControl1.SelectTab(tabPage2);

                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtcheckingid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        tabControl1.SelectTab(tabPage2);
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found","GridView", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("colorname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            Cursor.Current = Cursors.Default;
        }


     
       
        private void CheckingEntry_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }


        public void News()
        {
            GridLoad(); compload();
          
          
            empty();
        }
        private void empty()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows.Clear();
                dataGridView1.AllowUserToAddRows = false;
            }
            Class.Users.UserTime = 0;txtstyleref.Text = "";txtorderno.Text = ""; txtasptblpurid.Text = "";
            txtcheckingid.Text = ""; txtcheckno.Text = "";lblbarcode.Text = ""; lblcount.Text = "";txtlotno.Text = "";
            txtcheckid1.Text = "";txtbundle.Text = "";txtsize.Text = ""; txtbarcode.Text = "";
            GlobalVariables.New_Flg = false; GlobalVariables.Saves.Visible = true; pictureBox1.Image = null;
            combostyle.Text = "";combopono.Text = "";combocompcode.Text = ""; combocompname.Text = "";
            combobuyer.Text = ""; combobuyer.Text = "";combostyle.Text = "";
            txtcheckid1.Text = ""; comboprocess.Text = "";comboprocesstype.Text = "";
            checkactive.Checked = true; Class.Users.UserTime = 0;
            dataGridView1.Rows.Clear(); dataGridView1.Columns.Clear(); dateTimePicker1.Value = Convert.ToDateTime(System.DateTime.Now.ToString());
            lblcount.ForeColor = Color.White;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            dataGridView1.Rows.Clear();
            panel3.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
           
            Class.Users.TableName = "asptblchk";
            Class.Users.TableNameGrid = "asptblchkdet";
            GlobalVariables.HideCols = new string[] { "asptblchkdetid", "asptblchkid", "asptblchk1id", "compcode", "pono", "asptblpurdetid", "asptblpurid", "ProcessName", "processcheck" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 0, 100, 0, 0, 300, 100, 50, 0, 0 };
            Class.Users.Query = "select a.asptblchkdetid, a.asptblchkid,a.asptblchk1id,a.compcode, a.pono,a.asptblpurdet1id as QrCode,a.asptblpurdetid,a.asptblpurid,a.colorname,a.sizename,a.orderqty as qty,ProcessName,processcheck from asptblchkdet a where a.asptblchkdetid<0";
            CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblchkdet");


            //GlobalVariables.HideCols = new string[] { "asptblchkdetid", "asptblchkid", "asptblchk1id", "compcode", "pono", "asptblpurdetid", "asptblpurid", "ProcessName", "processcheck" };
            //GlobalVariables.WidthCols = new Int32[] { 0,0,0,0,0,100,300,100,100,80,0,0,50 };
            //Class.Users.Query = "select a.asptblchkdetid, a.asptblchkid,a.asptblchk1id,a.compcode, a.pono,a.asptblchkdet1id as QrCode,a.colorname,a.sizename,a.orderqty as qty,a.asptblpurdetid,a.asptblpurid,a.ProcessName,a.processcheck  from asptblchkdet a where a.asptblchkdetid<0;";
            //CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols);

        }


        private void fabricid(string text)
        {
            try
            {
                string sel = "select asptblfabmasid from  asptblfabmas where active='T' AND fabric='" + text + "' ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsizmas");
                DataTable dt = ds.Tables["asptblsizmas"];
                fabid = "";
                fabid = Convert.ToString(dt.Rows[0]["asptblfabmasid"].ToString());

            }
            catch (Exception EX)
            { }
        }


        public void compload()
        {
            try
            {
                
                DataTable dt = mas.comcode();
                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                 combocompcode.DataSource = dt;



            }
            catch (Exception EX)
            { }
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "SELECT a.asptblchkid,a.wdate, b.compcode,a.checkno, c.buyername,a.pono,d.stylename,a.lotno,a.orderqty,a.processtype, e.processname FROM  asptblchk a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename JOIN asptblpromas E ON E.asptblpromasid=A.processname    order by  a.asptblchkid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblchk");
                DataTable dt = ds.Tables["asptblchk"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblchkid"].ToString());
                        list.SubItems.Add(myRow["wdate"].ToString().Substring(0, 10));
                        list.SubItems.Add(myRow["checkno"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["buyername"].ToString());
                        list.SubItems.Add(myRow["lotno"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["processtype"].ToString());
                        list.SubItems.Add(myRow["processname"].ToString());
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
                    Cursor.Current = Cursors.WaitCursor;
                    dataGridView1.Rows.Clear(); 
                    GlobalVariables.New_Flg = true;
                    txtcheckingid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "SELECT a.asptblchkid,a.asptblchk1id,b.compcode,b.compname,a.checkno,a.wdate, c.buyercode,a.pono,d.stylename,a.lotno,a.bundle,a.size,a.active,a.processtype,e.processname,e.asptblpromasid,a.notes FROM  asptblchk a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename join asptblpromas e on e.asptblpromasid=a.processname where a.asptblchkid=" + txtcheckingid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblchk");
                    DataTable dt = ds.Tables["asptblchk"];
                    if (dt.Rows.Count > 0)
                    {
                        txtcheckingid.Text = Convert.ToString(dt.Rows[0]["asptblchkid"].ToString());
                        txtcheckid1.Text = Convert.ToString(dt.Rows[0]["asptblchk1id"].ToString());
                        txtcheckno.Text = Convert.ToString(dt.Rows[0]["checkno"].ToString());
                        dateTimePicker1.Text = dt.Rows[0]["wdate"].ToString();
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                        combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                        txtbundle.Text = Convert.ToString(dt.Rows[0]["bundle"].ToString());
                        txtsize.Text = Convert.ToString(dt.Rows[0]["size"].ToString());                       
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        comboprocess.Text = dt.Rows[0]["processname"].ToString();                
                        comboprocesstype.Text = dt.Rows[0]["processtype"].ToString();
                        combonotes.Text = dt.Rows[0]["notes"].ToString();
                        string po = CC.HideButton(combopono.Text, "asptblordclomas");
                        if (po != "")
                        {
                            if (Class.Users.HUserName == "VAIRAM" || Class.Users.HUserName == "ADMIN")
                            {
                                GlobalVariables.Saves.Visible = true;
                            }
                            else
                            {
                                GlobalVariables.Saves.Visible = false;
                              
                            }
                        }
                        else
                        {
                            GlobalVariables.Saves.Visible = true;
                        }
                        pono(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblchkid"].ToString());
                        ponoGrid(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblchkid"].ToString());
                        combonotes.Text = Convert.ToString(dt.Rows[0]["notes"].ToString());
                        Cursor.Current = Cursors.Default;
                    }
                    
                }

            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show(ex.ToString());
            }
            
            tabControl1.SelectTab(tabPage1); combocompcode.Select();
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {


                int item0 = 0;
                if (txtsearch.Text.Length >=5 && checkBox1.Checked==true)
                {
                    try
                    {
                        listView1.Items.Clear();
                        string sel1 = "SELECT a.asptblchkid,a.wdate, b.compcode,a.checkno, c.buyername,a.pono,d.stylename,a.lotno,a.orderqty,a.processtype, e.processname FROM  asptblchk a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename JOIN asptblpromas E ON E.asptblpromasid=A.processname join asptblchkdet f on f.asptblchkid=a.asptblchkid where f.BARCODE like'%"+txtsearch.Text+"%'     order by  a.asptblchkid desc";
                        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblchk");
                        DataTable dt = ds.Tables["asptblchk"];
                        if (dt != null)
                        {
                            int i = 1;
                            foreach (DataRow myRow in dt.Rows)
                            {
                                ListViewItem list = new ListViewItem();
                                list.SubItems.Add(i.ToString());
                                list.SubItems.Add(myRow["asptblchkid"].ToString());
                                list.SubItems.Add(myRow["wdate"].ToString().Substring(0, 10));
                                list.SubItems.Add(myRow["checkno"].ToString());
                                list.SubItems.Add(myRow["compcode"].ToString());
                                list.SubItems.Add(myRow["buyername"].ToString());
                                list.SubItems.Add(myRow["lotno"].ToString());
                                list.SubItems.Add(myRow["pono"].ToString());
                                list.SubItems.Add(myRow["orderqty"].ToString());
                                list.SubItems.Add(myRow["stylename"].ToString());
                                list.SubItems.Add(myRow["processtype"].ToString());
                                list.SubItems.Add(myRow["processname"].ToString());
                                list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                                listView1.Items.Add(list);
                                i++;
                            }
                            lbltotal.Text = "Total Count    :" + listView1.Items.Count;
                        }
                        else
                        {
                            MessageBox.Show("Invalid");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                if (txtsearch.Text.Length >=3 && checkBox1.Checked==false)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text) || item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[8].ToString().Contains(txtsearch.Text) || item.SubItems[10].ToString().Contains(txtsearch.Text) || item.SubItems[11].ToString().Contains(txtsearch.Text) || item.SubItems[12].ToString().Contains(txtsearch.Text))
                        {

                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
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
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                if(txtsearch.Text=="")
                {
                    checkBox1.Checked = false;
                    ListView ll = new ListView();

                    listView1.Items.Clear(); 
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = item.SubItems[0].Text;
                        list.SubItems.Add(item.SubItems[1].Text);
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
            //        string sel1 = "  SELECT  a.asptblchkid,a.colorname,a.active from asptblchk a  where a.colorname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblchk");
            //        DataTable dt = ds.Tables["asptblchk"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptblchkid"].ToString());
            //                list.SubItems.Add(myRow["colorname"].ToString());
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
            if (txtcheckingid.Text != "")
            {
                string sel1 = "select a.asptblchkid from asptblchk a join gtstatemast b on a.asptblchkid=b.country where a.asptblchkid='" + txtcheckingid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblchk");
                DataTable dt = ds.Tables["asptblchk"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcheckid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtcheckingid.Text != "")
                    {
                        string del = "delete from asptblchk where asptblchkid=" + txtcheckingid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptblchkdet where asptblchkdetid=" + txtcheckingid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + txtcheckid1.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtcheckid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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

        private void txtpurid1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void combotableno_SelectedIndexChanged(object sender, EventArgs e)
        {

           // sizeload(txttableno.Text);
        }


     
     
        private void butlistdelete_Click(object sender, EventArgs e)
        {




        }

     
        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void refeshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            empty(); GridLoad();   compload(); 
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {

                if (txtcheckingid.Text == "")
                {
                    News();
                   
                }
                combocompcode.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            {
                txtsearch.Select();

            }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void combobuyer_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboorderno_TextChanged(object sender, EventArgs e)
        {

        }




        public void color(string ord)
        {
            //try
            //{
            //    string sel1 = "SELECT distinct e.asptblcolmasid,e.colorname  FROM  asptblordentry a  join gtcompmast b on a.compcode=b.gtcompmastid join asptblchk c on c.orderno=a.orderno and c.compcode=b.gtcompmastid join asptblchkdet d on d.asptblchkid=c.asptblchkid join asptblcolmas e on e.asptblcolmasid=d.colorname where a.orderno='" + ord + "'  order by a.asptblordentryid desc";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordentry");
            //    DataTable dt = ds.Tables["asptblordentry"];
            //    combocolor.DataSource = null;

            //    combocolor.DisplayMember = "colorname";
            //    combocolor.ValueMember = "asptblcolmasid";
            //    combocolor.DataSource = dt;
            //}
            //catch (Exception EX)
            //{ }
        }
        public void ponoGrid(string s, string ss, string sss)
        {
     
            if (sss != "")
            {
                DataTable dt1 = new DataTable();
                dt1 = Utility.SQLQuery("select c.asptblchkdetid,a.asptblchkid,a.asptblchk1id,b.compcode,a.pono,c.asptblpurdet1id,c.asptblpurdetid,c.asptblpurid, g.colorname,h.SIZENAME,c.orderqty,c.processname from asptblchk a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblchkdet c on c.asptblchkid=a.asptblchkid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblcolmas g on g.asptblcolmasid=c.colorname join asptblsizmas h on h.ASPTBLSIZMASID=c.sizename  where  b.compcode='" + s + "'  and a.pono='" + ss + "' and a.asptblchkid='" + sss + "'");
                if (dt1.Rows.Count > 0)
                {
                    dataGridView1.Rows.Clear();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[i].Cells[0].Value = dt1.Rows[i]["asptblchkdetid"].ToString();
                        dataGridView1.Rows[i].Cells[1].Value = dt1.Rows[i]["asptblchkid"].ToString();
                        dataGridView1.Rows[i].Cells[2].Value = dt1.Rows[i]["asptblchk1id"].ToString();
                        dataGridView1.Rows[i].Cells[3].Value = dt1.Rows[i]["compcode"].ToString();
                        dataGridView1.Rows[i].Cells[4].Value = dt1.Rows[i]["pono"].ToString();
                        dataGridView1.Rows[i].Cells[5].Value = dt1.Rows[i]["asptblpurdet1id"].ToString();
                        dataGridView1.Rows[i].Cells[6].Value = dt1.Rows[i]["asptblpurdetid"].ToString();
                        dataGridView1.Rows[i].Cells[7].Value = dt1.Rows[i]["asptblpurid"].ToString();
                        dataGridView1.Rows[i].Cells[8].Value = dt1.Rows[i]["colorname"].ToString();
                        dataGridView1.Rows[i].Cells[9].Value = dt1.Rows[i]["sizename"].ToString();
                        dataGridView1.Rows[i].Cells[10].Value = dt1.Rows[i]["orderqty"].ToString();
                        dataGridView1.Rows[i].Cells[11].Value = dt1.Rows[i]["processname"].ToString();
                        //dataGridView1.Rows[i].Cells[8].Value = dt1.Rows[i]["comqty"].ToString();
                        //dataGridView1.Rows[i].Cells[9].Value = dt1.Rows[i]["lotqty"].ToString();
                        //dataGridView1.Rows[i].Cells[10].Value = dt1.Rows[i]["balqty"].ToString();
                        //dataGridView1.Rows[i].Cells[11].Value = dt1.Rows[i]["notes"].ToString();
                    }
                    CommonFunctions.SetRowNumber(dataGridView1);
                }
            }
        }
        public void pono(string s, string ss,string sss)
        {
            if (s != "" && ss != "" )
            {
                DataTable dt = new DataTable();
                dt = Utility.SQLQuery("select distinct b.compcode,b.compname,b.gtcompmastid, f.asptblbuymasid,f.buyercode,g.asptblstymasid, g.stylename ,h.orderqty, h.lotno,h.bundle,h.lotno,h.size,a.garmentimage,a.orderno,a.styleref,a.asptblpurid from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblpromas e on e.asptblpromasid=a.processname  join asptblbuymas f on f.asptblbuymasid=a.buyer   join asptblstymas  g on g.asptblstymasid=a.stylename join asptblprolot h on h.pono=a.pono and h.compcode=f.compcode   where  b.compcode='" + s + "' and a.pono='" + ss + "'");
                if (dt.Rows.Count > 0)
                {
                    combocompcode.DisplayMember = "compcode";
                    combocompcode.ValueMember = "gtcompmastid";
                    combocompcode.DataSource = dt;
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt;
                  
                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.ValueMember = "asptblbuymasid";
                    combobuyer.DataSource = dt;
                    combostyle.DisplayMember = "stylename";
                    combostyle.ValueMember = "asptblstymasid";  
                    combostyle.DataSource = dt;
                    txtorderno.Text = Convert.ToString(dt.Rows[0]["orderno"].ToString());
                    txtstyleref.Text = Convert.ToString(dt.Rows[0]["styleref"].ToString());
                    // dt = Utility.SQLQuery("select a.orderqty,a.lotno,a.bundle from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid    where  b.compcode='" + s + "' and a.pono='" + ss + "'");
                    txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                    txtlotno.Text = dt.Rows[0]["lotno"].ToString();
                    txtbundle.Text = dt.Rows[0]["bundle"].ToString();
                    txtsize.Text = dt.Rows[0]["size"].ToString();
                    txtasptblpurid.Text= dt.Rows[0]["asptblpurid"].ToString();
                    int c = dt.Rows.Count;
                    if (c >= 1)
                    {
                        byte[] stdbytes; Int64 std;
                        if (c >= 1 && dt.Rows[0]["garmentimage"].ToString() != "")
                        {
                            pictureBox1.Image = null; stdbytes = null;
                            stdbytes = (byte[])dt.Rows[0]["garmentimage"];
                            Image img = Models.Device.ByteArrayToImage(stdbytes);
                            pictureBox1.Image = img;
                        }
                        else
                        {
                            pictureBox1.Image = null; stdbytes = null;
                        }
                    }
                }

            }
            
        }
        void auto()
        {

            DataTable dt1 = CC.autonumberload(Class.Users.Finyear, Class.Users.HCompcode, screen, "asptblchk");
            if (dt1.Rows.Count > 0)
            {
                Class.Users.UserTime = 0;
                DataTable dt11 = CC.shortcode(Class.Users.Finyear, Class.Users.HCompcode, screen, "asptblchk");
                if (dt11.Rows.Count <= 0) { return; }
                else
                {
                    combocompname.Text = dt11.Rows[0]["COMPNAME"].ToString();
                     txtshortcode.Text = dt11.Rows[0]["shortcode"].ToString();
                    txtcheckid1.Text = dt1.Rows[0]["id"].ToString();
                    txtcheckno.Text = Class.Users.Finyear + "-" + txtshortcode.Text + "-" + txtcheckid1.Text;
                    p.asptblchk1id = Convert.ToInt64("0" + txtcheckid1.Text);
                    p.checkno = Convert.ToString(txtcheckno.Text);
                    Class.Users.UserTime = 0;
                }
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

            allip1.Items.Clear(); allip2.Items.Clear();
            if (combocompcode.SelectedIndex != -1)
            {
                if (Class.Users.HCompcode != "" && txtcheckingid.Text == "")
                {
                    auto();
                    
                }

            }
        }
        int totalcount = 0;

        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopono.SelectedIndex != -1)
            {
                auto(); txtbarcode.Text = "";
                pono(Class.Users.HCompcode, combopono.Text, txtcheckingid.Text);
                    CC.ds = Utility.ExecuteSelectQuery("select MIN(c.barcode) AS  MINID , MAX(c.barcode) MAXID,count(c.barcode) cnt from asptblpurdet1 a join asptblprolot b on a.pono=b.pono and b.compcode=a.compcode join asptblprolotdet c on c.asptblpurid=a.asptblpurid and c.barcode=a.barcode where a.pono='" + combopono.Text + "' ", "asptblprolot");//and a.processtype='" + comboprocesstype.Text+"'
                lblbarcode.Refresh();
                lblbarcode.Text = "'" + CC.ds.Tables[0].Rows[0]["MINID"].ToString() + "-" + CC.ds.Tables[0].Rows[0]["MAXID"].ToString() + "'";
                totalcount = Convert.ToInt32(CC.ds.Tables[0].Rows[0]["cnt"].ToString());
            }
        }

       

        private void combocolor_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboorderno.SelectedValue != null && combocolor.SelectedValue != null)
            //{
            //    colorsizeqty(comboorderno.Text, combocolor.Text);
            //    fabricsizeqty(comboorderno.Text, combocolor.Text); listView4.Items.Clear(); allip1.Items.Clear(); allip2.Items.Clear();
            //}
        }
   
        public void Prints()
        {
            
        }

        public void Searchs()
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

        private void butSave_Click(object sender, EventArgs e)
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void listView4_ItemChecked_1(object sender, ItemCheckedEventArgs e)
        {
            try
            {


                ListViewItem it = new ListViewItem();
                if (e.Item.Checked == true)
                {

                    e.Item.SubItems[8].Text = "Connected";
                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.SubItems[3].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
              
                    allip3delete.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[8].Text == "Connected")
                {
                    e.Item.SubItems[8].Text = "DisConnected";
                    e.Item.Checked = false;
                    for (int c = 0; c < allip3delete.Items.Count; c++)
                    {
                        if (e.Item.SubItems[2].Text == allip3delete.Items[c].SubItems[1].Text)
                        {
                           
                            allip3delete.Items[c].Remove();
                            c--;
                            

                        }
                    }
                    Cursor = Cursors.Default;
                }

            }
            catch (Exception ex) { MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); }
        }

     
        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();
            Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T' AND NOT processname='STITCHING'  AND NOT processname='CUTTING' order by 2  ", "asptblpromasid", "processname");

        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

        }




        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
           
        }



        private void tabControl5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void txtorderqty_TextChanged(object sender, EventArgs e)
        {

        }
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {

        }
        private void Tb_KeyPress(object sender, KeyPressEventArgs e)
        {           
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) || (sender as TextBox).Text.Length >= 8)
            {
                e.Handled = true;
            }
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
         
        }

        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 8)
                {
                    int tot = 0;
                    int col6, col7, col8;          
                    if (Convert.ToInt32("0" + dataGridView1.Rows[e.RowIndex].Cells[8].EditedFormattedValue) > 0 && Convert.ToInt32("0" + dataGridView1.Rows[e.RowIndex].Cells[6].EditedFormattedValue) > 0)
                    {
                        col8 = 0; col7 = 0;
                        col6 = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[6].EditedFormattedValue);
                        col8 = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[8].EditedFormattedValue);
                        col7 = col6 - col8;
                        dataGridView1.Rows[e.RowIndex].Cells[7].Value = col8.ToString();
                        dataGridView1.Rows[e.RowIndex].Cells[9].Value = col7.ToString();
                        
                    }
                }
            }
            catch (Exception ex) { }
        }

        private void dataGridView1_TabIndexChanged(object sender, EventArgs e)
        {

        }
        int rowcount = 0;

        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
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

        DataTable dt1 = new DataTable();
        private void BindGrid(DataGridView grid, DataTable dt1, string barcode,int col)
        {
            bool chek = mas.checkduplicate1(5, grid, barcode);

            if (chek == true)
            {
                grid.Rows.Add();
                rowcount = dataGridView1.Rows.Count - 1;
                grid.Rows[rowcount].Cells[0].Value = "0";
                grid.Rows[rowcount].Cells[1].Value = txtcheckingid.Text;
                grid.Rows[rowcount].Cells[2].Value = txtcheckid1.Text;
                grid.Rows[rowcount].Cells[3].Value = Class.Users.HCompcode;
                grid.Rows[rowcount].Cells[4].Value = combopono.Text;
                grid.Rows[rowcount].Cells[5].Value = dt1.Rows[0]["barcode"].ToString();
                grid.Rows[rowcount].Cells[6].Value = dt1.Rows[0]["asptblpurdetid"].ToString();
                grid.Rows[rowcount].Cells[7].Value = dt1.Rows[0]["asptblpurid"].ToString();
                grid.Rows[rowcount].Cells[8].Value = dt1.Rows[0]["colorname"].ToString();
                grid.Rows[rowcount].Cells[9].Value = dt1.Rows[0]["sizename"].ToString();
                grid.Rows[rowcount].Cells[10].Value = 1;
                grid.Rows[rowcount].Cells[11].Value = comboprocess.Text;
                grid.Rows[rowcount].DefaultCellStyle.BackColor = rowcount % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
               

            }
            else
            {
                Cursor = Cursors.Default;
                 lblcount.Refresh(); lblcount.Text = "Child Record Found " + txtbarcode.Text; 
            }
            grid.Sort(grid.Columns[col], ListSortDirection.Descending);
            grid.AllowUserToAddRows = false;
        }
      private DataTable  barcode(string type,string pono, string bar)
        {
            Class.Users.dt = null;
            if (type == "INWARD")
            {
                Class.Users.Query = "select  DISTINCT f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + Class.Users.HCompcode + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "'  AND F.ISSUETYPE='REWORK' AND F.REMARKS='STITCHING MISTAKE' AND F.RESTITCHING='T' AND F.CHECKING='F' UNION ALL select  DISTINCT f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + Class.Users.HCompcode + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "'  AND F.ISSUETYPE='DELIVERY' AND F.PANELMISTAKE='DELIVERY' AND F.STITCHING='T' AND F.CHECKING='F' ORDER BY 1 ";
                Class.Users.dt = CC.select(Class.Users.Query, "asptblpurdet1");
            }
            if (type == "DELIVERY")
            {
                Class.Users.Query = "select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid    where  b.compcode='" + Class.Users.HCompcode + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "'  AND F.CHECKING='T' AND F.ISSUETYPE='INWARD' ORDER BY 1";
                Class.Users.dt = CC.select(Class.Users.Query, "asptblpurdet1");
            }
            
            if (type == "REWORK")
            {
                Class.Users.Query = "select distinct  f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid JOIN asptblcutpanret g on g.pono=f.pono  where  b.compcode='" + Class.Users.HCompcode + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND F.RECHECKING='T' AND f.ISSUETYPE='CHECKING MISTAKE' AND f.remarks='CHECKING MISTAKE'  UNION ALL select distinct  f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid JOIN asptblcutpanret g on g.pono=f.pono  where  b.compcode='" + Class.Users.HCompcode + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND F.RECHECKING='F' AND F.DELIVERY='T' AND F.restitching='T' AND f.ISSUETYPE='DELIVERY' AND f.remarks='REWORK' UNION ALL select distinct  f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid JOIN asptblcutpanret g on g.pono=f.pono  where  b.compcode='" + Class.Users.HCompcode + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND f.ISSUETYPE='REWORK' AND f.remarks='STITCHING MISTAKE' AND F.RESTITCHING='T' AND F.RECHECKING='F' AND F.CHECKING='T' ORDER BY 1";
                Class.Users.dt = CC.select(Class.Users.Query, "asptblpurdet1");
            }
            
            return Class.Users.dt;
        }
        private void butadd_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtbarcode.Text != "")
                {
                    Class.Users.UserTime = 0;
                    string source = ""; int cnt2 = 1; Class.Users.UserTime = 0;
                    source = txtbarcode.Text.Trim();
                    string data = getBetween(source, "'", "'");
                    string[] data1 = data.Split('-'); rowcount = 0;
                    if (data1.Length == 2)
                    {
                        Cursor.Current = Cursors.WaitCursor; dt1 = null;
                        int a1 = Convert.ToInt32(data1[0]); dataGridView1.Rows.Clear();
                        int a2 = Convert.ToInt32(data1[1]);
                        int a4 = 1;
                        a4 += a2 - a1;
                        for (int aa = a1; aa <= a2; aa++)
                        {

                             DataTable dt = Utility.SQLQuery("select count(asptblchkdetID) CNT from asptblchkdet  where  barcode='" + aa.ToString() + "'  and compcode='" + combocompcode.SelectedValue + "' and  pono='" + combopono.Text + "'  and  processcheck='T' and  issuetype='" + comboprocesstype.Text + "' and INWARD='F'  and delivery='F' ");
                            string s = dt.Rows[0]["CNT"].ToString();
                            if (Convert.ToInt32(s) >= 1) { lblcount.Refresh(); lblcount.Text = "Child Record Found  .  " + a2.ToString(); Cursor.Current = Cursors.Default;  }
                            else
                            {
                               dt1= barcode(comboprocesstype.Text, combopono.Text, aa.ToString());                                
                                if (dt1.Rows.Count > 0)
                                {

                                    BindGrid(dataGridView1, dt1, aa.ToString(),5);
                                    lblcount.Refresh();
                                    lblcount.Text = a4 + " of " + cnt2.ToString();
                                    rowcount++; cnt2++;
                                   
                                }
                                else
                                {
                                    lblcount.Refresh(); /*lblcount.Text = " Invalid " + a4;*/
                                   
                                    lblcount.Text = a4 + " of " + cnt2.ToString();
                                    rowcount++; cnt2++;
                                    if (a4 == cnt2)
                                    {
                                        MessageBox.Show("Invalid Barcode");
                                        lblcount.Refresh(); lblcount.Text = " Invalid BarCode ";
                                        txtbarcode.Text = "";
                                    }
                                }
                            }
                        }
                        lblcount.Text = "Total Rows " + dataGridView1.Rows.Count.ToString();

                        //dataGridView1.Sort(dataGridView1.Columns[5], ListSortDirection.Descending);
                        Cursor = Cursors.Default;
                    }
                    else
                    {
                        if (txtbarcode.Text.Length >= 5)
                        {
                            DataTable dt = Utility.SQLQuery("select count(asptblchkdetID) CNT from asptblchkdet  where  barcode='" + txtbarcode.Text + "'  and compcode='" + combocompcode.SelectedValue + "' and  pono='" + combopono.Text + "'  and  processcheck='T' and  issuetype='" + comboprocesstype.Text + "' and INWARD='F'  and delivery='F' ");

                            string s = dt.Rows[0]["CNT"].ToString();
                            if (Convert.ToInt32(s) >= 1)
                            {
                                MessageBox.Show("Child Record Found");
                                MessageBox.Show("Child Record Found ", "Alert", MessageBoxButtons.OK); lblcount.Refresh(); lblcount.Text = "Child Record Found  .  " + data1[0].ToString(); Cursor.Current = Cursors.Default;  }
                            else
                            {
                               dt1= barcode(comboprocesstype.Text, combopono.Text, txtbarcode.Text);                               
                                if (dt1.Rows.Count > 0)
                                {
                                    BindGrid(dataGridView1, dt1, txtbarcode.Text,5);
                                    lblcount.Refresh(); lblcount.Text = "Count :" + dataGridView1.Rows.Count.ToString() + " Of " + totalcount.ToString();
                                    txtbarcode.Text = ""; txtbarcode.Select();

                                }
                                else
                                {
                                    lblcount.Refresh(); lblcount.Text = "Count :" + dataGridView1.Rows.Count.ToString() + " Of " + totalcount.ToString();
                                    MessageBox.Show("Invalid data " + txtbarcode.Text, "Alert", MessageBoxButtons.OK);
                                    txtbarcode.Text = ""; txtbarcode.Select();
                                }
                            }
                        }
                    }


                }
                CommonFunctions.SetRowNumber(dataGridView1);
            }catch { Exception ex; }
        }
    private void BindGrid(DataGridView grid,DataTable dt1)
        {
            grid.Rows.Add();
            rowcount = dataGridView1.Rows.Count - 1;
            grid.Rows[rowcount].Cells[0].Value = "0";
            grid.Rows[rowcount].Cells[1].Value = txtcheckingid.Text;
            grid.Rows[rowcount].Cells[2].Value = txtcheckid1.Text;
            grid.Rows[rowcount].Cells[3].Value = Class.Users.HCompcode;
            grid.Rows[rowcount].Cells[4].Value = combopono.Text;
            grid.Rows[rowcount].Cells[5].Value = dt1.Rows[0]["barcode"].ToString();
            grid.Rows[rowcount].Cells[6].Value = dt1.Rows[0]["asptblpurdetid"].ToString();
            grid.Rows[rowcount].Cells[7].Value = dt1.Rows[0]["asptblpurid"].ToString();
            grid.Rows[rowcount].Cells[8].Value = dt1.Rows[0]["colorname"].ToString();
            grid.Rows[rowcount].Cells[9].Value = dt1.Rows[0]["sizename"].ToString();
            grid.Rows[rowcount].Cells[10].Value = 1;
            grid.Rows[rowcount].Cells[11].Value = comboprocess.Text;
        }
        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 5)
            {
                mas.checkduplicate(e.ColumnIndex, dataGridView1);
            }
        }

        private void txtbarcode_TextChanged(object sender, EventArgs e)
        {
            if (txtcheckingid.Text != "")
            {
                lblcount.Refresh();
                lblcount.Text = "Unable to process";
            }
            else
            {
                butadd_Click(sender, e);
            }
            
        }

        private void comboprocess_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void txtbundle_TextChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void txtsize_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboprocesstype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopono.SelectedIndex != -1 && dataGridView1.Columns.Count > 0 && comboprocesstype.Text != "")
            {              
                CC.ds = Utility.ExecuteSelectQuery("select MIN(c.asptblpurdet1ID) AS  MINID , MAX(c.asptblpurdet1ID) MAXID,count(c.asptblpurdet1id) cnt from asptblpurdet1 a join asptblprolot b on a.pono=b.pono and b.compcode=a.compcode join asptblprolotdet c on c.asptblpurid=a.asptblpurid and c.asptblpurdet1id=a.asptblpurdet1id where a.pono='" + combopono.Text + "'  and B.processtype='" + comboprocesstype.Text+"'", "asptblprolotdet");//and a.processtype='" + comboprocesstype.Text+"'
                lblbarcode.Refresh();
                lblbarcode.Text = "'" + CC.ds.Tables[0].Rows[0]["MINID"].ToString() + "-" + CC.ds.Tables[0].Rows[0]["MAXID"].ToString() + "'       Total Rows  :" + CC.ds.Tables[0].Rows[0]["cnt"].ToString();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void comboprocesstype_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if(comboprocesstype.Text != "")
            {
                Class.Users.dt = null;
                if (comboprocesstype.Text == "INWARD")
                {
                    combopono.DataSource = null;
                    Class.Users.dt = CC.select("select DISTINCT X.asptblpurID, X.PONO from( select DISTINCT A.asptblpurID, A.PONO from asptblpur A  join gtcompmast b on a.compcode=b.gtcompmastid  join asptblprolot c on c.pono=a.pono and c.compcode=b.gtcompmastid  LEFT JOIN  asptblordclomas D ON A.PONO=D.PONO AND D.ACTIVE='T' WHERE D.PONO IS NULL  AND  C.DELIVERY='T' and c.processtype='DELIVERY' and b.compcode='" + Class.Users.HCompcode + "' UNION ALL select DISTINCT A.asptblpurID, A.PONO from asptblpur A join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=b.gtcompmastid  LEFT JOIN  asptblordclomas D ON A.PONO=D.PONO AND D.ACTIVE='T' WHERE D.PONO IS NULL  AND  C.DELIVERY='T' and  c.processtype='REWORK' and b.compcode='" + Class.Users.HCompcode+"') X  ORDER BY X.asptblpurID DESC", "ASPTBLPUR");

                    //Class.Users.dt = CC.select("select distinct a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=b.gtcompmastid  where  C.DELIVERY='T' and c.processtype='DELIVERY' OR c.processtype='REWORK' and b.compcode='" + Class.Users.HCompcode + "' order by a.asptblpurid desc", "ASPTBLPUR");
                    Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T'  AND NOT processname='CHECKING'  AND NOT processname='CUTTING' AND NOT processname='PACKING' order by 2  ", "asptblpromasid", "processname");
                    label7.Text = "Pre-Process";//combonotes.Enabled = false ;
                   // combonotes.Text = "";
                }
                if (comboprocesstype.Text == "DELIVERY")
                {
                    combopono.DataSource = null;
                    Class.Users.dt = CC.select("select distinct a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=b.gtcompmastid   RIGHT JOIN  asptblchk D ON D.PONO=C.PONO AND  D.compcode=A.compcode and D.issuetype='INWARD' LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where  a.active='T' AND E.PONO IS NULL and b.compcode='" + Class.Users.HCompcode + "' order by a.asptblpurid desc", "ASPTBLPUR");

                    //  Class.Users.dt = CC.select("select distinct a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=b.gtcompmastid and c.processtype='DELIVERY'  RIGHT JOIN  asptblchk D ON D.PONO=C.PONO AND  D.compcode=A.compcode and D.processtype='INWARD' where  a.active='T' and b.compcode='" + Class.Users.HCompcode + "' order by a.asptblpurid desc", "ASPTBLPUR");
                    Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T'  AND NOT processname='STITCHING'  AND NOT processname='CUTTING' AND NOT processname='CHECKING' order by 2  ", "asptblpromasid", "processname");
                    label7.Text = "Next-Process"; //combonotes.Enabled = true;
                }
                if (comboprocesstype.Text == "REWORK")
                {
                    combopono.DataSource = null;
                    Class.Users.dt = CC.select("SELECT distinct X.asptblpurid, X.PONO FROM (select distinct a.asptblpurid, a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=a.compcode and c.compcode=b.gtcompmastid  JOIN asptblcutpanretdet D ON D.pono=A.pono and D.issuetype='CHECKING MISTAKE'  AND  D.REMARKS='STITCHING MISTAKE'   AND  C.DELIVERY='T'  LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where  a.active='T' AND E.PONO IS NULL  and b.compcode='" + Class.Users.HCompcode + "'  union all select distinct a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=a.compcode and c.compcode=b.gtcompmastid   JOIN asptblcutpanret D ON D.pono=A.pono and D.issuetype='CHECKING MISTAKE'  AND  D.REMARKS='CHECKING MISTAKE' LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where  E.PONO IS NULL AND  b.compcode='" + Class.Users.HCompcode + "' ) X order by  x.asptblpurid desc ", "ASPTBLPUR");

                    //              Class.Users.dt = CC.select("SELECT distinct X.asptblpurid, X.PONO FROM (select distinct a.asptblpurid, a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=a.compcode and c.compcode=b.gtcompmastid   JOIN asptblchk D ON D.pono=A.pono and C.issuetype='DELIVERY'  AND  C.DELIVERY='T' where  a.active='T' and b.compcode='" + Class.Users.HCompcode + "'  union all select distinct a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.pono and c.compcode=a.compcode and c.compcode=b.gtcompmastid   JOIN asptblcutpanret D ON D.pono=A.pono and D.issuetype='CHECKING MISTAKE'  AND  D.REMARKS='CHECKING MISTAKE' where  b.compcode='" + Class.Users.HCompcode + "' ) X order by  x.asptblpurid desc ", "ASPTBLPUR");
                    Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T'  AND NOT processname='CHECKING' AND NOT processname='STITCHING'  AND NOT processname='CUTTING' order by 2  ", "asptblpromasid", "processname");
                    label7.Text = "Return"; //combonotes.Enabled = true;
                }

                if (Class.Users.dt != null && Class.Users.dt.Rows.Count>0)
                {
                    combopono.DisplayMember = "pono";
                    combopono.ValueMember = "pono";
                    combopono.DataSource = Class.Users.dt;
                }
                else
                {
                    combopono.DataSource = null;
                }
                dataGridView1.Rows.Clear();
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (deleteToolStripMenuItem.Text == "Delete")
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    dataGridView1.Rows.RemoveAt(row.Index);
                    dataGridView1.Refresh();
                }
            }
        }

        private void deleteAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (deleteAllToolStripMenuItem.Text == "Delete All")
            {
                dataGridView1.Rows.Clear();
            }
        }

        private void txtlotno_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblbarcode_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtbarcode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar <= '"' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtbarcode_MouseHover(object sender, EventArgs e)
        {
            txtbarcode.BackColor = Color.LightYellow;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            txtsearch.Text = "";
        }

        private void txtbarcode_MouseLeave(object sender, EventArgs e)
        {
            txtbarcode.BackColor = Color.White;
        }

        private void butDelete_Click(object sender, EventArgs e)
        {

        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }
    }
}

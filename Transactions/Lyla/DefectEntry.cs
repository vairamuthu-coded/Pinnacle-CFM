using Pinnacle.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;


namespace Pinnacle.Transactions.Lyla
{
    public partial class DefectEntry : Form,ToolStripAccess
    {
        public DefectEntry()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());

        }
        private string screen = "DefectEntry";
        Models.LYLA.CutPanelReturnModel p = new Models.LYLA.CutPanelReturnModel();
        private static DefectEntry _instance; string coid = "", siid = "", fabid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        public static DefectEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DefectEntry();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

        Models.CommonClass CC = new Models.CommonClass();

        
        private void DefectEntry_Load(object sender, EventArgs e)
        {     
            comboissuetype.Select();
          
        }
        string maxid = ""; string maxid1= "";
        
        public void colorid(string s)
        {
            try
            {               
                DataTable dt =CC.select("select asptblcolmasid,colorname from  asptblcolmas where active='T' and colorname='" + s + "'  order by 2","asptblcolmas");
                coid = "";
                coid = Convert.ToString(dt.Rows[0]["asptblcolmasid"].ToString());
            }
            catch (Exception EX)
            { }
        }
        public void sizeid(string s)
        {
            try
            {
               
                DataTable dt = CC.select("select asptblsizmasid from  asptblsizmas where active='T' AND sizename='" + s + "' order by 1", "asptblsizmas");

                if (dt.Rows.Count > 0)
                {
                    siid = "";
                    siid = Convert.ToString(dt.Rows[0]["asptblsizmasid"].ToString());
                }
            }
            catch (Exception EX)
            { }
        }
        public void Remarks(string s)
        {
            try
            {

                DataTable dt = CC.select("select asptblremmasid,remarks from asptblremmas WHERE active='T' AND remarks='" + s + "' order by 1", "asptblremmas");

                if (dt.Rows.Count > 0)
                {
                    siid = "";
                    siid = Convert.ToString(dt.Rows[0]["asptblremmasid"].ToString());
                }
            }
            catch (Exception EX)
            { }
        }
        public void Saves()
        {
            try
            {
                if (combodefecttype.Text == "")
                {
                    MessageBox.Show("pls select Defect Type..","Defect Type",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    combodefecttype.Focus();
                    return;
                }
                if (combonotes.Text == "")
                {
                    MessageBox.Show("pls select Notes Field..", "Notes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    combonotes.Focus();
                    return;
                }
                if (comboRemarks.SelectedValue == null)
                {
                    MessageBox.Show("pls select Remarks..", "Remarks", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    comboRemarks.Focus();
                    return;
                }
                if (dataGridView1.Rows.Count > 0 && comboRemarks.SelectedValue != null)
                {
                    Cursor.Current = Cursors.WaitCursor;
                 
                    p.asptblcutpanretid = Convert.ToInt64("0" + txtcutpanelid.Text);
                    p.asptblcutpanret1id = Convert.ToInt64("0" + txtcutpanel1id.Text);                 
                    p.finyear = Class.Users.Finyear;
                    p.panelno = Convert.ToString(txtpanelno.Text);
                    p.shortcode = Convert.ToString(txtshortcode.Text);
                    p.cutpaneldate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));                   
                    p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                    p.pono = Convert.ToString(combopono.Text);
                    p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                    p.lotno = Convert.ToString(txtlotno.Text.ToUpper());
                    p.bundle = Convert.ToString(txtbundle.Text.ToUpper());
                    p.processname = Convert.ToInt64("0"+comboprocess.SelectedValue);
                    p.issuetype = Convert.ToString(comboissuetype.Text);
                    p.Remarks = Convert.ToString(comboRemarks.SelectedValue);
                    p.notes = Convert.ToString(combonotes.Text);
                    p.defecttype = Convert.ToString(combodefecttype.Text);
                    if (comboissuetype.Text.Trim() == "STITCHING MISTAKE") { p.delivery = "F"; p.stitching = "T"; p.restitching = "T";p.issuetype = "STITCHING MISTAKE"; } else { p.stitching = "F"; }
                    if (comboissuetype.Text.Trim() == "CHECKING MISTAKE") { p.checking = "F"; p.rechecking = "T"; p.issuetype = "CHECKING MISTAKE"; } else { p.checking = "F";  }
                    p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
                    p.username = Class.Users.USERID;
                    p.createdby = Convert.ToString(Class.Users.HUserName);
                    p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                    p.modified = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                    p.modifiedby = Class.Users.HUserName;
                    p.ipaddress = Class.Users.IPADDRESS;
                    DataTable dt0 = CC.select("select asptblcutpanretid   from  asptblcutpanret   where  asptblcutpanret1id='" + p.asptblcutpanret1id+ "' and  compcode='" + p.compcode + "' and  pono='" + p.pono + "'  AND buyer='"+p.buyer+"' AND processname='"+p.processname+ "' AND lotno='"+p.lotno+ "' AND bundle='" + p.bundle+"' and  remarks='" + p.Remarks + "' and  issuetype='" + p.issuetype + "' and  cutting='" + p.cutting + "' and  STITCHING='" + p.stitching + "' and  checking='" + p.checking + "' and  restitching='" + p.restitching + "' and  rechecking='" + p.rechecking + "' and  DEFECTTYPE='" + p.defecttype + "' ", "asptblcutpanret");
                    if (dt0.Rows.Count != 0) { }
                    else if (dt0.Rows.Count != 0 && p.asptblcutpanretid == 0 || p.asptblcutpanretid == 0)
                    {
                      
                        auto();
                      
                        string ins = "insert into asptblcutpanret(asptblcutpanret1id,shortcode,panelno,finyear,cutpaneldate,compcode,stylename,orderqty,pono,buyer,processname,lotno,issuetype,remarks,cutting,STITCHING,checking,restitching,rechecking,compcode1,username,createdby,modifiedby,ipaddress,DEFECTTYPE,notes,modified)  VALUES('" + p.asptblcutpanret1id + "','" + p.shortcode + "','" + p.panelno + "','" + p.finyear + "','" + p.cutpaneldate + "','" + p.compcode + "','" + p.stylename + "','" + p.orderqty + "','" + p.pono + "','" + p.buyer + "','" + p.processname + "','" + p.lotno + "','" + p.issuetype + "','" + p.Remarks + "','" + p.cutting + "','" + p.stitching +"','" + p.checking + "','"+p.restitching+"','"+p.rechecking+"','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + p.modifiedby + "','" + p.ipaddress + "','"+p.defecttype+ "','" + p.notes + "',date_format('" + p.modified + "','%Y-%m-%d'))";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptblcutpanret1id) id1 , max(asptblcutpanretid) id  from  asptblcutpanret   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and PONO='" + combopono.Text + "'  ";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblcutpanret");
                        DataTable dt2 = ds2.Tables["asptblcutpanret"];
                        maxid = dt2.Rows[0]["id"].ToString();
                        maxid1 = dt2.Rows[0]["id1"].ToString();
                    }
                   else
                    {
                        //string up = "update  asptblcutpanret   set cutpaneldate=date_format('" + p.modified + "', '%Y-%m-%d'), compcode='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',modified=date_format('" + p.modified + "','%Y-%m-%d'), modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblcutpanretid='" + txtcutpanelid.Text + "'";
                        //Utility.ExecuteNonQuery(up);                    
                    }
                    Models.LYLA.CutPanelReturndetModel pp = new Models.LYLA.CutPanelReturndetModel();
                    int i = 0,j=1; int cnt = dataGridView1.Rows.Count;
                    if (cnt >= 0)
                    {

                        for (i = 0; i < cnt; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value == null)
                            {
                                dataGridView1.Rows[i].Cells[0].Value = 0;
                            }
                            else
                            {
                                pp.asptblcutpanretdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                              
                            }
                            pp.barcode = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[1].Value.ToString());
                            pp.finyear = Class.Users.Finyear; coid = "";
                            pp.asptblcutpanretid = Convert.ToInt64("0" + maxid);
                            pp.asptblcutpanret1id = Convert.ToInt64("0" + maxid1);                          
                            pp.asptblpurdetid = Convert.ToInt64("0" + dataGridView1.Rows[0].Cells[2].Value.ToString());
                            pp.asptblpurid = Convert.ToInt64("0" + dataGridView1.Rows[0].Cells[3].Value.ToString());
                            colorid(dataGridView1.Rows[i].Cells[6].Value.ToString()); 
                            pp.colorname = coid;
                            sizeid(dataGridView1.Rows[i].Cells[7].EditedFormattedValue.ToString());
                            pp.sizename = siid;
                            pp.pcs = 1; siid = "";
                            pp.issuetype = Convert.ToString(comboissuetype.Text);
                            p.Remarks = dataGridView1.Rows[i].Cells[10].EditedFormattedValue.ToString();           
                            if(p.Remarks == "")
                            {
                                p.Remarks = comboRemarks.SelectedValue.ToString();
                            }
                            DataTable dt1 = CC.select("select asptblcutpanretdetid   from  asptblcutpanretdet   where  barcode='" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "' and  compcode='" + p.compcode + "' and  pono='" + p.pono + "'  and  colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "' and issuetype='"+pp.issuetype + "' and  finyear='" + p.finyear + "' and remarks='"+p.Remarks+"' and  cutting='" + p.cutting + "' and  STITCHING='" + p.stitching + "' and  checking='" + p.checking + "' and  restitching='" + p.restitching + "' and  rechecking='" + p.rechecking + "' and modified=date_format('" + p.modified + "','%Y-%m-%d') ", "asptblcutpanretdet");
                            if (dt1.Rows.Count != 0) { }
                            else if (dt1.Rows.Count != 0 && pp.asptblcutpanretdetid == 0 || pp.asptblcutpanretdetid == 0 || p.asptblcutpanretid==0)
                            {
                                string ins1 = "insert into asptblcutpanretdet(asptblcutpanretid,asptblcutpanret1id,asptblpurdet1id,barcode,asptblpurdetid,asptblpurid,finyear,compcode,pono,colorname,sizename,pcs,issuetype,processcheck,Remarks,cutting,STITCHING,checking,restitching,rechecking,notes,modified) values('" + pp.asptblcutpanretid + "' ,'" + pp.asptblcutpanret1id + "' ,'" + pp.barcode + "' ,'"+pp.barcode+"','" + pp.asptblpurdetid + "','" + pp.asptblpurid + "','" + Class.Users.Finyear + "', '" + combocompcode.SelectedValue + "' ,'" + p.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.pcs + "','" + p.issuetype + "','T','" + p.Remarks + "','" + p.cutting + "','" + p.stitching + "','" + p.checking + "','"+p.restitching+"','"+p.rechecking+ "','" + p.notes + "',date_format('" + p.modified + "','%Y-%m-%d'));";
                                Utility.ExecuteNonQuery(ins1);                             

                                    if (comboissuetype.Text.Trim() == "STITCHING MISTAKE")
                                    {
                                        string up1 = "update  asptblpurdet1  set  remarks='" + p.Remarks + "', panelmistake='" + p.issuetype + "',ISSUETYPE='" + p.issuetype + "',restitching='F',delivery='T', processcheck='F'  where barcode='" + pp.barcode + "' and compcode='" + p.compcode + "' and finyear='" + p.finyear + "' and pono='" + p.pono + "'";
                                        Utility.ExecuteNonQuery(up1);
                                    }
                                    if (comboissuetype.Text.Trim() == "CHECKING MISTAKE")
                                    {
                                        string up1 = "update  asptblpurdet1  set remarks='" + p.Remarks + "', panelmistake='" + p.issuetype + "',ISSUETYPE='" + p.issuetype + "',rechecking='" + p.rechecking + "' ,delivery='T'  , processcheck='F' where barcode='" + pp.barcode + "' and compcode='" + p.compcode + "' and finyear='" + p.finyear + "' and pono='" + p.pono + "'";
                                        Utility.ExecuteNonQuery(up1);
                                    }


                            }
                            else
                            {
                                //string up = "update  asptblcutpanretdet   set notes='"+combonotes.Text+"' , modified=date_format('" + p.modified + "','%Y-%m-%d') where asptblcutpanretdetid='" + pp.asptblcutpanretdetid + "'";
                                //Utility.ExecuteNonQuery(up);
                            }
                            
                            lblcount.Refresh();
                        }
                    }
                    if (txtcutpanelid.Text == "")
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Record Saved Successfully " + txtcutpanelid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        tabControl1.SelectTab(tabPage2);

                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtcutpanelid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        tabControl1.SelectTab(tabPage2);
                    }

                }
                else
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("Invalid Grid Record " + dataGridView1.Rows.Count.ToString(), " No Data Found ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show(ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            Cursor.Current = Cursors.Default;
        }

        private void DefectEntry_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       


     

       
        public void News()
        {
           
            GridLoad();
            compload();
            empty();

        }
        private void empty()
        {
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows.Clear();
                dataGridView1.AllowUserToAddRows = false;
            }
            GlobalVariables.Saves.Visible = true;txtorderno.Text = "";txtstyleref.Text = "";
            txtcutpanelid.Text = ""; txtpanelno.Text = "";combodefecttype.Text = ""; txtbarcode.Text = ""; lblcount.Text = "";txtlotno.Text = "";txtshortcode.Text = "";
            txtcutpanel1id.Text = "";txtbundle.Text = "";txtsize.Text = "";txtorderqty.Text = "";lblbarcode.Text = "";
            GlobalVariables.New_Flg = false; Class.Users.UserTime = 0;
            combostyle.Text = "";combopono.Text = "";combocompcode.Text = ""; combocompname.Text = "";
            combobuyer.Text = ""; combopono.Text = "";combobuyer.Text = "";combostyle.Text = "";combopono.Text = "";
            txtcutpanel1id.Text = ""; comboprocess.Text = "";comboissuetype.Text = "";          
            dataGridView1.Rows.Clear(); dataGridView1.Columns.Clear(); dateTimePicker1.Value = Convert.ToDateTime(System.DateTime.Now.ToString());
            lblcount.ForeColor = Color.White;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            panel2.BackColor = Class.Users.BackColors;
            dataGridView1.Rows.Clear();
            panel3.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
           
            listView1.Font = Class.Users.FontName;
            Class.Users.TableName = "asptblcutpanret";
            Class.Users.TableNameGrid = "asptblcutpanretdet";
            GlobalVariables.HideCols = new string[] {"asptblcutpanretdetid", "asptblpurdetid","asptblpurid", "compcode", "pono" };
            GlobalVariables.WidthCols = new Int32[] {0,150,0, 0, 0,0,100, 80, 100, 125, 500 };
            Class.Users.Query = "select a.asptblcutpanretdetid, a.asptblpurdet1id as qrcode,a.asptblpurdetid,a.asptblpurid, a.compcode,a.pono,a.colorname,a.sizename,a.Pcs,a.issuetype,a.Remarks from asptblcutpanretdet a where a.asptblcutpanretdetid<0";
            CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblcutpanretdet");
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
      
        void processload(string pro1)
        {
            string pro = "";
            //if (pro1.Trim() == "PANEL MISTAKE")
            //{
            //    pro = "CUTTING";
            //}
            if (pro1.Trim() == "STITCHING MISTAKE")
            {

                pro = "STITCHING";
            }
            if (pro1.Trim() == "CHECKING MISTAKE")
            {

                pro = "CHECKING";
            }
            Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T' AND processname='" + pro + "'    order by 2  ", "asptblpromasid", "processname");

        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "SELECT a.asptblcutpanretid,a.cutpaneldate,a.panelno, b.compcode,c.buyername,a.lotno,a.pono,a.orderqty,d.stylename,a.ISSUETYPE,E.processname FROM  asptblcutpanret a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename JOIN asptblpromas E ON E.asptblpromasid=A.processname  order by  a.asptblcutpanretid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcutpanret");
                DataTable dt = ds.Tables["asptblcutpanret"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblcutpanretid"].ToString());
                        list.SubItems.Add(myRow["cutpaneldate"].ToString().Substring(0,10));
                        list.SubItems.Add(myRow["panelno"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["buyername"].ToString());
                        list.SubItems.Add(myRow["lotno"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());                      
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["ISSUETYPE"].ToString());                       
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
            if (listView1.Items.Count > 0)
            {
                Cursor.Current = Cursors.WaitCursor; 
                dataGridView1.Rows.Clear(); 
                GlobalVariables.New_Flg = true;
                txtcutpanelid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                string sel1 = "SELECT aa.asptblcutpanretid,aa.asptblcutpanret1id, a.asptblpurid,a.asptblpur1id,a.pono,a.shortcode, aa.cutpaneldate,a.orderqty,a.bundle,a.size, c.compcode,c.compname,c.gtcompmastid, d.buyercode as buyer,d.buyercode, d.asptblbuymasid,e.sizegroup,e.asptblsizgrpid,f.stylename, f.asptblstymasid, g.processname,g.asptblpromasid ,a.orderqty, a.lotno,aa.ISSUETYPE,aa.remarks,aa.DEFECTTYPE,aa.notes FROM   asptblpur  a  join  asptblcutpanret aa on  a.pono=aa.pono join gtcompmast c on a.compcode=c.gtcompmastid join asptblbuymas d on d.asptblbuymasid=a.buyer  join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblstymas f on f.asptblstymasid=a.stylename join asptblpromas g on g.asptblpromasid=a.processname where aa.asptblcutpanretid='" + txtcutpanelid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                DataTable dt = ds.Tables["asptblpur"];
                if (dt.Rows.Count > 0)
                {
                    txtcutpanelid.Text = Convert.ToString(dt.Rows[0]["asptblcutpanretid"].ToString());
                    txtcutpanel1id.Text = Convert.ToString(dt.Rows[0]["asptblcutpanret1id"].ToString());                   
                    txtasptblpurid.Text = Convert.ToString(dt.Rows[0]["asptblpurid"].ToString());
                    txtpanelno.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                    txtshortcode.Text = dt.Rows[0]["shortcode"].ToString();
                    txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(dt.Rows[0]["cutpaneldate"].ToString().Substring(0, 10));
                    combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                    combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                    combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                    comboissuetype.Text = dt.Rows[0]["issuetype"].ToString();
                    combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                    combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                    txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                    txtbundle.Text = Convert.ToString(dt.Rows[0]["bundle"].ToString());                 
                    comboprocess.Text = dt.Rows[0]["processname"].ToString();                  
                    comboRemarks.Text = dt.Rows[0]["remarks"].ToString();
                    combodefecttype.Text = dt.Rows[0]["DEFECTTYPE"].ToString();
                    combonotes.Text = Convert.ToString(dt.Rows[0]["notes"].ToString());
                    string po = CC.HideButton(combopono.Text, "asptblchk");
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
                    if (Convert.ToInt64(dt.Rows[0]["asptblcutpanretid"].ToString())>0)
                    {                   
                     
                        dt1 = Utility.SQLQuery("SELECT  b.asptblcutpanretdetid,b.barcode, b.asptblpurdetid,b.asptblpurid,c.compcode, a.pono,e.colorname,f.SIZENAME,a.ISSUETYPE,a.remarks FROM asptblcutpanret A JOIN ASptblcutpanretdet b on a.asptblcutpanretid=b.asptblcutpanretid  join gtcompmast c on c.gtcompmastid=a.compcode  join asptblcolmas e on e.asptblcolmasid=b.colorname join asptblsizmas f on f.asptblsizmasid=b.sizename  where  c.compcode='" + Convert.ToString(dt.Rows[0]["compcode"].ToString()) + "'  and a.pono='" + Convert.ToString(dt.Rows[0]["pono"].ToString()) + "' and B.asptblcutpanretid='" + Convert.ToInt64("0"+dt.Rows[0]["asptblcutpanretid"].ToString()) + "' ORDER BY 2");
                        if (dt1.Rows.Count > 0)
                        {
                            GlobalVariables.New_Flg = true;
                            int a2 = 0;
                            Cursor.Current = Cursors.WaitCursor;
                            BindGrid(dataGridView1, dt1, "", a2);
                            Cursor.Current = Cursors.Default;

                        }
                        CommonFunctions.SetRowNumber(dataGridView1);

                    }
                    Cursor.Current = Cursors.Default;
                }
               
            }
            
            tabControl1.SelectTab(tabPage1); combocompcode.Select();
           

        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {


                int item0 = 0;
                if (txtsearch.Text.Length > 1)
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
                else
                {

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
          
        }

        public void Deletes()
        {
            if (txtcutpanelid.Text != "")
            {
                string sel1 = "select a.asptblcutpanretdetid from asptblcutpanretdet a join asptblpurdet b on a.asptblpurdet1id=b.asptblpurdet1id where a.asptblcutpanretdetid='" + txtcutpanelid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcutpanretdet");
                DataTable dt = ds.Tables["asptblcutpanretdet"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcutpanel1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    if (txtcutpanelid.Text != "")
                    {
                        string del = "delete from asptblcutpanretdet where asptblcutpanretdetid=" + txtcutpanelid.Text;
                        Utility.ExecuteNonQuery(del);                      
                        MessageBox.Show("Record Deleted Successfully " + txtcutpanel1id.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtcutpanel1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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

     

       
        ListView allip = new ListView();
        ListView allip1 = new ListView();
        ListView allip2 = new ListView();
        ListView allip3delete = new ListView();
    
     
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

                if (txtcutpanelid.Text == "")
                {
                   // News();
                   
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
            //    string sel1 = "SELECT distinct e.asptblcolmasid,e.colorname  FROM  asptblordentry a  join gtcompmast b on a.compcode=b.gtcompmastid join asptblpur c on c.orderno=a.orderno and c.compcode=b.gtcompmastid join asptblpurdet d on d.asptblpurid=c.asptblpurid join asptblcolmas e on e.asptblcolmasid=d.colorname where a.orderno='" + ord + "'  order by a.asptblordentryid desc";
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
                dt1 = Utility.SQLQuery("select a.asptblpurdet1id as qrcode, a.asptblpurdetid,a.asptblpurid,a.asptblpur1id, a.compcode,a.pono,a.colorname,a.sizename,a.Pcs,a.ProcessName,a.Remarks from asptblpurdet1 a JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID  JOIN asptblcolmas C ON C.COLORNAME=A.COLORNAME JOIN ASPTBLSIZMAS D ON D.SIZENAME=A.SIZENAME  where  b.compcode='" + s + "'  and a.pono='" + ss + "' and a.asptblpurdetid='" + sss + "'");
                if (dt1.Rows.Count > 0)
                {
                    dataGridView1.Rows.Clear();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[i].Cells[0].Value = dt1.Rows[i]["qrcode"].ToString();
                        dataGridView1.Rows[i].Cells[1].Value = dt1.Rows[i]["asptblpurdetid"].ToString();
                        dataGridView1.Rows[i].Cells[2].Value = dt1.Rows[i]["asptblpurid"].ToString();
                        dataGridView1.Rows[i].Cells[2].Value = dt1.Rows[i]["asptblpur1id"].ToString();
                        dataGridView1.Rows[i].Cells[3].Value = dt1.Rows[i]["compcode"].ToString();
                        dataGridView1.Rows[i].Cells[4].Value = dt1.Rows[i]["pono"].ToString();
                        dataGridView1.Rows[i].Cells[5].Value = dt1.Rows[i]["colorname"].ToString();
                        dataGridView1.Rows[i].Cells[6].Value = dt1.Rows[i]["sizename"].ToString();
                        dataGridView1.Rows[i].Cells[7].Value = dt1.Rows[i]["pcs"].ToString();
                        dataGridView1.Rows[i].Cells[8].Value = dt1.Rows[i]["ProcessName"].ToString();
                        dataGridView1.Rows[i].Cells[9].Value = dt1.Rows[i]["Remarks"].ToString();                       
                    }
                    CommonFunctions.SetRowNumber(dataGridView1);
                    lblcount.Refresh();lblcount.Text ="Total Rows :"+ dataGridView1.Rows.Count.ToString();
                }
            }
        }
        public void pono(string s, string ss,string sss)
        {
            if (s != "" && ss != "")
            {
                DataTable dt = new DataTable();
                dt = Utility.SQLQuery("SELECT  c.compcode,c.compname,c.gtcompmastid, d.buyercode,d.asptblbuymasid,e.sizegroup,e.asptblsizgrpid,f.stylename, f.asptblstymasid, g.processname,g.asptblpromasid ,a.orderqty, a.lotno,'' bundle FROM   asptblpur  a join gtcompmast c on a.compcode=c.gtcompmastid join asptblbuymas d on d.asptblbuymasid=a.buyer  join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup  join asptblstymas f on f.asptblstymasid=a.stylename join asptblpromas g on g.asptblpromasid=a.processname  where  c.compcode='" + s + "' and a.pono='" + ss + "'");
                if (dt.Rows.Count > 0)
                {
                   
             
                    combocompcode.DisplayMember = "compcode";
                    combocompcode.ValueMember = "gtcompmastid";
                    combocompcode.DataSource = dt;
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt;
                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.ValueMember = "asptblbuymasid";
                    combobuyer.DataSource = dt;
                    combostyle.DisplayMember = "stylename";
                    combostyle.ValueMember = "asptblstymasid";  
                    combostyle.DataSource = dt;
                    comboprocess.DisplayMember = "processname";
                    comboprocess.ValueMember = "asptblpromasid";
                    comboprocess.DataSource = dt;

                    dt = Utility.SQLQuery("select a.orderqty,a.lotno from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid    where  b.compcode='" + s + "' and a.pono='" + ss + "'");
                    txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                    txtlotno.Text = dt.Rows[0]["lotno"].ToString();

                }

            }
           
        }
        void auto()
        {
            
            DataTable dt1 = CC.autonumberload(Class.Users.Finyear, Class.Users.HCompcode, screen, "asptblcutpanret");
            if (dt1.Rows.Count > 0)
            {
                DataTable dt11 = CC.shortcode(Class.Users.Finyear, Class.Users.HCompcode, screen, "asptblcutpanret");
                if (dt11.Rows.Count < 0) { return; }
                else
                {
                    txtcutpanel1id.Text = "";
                    combocompcode.Text = dt11.Rows[0]["compcode"].ToString();
                    combocompname.Text = dt11.Rows[0]["COMPNAME"].ToString();
                     txtshortcode.Text = dt11.Rows[0]["shortcode"].ToString();
                    txtcutpanel1id.Text = dt1.Rows[0]["id"].ToString();                   
                    p.asptblcutpanret1id = Convert.ToInt64("0" + txtcutpanel1id.Text);
                   p.panelno=
                    txtpanelno.Text = Class.Users.Finyear + "-" + txtshortcode.Text + "-" + txtcutpanel1id.Text;
                    p.panelno = txtpanelno.Text;
                   Class.Users.UserTime = 0;
                }
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

            allip1.Items.Clear(); allip2.Items.Clear();
            
        }

        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopono.Text != "")
            {
                txtbarcode.Text = "";
                string sel1 = "SELECT a.asptblpurid,a.asptblpur1id,a.pono,a.shortcode, a.podate,a.orderqty,a.bundle,a.size, c.compcode,c.compname,c.gtcompmastid, d.buyercode as buyer,d.buyercode, d.asptblbuymasid,e.sizegroup,e.asptblsizgrpid,f.stylename, f.asptblstymasid, g.processname,g.asptblpromasid ,a.orderqty, a.lotno,a.processtype,a.orderno,a.styleref,a.garmentimage FROM   asptblpur  a join gtcompmast c on a.compcode=c.gtcompmastid join asptblbuymas d on d.asptblbuymasid=a.buyer  join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblstymas f on f.asptblstymasid=a.stylename join asptblpromas g on g.asptblpromasid=a.processname where a.PONO='" + combopono.Text+"' and c.compcode='"+combocompcode.Text+"' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                    DataTable dt = ds.Tables["asptblpur"];
                if (dt.Rows.Count > 0)
                {
                    pono(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblpurid"].ToString());
                    txtorderno.Text = Convert.ToString(dt.Rows[0]["orderno"].ToString());
                    txtstyleref.Text = Convert.ToString(dt.Rows[0]["styleref"].ToString());
                    txtpanelno.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                    txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                    dateTimePicker1.Value =Convert.ToDateTime(dt.Rows[0]["podate"].ToString().Substring(0, 10));
                  
                    combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                    combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                    combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                    comboprocess.Text = Convert.ToString(dt.Rows[0]["processname"].ToString());
                    combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                    combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                    txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                    txtbundle.Text = Convert.ToString(dt.Rows[0]["bundle"].ToString());
                    txtsize.Text = Convert.ToString(dt.Rows[0]["size"].ToString());
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
                // ponoGrid(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblpurdet1id"].ToString());

                Cursor.Current = Cursors.Default;
                
                if (combopono.SelectedIndex != -1)
                {                   
                        DataTable dt1 = CC.select("select MIN(A.barcode) AS  MINID , MAX(A.barcode) MAXID,count(a.barcode) cnt from asptblpurdet1 a where a.pono='" + combopono.Text + "'", "asptblpurdet1");
                    if (dt1.Rows[0]["MINID"].ToString() == "") { lblcount.Refresh(); lblcount.Text = "No Data Found.."; lblbarcode.Refresh(); lblbarcode.Text = "0"; }
                    else
                    {
                        lblbarcode.Refresh();
                        lblbarcode.Text = "'" + dt1.Rows[0]["MINID"].ToString() + "-" + dt1.Rows[0]["MAXID"].ToString() + "'";
                        totalcount = Convert.ToInt32(dt1.Rows[0]["cnt"].ToString());
                        if (GlobalVariables.New_Flg == false)
                        {
                            auto();
                        }
                    }
                }
            }
            dataGridView1.Rows.Clear();
            Cursor.Current = Cursors.Default;
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
            Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T'  order by 2  ", "asptblpromasid", "processname");
            Utility.Load_Combo(comboRemarks, "select asptblremmasid,remarks from asptblremmas WHERE active='T'  order by 2  ", "asptblremmasid", "remarks");

           // bindcombo();
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
           // CommonFunctions.SetRowNumber(dataGridView1);
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
        int aa = 0, a3 = 1, a1 = 0, cnt2 = 1;
        private void BindGrid(DataGridView grid, DataTable dt1, string barcode,int a2)
        {
           


            if (txtcutpanelid.Text == "" && GlobalVariables.New_Flg == false)
            {
                bool chek = mas.checkduplicate1(1, grid, barcode);

                if (chek == true)
                {
                    grid.Rows.Add();
                    rowcount = dataGridView1.Rows.Count - 1;
                    grid.Rows[rowcount].Cells[1].Value = dt1.Rows[0]["barcode"].ToString();
                    grid.Rows[rowcount].Cells[2].Value = dt1.Rows[0]["asptblpurdetid"].ToString();
                    grid.Rows[rowcount].Cells[3].Value = dt1.Rows[0]["asptblpurid"].ToString();
                    grid.Rows[rowcount].Cells[4].Value = dt1.Rows[0]["compcode"].ToString();
                    grid.Rows[rowcount].Cells[5].Value = dt1.Rows[0]["pono"].ToString();
                    grid.Rows[rowcount].Cells[6].Value = dt1.Rows[0]["colorname"].ToString();
                    grid.Rows[rowcount].Cells[7].Value = dt1.Rows[0]["sizename"].ToString();
                    grid.Rows[rowcount].Cells[8].Value = 1;
                    grid.Rows[rowcount].Cells[9].Value = comboissuetype.Text;
                    dataGridView1.Rows[rowcount].Cells[10].Value = comboRemarks.Text;
                    grid.Rows[rowcount].DefaultCellStyle.BackColor = rowcount % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                    //lblcount.Refresh();
                    //lblcount.Text = rowcount + " of " + rowcount.ToString();
                }
                else
                {
                    Cursor = Cursors.Default;
                    return;
                }
               
            }
            else
            {
                aa = 0; a3 = 1; a1 = 0; rowcount = 0; cnt2 = 0;
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    grid.Rows.Add();
                    grid.Rows[i].Cells[0].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblcutpanretdetid"].ToString());
                    grid.Rows[i].Cells[1].Value = dt1.Rows[i]["barcode"].ToString();
                    grid.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblpurdetid"].ToString());
                    grid.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblpurid"].ToString());
                    grid.Rows[i].Cells[4].Value = dt1.Rows[i]["compcode"].ToString();
                    grid.Rows[i].Cells[5].Value = dt1.Rows[i]["pono"].ToString();
                    grid.Rows[i].Cells[6].Value = dt1.Rows[i]["colorname"].ToString();
                    grid.Rows[i].Cells[7].Value = dt1.Rows[i]["sizename"].ToString();
                    grid.Rows[i].Cells[8].Value = 1;
                    grid.Rows[i].Cells[9].Value = dt1.Rows[i]["ISSUETYPE"].ToString();
                    grid.Rows[i].Cells[10].Value = dt1.Rows[i]["REMARKS"].ToString();                   
                  
                    cnt2++; rowcount++;
                    grid.Sort(grid.Columns[1], ListSortDirection.Descending);
                    lblcount.Refresh();
                    lblcount.Text = cnt2 + " of " + rowcount.ToString();
                }

            }

            grid.Sort(grid.Columns[a2], ListSortDirection.Descending);
            grid.AllowUserToAddRows = false;
        }
        private DataTable barcode(string type,string com, string pono, string bar)
        {
            //if (type == "PANEL MISTAKE")
            //{
            //    Class.Users.dt1 = Utility.SQLQuery("select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.DELIVERY='F' AND f.inward='F' AND F.CUTTING='F' ");

            //}
            if (type == "STITCHING MISTAKE")
            {
                Class.Users.dt1 = Utility.SQLQuery("select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.STITCHING='T' AND F.ISSUETYPE='INWARD' AND NOT F.panelmistake='CHECKING INWARD' UNION ALL select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.RESTITCHING='T' AND F.ISSUETYPE='REWORK'  AND NOT F.PANELMISTAKE='REWORK' AND NOT F.PANELMISTAKE='CHECKING DELIVERY' UNION ALL select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.restitching='T' AND F.ISSUETYPE='STITCHING MISTAKE' UNION ALL select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "'  AND F.ISSUETYPE='DELIVERY' AND NOT F.PANELMISTAKE='CHECKING DELIVERY' AND NOT F.PANELMISTAKE='DELIVERY'");

            }
            if (type == "CHECKING MISTAKE")
            {
                Class.Users.dt1 = Utility.SQLQuery("select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid   where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.ISSUETYPE='INWARD' AND F.CHECKING='T' UNION ALL select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid   where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.ISSUETYPE='REWORK' AND F.RECHECKING='T' AND NOT F.PANELMISTAKE='CHECKING DELIVERY'  UNION ALL select distinct f.barcode, c.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode, a.pono,d.colorname,e.SIZENAME from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.ISSUETYPE='DELIVERY' AND NOT F.PANELMISTAKE='CHECKING DELIVERY' AND NOT F.PANELMISTAKE='DELIVERY'");

            }

            return Class.Users.dt1;
        }
        int totalcount = 0;
        private void butadd_Click(object sender, EventArgs e)
        {

            if (txtbarcode.Text.Length >= 5)
            {
                aa = 0; a3 = 1; a1 = 0; rowcount = 1; cnt2 = 1; int a2 = 0;
                string source = ""; Class.Users.UserTime = 0;
                source = txtbarcode.Text.Trim();
                string data = getBetween(source, "'", "'"); rowcount = 0;
                string[] data1 = data.Split('-');
                if (data1.Length == 2)
                {

                    a1 = Convert.ToInt32(data1[0]); dataGridView1.Rows.Clear();
                    a2 = Convert.ToInt32(data1[1]);
                    a3 = 1;
                    a3 += a2 - a1; Cursor.Current = Cursors.WaitCursor;
                    for (aa = a1; aa <= a2; aa++)
                    {
                        DataTable dt = CC.select("select count(asptblcutpanretDETID) CNT from  asptblcutpanretDET   where  barcode='" + aa.ToString() + "' and  compcode='" + combocompcode.SelectedValue + "' and  pono='" + combopono.Text + "' AND RECHECKING='T' and issuetype='" + comboissuetype.Text + "' ", "asptblcutpanretdet");
                        string s = dt.Rows[0]["CNT"].ToString();
                        if (Convert.ToInt32(s) >= 1) { lblcount.Refresh(); lblcount.Text = "Child Record Found  .  " + a2.ToString(); lblcount.Text = a3 + " of " + cnt2.ToString(); cnt2++; }
                        else
                        {
                            dt1 = barcode(comboissuetype.Text, combocompcode.Text, combopono.Text, aa.ToString());
                            if (dt1.Rows.Count > 0)
                            {

                                BindGrid(dataGridView1, dt1, aa.ToString(), 1);
                                lblcount.Refresh();
                                lblcount.Text = a3 + " of " + cnt2.ToString();
                                cnt2++;
                            }
                            else
                            {
                                lblcount.Refresh();
                                lblcount.Text = a3 + " of " + cnt2.ToString();
                                cnt2++;
                                if (a3 == cnt2)
                                {
                                    Cursor.Current = Cursors.Default;
                                    MessageBox.Show("Invalid Barcode");
                                    lblcount.Refresh(); lblcount.Text = " Invalid BarCode ";
                                    CommonFunctions.SetRowNumber(dataGridView1);
                                    txtbarcode.Text = ""; return;
                                }

                            }
                        }
                    }

                    lblcount.Text = "Total Rows " + dataGridView1.Rows.Count.ToString();
                    CommonFunctions.SetRowNumber(dataGridView1);
                    Cursor.Current = Cursors.Default;
                }
                else
                {
                    if (txtbarcode.Text.Length >= 5)
                    {
                        bool chek = mas.checkduplicate1(1, dataGridView1, txtbarcode.Text);
                        if (chek == true)
                        {
                            try
                            {
                                DataTable dt = CC.select("select count(asptblcutpanretDETID) CNT from  asptblcutpanretDET   where  barcode='" + txtbarcode.Text + "' and  compcode='" + combocompcode.SelectedValue + "' and  pono='" + combopono.Text + "' AND RECHECKING='T' and issuetype='" + comboissuetype.Text + "' ", "asptblcutpanretdet");
                                string s = dt.Rows[0]["CNT"].ToString();
                                if (Convert.ToInt32(s) >= 1)
                                {                                  
                                    MessageBox.Show("Child Record Found "); MessageBox.Show("Child Record Found ");
                                    lblcount.Refresh(); lblcount.Text = "Child Record Found  .  "; txtbarcode.Text = "";
                                    Cursor.Current = Cursors.Default; return;
                                }
                                else
                                {

                                    dt1 = barcode(comboissuetype.Text, combocompcode.Text, combopono.Text, txtbarcode.Text);
                                    if (dt1.Rows.Count > 0)
                                    {
                                       
                                        BindGrid(dataGridView1, dt1, txtbarcode.Text, 1);
                                        lblcount.Refresh(); lblcount.Text = "Count :" + dataGridView1.Rows.Count.ToString() + " Of " + totalcount.ToString();
                                        Cursor.Current = Cursors.Default;
                                        txtbarcode.Text = ""; txtbarcode.Select();

                                    }
                                    else
                                    {
                                        MessageBox.Show("Invalid data  " + txtbarcode.Text, "Alert", MessageBoxButtons.OK);
                                        lblcount.Refresh(); lblcount.Text = "Count :" + dataGridView1.Rows.Count.ToString() + " Of " + totalcount.ToString();
                                        txtbarcode.Text = ""; txtbarcode.Select();
                                    }

                                }
                            }
                            catch (Exception EX) { MessageBox.Show(EX.Message.ToString()); }
                        }
                        else
                        {
                            txtbarcode.Text = "";
                        }
                    }
                }
            }
          
            Cursor.Current = Cursors.Default;
        }
     
        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                mas.checkduplicate(e.ColumnIndex, dataGridView1);
            }
        }
      
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            //    if (this.dataGridView1.Columns[e.ColumnIndex].HeaderText == "REMARKS")
            //    {
            //        combo = new DataGridViewComboBoxCell();
            //        this.combo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            //        bindcombo();
            //        this.dataGridView1[e.ColumnIndex, e.RowIndex] = combo;
            //        this.dataGridView1.Columns[e.ColumnIndex].ReadOnly = false;

            //    }
            //    else
            //    {
            //        this.dataGridView1.Columns[e.ColumnIndex].ReadOnly = true;
            //    }
            //}
            //catch { }
        }
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //ComboBox combo = e.Control as ComboBox;

            //if (combo != null)
            //{
            //    combo.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
            //    combo.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);

            //}
        }
        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //ComboBox cb = (ComboBox)sender;          
            //item = cb.Text;
            
        }
        private void txtbarcode_TextChanged(object sender, EventArgs e)
        {
            if (txtcutpanelid.Text != "")
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

        private void contextMenuStrip1_Opening_1(object sender, CancelEventArgs e)
        {

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

        private void panel4_Paint(object sender, PaintEventArgs e)
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

        private void txtbarcode_MouseLeave(object sender, EventArgs e)
        {
            txtbarcode.BackColor = Color.White;
        }

        private void comboprocesstype_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboprocesstype.Text == "INWARD")
            //{
            //    DataTable dt0 = Utility.SQLQuery("select count(A.asptblcutpanretdetid) CNT,max(A.asptblcutpanretdetid) maxid from asptblcutpanretdet A WHERE a.compcode='" + combocompcode.SelectedValue + "' and A.pono='" + combopono.Text + "'  and A.processtype='INWARD' and a.processcheck='T' ");// and A.processtype='INWARD'
            //    if (Convert.ToInt32("0" + dt0.Rows[0]["CNT"].ToString()) <0 || dt0.Rows[0]["maxid"].ToString() != "")
            //    {
            //        lblcount.Refresh();
            //        lblcount.Text = "";
            //    }
            //    else { combobuyer.Text = "";lblbarcode.Text = "";lblcount.Text = "Invalid Record .."; comboprocess.Text = "";combostyle.Text = "";}
                
            //}
            //else
            //{
            //    DataTable dt0 = Utility.SQLQuery("select count(A.asptblcutpanretdetid) CNT,max(A.asptblcutpanretdetid) maxid from asptblcutpanretdet A WHERE a.compcode='" + combocompcode.SelectedValue + "' and A.pono='" + combopono.Text + "'  and A.processtype='INWARD' and a.processcheck='T' ");// and A.processtype='INWARD'
            //    if (Convert.ToInt32("0" + dt0.Rows[0]["CNT"].ToString()) < 0 || dt0.Rows[0]["maxid"].ToString() != "")
            //    {
            //        Utility.Load_Combo(combopono, "select a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid JOIN asptblcutpanretdet C on c.pono=a.pono and c.compcode=b.gtcompmastid where  a.active='T' and b.compcode='" + combocompcode.Text + "'  order by a.asptblpurid desc", "pono", "pono");
            //        lblcount.Text = "";
            //    }
            //    else { combobuyer.Text = ""; lblbarcode.Text = ""; lblcount.Text = "Invalid Record .."; comboprocess.Text = ""; combostyle.Text = ""; }

            //}

        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }

        private void comboissuetype_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboRemarks.Enabled = false;
            comboRemarks.Enabled = false;
            butheader.Text = comboissuetype.Text + " - ENTRY";
            if (comboissuetype.Text == "CHECKING MISTAKE")
            {
                comboRemarks.Enabled = true; combopono.DataSource = null;
                Class.Users.dt = CC.select("select distinct A.asptblpurID, A.pono,B.compcode from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblchk c on c.pono=a.PONO   AND B.GTCOMPMASTID=C.COMPCODE   AND a.active='T' JOIN ASPTBLCHK d on d.pono=a.pono and d.pono=c.pono  LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where   E.PONO IS NULL and  B.compcode='" + Class.Users.HCompcode + "' AND C.ISSUETYPE='INWARD' order by A.asptblpurID desc", "asptblpur");

                if (Class.Users.dt.Rows.Count <= 0)
                {
                   
                    Class.Users.dt = CC.select("select distinct C.asptblpurID,C.pono,B.compcode from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblpurdet1 c on c.pono=a.PONO   AND B.GTCOMPMASTID=C.COMPCODE   AND a.active='T' JOIN ASPTBLCHK d on d.pono=a.pono and d.pono=c.pono LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where   E.PONO IS NULL and  B.compcode='" + Class.Users.HCompcode + "' AND C.DELIVERY='T' and C.ISSUETYPE='INWARD' order by C.pono desc", "asptblpur");
                }
                Utility.Load_Combo(comboRemarks, "select asptblremmasid,remarks from asptblremmas WHERE active='T' AND NOT remarks='PANEL MISTAKE' order by 2  ", "remarks", "remarks");

            }
            
            if (comboissuetype.Text == "STITCHING MISTAKE")
            {
                combopono.DataSource = null;
                Utility.Load_Combo(comboRemarks, "select asptblremmasid,remarks from asptblremmas WHERE active='T' AND NOT remarks='PANEL MISTAKE' AND NOT remarks='CHECKING MISTAKE' order by 2  ", "remarks", "remarks");
               
                Class.Users.dt = CC.select(" select distinct x.asptblpurid, x.pono from ( select a.asptblpurid,a.pono,B.compcode,C.ISSUETYPE from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblprolot c on c.pono=a.PONO   AND B.GTCOMPMASTID=C.COMPCODE  AND a.active='T' and c.issuetype='INWARD' LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where   E.PONO IS NULL and  b.compcode='" + Class.Users.HCompcode + "'  UNION ALL select a.asptblpurid,a.pono,B.compcode,C.ISSUETYPE from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid join asptblchk c on c.pono=a.PONO AND B.GTCOMPMASTID=C.COMPCODE  AND a.active='T' and c.processtype='DELIVERY' LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where   E.PONO IS NULL and  b.compcode='" + Class.Users.HCompcode + "'   ) X     order by x.asptblpurid desc", "asptblpur");//and X.issuetype='" + comboissuetype.Text.Trim() + "'//  UNION ALL select A.pono,B.compcode,A.ISSUETYPE from asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid  AND a.active='T' 
            }
           
            if (Class.Users.dt.Rows.Count > 0)
            {
                combopono.ValueMember = "pono";
                combopono.DisplayMember = "pono";
                combopono.DataSource = Class.Users.dt;
            }
            else
            {
                combopono.DataSource = null;
                butheader.Text = comboissuetype.Text + " - ENTRY";
                empty();
            }

            processload(comboissuetype.Text);
            combopono_SelectedIndexChanged(sender,e);
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

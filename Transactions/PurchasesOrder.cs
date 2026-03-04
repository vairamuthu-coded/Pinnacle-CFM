using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions
{
    public partial class PurchasesOrder : Form, ToolStripAccess
    {
        public PurchasesOrder()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            // tabControl1.SelectTab(tabPage2);


        }

        private static PurchasesOrder _instance; string coid = "", siid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        public static PurchasesOrder Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PurchasesOrder();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void usercheck(string s, string ss, string sss)
        {


        }
        private DataTable autonumberload(string y, string com, string scr)
        {

            DataTable dt1 = new DataTable();

            string sel1 = "select max(a.asptblpurid1)+1 as id,a.shortcode from asptblpur a join gtcompmast b on a.compcode = b.gtcompmastid  where a.finyear='" + y + "' and b.compcode='" + com + "' group by a.shortcode";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblpur");
            dt1 = ds1.Tables["asptblpur"];
            if (dt1.Rows.Count <= 0)
            {
                sel1 = "";
                sel1 = "select max(a.sequenceno)+1 as id,a.shortcode from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptbluserrights c on c.userrightsid=a.screen where a.finyear='" + y + "' and b.compcode='" + com + "' AND C.MENUNAME='" + scr + "' group by a.shortcode";
                ds1 = Utility.ExecuteSelectQuery(sel1, "asptblautogeneratemas");
                dt1 = ds1.Tables["asptblautogeneratemas"];

            }

            return dt1;
        }
        private void PurchasesOrder_Load(object sender, EventArgs e)
        {
         
            txtsearch.Select();
           
        }
        private void combostyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            sizeload(combosizegroup.Text);
        }
        public void News()
        {
           
         
            empty();
 
          
           // tabControl1.SelectTab(tabPage1);

        }

        private void empty()
        {
            txtasptblpurid.Text = ""; combocompcode.Text = ""; combocompname.Text = "";
            txtasptblpur1id.Text = ""; txtpono.Text = ""; GlobalVariables.New_Flg = false;
            combosizegroup.Enabled = true; combobuyer.Text = "";combosizegroup.Text = "";
            mas.ColIndex.Clear(); mas.SizeIndex.Clear(); allip2.Items.Clear();
            dataGridView1.Rows.Clear(); dataGridView1.Columns.Clear();
            allip1.Items.Clear(); checksizeall.Checked = false;
            txtasptblpur1id.Text = ""; Class.Users.UserTime = 0;
            checkactive.Checked = true; txtorderqty.Text = "";
            butheader.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            listView3.Font = Class.Users.FontName;
            foreach (ListViewItem l2 in listView2.Items)
                l2.Checked = false;
            foreach (ListViewItem l3 in listView3.Items)
                l3.Checked = false; panelcolorsize.Visible = false;
            GridLoad();
            compload(); colorload();
            Utility.Load_Combo(combobuyer, "select asptblbuymasid,buyercode from asptblbuymas where  active='T' order by 1", "asptblbuymasid", "buyercode");
            Utility.Load_Combo(combosizegroup, "select asptblsizgrpid,sizegroup from asptblsizgrp where  active='T' order by 1", "asptblsizgrpid", "sizegroup");
            Utility.Load_Combo(combostyle, "select asptblstymasid,stylename from asptblstymas where  active='T' order by 1", "asptblstymasid", "stylename");


            tabControl2.SelectTab(tabPage3);
        }
        public void Saves()
        {


            if (txtasptblpur1id.Text == "" && dataGridView1.Rows.Count < 0)
            {
                MessageBox.Show("colorname Name is empty " + " Alert " + txtasptblpur1id.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (txtasptblpur1id.Text != "" && dataGridView1.Rows.Count > 0)
            {
                string maxid = ""; Models.PurchasesModel p = new Models.PurchasesModel();

                p.asptblpurid = Convert.ToInt64("0" + txtasptblpurid.Text);
                p.asptblpur1id = Convert.ToInt64("0" + txtasptblpur1id.Text);
                p.shortcode = Convert.ToString(txtshortcode.Text);
                p.finyear = Class.Users.Finyear;
                p.podate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                p.pono = Convert.ToString(txtpono.Text);
                p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
                p.sizegroup = Convert.ToInt64("0" + combosizegroup.SelectedValue);
                p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                p.compcode1 = Class.Users.COMPCODE;
                p.username = Class.Users.USERID;
                p.createdby = Convert.ToString(Class.Users.HUserName);
                p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                p.modified = Convert.ToString(System.DateTime.Now.ToString());
                p.modifiedby = Class.Users.HUserName;
                p.ipaddress = Class.Users.IPADDRESS;
                if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; }
                if (checkpo.Checked == true) { p.pocancel = "T"; } else { p.pocancel = "F"; }
                if (GlobalVariables.New_Flg == false)
                {
                    DataTable dt1 = autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName);
                    if (dt1.Rows.Count > 0)
                    {
                        txtasptblpur1id.Text = dt1.Rows[0]["id"].ToString();
                        txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
                        txtpono.Text = Class.Users.Finyear + "-" + dt1.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
                        p.asptblpurid = Convert.ToInt64("0" + txtasptblpur1id.Text);
                        p.pono = Convert.ToString(txtpono.Text);
                        p.shortcode = Convert.ToString(txtshortcode.Text);
                    }
                    string ins = "insert into asptblpur(asptblpurid1,shortcode,finyear,podate,compcode,stylename,orderqty,pono,buyer,sizegroup,pocancel,active,compcode1,username,createdby,modifiedby,ipaddress)  VALUES('" + p.asptblpur1id + "','" + p.shortcode + "','" + p.finyear + "','" + p.podate + "','" + p.compcode + "','" + p.stylename + "','" + p.orderqty + "','" + p.pono + "','" + p.buyer + "','" + p.sizegroup + "','" + p.pocancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + p.modifiedby + "','" + p.ipaddress + "')";
                    Utility.ExecuteNonQuery(ins);

                    string sel2 = "select max(asptblpur1id) id    from  asptblpur   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
                    DataTable dt2 = ds2.Tables["asptblpur"];
                    maxid = dt2.Rows[0]["id"].ToString();
                }
                if (GlobalVariables.New_Flg == true)
                {
                    string up = "update  asptblpur  set   asptblpur1id='" + p.asptblpur1id + "' ,shortcode='" + p.shortcode + "',finyear='" + p.finyear + "' ,  podate='" + p.podate + "' , compcode ='" + p.compcode + "' ,stylename='" + p.stylename + "' , orderqty='" + p.orderqty + "', buyer='" + p.buyer + "' , pono='" + p.pono + "' ,  sizegroup='" + p.sizegroup + "'  , pocancel='" + p.pocancel + "' , active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblpurid='" + txtasptblpurid.Text + "'";
                    Utility.ExecuteNonQuery(up);
                    maxid = txtasptblpurid.Text;
                }
                Models.PurchasesModeldetail pp = new Models.PurchasesModeldetail();
                int i = 0, j = 0;
                if (dataGridView1.Rows.Count >= 0)
                {

                    for (i = 0; i < dataGridView1.Columns.Count - 1; i++)
                    {

                        for (j = 0; j < dataGridView1.Rows.Count; j++)
                        {
                            colorid(dataGridView1.Rows[j].Cells[0].Value.ToString());
                            pp.colorname = coid;
                            sizeid(dataGridView1.Columns[i + 1].HeaderText.ToString());
                            pp.sizename = siid;
                            pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[j].Cells[i + 1].Value.ToString());
                            pp.asptblpurdetid = Convert.ToInt64("0" + dataGridView1.Rows[j].Cells[1].Value.ToString());
                            pp.asptblpurid = Convert.ToInt64("0" + txtasptblpurid.Text);
                            pp.asptblpur1id = Convert.ToInt64(txtasptblpur1id.Text);
                            pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                            pp.pono = Convert.ToString(txtpono.Text);
                            string sel1 = "select asptblpurdetid   from  asptblpurdet   where  asptblpurid='" + pp.asptblpurid + "' and asptblpur1id='" + pp.asptblpur1id + "' and compcode='" + combocompcode.SelectedValue + "' and  pono='" + pp.pono + "'  and  colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "' ";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                            DataTable dt1 = ds1.Tables["asptblpur"];
                            if (dt1.Rows.Count <= 0)
                            {
                                string sel2 = "select max(asptblpurid) id    from  asptblpur   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and pono='" + txtpono.Text + "'";
                                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
                                DataTable dt2 = ds2.Tables["asptblpur"];
                                maxid = dt2.Rows[0]["id"].ToString();
                                string ins1 = "insert into asptblpurdet(asptblpurid,asptblpur1id,compcode,pono,colorname,sizename,orderqty) values('" + maxid.ToString() + "' ,'" + txtasptblpur1id.Text + "' , '" + combocompcode.SelectedValue + "' ,'" + pp.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.orderqty + "');";
                                Utility.ExecuteNonQuery(ins1);

                            }
                            else
                            {                               
                                if (dt1.Rows.Count == 1)
                                {
                                    string up1 = "update  asptblpurdet  set asptblpurid='" + pp.asptblpurid + "' ,asptblpur1id='" + pp.asptblpur1id + "', compcode='" + p.compcode + "' , pono='" + pp.pono + "' , colorname='" + pp.colorname + "',sizename='" + pp.sizename + "', orderqty='" + pp.orderqty + "' where asptblpurdetid='" + dt1.Rows[0]["asptblpurdetid"].ToString() + "'";
                                    Utility.ExecuteNonQuery(up1);
                                }
                                else
                                {
                                    MessageBox.Show("Invalid Update" + sel1.ToString());
                                }
                            }

                        }

                    }



                }

                if (txtasptblpurid.Text == "")
                {
                    MessageBox.Show("Record Saved Successfully " + txtasptblpurid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    empty(); GridLoad();
                    tabControl1.SelectTab(tabPage2);
                }
                if (txtasptblpurid.Text != "")
                {
                    MessageBox.Show("Record Updated Successfully " + txtasptblpurid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    empty(); GridLoad();

                    tabControl1.SelectTab(tabPage2);
                }
            }
            else
            {
                MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


            }


        }

        private void PurchasesOrder_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }





        public void colorid(string s)
        {
            try
            {
                string sel = "select asptblcolmasid,colorname from  asptblcolmas where active='T' and colorname='" + s + "'  order by 2";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblcolmas");
                DataTable dt = ds.Tables["asptblcolmas"];
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
                string sel = "select asptblsizmasid from  asptblsizmas where sizename='" + s + "' ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsizmas");
                DataTable dt = ds.Tables["asptblsizmas"];
                if (dt.Rows.Count > 0)
                {
                    siid = "";
                    siid = Convert.ToString(dt.Rows[0]["asptblsizmasid"].ToString());
                }
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
                combocompcode.Text = "";

            }
            catch (Exception EX)
            { }
        }



        private void colorload()
        {
            try
            {
                listView2.Items.Clear();
                string sel1 = " SELECT c.asptblcolmasid,c.colorname  FROM   asptblcolmas c   where  c.active='T' order by 2 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcolmas");
                DataTable dt = ds.Tables["asptblcolmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblcolmasid"].ToString());
                        list.SubItems.Add(myRow["colorname"].ToString());
                        list.SubItems.Add("");
                        this.listfilter2.Items.Add((ListViewItem)list.Clone());
                        listView2.Items.Add(list);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                    }
                    lbltotalcolor.Text = "Total Count    :" + listView2.Items.Count;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void sizeload(string sizgroup)
        {
            try
            {
                if (sizgroup != "")
                {
                    listView3.Items.Clear(); 
                    string sel1 = "   SELECT c.asptblsizmasid, c.sizename   FROM  asptblsizgrpDet a join asptblsizgrp b on a.asptblsizgrpid=b.asptblsizgrpid join asptblsizmas c on c.asptblsizmasid=a.sizename  where b.sizegroup='" + sizgroup + "'  order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizmas");
                    DataTable dt = ds.Tables["asptblsizmas"];
                    if (dt.Rows.Count > 0)
                    {
                        int i = 1;
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(myRow["asptblsizmasid"].ToString());
                            list.SubItems.Add(myRow["sizename"].ToString());
                            list.SubItems.Add("");
                            listView3.Items.Add(list);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            i++;
                        }
                        lbltotalcolor.Refresh();
                        lbltotalcolor.Text = "Total Count    :" + listView3.Items.Count;
                    }
                    panelcolorsize.Visible = true;
                    panelcolorsize.Width = 450;
                    panelcolorsize.Height = 400;
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
                string sel1 = "SELECT a.asptblpurid,b.compcode,a.podate,a.pono,e.sizegroup,a.orderqty, a.active FROM  asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup   order by  a.asptblpurid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                DataTable dt = ds.Tables["asptblpur"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblpurid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["podate"].ToString());
                        list.SubItems.Add(myRow["podate"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());
                        list.SubItems.Add(myRow["sizegroup"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

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
                    empty(); panelcolorsize.Visible = true;
                    txtasptblpurid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select a.asptblpurid,a.asptblpur1id,a.finyear,a.podate,b.compcode,a.stylename,a.pono,b.compname,c.buyercode,a.orderqty,d.sizegroup,a.pocancel,a.active  from  asptblpur a join gtcompmast b on b.gtcompmastid=a.compcode join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblsizgrp d on d.asptblsizgrpid=a.sizegroup   where a.asptblpurid=" + txtasptblpurid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                    DataTable dt = ds.Tables["asptblpur"];
                    if (dt.Rows.Count > 0)
                    {
                        GlobalVariables.New_Flg = true;
                       
                        txtasptblpurid.Text = Convert.ToString(dt.Rows[0]["asptblpurid"].ToString());
                        txtasptblpur1id.Text = Convert.ToString(dt.Rows[0]["asptblpurid1"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        txtpono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        combostyle.SelectedValue = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());                       
                        combosizegroup.Text = Convert.ToString(dt.Rows[0]["sizegroup"].ToString());
                        sizeload(combosizegroup.Text);
                        txtorderqty.Text = Convert.ToString(dt.Rows[0]["orderqty"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                      
                        string sel2 = "select distinct e.asptblsizmasid,  d.colorname from asptblpurdet a join asptblpur b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizmas e on e.asptblsizmasid=a.sizename  where c.compcode='" + combocompcode.Text + "' and a.asptblpurid='" + txtasptblpurid.Text + "' order by 2,1";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
                        DataTable dt2 = ds2.Tables["asptblpur"];
                     
                        foreach (DataRow row1 in dt2.Rows)
                        {
                            foreach (ListViewItem item0 in listView2.Items)
                            {
                                if (item0.SubItems[3].Text == row1["colorname"].ToString())
                                {
                                    item0.Checked = true;
                                }
                            }
                        }
                        string sel4 = "select distinct d.ASPTBLSIZMASID,e.colorname,  d.SIZENAME from asptblpur a join asptblpurdet b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=a.compcode join asptblsizmas d on d.ASPTBLSIZMASID=b.sizename join asptblcolmas e on e.asptblcolmasid=b.colorname  where c.compcode='" + combocompcode.Text + "' AND A.asptblpurid='" + txtasptblpurid.Text + "' order by 2,1";
                        DataSet ds4 = Utility.ExecuteSelectQuery(sel4, "asptblpur");
                        DataTable dt4 = ds4.Tables["asptblpur"];                       
                        foreach (DataRow row in dt4.Rows)
                        {
                            foreach (ListViewItem item1 in listView3.Items)
                            {
                                if (item1.SubItems[3].Text == row["sizename"].ToString())
                                {
                                    item1.Checked = true;
                                  
                                }
                            }
                        }
                      
                        butAdd_Click(sender, e);

                        panelcolorsize.Visible = false;
                        int i = 0, j = 0;
                        for (i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            string sel5 = "select  d.ASPTBLSIZMASID,  d.SIZENAME,b.orderqty,b.asptblpurdetid from asptblpur a join asptblpurdet b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=a.compcode join asptblsizmas d on d.ASPTBLSIZMASID=b.sizename join asptblcolmas e on e.asptblcolmasid=b.colorname  where c.compcode='" + combocompcode.Text + "' AND A.asptblpurid='" + txtasptblpurid.Text + "' and e.colorname='" + dataGridView1.Rows[i].Cells[0].EditedFormattedValue.ToString() + "'  order by 1";//and d.sizename='" + dataGridView1.Columns[i + 1].HeaderText.ToString() + "'
                            DataSet ds5 = Utility.ExecuteSelectQuery(sel5, "asptblpur");
                            DataTable dt5 = ds5.Tables["asptblpur"];
                            for (j = 0; j < dt5.Rows.Count; j++)
                            {

                                dataGridView1.Rows[i].Cells[j + 1].Value = dt5.Rows[j]["orderqty"].ToString();
                            }
                        }
                        
                    }
                    else
                    {
                        MessageBox.Show("Invalid");
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in ListView");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            tabControl1.SelectTab(tabPage1);
            combocompcode.Select();
        }

        private void butAdd_Click(object sender, EventArgs e)
        {
          
            string[] col = { "COLORNAME", "P-OUT" };
            GlobalVariables.WidthCols = new Int32[] {400};
            if (allip1.Items.Count >0 && allip2.Items.Count>0)
            {
                int cnt = 0;
                cnt = col.Length;
                foreach (string str in col)
                {
                    mas.SizeIndex.Add(str);
                }
                foreach (ListViewItem item1 in allip1.Items)
                {
                    mas.ColIndex.Add(item1.SubItems[2].Text);
                }

                foreach (ListViewItem item2 in allip2.Items)
                {
                    mas.SizeIndex.Add(item2.SubItems[2].Text);
                }

                if (mas.ColIndex.Count > 0)
                {
                    CommonFunctions.AddColumn(dataGridView1, mas.GridHeader.ToArray(), mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(),GlobalVariables.WidthCols,cnt);
                }
              
                mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
            }
           
        }

        private void butclear_Click(object sender, EventArgs e)
        {
            
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
                        if (item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[6].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[8].ToString().Contains(txtsearch.Text))
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
                    listView1.Items.Clear(); listView1.BackColor = System.Drawing.Color.LightGray;
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
                        list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);
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
            //        string sel1 = "  SELECT  a.asptblpurid,a.colorname,a.active from asptblpur a  where a.colorname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
            //        DataTable dt = ds.Tables["asptblpur"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptblpurid"].ToString());
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
            if (txtasptblpurid.Text != "")
            {
                string sel1 = "select a.asptblpurid from asptblpur a join asptbllay b on a.pono=b.pono where a.asptblpurid='" + txtasptblpurid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                DataTable dt = ds.Tables["asptblpur"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtasptblpur1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtasptblpurid.Text != "")
                    {
                        string del = "delete from asptblpur where asptblpurid=" + txtasptblpurid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptblpurdet where asptblpurid=" + txtasptblpurid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + txtasptblpur1id.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtasptblpur1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                }
            }
        }
        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void txtpurid1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void combosizegroup_SelectedIndexChanged(object sender, EventArgs e)
        {

            sizeload(combosizegroup.Text);
           
        }



        private void checkall_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //foreach (ListViewItem item in listView4.Items)
                //{
                //    item.Checked = true;
                //}

                //if (checkall.Checked == true)
                //{

                //    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                //    if (confirmation == DialogResult.Yes)
                //    {
                //        int i = 0;
                //        allip1.Items.Clear(); allip2.Items.Clear();
                //        listView4.Items.Clear();
                //        checkall.Checked = false;
                //        checksizeall.Checked = false;
                //        //listView2.Items.Clear();
                //        //listView3.Items.Clear();
                //        //colorload();
                //        //sizeload(combosizegroup.Text);


                //    }
                //}
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
                    e.Item.SubItems[4].Text = "checked";

                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.SubItems[3].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip1.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[4].Text == "checked")
                {
                    e.Item.SubItems[4].Text = "";
                    e.Item.Checked = false;
                    for (int c = 0; c < allip1.Items.Count; c++)
                    {
                        if (e.Item.SubItems[3].Text == allip1.Items[c].SubItems[2].Text)
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

                    e.Item.SubItems[4].Text = "checked";

                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.SubItems[3].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip2.Items.Add(it);

                }
                if (e.Item.Checked == false && e.Item.SubItems[4].Text == "checked")
                {
                    e.Item.SubItems[4].Text = "";
                    for (int c = 0; c < allip2.Items.Count; c++)
                    {
                        if (e.Item.SubItems[3].Text == allip2.Items[c].SubItems[2].Text)
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



        private void checksizeall_CheckedChanged(object sender, EventArgs e)
        {
            try
            {

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
            //try
            //{
            //    int item0 = 0; listView4.Items.Clear();int i = listfilter4.Items.Count;
            //    if (txtpursearch.Text.Length >= 1)
            //    {

            //        foreach (ListViewItem item in listfilter4.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (item.SubItems[3].ToString().Contains(txtpursearch.Text)  || item.SubItems[4].ToString().Contains(txtpursearch.Text))
            //            {
            //                list.Text = item.SubItems[0].Text;
            //                list.SubItems.Add(item.SubItems[1].Text);
            //                list.SubItems.Add(item.SubItems[2].Text);
            //                list.SubItems.Add(item.SubItems[3].Text);
            //                list.SubItems.Add(item.SubItems[4].Text);
            //                list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
            //                listView4.Items.Add(list);
            //            }
            //            item0++;i++;
            //        }
            //        lbltotalall.Text = "Total Count: " + listView4.Items.Count;
            //    }
            //    else
            //    {
            //        foreach (ListViewItem item in listfilter4.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            list.Text = item.SubItems[0].Text;
            //            list.SubItems.Add(item.SubItems[1].Text);
            //            list.SubItems.Add(item.SubItems[2].Text);
            //            list.SubItems.Add(item.SubItems[3].Text);
            //            list.SubItems.Add(item.SubItems[4].Text);
            //            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
            //            listView4.Items.Add(list);
            //            item0++; i++;
            //        }

            //        lbltotalall.Text = "Total Count: " + listView4.Items.Count;
            //    }



            //}
            //catch (Exception ex)
            //{
            //    //MessageBox.Show("---" + ex.ToString());
            //}
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void refeshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            empty(); GridLoad(); // sizegroupload();
            compload();// stylenameload();
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            //if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            //{

            //    if (txtpurid.Text == "")
            //    {
            //        empty();

            //    }
            //    combocompcode.Select();
            //}
            //if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            //{
            //    txtsearch.Select();

            //}
        }

        private void tabControl3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (tabControl3.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            //{

            //    if (txtpurid.Text == "")
            //    {
            //        empty();

            //    }
            //    combocompcode.Select();
            //}
            //if (tabControl3.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            //{
            //    txtsearch.Select();

            //}
        }

        private void txtcolorsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; int i = 1;
                if (txtcolorsearch.Text.Length >= 1)
                {
                    foreach (ListViewItem item in listfilter2.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtcolorsearch.Text))
                        {
                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView2.Items.Add(list);
                        }
                        item0++; i++;
                    }
                    lbltotalcolor.Text = "Total Count: " + listView2.Items.Count;
                }
                else
                {

                    listView2.Items.Clear(); i = 1;
                    foreach (ListViewItem item in listfilter2.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = item.SubItems[0].Text;
                        list.SubItems.Add(item.SubItems[1].Text);
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listView2.Items.Add(list);
                        item0++; i++;
                    }
                    lbltotalcolor.Text = "Total Count: " + listView2.Items.Count;
                }



            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }
        }

        private void combocompname_SelectedIndexChanged(object sender, EventArgs e)
        {



        }

        private void refeshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {


        }

        private void refreshToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            // compnameload();

        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = mas.findcomcode(combocompcode.Text);
            combocompname.DisplayMember = "compname";
            combocompname.ValueMember = "gtcompmastid";
            combocompname.DataSource = dt;

            if (combocompcode.Text != "" && txtasptblpurid.Text == "")
            {
                DataTable dt1 = autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName);
                if (dt1.Rows.Count > 0)
                {
                    txtasptblpur1id.Text = dt1.Rows[0]["id"].ToString();
                    txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
                    txtpono.Text = Class.Users.Finyear + "-" + dt1.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
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



        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CommonFunctions.SetRowNumber(dataGridView1);
        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();
        }


        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column1_KeyPress);
            if (dataGridView1.CurrentCell.ColumnIndex > 0) //Desired Column
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column1_KeyPress);
                }
            }

        }
        private void Column1_KeyPress(object sender, KeyPressEventArgs e)
        {            
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) || (sender as TextBox).Text.Length >= 8)
            {
                e.Handled = true;
            }
            
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void butshow_Click(object sender, EventArgs e)
        {
           
        }

        private void butclose_Click(object sender, EventArgs e)
        {
            panelcolorsize.Visible = false;
            butAdd_Click(sender,e);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        public void ReadOnlys()
        {
        }
    }
}

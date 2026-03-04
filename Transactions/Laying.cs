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
    public partial class Laying : Form,ToolStripAccess
    {
        public Laying()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
          
      
        }

        private static Laying _instance; string coid = "", siid = "", fabid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        DataGridViewComboBoxColumn UomColumn = new DataGridViewComboBoxColumn();
        public static Laying Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new Laying();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        private void Laying_Load(object sender, EventArgs e)
        {
          
            Utility.Load_Combo(combotableno, "select asptbltabmasid,tableno from asptbltabmas where active='T' order by 2", "asptbltabmasid", "tableno");
            compload();
            Class.Users.TableNameGrid = "asptbllaydet";
            GlobalVariables.HideCols = new string[] { "asptbllaydetid", "asptbllayid", "asptbllayid1", "compcode", "pono", "prodno" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 0, 300, 200, 100, 100, 100, 100, 100 };



            sel4 = "select  a.asptbllaydetid,a.asptbllayid,a.asptbllayid1, a.pono ,a.prodno,a.fabric, a.colorname ,a.sizename as size,a.markerno,a.orderqty as qty  from  asptbllaydet  a  join asptblcolmas b on b.asptblcolmasid=a.colorname where a.asptbllaydetid< 0";
            CommonFunctions.AddGridColumn(dataGridView1, sel4, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptbllaydet");
            UomColumn.Name = "UOM";
            dataGridView1.Columns.Insert(10, UomColumn);
            UomColumn.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            UomColumn.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
        }

        string maxid = "";
        public void Saves()
        {

            Models.LayCuttingModel p = new Models.LayCuttingModel();
            p.asptbllayid = Convert.ToInt64("0" + txtlayid.Text);
            p.asptbllayid1 = Convert.ToInt64("0" + txtlayid1.Text);
            p.docid = Convert.ToString(txtdocid.Text);
            p.finyear = Class.Users.Finyear;
            p.laydate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
            p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
            p.prodno = Convert.ToString(comboprodno.Text);
            p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
            p.shortcode = Convert.ToString(txtshortcode.Text);
            p.pono = Convert.ToString(combopono.Text);
            p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
            p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
            p.layno = Convert.ToString("0" + txtlayno.Text);
            p.tableno = Convert.ToInt64(combotableno.SelectedValue);
            p.markerno = Convert.ToString(txtmarkerno.Text);
            p.lotno = Convert.ToString(combolotno.Text);
            p.compcode1 = Class.Users.COMPCODE;
            p.username = Class.Users.USERID;
            p.createdby = Convert.ToString(Class.Users.HUserName);
            p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
            p.modified = Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            p.modifiedby = Class.Users.HUserName;
            p.ipaddress = Class.Users.IPADDRESS;
            if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; checkactive.Checked = false; }
            if (checklaycancel.Checked == true) { p.laycancel = "T"; } else { p.laycancel = "F"; checklaycancel.Checked = false; }
            if (GlobalVariables.New_Flg == false)
            {
                DataTable dt1 = autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName);
                if (dt1.Rows.Count > 0)
                {
                    txtlayid1.Text = dt1.Rows[0]["id"].ToString();
                    txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
                    txtdocid.Text = Class.Users.Finyear + "-" + dt1.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
                    p.docid = txtdocid.Text;
                    p.asptbllayid1 = Convert.ToInt64("0" + txtlayid1.Text);
                    p.layno = Convert.ToString(txtlayno.Text);
                    p.shortcode = Convert.ToString(txtshortcode.Text);
                }
                string ins = "insert into asptbllay(asptbllayid1,shortcode,docid,finyear,laydate,prodno,layno,compcode,buyer,pono,orderqty,stylename,lotno,markerno,tableno,laycancel,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.asptbllayid1 + "','" + p.shortcode + "','" + p.docid + "','" + p.finyear + "','" + p.laydate + "','" + p.prodno + "','" + p.layno + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.lotno + "','" + p.markerno + "' ,'" + p.tableno + "','" + p.laycancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + p.modified + "','" + p.modifiedby + "','" + p.ipaddress + "')";
                Utility.ExecuteNonQuery(ins);

                GridViewLoad();
            }

            if (GlobalVariables.New_Flg == true)
            {

                string up = "update  asptbllay   set  asptbllayid1='" + p.asptbllayid1 + "' ,shortcode='" + p.shortcode + "' ,docid='" + p.docid + "', finyear='" + p.finyear + "' ,  laydate='" + p.laydate + "' ,layno='" + p.layno + "', compcode ='" + p.compcode + "' , buyer='" + p.buyer + "' , pono='" + p.pono + "' ,  orderqty ='" + p.orderqty + "' , stylename='" + p.stylename + "' ,lotno='" + p.lotno + "', markerno='" + p.markerno + "'  ,tableno='" + p.tableno + "', laycancel='" + p.laycancel + "' , active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',modified='" + Convert.ToDateTime(p.modified).ToString("yyyy-MM-dd hh:mm:ss") + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbllayid='" + txtlayid.Text + "'";
                Utility.ExecuteNonQuery(up);

                GridViewLoad();
            }

            if (txtlayid.Text == "")
            {
                MessageBox.Show("Record Saved Successfully " + txtlayid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                GridLoad(); empty();

                tabControl1.SelectTab(tabPage2);
            }
            else
            {
                MessageBox.Show("Record Updated Successfully " + txtlayid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                GridLoad(); empty();

                tabControl1.SelectTab(tabPage2);
            }
        }
        private void GridViewLoad()
        {
            Models.LayCuttingModel.LayCuttingModeldet pp = new Models.LayCuttingModel.LayCuttingModeldet();
            int i = 0;

            if (dataGridView1.Rows.Count >= 0)
            {
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    pp.asptbllaydetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value);
                    pp.asptbllayid = Convert.ToInt64("0" + txtlayid.Text);
                    pp.asptbllayid1 = Convert.ToInt64("0" + txtlayid1.Text);
                    pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    pp.prodno = Convert.ToString(comboprodno.Text);
                    pp.pono = Convert.ToString(combopono.Text);
                    fabricid(dataGridView1.Rows[i].Cells[4].Value.ToString());
                    colorid(dataGridView1.Rows[i].Cells[5].Value.ToString());
                    sizeid(dataGridView1.Rows[i].Cells[6].Value.ToString());                  
                    pp.colorname = coid;
                    pp.sizename = siid;
                    pp.fabric = fabid;
                    pp.markerno = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value.ToString());
                    pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[8].Value.ToString());
                    pp.uom = Convert.ToString(dataGridView1.Rows[i].Cells[9].Value.ToString());
                    string sel = "select  a.asptbllaydetid   from  asptbllaydet  a where a.asptbllayid1='" + pp.asptbllayid1 + "'  and  a.compcode='" + pp.compcode + "'   and a.pono='" + pp.pono + "' and a.fabric='" + pp.fabric + "' and a.colorname='" + pp.colorname + "' and a.sizename='" + pp.sizename + "'  and a.markerno='" + pp.markerno + "'  and a.orderqty='" + pp.orderqty + "' and a.uom='" + pp.uom + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllay");
                    DataTable dt = ds.Tables["asptbllay"];
                    if (dt.Rows.Count != 0)
                    {

                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt64("0" + pp.asptbllaydetid) == 0 || Convert.ToInt64("0" + pp.asptbllaydetid) == 0)
                    {
                        string sel2 = "select max(asptbllayid) id,asptbllayid1    from  asptbllay   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and docid='" + txtdocid.Text + "' group by asptbllayid1";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbllay");
                        DataTable dt2 = ds2.Tables["asptbllay"];
                        if (dt2 == null) { }
                        {
                            pp.asptbllayid = Convert.ToInt64("0" + dt2.Rows[0]["id"].ToString());
                            pp.asptbllayid1 = Convert.ToInt64("0" + dt2.Rows[0]["asptbllayid1"].ToString());
                        }
                        string ins1 = "insert into asptbllaydet(asptbllayid,asptbllayid1,compcode,pono,colorname,sizename,fabric,markerno,orderqty,uom) values('" + pp.asptbllayid + "' ,'" + pp.asptbllayid1 + "' ,'" + pp.compcode + "' ,'" + pp.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.fabric + "' ,'" + pp.markerno + "','" + pp.orderqty + "' ,'" + pp.uom + "');";
                        Utility.ExecuteNonQuery(ins1);
                    }
                    else
                    {
                        string up1 = "update  asptbllaydet  set asptbllayid='" + pp.asptbllayid + "' ,asptbllayid1='" + pp.asptbllayid1 + "' , compcode='" + pp.compcode + "' , pono='" + pp.pono + "' , colorname='" + pp.colorname + "',sizename='" + pp.sizename + "' ,fabric='" + pp.fabric + "' , markerno='" + pp.markerno + "',orderqty='" + pp.orderqty + "' ,uom='" + pp.uom + "' where asptbllaydetid='" + pp.asptbllaydetid + "'";
                        Utility.ExecuteNonQuery(up1);
                    }


                }
            }
        }

       
        private void Laying_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       
        public void News()
        {

           empty();GridLoad();fabricload();
        }
        string sel4 = "";
        private void empty()
        {
            txtlayid.Text = ""; txtorderqty.Text = ""; txtdocid.Text = ""; txtlayno.Text = "";
            txtlayid1.Text = ""; comboprodno.Text = ""; combocompname.Text = ""; combopono.Text = "";
            combotableno.Text = ""; txtlayid1.Text = ""; combocompcode.Enabled = true;
            GlobalVariables.New_Flg = false;
            combocompcode.Text = "";
            combostyle.Text = "";
            combobuyer.Text = ""; panel9.Visible = false;
            txtlayid1.Text = "";
            checkactive.Checked = true;
            dataGridView1.Rows.Clear(); combolotno.Text = ""; txtmarkerno.Text = "";
            checklaycancel.Checked = false;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
          
            panel6.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            panel10.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            listView3.Font = Class.Users.FontName;          
            allip1.Items.Clear(); allip2.Items.Clear(); checksizeall.Checked = false;
            foreach (ListViewItem item in listView3.Items)
            {
                item.Checked = false;
            }
            foreach (ListViewItem item in listView2.Items)
            {
                item.Checked = false;
            }
            uomload(); CommonFunctions.SetRowNumber(dataGridView1);
        }

        public void colorid(string s)
        {
            try
            {
                string sel = "select asptblcolmasid from  asptblcolmas where colorname='" + s + "' ";
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
                string sel = "select asptblsizmasid from  asptblsizmas where sizename='" + s + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsizmas");
                DataTable dt = ds.Tables["asptblsizmas"];
                siid = "";
                siid = Convert.ToString(dt.Rows[0]["asptblsizmasid"].ToString());

            }
            catch (Exception EX)
            { }
        }
        private void fabricid(string text)
        {
            try
            {
                string sel = "select asptblfabmasid from  asptblfabmas where fabric='" + text + "' ";
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
        public void uomload()
        {
            try
            {
              
               string[] items1=new string[]{"","Kgs","Mtrs" };

                if (items1.Length > UomColumn.Items.Count)
                {
                    UomColumn.Items.Clear();
                    foreach (string str in items1)
                    {
                        UomColumn.Items.Add(str);
                    }
                }


            }
            catch (Exception EX)
            { }
        }
        private void fabricload()
        {
            try
            {
                listView3.Items.Clear();
                GlobalVariables.Dt = Utility.SQLQuery("select a.asptblfabmasid, a.fabric from asptblfabmas a   where A.ACTIVE='T'  order by 1");
                if (GlobalVariables.Dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in GlobalVariables.Dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();

                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblfabmasid"].ToString());
                        list.SubItems.Add(myRow["fabric"].ToString());                       
                        list.SubItems.Add("");                     
                        listView3.Items.Add(list);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                    }
                    lbltotalsize.Text = "Total Count    :" + listView3.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void colorload(string s,string ss)
        {
            try
            {
                listView2.Items.Clear();
                GlobalVariables.Dt = Utility.SQLQuery("select  d.colorname, e.SIZENAME, a.orderqty,c.lotqty,c.balqty from asptblprolot a join gtcompmast b on a.compcode = b.gtcompmastid  join asptblprolotdet c on c.asptblprolotid = a.asptblprolotid  and c.compcode = a.compcode and c.compcode = b.gtcompmastid join asptblcolmas d on d.asptblcolmasid = c.colorname  join asptblsizmas e on e.asptblsizmasid = c.sizename join asptblbuymas f on f.asptblbuymasid = a.buyer   where  b.compcode = '" + s + "' and a.prodno ='" + ss+"' order by 1");
                if (GlobalVariables.Dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in GlobalVariables.Dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["colorname"].ToString());
                        list.SubItems.Add(myRow["sizename"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["lotqty"].ToString());
                        list.SubItems.Add(myRow["balqty"].ToString());
                        list.SubItems.Add("");                       
                        listView2.Items.Add(list);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                    }
                    lbltotalsize.Text = "Total Count    :" + listView2.Items.Count;
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
                string sel1 = "SELECT a.asptbllayid, a.laydate,a.docid,b.compcode,c.buyername,a.layno,a.pono,a.orderqty,d.stylename,e.tableno,a.active FROM  asptbllay a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename join asptbltabmas e on e.asptbltabmasid=a.tableno   order by  a.asptbllayid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllay");
                DataTable dt = ds.Tables["asptbllay"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbllayid"].ToString());
                        list.SubItems.Add(myRow["laydate"].ToString().Substring(0,10));
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["buyername"].ToString());
                        list.SubItems.Add(myRow["layno"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());                      
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["tableno"].ToString());
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
                    GlobalVariables.New_Flg = true;
                    txtlayid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    combocompcode.Enabled = false;
                    string sel1 = " select a.asptbllayid,a.asptbllayid1,a.docid,a.prodno,a.layno,a.finyear,a.laydate,b.compcode,b.compname,e.buyername,a.pono,a.orderqty, c.stylename,a.lotno, d.tableno,a.markerno,a.laycancel,a.active  from  asptbllay a  join gtcompmast b on b.gtcompmastid=a.compcode   join asptblstymas c on c.asptblstymasid=a.stylename  join asptbltabmas d on d.asptbltabmasid=a.tableno   join asptblbuymas e on e.asptblbuymasid=a.buyer   where a.asptbllayid=" + txtlayid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllay");
                    DataTable dt = ds.Tables["asptbllay"];
                    if (dt.Rows.Count > 0)
                    {
                        txtlayid.Text = Convert.ToString(dt.Rows[0]["asptbllayid"].ToString());
                        txtlayid1.Text = Convert.ToString(dt.Rows[0]["asptbllayid1"].ToString());
                        txtdocid.Text = Convert.ToString(dt.Rows[0]["docid"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        comboprodno.Text = Convert.ToString(dt.Rows[0]["prodno"].ToString());
                        combopono_SelectedIndexChanged(sender, e);
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        txtlayno.Text = Convert.ToString(dt.Rows[0]["layno"].ToString());
                        combolotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                        combotableno.Text = Convert.ToString(dt.Rows[0]["tableno"].ToString());
                        txtmarkerno.Text = Convert.ToString(dt.Rows[0]["markerno"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        if (dt.Rows[0]["laycancel"].ToString() == "T") { checklaycancel.Checked = true; } else { checklaycancel.Checked = true; checklaycancel.Checked = false; }
                    }

                    string sel2 = "select a.asptbllaydetid,a.asptbllayid,a.asptbllayid1,c.compcode, b.pono, b.prodno,f.fabric, d.colorname,e.sizename,a.markerno,a.orderqty,a.uom from asptbllaydet a join asptbllay b on a.asptbllayid=b.asptbllayid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizmas e on e.asptblsizmasid=a.sizename join asptblfabmas f on f.ASPTBLFABMASID=a.fabric where a.asptbllayid=" + txtlayid.Text;
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbllay");
                    DataTable dt2 = ds2.Tables["asptbllay"];
                    dataGridView1.Rows.Clear(); listfilter4.Items.Clear();
                    int i = 1;
         
                    for (i = 0; i < dt2.Rows.Count; i++)
                    {

                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[i].Cells[0].Value = dt2.Rows[i]["asptbllaydetid"].ToString();
                        dataGridView1.Rows[i].Cells[1].Value = dt2.Rows[i]["asptbllayid"].ToString();
                        dataGridView1.Rows[i].Cells[2].Value = dt2.Rows[i]["asptbllayid1"].ToString();
                        dataGridView1.Rows[i].Cells[3].Value = combopono.Text;
                        dataGridView1.Rows[i].Cells[4].Value = combopono.Text;
                        dataGridView1.Rows[i].Cells[5].Value = dt2.Rows[i]["Fabric"].ToString();
                        dataGridView1.Rows[i].Cells[6].Value = dt2.Rows[i]["ColorName"].ToString();
                        dataGridView1.Rows[i].Cells[7].Value = dt2.Rows[i]["SizeName"].ToString();
                        dataGridView1.Rows[i].Cells[8].Value = txtmarkerno.Text;
                        dataGridView1.Rows[i].Cells[9].Value = dt2.Rows[i]["OrderQty"].ToString();
                        dataGridView1.Rows[i].Cells[10].Value = dt2.Rows[i]["UOM"].ToString();

                    }
                    lbltotal.Text = "Total Count    :" + dataGridView1.Rows.Count;


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            CommonFunctions.SetRowNumber(dataGridView1);
            tabControl1.SelectTab(tabPage1); combocompcode.Select();
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
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text) || item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[3].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text) || item.SubItems[6].ToString().Contains(txtsearch.Text))
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
            //        string sel1 = "  SELECT  a.asptbllayid,a.colorname,a.active from asptbllay a  where a.colorname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllay");
            //        DataTable dt = ds.Tables["asptbllay"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptbllayid"].ToString());
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
            if (txtlayid.Text != "")
            {
                string sel1 = "select a.asptbllayid from asptbllay a join gtstatemast b on a.asptbllayid=b.country where a.asptbllayid='" + txtlayid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllay");
                DataTable dt = ds.Tables["asptbllay"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtlayid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtlayid.Text != "")
                    {
                        string del = "delete from asptbllay where asptbllayid=" + txtlayid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptbllaydet where asptbllaydetid=" + txtlayid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + txtlayid1.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtlayid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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

        private void butlistinsert_Click(object sender, EventArgs e)
        {

            if (allip1.Items.Count > 0 && allip2.Items.Count > 0)
            {
                panel9.Visible = false;
                int i = 0 ,j = 0, k = 0;             
                dataGridView1.Rows.Clear();
               int cnt = 0;
                for (i = 0; i < allip1.Items.Count; i++)
                {
                    for (j = 0; j < allip2.Items.Count; j++)
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[cnt].Cells[0].Value = "0";
                        dataGridView1.Rows[cnt].Cells[1].Value = txtlayid.Text;
                        dataGridView1.Rows[cnt].Cells[2].Value = txtlayid1.Text;
                        dataGridView1.Rows[cnt].Cells[3].Value = combopono.Text;
                        dataGridView1.Rows[cnt].Cells[4].Value = comboprodno.Text;
                        dataGridView1.Rows[cnt].Cells[5].Value = allip1.Items[i].SubItems[3].Text;
                        dataGridView1.Rows[cnt].Cells[6].Value = allip2.Items[j].SubItems[1].Text;
                        dataGridView1.Rows[cnt].Cells[7].Value = allip2.Items[j].SubItems[2].Text;
                        dataGridView1.Rows[cnt].Cells[8].Value = txtmarkerno.Text;
                        dataGridView1.Rows[cnt].Cells[9].Value = allip2.Items[j].SubItems[3].Text;
                        dataGridView1.Rows[cnt].Cells[10].Value = "";
                        cnt++;
                    }
                }

                //if (txtlayid.Text != "")
                //{
                //    string del = "delete from asptbllaydet a where a.pono='" + combopono.Text + "' and a.asptbllayid='" + txtlayid.Text + "'";
                //    Utility.ExecuteNonQuery(del);
                //}

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
            else
            {
                MessageBox.Show("pls Select Fabric  and Color Grid");
            }
            lbltotalall.Text = "Total Count: " + dataGridView1.Rows.Count;
           // CommonFunctions.SetRowNumber(dataGridView1);
        }

        private void checkall_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
               
                if (checkall.Checked == true)
                {

                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        int i = 0;
                        allip1.Items.Clear(); allip2.Items.Clear();
                        dataGridView1.Rows.Clear();
                        checkall.Checked = false;
                        checksizeall.Checked = false;

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
        ListView allip3delete = new ListView();
        private void listView2_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {


                ListViewItem it = new ListViewItem();
                if (e.Item.Checked == true)
                {
                    e.Item.SubItems[7].Text = "✔";
                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.SubItems[3].Text);
                    it.SubItems.Add(e.Item.SubItems[4].Text);
                    it.SubItems.Add(e.Item.SubItems[5].Text);
                    it.SubItems.Add(e.Item.SubItems[6].Text);
                    it.SubItems.Add(e.Item.SubItems[7].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip2.Items.Add(it);
                }
                if (e.Item.Checked == false && e.Item.SubItems[7].Text == "✔")
                {
                    e.Item.SubItems[7].Text = "✖";
                    e.Item.Checked = false;
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

        private void listView3_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                ListViewItem it = new ListViewItem();
                if (e.Item.Checked == true)
                {

                    e.Item.SubItems[4].Text = "✔";
                    it.SubItems.Add(e.Item.SubItems[1].Text);
                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.SubItems[3].Text);                  
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip1.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[4].Text == "✔")
                {
                    e.Item.SubItems[4].Text = "✖";
                    for (int c = 0; c < allip1.Items.Count; c++)
                    {
                        if (e.Item.SubItems[2].Text == allip1.Items[c].SubItems[2].Text)
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
                        foreach (DataGridViewRow eachItem in dataGridView1.Rows)
                        {
                            string ID = eachItem.Cells[2].EditedFormattedValue.ToString();
                            if (Convert.ToInt64("0" + ID) > 1)
                            {
                                string del = "delete from asptbllaydet where asptbllaydetid=" + ID;
                                Utility.ExecuteNonQuery(del);
                            }
                            dataGridView1.Rows.Remove(eachItem);
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
                //    allip2.Items.Add(it);
                //}
                if (checksizeall.Checked == true)
                {
                    foreach (ListViewItem item in listView2.Items)
                    {
                        item.Checked = true;
                    }
                }
                else
                {
                    foreach (ListViewItem item in listView2.Items)
                    {
                        item.Checked = false;
                    }
                }
                //if (checksizeall.Checked == true)
                //{
                //    foreach (ListViewItem item in listView3.Items)
                //    {
                //        item.Checked = true;
                //    }
                //}
                //else
                //{
                //    foreach (ListViewItem item in listView3.Items)
                //    {
                //        item.Checked = false;
                //    }
                //}

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
                int item0 = 0; dataGridView1.Rows.Clear();
                if (txtpursearch.Text.Length >= 1)
                {

                    foreach (DataGridViewRow item in dataGridView1.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.Cells[2].ToString().Contains(txtpursearch.Text) || item.Cells[3].ToString().Contains(txtpursearch.Text))
                        {

                           
                            //list.Text = item.SubItems[0].Text;
                            //list.SubItems.Add(item.SubItems[1].Text);
                            //list.SubItems.Add(item.SubItems[2].Text);
                            //list.SubItems.Add(item.SubItems[3].Text);
                            //list.SubItems.Add(item.SubItems[4].Text);
                            //list.SubItems.Add(item.SubItems[5].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            dataGridView1.Rows.Add(list);
                        }
                        item0++;
                    }
                    lbltotalall.Text = "Total Count: " + dataGridView1.Rows.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    item0 = listfilter.Items.Count;
                    dataGridView1.Rows.Clear();
                    foreach (ListViewItem item in listfilter4.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = item.SubItems[0].Text;
                        list.SubItems.Add(item.SubItems[1].Text);
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.SubItems.Add(item.SubItems[5].Text);
                        list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        dataGridView1.Rows.Add(list);

                        item0++;
                    }

                    lbltotalall.Text = "Total Count: " + dataGridView1.Rows.Count;
                }



            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }
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

                if (txtlayid.Text == "")
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
        public void pono(string s,string ss)
        {
            DataTable Dt = new DataTable();
            if (s != "" && ss != "")               
            {
                Dt = Utility.SQLQuery("select distinct e.asptblstymasid,e.stylename,f.asptblbuymasid,f.buyercode,a.orderqty,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblstymas e on e.asptblstymasid=a.stylename join asptblbuymas f on f.asptblbuymasid=a.buyer   join asptblprolot g on g.pono=a.pono and g.compcode=a.compcode and g.compcode=b.gtcompmastid  where  b.compcode='" + s + "' and g.prodno='" + ss + "'");
                if (Dt.Rows.Count > 0)
                {
                    combostyle.DisplayMember = "stylename";
                    combostyle.ValueMember = "asptblstymasid";
                    combostyle.DataSource = Dt;

                    combostyle.DisplayMember = "stylename";
                    combostyle.ValueMember = "asptblstymasid";
                    combostyle.DataSource = Dt;

                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.ValueMember = "asptblbuymasid";
                    combobuyer.DataSource = Dt;
                    txtorderqty.Text = Dt.Rows[0]["orderqty"].ToString();
                    combopono.DisplayMember = "pono";
                    combopono.ValueMember = "pono";
                    combopono.DataSource = Dt;
                }
            }
        }
        private DataTable autonumberload(string y, string com, string scr)
        {
            DataTable dt1 = new DataTable();
            string sel1 = "select max(a.asptbllayid1)+1 as id,a.shortcode from asptbllay a join gtcompmast b on a.compcode = b.gtcompmastid  where a.finyear='" + y + "' and b.compcode='" + com + "' group by a.shortcode";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbllay");
            dt1 = ds1.Tables["asptbllay"];
            if (dt1.Rows.Count <= 0)
            {
                sel1 = "";
                sel1 = "select max(a.sequenceno)+1 as id,a.shortcode from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptbluserrights c on c.userrightsid=a.screen where a.finyear='" + y + "' and b.compcode='" + com + "' AND C.MENUNAME='" + scr + "' group by a.shortcode";
                ds1 = Utility.ExecuteSelectQuery(sel1, "asptblautogeneratemas");
                dt1 = ds1.Tables["asptblautogeneratemas"];
            }

            return dt1;
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

            allip1.Items.Clear(); allip2.Items.Clear();
            if (((ComboBox)sender).SelectedIndex != -1)
            {
                if (sender == combocompcode)
                {
                    DataTable Dt = new DataTable();
                    Int64 c = Convert.ToInt64(combocompcode.SelectedValue.ToString());
                    Dt = Utility.SQLQuery("select distinct g. prodno,b.compname from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblstymas e on e.asptblstymasid=a.stylename join asptblbuymas f on f.asptblbuymasid=a.buyer   join asptblprolot g on g.pono=a.pono and g.compcode=a.compcode and g.compcode=b.gtcompmastid  where  b.compcode='" + combocompcode.Text + "' ");

                    // GlobalVariables.Dt = Utility.SQLQuery("select A.asptblpurid,A.PONO,B.COMPNAME  from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + combocompcode.Text + "'");
                    if (Dt.Rows.Count > 0)
                    {
                        comboprodno.DisplayMember = "prodno";
                        comboprodno.ValueMember = "prodno";
                        combocompname.Text = Dt.Rows[0]["COMPNAME"].ToString();
                        comboprodno.DataSource = Dt;
                        if (combocompcode.Text != "" && txtlayid.Text == "")
                        {
                            DataTable dt1 = autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName);
                            if (dt1.Rows.Count > 0)
                            {
                                txtlayid1.Text = dt1.Rows[0]["id"].ToString();
                                txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
                                txtdocid.Text = Class.Users.Finyear + "-" + dt1.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
                            }
                        }
                    }
                    else
                    {
                        comboprodno.DataSource = null;
                    }
                }
            }
        }
        private void combopono_SelectedIndexChanged_1(object sender, EventArgs e)
        {
           DataTable Dt = Utility.SQLQuery("select distinct g.lotno from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblstymas e on e.asptblstymasid=a.stylename join asptblbuymas f on f.asptblbuymasid=a.buyer   join asptblprolot g on g.pono=a.pono and g.compcode=a.compcode and g.compcode=b.gtcompmastid  where  b.compcode='" + combocompcode.Text + "' and g.prodno='" + comboprodno.Text + "' and g.pono='" + combopono.Text + "'");
            if (Dt.Rows.Count > 0)
            {
        
                combolotno.DisplayMember = "lotno";
                combolotno.ValueMember = "lotno";
                combolotno.DataSource = Dt;
            }
        }
        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboprodno.Text != "")
            {
             
                pono(combocompcode.Text, comboprodno.Text);
                colorload(combocompcode.Text, comboprodno.Text);
                fabricload();
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
        private void colorsizeqty(string ord,string col)
        {
            try
            {
                if (ord != "" && col != "")
                {
                    listView2.Items.Clear();
                    string sel1 = "select f.colorname,g.sizename,c.shipqty  from asptblordentry a join asptblordentrydet b on a.asptblordentryid=b.asptblordentryid join asptblordentrysubdet c on c.asptblordentryid=a.asptblordentryid and c.asptblordentryid=b.asptblordentryid AND b.indexno=c.indexno join gtcompmast e on e.gtcompmastid=a.compcode and e.gtcompmastid=b.compcode join asptblcolmas f on f.asptblcolmasid=b.colorname join asptblsizmas g on g.asptblsizmasid=c.sizename where  a.orderno='" + ord + "' and f.colorname='" + col + "' order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizmas");
                    DataTable dt = ds.Tables["asptblsizmas"];
                    if (dt.Rows.Count > 0)
                    {
                        int i = 1;
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(myRow["colorname"].ToString());
                            list.SubItems.Add(myRow["sizename"].ToString());
                            list.SubItems.Add(myRow["shipqty"].ToString());
                            list.SubItems.Add("");
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView2.Items.Add(list);
                            i++;
                        }
                        lbltotalsize.Text = "Total Count    :" + listView2.Items.Count;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void fabricsizeqty(string ord, string col)
        {
            try
            {
                if (ord != "" && col != "")
                {
                    listView3.Items.Clear();
                    string sel1 = "select distinct h.fabric  from asptblordentry a  join asptblordentrydet b on a.asptblordentryid=b.asptblordentryid   join asptblordentrysubdet c on c.asptblordentryid=a.asptblordentryid and c.asptblordentryid=b.asptblordentryid  AND b.indexno=c.indexno  join gtcompmast e on e.gtcompmastid=a.compcode and e.gtcompmastid=b.compcode join asptblcolmas f on f.asptblcolmasid=C.colorname AND f.asptblcolmasid=b.colorname join asptblsizmas g on g.asptblsizmasid=c.sizename join asptblfabricmas h on h.asptblfabricmasid=b.fabric and h.asptblfabricmasid=c.fabric  where  a.orderno='" + ord + "' and f.colorname='" + col + "' order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizmas");
                    DataTable dt = ds.Tables["asptblsizmas"];
                    if (dt.Rows.Count > 0)
                    {
                        int i = 1;
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(myRow["fabric"].ToString());
                            list.SubItems.Add("");
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView3.Items.Add(list);
                            i++;
                        }
                        lbltotalsize.Text = "Total Count    :" + listView3.Items.Count;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void combolotno_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void butshow_Click(object sender, EventArgs e)
        {
            if (panel9.Visible == false)
            {
                panel9.Height = 450;
                panel9.Width = 455;
                panel9.Visible = true;
            }
            else
            {
                panel9.Visible = false;
            }
        }

        private void btnclose_Click(object sender, EventArgs e)
        {
            panel9.Visible = false;
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            
        }

        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //dataGridView1.CheckedIndexCollection checkedItems = dataGridView1.CheckedIndices;
            //while (checkedItems.Count > 0)
            //{
            //    dataGridView1.Rows.RemoveAt(checkedItems[0]);
            //    foreach (ListViewItem item in allip3delete.Items)
            //    {
            //        string del = "delete from asptbllaydet where asptbllaydetid='" + Convert.ToInt64("0" + item.SubItems[1].Text) + "'";
            //         Utility.ExecuteNonQuery(del);
            //    }
            //}
           
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions
{
    public partial class BundleGeneration : Form,ToolStripAccess
    {
        public BundleGeneration()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            //tabControl1.SelectTab(tabPage2);
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;           
            panel3.BackColor = Class.Users.BackColors;            
       
             
        }

        private static BundleGeneration _instance; string coid = "", siid = "", fabid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView(); string sel4, sel5 = "";
        public static BundleGeneration Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new BundleGeneration();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        private DataTable autonumberload(string y, string com, string scr)
        {
            DataTable dt1 = new DataTable();
            string sel1 = "select max(a.asptblbunid1)+1 as id,a.shortcode,b.compname from asptblbun a join gtcompmast b on a.compcode = b.gtcompmastid  where a.finyear='" + y + "' and b.compcode='" + com + "' group by a.shortcode,b.compname";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblbun");
            dt1 = ds1.Tables["asptblbun"];
            if (dt1.Rows.Count <= 0)
            {
                sel1 = "";
                sel1 = "select max(a.sequenceno)+1 as id,a.shortcode,b.compname from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptbluserrights c on c.userrightsid=a.screen where a.finyear='" + y + "' and b.compcode='" + com + "' AND C.MENUNAME='" + scr + "' group by a.shortcode,b.compname";
                ds1 = Utility.ExecuteSelectQuery(sel1, "asptblautogeneratemas");
                dt1 = ds1.Tables["asptblautogeneratemas"];
            }

            return dt1;
        }
        private void BundleGeneration_Load(object sender, EventArgs e)
        {
             GridLoad();
           
            //Utility.Load_Combo(comboshiftname, "select asptblshitypeid,shiftno from asptblshitype where active='T' order by 2", "asptblshitypeid", "shiftno");
          
        }
        string maxid = "";
        public void Saves()
        {
            try
            {



                Models.BundleModel p = new Models.BundleModel();
                p.asptblbunid = Convert.ToInt64("0" + txtbunid.Text);
                p.asptblbunid1 = Convert.ToInt64("0" + txtbunid1.Text);
                p.docid = Convert.ToString(txtdocid.Text);
                p.shortcode = Convert.ToString(txtshortcode.Text);
                p.finyear = Class.Users.Finyear;
                p.docdate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                p.compname = Convert.ToInt64(combocompcode.SelectedValue);
                p.pono = Convert.ToString(combopono.Text);
                p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);                    
                p.cutno = Convert.ToString(txtcutno.Text);
                p.lotno = Convert.ToString(txtlotno.Text);
                p.compcode1 = Class.Users.COMPCODE;
                p.username = Class.Users.USERID;
                p.createdby = Convert.ToString(Class.Users.HUserName);
                p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                p.modified = Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                p.modifiedby = Class.Users.HUserName;
                p.ipaddress = Class.Users.IPADDRESS;
                if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; checkactive.Checked = false; }
                if (GlobalVariables.New_Flg == false)
                {
                    string ins = "insert into asptblbun(asptblbunid1,shortcode,docid,finyear,docdate,compcode,buyer,pono,orderqty,stylename,lotno,layno,cutno,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.asptblbunid1 + "','" + p.shortcode + "','" + p.docid + "','" + p.finyear + "','" + p.docdate + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.lotno + "','" + p.layno + "' ,'" + p.cutno + "' ,'" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + p.modified + "','" + p.modifiedby + "','" + p.ipaddress + "');";
                    Utility.ExecuteNonQuery(ins);
                    GridViewLoad();
                }
                if (GlobalVariables.New_Flg == true)
                {
                    string up = "update  asptblbun   set  asptblbunid1='" + p.asptblbunid1 + "',shortcode='"+p.shortcode+ "',docid='" + p.docid + "' ,finyear='" + p.finyear + "' ,  docdate='" + p.docdate + "' , compcode ='" + p.compcode + "' ,buyer='" + p.buyer + "' , pono='" + p.pono + "' ,  orderqty ='" + p.orderqty + "' , stylename='" + p.stylename + "' ,lotno='" + p.lotno + "', layno='" + p.layno + "',cutno='" + p.cutno + "'  ,active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',modified='" + Convert.ToDateTime(p.modified).ToString("yyyy-MM-dd hh:mm:ss") + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblbunid='" + txtbunid.Text + "'";
                    Utility.ExecuteNonQuery(up);
                    GridViewLoad();
                }
                if (txtbunid.Text == "")
                {
                    MessageBox.Show("Record Saved Successfully " + txtbunid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tabControl1.SelectTab(tabPage2);
                }
                else
                {
                    MessageBox.Show("Record Updated Successfully " + txtbunid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tabControl1.SelectTab(tabPage2);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("colorname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }


        private void GridViewLoad()
        {
            Models.BundleModel.BundleModelDet pp = new Models.BundleModel.BundleModelDet();
            int i = 0;

            if (dataGridView2.Rows.Count >= 0)
            {
                for (i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    pp.asptblbundetid = Convert.ToInt64("0" + dataGridView2.Rows[i].Cells[0].Value);
                    pp.asptblbunid = Convert.ToInt64("0" + txtbunid.Text);
                    pp.asptblbunid1 = Convert.ToInt64("0" + txtbunid1.Text);
                    pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    pp.pono = Convert.ToString(combopono.Text);
                    pp.layno = Convert.ToString(combolayno.Text);
                    fabricid(dataGridView2.Rows[i].Cells[6].Value.ToString());
                    colorid(dataGridView2.Rows[i].Cells[7].Value.ToString());
                    sizeid(dataGridView2.Rows[i].Cells[8].Value.ToString());
                    pp.colorname = coid;
                    pp.sizename = siid;
                    pp.fabric = fabid;
                    pp.markerno = Convert.ToInt64(dataGridView2.Rows[i].Cells[9].Value.ToString());
                    pp.bundleno = Convert.ToInt64(dataGridView2.Rows[i].Cells[10].Value.ToString());
                    pp.fromply = Convert.ToInt64(dataGridView2.Rows[i].Cells[11].Value.ToString());
                    pp.toply = Convert.ToInt64(dataGridView2.Rows[i].Cells[12].Value.ToString());
                    pp.cutqty = Convert.ToInt64("0" + dataGridView2.Rows[i].Cells[13].Value.ToString());
                    pp.notes = Convert.ToString(i.ToString() + "-" + dataGridView2.Rows[i].Cells[14].EditedFormattedValue.ToString());
                    string sel = "select  a.asptblbundetid   from  asptblbundet  a where a.asptblbunid1='" + pp.asptblbunid1 + "'  and  a.compcode='" + pp.compcode + "'   and a.pono='" + pp.pono + "' and a.layno='" + pp.layno + "'  and a.fabric='" + pp.fabric + "' and a.colorname='" + pp.colorname + "' and a.sizename='" + pp.sizename + "'  and a.markerno='" + pp.markerno + "'  and a.bundleno='" + pp.bundleno + "' and a.fromply='" + pp.fromply + "' and a.toply='" + pp.toply + "' and a.cutqty='" + pp.cutqty + "' and a.notes='" + pp.notes + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllay");
                    DataTable dt = ds.Tables["asptbllay"];
                    if (dt.Rows.Count != 0)
                    {

                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt64("0" + pp.asptblbundetid) == 0 || Convert.ToInt64("0" + pp.asptblbundetid) == 0)
                    {
                        string sel2 = "select max(asptblbunid) id,asptblbunid1    from  asptblbun   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and docid='" + txtdocid.Text + "' group by asptblbunid1";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblbun");
                        DataTable dt2 = ds2.Tables["asptblbun"];
                        if (dt2 == null) { }
                        {
                            pp.asptblbunid = Convert.ToInt64("0" + dt2.Rows[0]["id"].ToString());
                            pp.asptblbunid1 = Convert.ToInt64("0" + dt2.Rows[0]["asptblbunid1"].ToString());
                        }
                        string ins1 = "insert into asptblbundet(asptblbunid,asptblbunid1,compcode,pono,layno,Fabric,colorname,sizename,markerno,bundleno, fromply,toply,cutqty,notes) values('" + pp.asptblbunid + "','" + pp.asptblbunid1 + "','" + pp.compcode + "','" + pp.pono + "','" + pp.layno + "','" + pp.fabric + "','" + pp.colorname + "','" + pp.sizename + "','" + pp.markerno + "','" + pp.bundleno + "','" + pp.fromply + "','" + pp.toply + "','" + pp.cutqty + "','" + pp.notes + "')";
                        Utility.ExecuteNonQuery(ins1);
                    }
                    else
                    {
                        string up1 = "update  asptblbundet  set asptblbunid='"+pp.asptblbunid+ "',asptblbunid1='" + pp.asptblbunid1 + "',compcode='" + pp.compcode + "',pono='" + pp.pono + "',layno='" + pp.layno + "',Fabric='" + pp.fabric + "',colorname='" + pp.colorname + "',sizename='" + pp.sizename + "',markerno='" + pp.markerno + "',bundleno='" + pp.bundleno + "', fromply='" + pp.fromply + "',toply='" + pp.toply + "',cutqty='" + pp.cutqty + "',notes='" + pp.notes + "' where asptblbundetid='" + pp.asptblbundetid + "'";
                        Utility.ExecuteNonQuery(up1);
                    }


                }
            }
        }

        private void BundleGeneration_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       
        public void News()
        {

           empty();
            compload(); txtsearch.Select();
        }
        private void empty()
        {
            txtbunid.Text = ""; txtorderqty.Text = "";
            txtbunid1.Text = ""; combopono.Text = ""; combolayno.Text = ""; combocompname.Text = "";
            dataGridView2.Rows.Clear(); dataGridView3.Rows.Clear(); dataGridView2.Columns.Clear(); dataGridView3.Columns.Clear();
            combocompcode.Text = "";
            combostyle.Text = "";
            combobuyer.Text = "";
            txtbunid1.Text = "";
            checkactive.Checked = true;
            txtlotno.Text = ""; txtcutno.Text = "";
            checklaycancel.Checked = false;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            GlobalVariables.HideCols = new string[] { "ASPTBLSIZMASID", "ASPTBLPURDETID", "ASPTBLPURID", "ASPTBLPURID1", "ASPTBLBUNDETID", "ASPTBLBUNID", "ASPTBLBUNID1", "COMPCODE" };
                panel3.BackColor = Class.Users.BackColors;
            GlobalVariables.WidthCols = new Int32[] { 50, 50, 50, 50,50, 250, 300 };
            GlobalVariables.New_Flg = false;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            dataGridView3.Font = Class.Users.FontName;
            dataGridView2.Font = Class.Users.FontName;
            mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
            Class.Users.Query = ""; Class.Users.TableNameGrid = "";
            sel4 = "select  h.ASPTBLSIZMASID, c.asptblpurdetid,a.asptblpurid,a.asptblpurid1,b.compcode, d.layno, f.FABRIC,g.colorname,h.SIZENAME,e.markerno,A.orderqty,''AS NOTES from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptbllay d on d.pono=a.pono   join asptbllaydet e on e.asptbllayid=d.asptbllayid and e.asptbllayid1=d.asptbllayid1  and e.colorname=c.colorname and e.sizename=c.sizename  join asptblfabmas f on f.asptblfabmasid=e.FABRIC join asptblcolmas g on g.asptblcolmasid=e.colorname join asptblsizmas h on h.ASPTBLSIZMASID=e.sizename   order by 1";
            sel5 = "select  h.ASPTBLSIZMASID, '' as asptblbundetid,''asptblbunid,'' as asptblbunid1,b.compcode, d.layno, f.FABRIC,g.colorname,h.SIZENAME,e.markerno,''AS BUNDLENO,'' FROMPLAY,'' TOPLAY,'' ORDERQTY,'' as NOTES from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptbllay d on d.pono=a.pono   join asptbllaydet e on e.asptbllayid=d.asptbllayid and e.asptbllayid1=d.asptbllayid1  and e.colorname=c.colorname and e.sizename=c.sizename  join asptblfabmas f on f.asptblfabmasid=e.FABRIC join asptblcolmas g on g.asptblcolmasid=e.colorname join asptblsizmas h on h.ASPTBLSIZMASID=e.sizename   order by 1";

            Class.Users.Query = sel4;
            Class.Users.TableNameGrid = "asptblbun";


            CommonFunctions.AddGridColumn(dataGridView3, sel4, GlobalVariables.HideCols,GlobalVariables.WidthCols, "asptblbun");
            CommonFunctions.AddGridColumn(dataGridView2, sel5, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblbun");
            DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
            checkBoxColumn.HeaderText = "Check";
            checkBoxColumn.Name = "Check";
            checkBoxColumn.Width = 50;
            dataGridView3.Columns.Insert(0, checkBoxColumn);
    
        }

        public void colorid(string s)
        {
            try
            {
                string sel = "select asptblcolmasid from  asptblcolmas where colorname='" + s + "' ;";
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
                string sel = "select asptblfabmasid from  asptblfabmas where fabric='" + text + "' ;";
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
                string sel1 = "SELECT a.asptblbunid,b.compcode,c.buyername,a.pono,a.orderqty,d.stylename,a.lotno,a.active FROM  asptblbun a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename    order by  a.asptblbunid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbun");
                DataTable dt = ds.Tables["asptblbun"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblbunid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["buyername"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());                      
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["lotno"].ToString());
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
                    txtbunid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblbunid,a.docid,a.finyear,a.docdate,a.layno,b.compcode,b.compname,e.buyername,a.pono,a.orderqty, c.stylename,a.lotno,a.cutno, a.cutcancel,a.active  from  asptblbun a  join gtcompmast b on b.gtcompmastid=a.compcode   join asptblstymas c on c.asptblstymasid=a.stylename    join asptblbuymas e on e.asptblbuymasid=a.buyer   where a.asptblbunid=" + txtbunid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbun");
                    DataTable dt = ds.Tables["asptblbun"];
                    if (dt.Rows.Count > 0)
                    {
                        txtbunid.Text = Convert.ToString(dt.Rows[0]["asptblbunid"].ToString());
                        txtdocid.Text = Convert.ToString(dt.Rows[0]["docid"].ToString());
                        dateTimePicker1.Value = Convert.ToDateTime(dt.Rows[0]["docdate"].ToString().Substring(0,10));
                        combolayno.Text = Convert.ToString(dt.Rows[0]["layno"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());                      
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyername"].ToString());
                        combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        layno(combocompcode.Text, combopono.Text);
                       
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        txtorderqty.Text = Convert.ToString(dt.Rows[0]["orderqty"].ToString());                       
                        combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                       
                        txtcutno.Text = Convert.ToString(dt.Rows[0]["cutno"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        if (dt.Rows[0]["cutcancel"].ToString() == "T") { checklaycancel.Checked = true; } else { checklaycancel.Checked = true; checklaycancel.Checked = false; }


                       
                    }
                    string sel2 = "select a.asptblbundetid,a.asptblbunid,a.asptblbunid1,c.compcode,a.pono,a.layno, f.fabric, d.colorname,e.sizename,a.markerno,a.bundleno,a.fromply,a.toply, a.cutqty,a.notes from asptblbundet a join asptblbun b on a.asptblbunid=b.asptblbunid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizmas e on e.asptblsizmasid=a.sizename join asptblfabmas f on f.ASPTBLFABMASID=a.fabric where a.asptblbunid=" + txtbunid.Text;
                    dataGridView2.Rows.Clear();
                    CommonFunctions.AddColumnDynamic1(dataGridView2, sel2);
                   // string sel2 = "select a.asptblbundetid,a.asptblbunid,a.asptblbunid1,c.compcode,a.pono,a.layno, f.fabric, d.colorname,e.sizename,a.markerno,a.bundleno,a.fromply,a.toply, a.cutqty,a.notes from asptblbundet a join asptblbun b on a.asptblbunid=b.asptblbunid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizmas e on e.asptblsizmasid=a.sizename join asptblfabmas f on f.ASPTBLFABMASID=a.fabric where a.asptblbunid=" + txtbunid.Text;
                   // DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblbun");
                   // DataTable dt2 = ds2.Tables["asptblbun"];
                    
                    //if (dt2.Rows.Count > 0)
                    //{
                    //    int i = 0;
                    //    for (i= 0;i < dt2.Rows.Count;i++)
                    //    {

                    //        dataGridView2.Rows.Add();
                    //        dataGridView2.Rows[i].Cells[0].Value = dt2.Rows[i]["asptblbundetid"].ToString();
                    //        dataGridView2.Rows[i].Cells[1].Value = dt2.Rows[i]["asptblbunid"].ToString();
                    //        dataGridView2.Rows[i].Cells[2].Value = dt2.Rows[i]["asptblbunid1"].ToString();
                    //        dataGridView2.Rows[i].Cells[3].Value = dt2.Rows[i]["compcode"].ToString();
                    //        dataGridView2.Rows[i].Cells[4].Value = dt2.Rows[i]["pono"].ToString();
                    //        dataGridView2.Rows[i].Cells[5].Value = dt2.Rows[i]["layno"].ToString();
                    //        dataGridView2.Rows[i].Cells[6].Value = dt2.Rows[i]["fabric"].ToString();
                    //        dataGridView2.Rows[i].Cells[7].Value = dt2.Rows[i]["colorname"].ToString();
                    //        dataGridView2.Rows[i].Cells[8].Value = dt2.Rows[i]["sizename"].ToString();
                    //        dataGridView2.Rows[i].Cells[9].Value = dt2.Rows[i]["markerno"].ToString();
                    //        dataGridView2.Rows[i].Cells[10].Value = dt2.Rows[i]["bundleno"].ToString();
                    //        dataGridView2.Rows[i].Cells[11].Value = dt2.Rows[i]["fromply"].ToString();
                    //        dataGridView2.Rows[i].Cells[12].Value = dt2.Rows[i]["toply"].ToString();
                    //        dataGridView2.Rows[i].Cells[13].Value = dt2.Rows[i]["cutqty"].ToString();
                    //        dataGridView2.Rows[i].Cells[14].Value = dt2.Rows[i]["notes"].ToString();
                            
                    //    }
                    //    CommonFunctions.SetRowNumber(dataGridView2);

                    //}
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
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
                        if (item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text) || item.SubItems[6].ToString().Contains(txtsearch.Text) || item.SubItems[8].ToString().Contains(txtsearch.Text))
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
            if (txtbunid.Text != "")
            {
                string sel1 = "select a.asptblbunid from asptblbun a join gtstatemast b on a.asptblbunid=b.country where a.asptblbunid='" + txtbunid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbun");
                DataTable dt = ds.Tables["asptblbun"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtbunid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtbunid.Text != "")
                    {
                        string del = "delete from asptblbun where asptblbunid=" + txtbunid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptblbundet where asptblbundetid=" + txtbunid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + txtbunid1.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtbunid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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
            //if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            //{

            //    if (txtbunid.Text == "")
            //    {
            //        News();
                   
            //    }
            //    combocompcode.Select();
            //}
            //if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            //{
            //    txtsearch.Select();

            //}
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
        public void layno(string s,string ss)
        {
            if (s != "" && ss != "")
            {
                 Utility.Load_Combo(combolayno, "select a.asptbllayid, a.docid from asptbllay a join gtcompmast b on a.compcode=b.gtcompmastid   where  b.compcode='" + s + "' and a.pono='" + ss + "' order by a.docid desc", "asptbllayid", "docid");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {          

            GlobalVariables.WidthCols = new Int32[] { 50, 50, 50, 50, 50, 250, 350, 350 };

            CommonFunctions.AddColumnDynamic1(dataGridView3, sel4);
                
            
        }
        public void layno(string s, string ss,string sss)
        {
            if (s != "" && ss != "" && sss != "")
            {

                string sel = "select distinct e.asptblstymasid,e.stylename,f.asptblbuymasid,f.buyercode,a.orderqty,h.layno,h.lotno from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblstymas e on e.asptblstymasid=a.stylename join asptblbuymas f on f.asptblbuymasid=a.buyer join asptblprolot g on g.pono=a.pono and g.compcode=b.gtcompmastid  join asptbllay h on h.pono=g.pono and h.compcode = g.compcode and h.compcode = a.compcode and h.compcode = b.gtcompmastid  and h.pono   where  b.compcode='" + combocompcode.Text + "' and a.pono='" + combopono.Text + "' and h.docid='" + combolayno.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpur");
                DataTable dt3 = ds.Tables["asptblpur"];
                if (dt3.Rows.Count > 0)
                {
                    combostyle.DisplayMember = "stylename";
                    combostyle.ValueMember = "asptblstymasid";
                    combostyle.DataSource = dt3;

                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.ValueMember = "asptblbuymasid";
                    combobuyer.DataSource = dt3;
                    txtorderqty.Text = dt3.Rows[0]["orderqty"].ToString();

                    txtlotno.Text = dt3.Rows[0]["lotno"].ToString();
            
                }
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

            allip1.Items.Clear(); allip2.Items.Clear();
            if (combocompcode.Text != "" && dataGridView3.Columns.Count>0)
            {
                Utility.Load_Combo(combopono, "select A.PONO from asptbllay a join gtcompmast b on a.compcode=b.gtcompmastid   where  b.compcode='" + combocompcode.Text + "'  order by 1", "PONO", "PONO");


                if (combocompcode.Text != "" && txtbunid.Text == "")
                {
                    DataTable dt1 = autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName);
                    if (dt1.Rows.Count > 0)
                    {
                        combocompname.Text = dt1.Rows[0]["COMPNAME"].ToString();
                        txtbunid1.Text = dt1.Rows[0]["id"].ToString();
                        txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
                        txtdocid.Text = Class.Users.Finyear + "-" + dt1.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
                    }
                    else
                    {
                        combopono.DataSource = null;
                        txtbunid1.Text = "";
                        txtshortcode.Text = "";
                        txtdocid.Text = "";

                    }

                }
            }
        }

        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopono.SelectedIndex != -1 && combocompcode.Text != "")
            {
                layno(combocompcode.Text, combopono.Text);

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

        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            //if (allip3delete.Items.Count > 0)
            //{
            //    foreach (ListViewItem item in allip3delete.Items)
            //    {
            //        //string sel1 = "select a.asptblbundetid from asptblbundet a join GTCOMPMAST b on a.compcode=b.gtcompmastid where a.asptblbundetid='" + item.SubItems[1].Text + "'";
            //        //DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblbundet");
            //        //DataTable dt = ds.Tables["asptblbundet"];
            //        //if (dt.Rows.Count > 0)
            //        //{
            //        //    MessageBox.Show("Child Record Found.Can Not Delete." + combopono.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            //        //}
            //        //else
            //        //{

            //            string del = "delete from asptblbundet where asptblbundetid='" + Convert.ToInt64("0" + item.SubItems[1].Text) + "'";
            //            Utility.ExecuteNonQuery(del);

            //        MessageBox.Show("Record Deleted Successfully " + item.SubItems[1].Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //        //}
            //    }
            //}
        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
     
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
           
            //if (dataGridView1.Columns[e.ColumnIndex].Name == "Notes")
            //{
            //    CommonFunctions.SetRowNumber(dataGridView1);
            //}
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void combolayno_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combolayno.Text != "")
            {
                  layno(combocompcode.Text, combopono.Text, combolayno.Text);
                button2_Click(sender,e);
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void comboshiftname_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void txtcutno_TextChanged(object sender, EventArgs e)
        {

        }

        private void butbundlegenerate_Click(object sender, EventArgs e)
        {



            //if (dataGridView2.Columns.Count > 1)
            //{
            //    CommonFunctions.AddColumnDynamic1(dataGridView2, Class.Users.Query);
            //}



            DataTable dt5 = Utility.SQLQuery("select  h.ASPTBLSIZMASID, '0' as asptblbundetid,'" + Convert.ToUInt64("0" + txtbunid1.Text) + "' as asptblbunid,'" + Convert.ToUInt64("0" + txtbunid.Text) + "' as asptblbunid1,b.compcode, d.layno, f.FABRIC,g.colorname,h.SIZENAME,e.markerno,''AS BUNDLENO,'' FROMPLAY,'' TOPLAY,'' ORDERQTY,'' notes from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptbllay d on d.pono=a.pono   join asptbllaydet e on e.asptbllayid=d.asptbllayid and e.asptbllayid1=d.asptbllayid1  and e.colorname=c.colorname and e.sizename=c.sizename  join asptblfabmas f on f.asptblfabmasid=e.FABRIC join asptblcolmas g on g.asptblcolmasid=e.colorname join asptblsizmas h on h.ASPTBLSIZMASID=e.sizename  where  b.compcode='" + combocompcode.Text + "' and a.pono='" + combopono.Text + "' order by 1");
            decimal totcount = 0;
            Int32 ordqty = 0;//=Convert.ToInt32("0" + txtorderqty.Text);
            Int32 bunqty=Convert.ToInt32("0" + txtbundelsizes.Text);
            //totcount = ordqty / bunqty;
            decimal tot;//= totcount * bunqty;
            decimal lasttot;// = ordqty - tot;
            int fromplay = 275; int tolay = 350; string message = string.Empty;
            if (dataGridView2.Columns.Count > 1)
            {
                dataGridView2.Rows.Clear(); int i = 0, h = 0, cnt = 0;
               
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    bool isSelected = Convert.ToBoolean(row.Cells["Check"].Value);
                    if (isSelected)
                    {

                        totcount = 0; tot = 0; lasttot = 0;
                        ordqty = Convert.ToInt32("0" + row.Cells[11].Value.ToString());
                        totcount = ordqty / bunqty;
                        tot = totcount * bunqty;
                        lasttot = ordqty - tot;
                        for (i = 0; i <= totcount; i++)
                        {
                            dataGridView2.Rows.Add();
                            dataGridView2.Rows[cnt].Cells[1].Value = dt5.Rows[h]["asptblbundetid"].ToString();
                            dataGridView2.Rows[cnt].Cells[2].Value = dt5.Rows[h]["asptblbunid"].ToString();
                            dataGridView2.Rows[cnt].Cells[3].Value = dt5.Rows[h]["asptblbunid1"].ToString();
                            dataGridView2.Rows[cnt].Cells[4].Value = dt5.Rows[h]["compcode"].ToString();
                            dataGridView2.Rows[cnt].Cells[5].ReadOnly = false;
                            dataGridView2.Rows[cnt].Cells[5].Value = row.Cells[6].EditedFormattedValue.ToString();
                            dataGridView2.Rows[cnt].Cells[6].ReadOnly = false;
                            dataGridView2.Rows[cnt].Cells[6].Value = row.Cells[7].EditedFormattedValue.ToString();
                            dataGridView2.Rows[cnt].Cells[7].ReadOnly = false;
                            dataGridView2.Rows[cnt].Cells[7].Value = row.Cells[8].EditedFormattedValue.ToString();
                            dataGridView2.Rows[cnt].Cells[8].ReadOnly = false;
                            dataGridView2.Rows[cnt].Cells[8].Value = row.Cells[9].EditedFormattedValue.ToString();
                            dataGridView2.Rows[cnt].Cells[9].ReadOnly = false;
                            dataGridView2.Rows[cnt].Cells[9].Value = row.Cells[10].EditedFormattedValue.ToString();
                            dataGridView2.Rows[cnt].Cells[10].Value = txtcutno.Text.ToString();
                            dataGridView2.Rows[cnt].Cells[11].Value = fromplay.ToString();
                            dataGridView2.Rows[cnt].Cells[12].Value = tolay.ToString();     
                            if (i == totcount && lasttot >= 1)
                            {

                                dataGridView2.Rows[cnt].Cells[13].Value = lasttot;
                                dataGridView2.Rows[cnt].DefaultCellStyle.Font = new Font(Class.Users.FontName.FontFamily, Class.Users.FontName.Size, FontStyle.Bold);
                            }
                            else
                            {

                                dataGridView2.Rows[cnt].Cells[13].Value = txtbundelsizes.Text;
                            }

                            cnt++;
                        }
                        h++;
                    }
                }
            }
           
            CommonFunctions.SetRowNumber(dataGridView2);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (combolayno.Text != "")
            {
                Transactions.PopUpWindow pop = new PopUpWindow();
                Class.Users.GridID = Convert.ToString(txtbunid.Text);
                Class.Users.NonGridID = Convert.ToString(txtbunid1.Text);
                Class.Users.CompCode1 = Convert.ToString(combocompcode.Text);
                Class.Users.PoNo = Convert.ToString(combopono.Text);
                Class.Users.DocID = Convert.ToString(combolayno.Text);
                pop.Show();
            }
        }

  

     

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }
    }
}

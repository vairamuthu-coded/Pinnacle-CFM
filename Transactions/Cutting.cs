using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Transactions
{
    public partial class Cutting : Form,ToolStripAccess
    {
        public Cutting()
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
            panel6.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;           
        }

        private static Cutting _instance; string coid = "", siid = "", fabid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        public static Cutting Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new Cutting();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        private DataTable autonumberload(string y, string com, string scr)
        {
            DataTable dt1 = new DataTable();
            string sel1 = "select max(a.asptblcutid1)+1 as id,a.shortcode,b.compname from asptblcut a join gtcompmast b on a.compcode = b.gtcompmastid  where a.finyear='" + y + "' and b.compcode='" + com + "' group by a.shortcode,b.compname";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblcut");
            dt1 = ds1.Tables["asptblcut"];
            if (dt1.Rows.Count <= 0)
            {
                sel1 = "";
                sel1 = "select max(a.sequenceno)+1 as id,a.shortcode,b.compname from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptbluserrights c on c.userrightsid=a.screen where a.finyear='" + y + "' and b.compcode='" + com + "' AND C.MENUNAME='" + scr + "' group by a.shortcode,b.compname";
                ds1 = Utility.ExecuteSelectQuery(sel1, "asptblautogeneratemas");
                dt1 = ds1.Tables["asptblautogeneratemas"];
            }

            return dt1;
        }
        private void Cutting_Load(object sender, EventArgs e)
        {
             GridLoad();compload(); txtsearch.Select();
            Utility.Load_Combo(comboshiftname, "select asptblshitypeid,shiftno from asptblshitype where active='T' order by 2", "asptblshitypeid", "shiftno");
           
        }
        string maxid = "";
        public void Saves()
        {
            try
            {



                Models.CuttingModel p = new Models.CuttingModel();
                p.asptblcutid = Convert.ToInt64("0" + txtcutid.Text);
                p.asptblcutid1 = Convert.ToInt64("0" + txtcutid1.Text);
                p.docid = Convert.ToString(txtdocid.Text);
                p.shortcode = Convert.ToString(txtshortcode.Text);
                p.finyear = Class.Users.Finyear;
                p.docdate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                p.pono = Convert.ToString(combopono.Text);
                p.compname = Convert.ToInt64(combocompcode.SelectedValue);
                p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
                p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                p.layno = Convert.ToString(combolayno.Text);
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
                if (checklaycancel.Checked == true) { p.cutcancel = "T"; } else { p.cutcancel = "F"; checklaycancel.Checked = false; }
                if (GlobalVariables.New_Flg == false)
                {

                    string ins = "insert into asptblcut(asptblcutid1,shortcode,docid,finyear,docdate,compcode,buyer,pono,orderqty,stylename,lotno,layno,cutno,shiftno,cutcancel,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.asptblcutid1 + "','" + p.shortcode + "','" + p.docid + "','" + p.finyear + "','" + p.docdate + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.lotno + "','" + p.layno + "' ,'" + p.cutno + "' ,'" + p.shiftno + "','" + p.cutcancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + p.modified + "','" + p.modifiedby + "','" + p.ipaddress + "');";
                    Utility.ExecuteNonQuery(ins);
                    GridViewLoad();
                }
                if (GlobalVariables.New_Flg == true)
                {
                    
                    string up = "update  asptblcut   set  asptblcutid1='" + p.asptblcutid1 + "',shortcode='"+p.shortcode+ "',docid='" + p.docid + "' ,finyear='" + p.finyear + "' ,  docdate='" + p.docdate + "' , compcode ='" + p.compcode + "' ,buyer='" + p.buyer + "' , pono='" + p.pono + "' ,  orderqty ='" + p.orderqty + "' , stylename='" + p.stylename + "' ,lotno='" + p.lotno + "', layno='" + p.layno + "',cutno='" + p.cutno + "'  ,shiftno='" + p.shiftno + "', laycancel='" + p.cutcancel + "' , active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',modified='" + Convert.ToDateTime(p.modified).ToString("yyyy-MM-dd hh:mm:ss") + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblcutid='" + txtcutid.Text + "'";
                    Utility.ExecuteNonQuery(up);
                    GridViewLoad();
                }
                if (txtcutid.Text == "")
                {
                    MessageBox.Show("Record Saved Successfully " + txtcutid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tabControl1.SelectTab(tabPage2);
                }
                else
                {
                    MessageBox.Show("Record Updated Successfully " + txtcutid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            Models.CuttingModel.CuttingModeldet pp = new Models.CuttingModel.CuttingModeldet();
            int i = 0;

            if (dataGridView1.Rows.Count >= 0)
            {
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    pp.asptblcutdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value);
                    pp.asptblcutid = Convert.ToInt64("0" + txtcutid.Text);
                    pp.asptblcutid1 = Convert.ToInt64("0" + txtcutid1.Text);
                    pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);                   
                    pp.pono = Convert.ToString(combopono.Text);
                    fabricid(dataGridView1.Rows[i].Cells[5].Value.ToString());
                    colorid(dataGridView1.Rows[i].Cells[6].Value.ToString());
                    sizeid(dataGridView1.Rows[i].Cells[7].Value.ToString());
                    pp.colorname = coid;
                    pp.sizename = siid;
                    pp.fabric = fabid;
                    pp.markerno = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value.ToString());
                    pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[9].Value.ToString());
                    pp.notes = Convert.ToString(dataGridView1.Rows[i].Cells[10].Value.ToString());

                    string sel = "select  a.asptblcutdetid   from  asptblcutdet  a where a.asptblcutid1='" + pp.asptblcutid1 + "'  and  a.compcode='" + pp.compcode + "'   and a.pono='" + pp.pono + "' and a.fabric='" + pp.fabric + "' and a.colorname='" + pp.colorname + "' and a.sizename='" + pp.sizename + "'  and a.markerno='" + pp.markerno + "'  and a.orderqty='" + pp.orderqty + "' and a.notes='" + pp.notes + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllay");
                    DataTable dt = ds.Tables["asptbllay"];
                    if (dt.Rows.Count != 0)
                    {

                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt64("0" + pp.asptblcutdetid) == 0 || Convert.ToInt64("0" + pp.asptblcutdetid) == 0)
                    {
                        string sel2 = "select max(asptblcutid) id,asptblcutid1    from  asptblcut   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and docid='" + txtdocid.Text + "' group by asptblcutid1";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblcut");
                        DataTable dt2 = ds2.Tables["asptblcut"];
                        if (dt2 == null) { }
                        {
                            pp.asptblcutid = Convert.ToInt64("0" + dt2.Rows[0]["id"].ToString());
                            pp.asptblcutid1 = Convert.ToInt64("0" + dt2.Rows[0]["asptblcutid1"].ToString());
                        }
                        string ins1 = "insert into asptblcutdet(asptblcutid,asptblcutid1,compcode,pono,colorname,sizename,fabric,markerno,orderqty,notes) values('" + pp.asptblcutid + "' ,'" + pp.asptblcutid1 + "' ,'" + pp.compcode + "' ,'" + pp.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.fabric + "' ,'" + pp.markerno + "','" + pp.orderqty + "','" + pp.notes + "' );";
                        Utility.ExecuteNonQuery(ins1);
                    }
                    else
                    {
                        string up1 = "update  asptblcutdet  set asptblcutid='" + pp.asptblcutid + "' ,asptblcutid1='" + pp.asptblcutid1 + "' , compcode='" + pp.compcode + "' , pono='" + pp.pono + "' , colorname='" + pp.colorname + "',sizename='" + pp.sizename + "' ,fabric='" + pp.fabric + "' , markerno='" + pp.markerno + "',orderqty='" + pp.orderqty + "' ,notes='" + pp.notes + "' where asptblcutdetid='" + pp.asptblcutdetid + "'";
                        Utility.ExecuteNonQuery(up1);
                    }


                }
            }
        }

        private void Cutting_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       
        public void News()
        {

           empty();
        }
        string sel4 = "";
        private void empty()
        {
            txtcutid.Text = ""; txtorderqty.Text = "";
            txtcutid1.Text = ""; combopono.Text = "";combolayno.Text = ""; combocompname.Text = "";
            comboshiftname.Text = "";

            combocompcode.Text = "";
            combostyle.Text = "";
            combobuyer.Text = "";
            txtcutid1.Text = "";
            checkactive.Checked = true;

            txtlotno.Text = ""; txtcutno.Text = "";
            checklaycancel.Checked = false; 
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            dataGridView1.Rows.Clear();
            GlobalVariables.HideCols = new string[] { "asptblcutdetid", "asptblcutid","asptblcutid1", "compcode" };
                  
            panel3.BackColor = Class.Users.BackColors;           
            panel6.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            GlobalVariables.New_Flg = false;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            dataGridView1.Font = Class.Users.FontName;
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 0, 100,250,250 };
            Class.Users.TableNameGrid = "asptblbun";
            sel4 = "select  h.ASPTBLSIZMASID, '0' as asptblcutdetid,'" + Convert.ToUInt64("0" + txtcutid.Text) + "' as asptblcutid,'" + Convert.ToUInt64("0" + txtcutid1.Text) + "' as asptblcutid1,b.compcode, d.layno as layno1, f.FABRIC,g.colorname,h.SIZENAME,e.markerno,'' cutqty,'' notes from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptbllay d on d.pono=a.pono   join asptbllaydet e on e.asptbllayid=d.asptbllayid and e.asptbllayid1=d.asptbllayid1  and e.colorname=c.colorname and e.sizename=c.sizename  join asptblfabmas f on f.asptblfabmasid=e.FABRIC join asptblcolmas g on g.asptblcolmasid=e.colorname join asptblsizmas h on h.ASPTBLSIZMASID=e.sizename  where  b.compcode='" + combocompcode.Text + "' and a.pono='" + combopono.Text + "' and d.docid='" + txtdocid.Text + "' order by 1";

            CommonFunctions.AddGridColumn(dataGridView1, sel4, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblbun");
            CommonFunctions.AddGridColumn(dataGridView2, sel4, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblbun");
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
                string sel1 = "SELECT a.asptblcutid,b.compcode,c.buyername,a.pono,a.orderqty,d.stylename,a.lotno,a.active FROM  asptblcut a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename    order by  a.asptblcutid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcut");
                DataTable dt = ds.Tables["asptblcut"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblcutid"].ToString());
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
                    txtcutid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblcutid,a.docid,a.finyear,a.docdate,a.layno,b.compcode,b.compname,e.buyername,a.pono,a.orderqty, c.stylename,a.lotno, d.shiftno,a.cutno, a.cutcancel,a.active  from  asptblcut a  join gtcompmast b on b.gtcompmastid=a.compcode   join asptblstymas c on c.asptblstymasid=a.stylename  join asptblshitype d on d.asptblshitypeid=a.shiftno   join asptblbuymas e on e.asptblbuymasid=a.buyer   where a.asptblcutid=" + txtcutid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcut");
                    DataTable dt = ds.Tables["asptblcut"];
                    if (dt.Rows.Count > 0)
                    {
                        txtcutid.Text = Convert.ToString(dt.Rows[0]["asptblcutid"].ToString());
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
                        comboshiftname.Text = Convert.ToString(dt.Rows[0]["shiftno"].ToString());
                        txtcutno.Text = Convert.ToString(dt.Rows[0]["cutno"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        if (dt.Rows[0]["cutcancel"].ToString() == "T") { checklaycancel.Checked = true; } else { checklaycancel.Checked = true; checklaycancel.Checked = false; }


                       
                    }
                    string sel2 = "select a.asptblcutdetid,a.asptblcutid,a.asptblcutid1,c.compcode,a.pono, f.fabric, d.colorname,e.sizename,a.markerno,a.orderqty,a.notes from asptblcutdet a join asptblcut b on a.asptblcutid=b.asptblcutid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizmas e on e.asptblsizmasid=a.sizename join asptblfabmas f on f.ASPTBLFABMASID=a.fabric where a.asptblcutid=" + txtcutid.Text;
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblcut");
                    DataTable dt2 = ds2.Tables["asptblcut"];
                    dataGridView1.Rows.Clear();
                    if (dt2.Rows.Count > 0)
                    {
                        int i = 0;
                        for (i= 0;i < dt2.Rows.Count;i++)
                        {

                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[i].Cells[1].Value = dt2.Rows[i]["asptblcutdetid"].ToString();
                            dataGridView1.Rows[i].Cells[2].Value = dt2.Rows[i]["asptblcutid"].ToString();
                            dataGridView1.Rows[i].Cells[3].Value = dt2.Rows[i]["asptblcutid1"].ToString();
                            dataGridView1.Rows[i].Cells[4].Value = dt2.Rows[i]["compcode"].ToString();
                            dataGridView1.Rows[i].Cells[5].Value = dt2.Rows[i]["pono"].ToString();
                            dataGridView1.Rows[i].Cells[6].Value = dt2.Rows[i]["fabric"].ToString();
                            dataGridView1.Rows[i].Cells[7].Value = dt2.Rows[i]["colorname"].ToString();
                            dataGridView1.Rows[i].Cells[8].Value = dt2.Rows[i]["sizename"].ToString();
                            dataGridView1.Rows[i].Cells[9].Value = dt2.Rows[i]["markerno"].ToString();
                            dataGridView1.Rows[i].Cells[10].Value = dt2.Rows[i]["orderqty"].ToString();
                            dataGridView1.Rows[i].Cells[11].Value = dt2.Rows[i]["notes"].ToString();
                            
                        }
                        CommonFunctions.SetRowNumber(dataGridView1);

                    }
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
            //try
            //{
            //    if (txtsearch.Text.ToUpper() != "")
            //    {
            //        listView1.Items.Clear(); int iGLCount = 1;
            //        string sel1 = "  SELECT  a.asptblcutid,a.colorname,a.active from asptblcut a  where a.colorname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcut");
            //        DataTable dt = ds.Tables["asptblcut"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptblcutid"].ToString());
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
            if (txtcutid.Text != "")
            {
                string sel1 = "select a.asptblcutid from asptblcut a join gtstatemast b on a.asptblcutid=b.country where a.asptblcutid='" + txtcutid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcut");
                DataTable dt = ds.Tables["asptblcut"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcutid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtcutid.Text != "")
                    {
                        string del = "delete from asptblcut where asptblcutid=" + txtcutid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptblcutdet where asptblcutdetid=" + txtcutid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + txtcutid1.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtcutid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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

                if (txtcutid.Text == "")
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
        public void layno(string s,string ss)
        {
            if (s != "" && ss != "")
            {
                 Utility.Load_Combo(combolayno,"select  a.docid from asptbllay a join gtcompmast b on a.compcode=b.gtcompmastid   where  b.compcode='" + s + "' and a.pono='" + ss + "' order by a.docid desc", "docid", "docid");
            }
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
                    if (dataGridView1.Columns.Count > 1)
                    {
                       

                        GlobalVariables.WidthCols = new Int32[] { 50, 50, 50, 50, 50, 250, 350, 350 };
                        sel4 = "select  h.ASPTBLSIZMASID, '0' as asptblcutdetid,'" + Convert.ToUInt64("0" + txtcutid.Text) + "' as asptblcutid,'" + Convert.ToUInt64("0" + txtcutid1.Text) + "' as asptblcutid1,b.compcode, d.layno as layno1, f.FABRIC,g.colorname,h.SIZENAME,e.markerno,'' cutqty,'' notes from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptbllay d on d.pono=a.pono   join asptbllaydet e on e.asptbllayid=d.asptbllayid and e.asptbllayid1=d.asptbllayid1  and e.colorname=c.colorname and e.sizename=c.sizename  join asptblfabmas f on f.asptblfabmasid=e.FABRIC join asptblcolmas g on g.asptblcolmasid=e.colorname join asptblsizmas h on h.ASPTBLSIZMASID=e.sizename  where  b.compcode='" + s + "' and a.pono='" + ss + "' and d.docid='" + sss + "' order by 1";
                        CommonFunctions.AddColumnDynamic1(dataGridView1, sel4);
                    }
                    //DataTable dt4 = Utility.SQLQuery("select  h.ASPTBLSIZMASID, '0' as asptblcutdetid,'" + Convert.ToUInt64("0" + txtcutid.Text) + "' as asptblcutid,'" + Convert.ToUInt64("0" + txtcutid1.Text) + "' as asptblcutid1,b.compcode, d.layno as layno1, f.FABRIC,g.colorname,h.SIZENAME,e.markerno,'' cutqty,'' notes from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptbllay d on d.pono=a.pono   join asptbllaydet e on e.asptbllayid=d.asptbllayid and e.asptbllayid1=d.asptbllayid1  and e.colorname=c.colorname and e.sizename=c.sizename  join asptblfabmas f on f.asptblfabmasid=e.FABRIC join asptblcolmas g on g.asptblcolmasid=e.colorname join asptblsizmas h on h.ASPTBLSIZMASID=e.sizename  where  b.compcode='" + s + "' and a.pono='" + ss + "' and d.docid='" + sss + "' order by 1");
                    //if (dt4.Rows.Count > 0)
                    //{
                    //    dataGridView1.Rows.Clear();
                    //    dataGridView2.DataSource = dt4;
                    //    for (int i = 0; i < dt4.Rows.Count; i++)
                    //    {
                    //        dataGridView1.Rows.Add();
                    //        dataGridView1.Rows[i].Cells[0].Value = dt4.Rows[i]["asptblcutdetid"].ToString();
                    //        dataGridView1.Rows[i].Cells[1].Value = dt4.Rows[i]["asptblcutid"].ToString();
                    //        dataGridView1.Rows[i].Cells[2].Value = dt4.Rows[i]["asptblcutid1"].ToString();
                    //        dataGridView1.Rows[i].Cells[3].Value = dt4.Rows[i]["compcode"].ToString();
                    //        dataGridView1.Rows[i].Cells[4].Value = dt4.Rows[i]["layno1"].ToString();
                    //        dataGridView1.Rows[i].Cells[5].Value = dt4.Rows[i]["FABRIC"].ToString();
                    //        dataGridView1.Rows[i].Cells[6].Value = dt4.Rows[i]["colorname"].ToString();
                    //        dataGridView1.Rows[i].Cells[7].Value = dt4.Rows[i]["SIZENAME"].ToString();
                    //        dataGridView1.Rows[i].Cells[8].Value = dt4.Rows[i]["markerno"].ToString();
                    //        dataGridView1.Rows[i].Cells[9].Value = dt4.Rows[i]["cutqty"].ToString();
                    //        dataGridView1.Rows[i].Cells[10].Value = dt4.Rows[i]["Notes"].ToString();
                    //    }
                    //    CommonFunctions.SetRowNumber(dataGridView1);
                    //}
                    //else
                    //{
                    //    do
                    //    {
                    //        for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
                    //    }
                    //    while (dataGridView1.Rows.Count > 1);


                    //}
                }
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

            allip1.Items.Clear(); allip2.Items.Clear();
            if (combocompcode.SelectedIndex != -1)
            {
                Utility.Load_Combo(combopono, "select A.PONO from asptbllay a join gtcompmast b on a.compcode=b.gtcompmastid   where  b.compcode='" + combocompcode.Text + "'  order by 1", "PONO", "PONO");

                //DataTable dt = Utility.SQLQuery("select B.COMPNAME  from  join gtcompmast b where b.compcode='" + combocompcode.Text + "'");
                //if (dt.Rows.Count > 0)
                //{

                //    combocompname.Text = dt.Rows[0]["COMPNAME"].ToString();
                if (combocompcode.Text != "" && txtcutid.Text == "")
                {
                    DataTable dt1 = autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName);
                    if (dt1.Rows.Count > 0)
                    {
                        combocompname.Text = dt1.Rows[0]["COMPNAME"].ToString();
                        txtcutid1.Text = dt1.Rows[0]["id"].ToString();
                        txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
                        txtdocid.Text = Class.Users.Finyear + "-" + dt1.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
                    }
                }
                //}
                // else
                //{
                //    combopono.DataSource = null;
                //    txtcutid1.Text = "";
                //    txtshortcode.Text = "";
                //    txtdocid.Text = "";

                //}

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
            //        //string sel1 = "select a.asptblcutdetid from asptblcutdet a join GTCOMPMAST b on a.compcode=b.gtcompmastid where a.asptblcutdetid='" + item.SubItems[1].Text + "'";
            //        //DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcutdet");
            //        //DataTable dt = ds.Tables["asptblcutdet"];
            //        //if (dt.Rows.Count > 0)
            //        //{
            //        //    MessageBox.Show("Child Record Found.Can Not Delete." + combopono.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            //        //}
            //        //else
            //        //{

            //            string del = "delete from asptblcutdet where asptblcutdetid='" + Convert.ToInt64("0" + item.SubItems[1].Text) + "'";
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
                
               
            }
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

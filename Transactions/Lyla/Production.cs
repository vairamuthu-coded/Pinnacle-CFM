using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace Pinnacle.Transactions.Lyla
{
    public partial class Production : Form,ToolStripAccess
    {
        public Production()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            tabControl1.SelectTab(tabPage2);

           
        }

        private static Production _instance; string coid = "", siid = "", fabid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        public static Production Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new Production();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

        Models.CommonClass CC = new Models.CommonClass();

        
        private void Production_Load(object sender, EventArgs e)
        {
          
          
           
            combocompcode.Select();
        }
        string maxid = ""; 
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
        public void Saves()
        {
            try
            {

                Models.ProductionLot p = new Models.ProductionLot();
                p.asptblprolotid = Convert.ToInt64("0" + txtprodid.Text);
                p.asptblprolotid1 = Convert.ToInt64("0" + txtprodid1.Text);
                p.finyear = Class.Users.Finyear;
                p.shortcode = Convert.ToString(txtshortcode.Text);
                p.wdate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                p.prodno = Convert.ToString(txtprodno.Text);
                p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                p.pono = Convert.ToString(combopono.Text);
                p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                p.lotno = Convert.ToString(txtlotno.Text.ToUpper());
                p.bundle = Convert.ToString(txtbundle.Text);
                p.size = Convert.ToString(txtsize.Text.ToUpper());
                p.compcode1 = Class.Users.COMPCODE;
                p.orderqty = Convert.ToInt64(txtorderqty.Text);
                p.username = Class.Users.USERID;
                p.createdby = Convert.ToString(Class.Users.HUserName);
                p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                p.modified = Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                p.modifiedby = Class.Users.HUserName;
                p.ipaddress = Class.Users.IPADDRESS;
                if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; }
                if (checkcancel.Checked == true) { p.productioncancel = "T"; } else { p.productioncancel = "F"; }
                if (GlobalVariables.New_Flg == false)
                {
                    auto();
                    string ins = "insert into asptblprolot(asptblprolotid1,finyear,shortcode,wdate,prodno,compcode,buyer,pono,orderqty,stylename,lotno,bundle,size,productioncancel,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.asptblprolotid1 + "','" + p.finyear + "','" + p.shortcode + "','" + p.wdate + "','" + p.prodno + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.lotno + "','" + p.bundle + "','" + p.size + "','" + p.productioncancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "',date_format('" + p.modified + "','%Y-%m-%d'),'" + p.modifiedby + "','" + p.ipaddress + "');";
                    Utility.ExecuteNonQuery(ins);
                } 
                if (GlobalVariables.New_Flg == true)
                {
                    string up = "update  asptblprolot   set  asptblprolotid1='" + p.asptblprolotid1 + "',finyear='" + p.finyear + "' ,shortcode='"+p.shortcode+"',  wdate='" + p.wdate + "' , prodno ='" + p.prodno + "' ,compcode ='" + p.compcode + "' , buyer='" + p.buyer + "' , pono='" + p.pono + "' ,  orderqty ='" + p.orderqty + "' , stylename='" + p.stylename + "' ,lotno='" + p.lotno + "',bundle='" + p.bundle + "' ,size='" + p.size + "', productioncancel='" + p.productioncancel + "' , active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',modified=date_format('" + p.modified + "','%Y-%m-%d'), modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblprolotid='" + txtprodid.Text + "'";
                    Utility.ExecuteNonQuery(up);
                }
                Models.ProductionLot.ProductionLotDet pp = new Models.ProductionLot.ProductionLotDet();
                int i = 0;
                if (dataGridView1.Rows.Count >= 0)
                {
                    string sel2 = "select max(asptblprolotid) id    from  asptblprolot   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and prodno='" + txtprodno.Text + "'";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblprolot");
                    DataTable dt2 = ds2.Tables["asptblprolot"];
                    maxid = dt2.Rows[0]["id"].ToString();
                   
                    for (i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value == null)
                        {
                            dataGridView1.Rows[i].Cells[0].Value = 0;
                        }
                        else
                        {
                            pp.asptblprolotdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                        }
                        pp.asptblprolotid = Convert.ToInt64("0" + maxid);
                        pp.asptblprolotid1 = Convert.ToInt64(maxid);
                        pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                        pp.pono = Convert.ToString(combopono.Text);
                        pp.asptblpurdet1id = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[5].Value.ToString());
                        pp.asptblpurdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[6].Value.ToString());
                        pp.asptblpurid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[7].Value.ToString());
                        colorid(dataGridView1.Rows[i].Cells[8].Value.ToString());
                        pp.colorname = coid;
                        sizeid(dataGridView1.Rows[i].Cells[9].Value.ToString());
                        pp.sizename = siid;
                        pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[10].Value.ToString());
                        pp.process = Convert.ToString(comboprocess.SelectedValue);
                        pp.processcheck = "T";
                        pp.finyear = Class.Users.Finyear;
                        string sel1 = "select asptblprolotdetid   from  asptblprolotdet   where  asptblpurdet1id='" + pp.asptblprolotid + "' and asptblpurdetid='" + pp.asptblpurdetid + "' and asptblpurid='" + pp.asptblpurid + "' and asptblprolotid='" + pp.asptblprolotid + "' and asptblprolotid1='" + pp.asptblprolotid1 + "' and compcode='" + combocompcode.SelectedValue + "' and  pono='" + pp.pono + "'  and  colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "' and processname='" + pp.process + "' and  processcheck='" + pp.processcheck + "' and  finyear='" + pp.finyear + "' ";
                        DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                        DataTable dt1 = ds1.Tables["asptblprolot"];
                        if (dt1.Rows.Count != 0 && pp.asptblprolotdetid == 0 || pp.asptblprolotdetid == 0)
                        {
                          
                            string ins1 = "insert into asptblprolotdet(asptblpurdet1id,asptblpurdetid,asptblpurid,asptblprolotid,asptblprolotid1,compcode,pono,colorname,sizename,orderqty,processname,processcheck,finyear) values('" + pp.asptblpurdet1id + "','" + pp.asptblpurdetid + "','" + pp.asptblpurid + "','" + maxid.ToString() + "' ,'" + pp.asptblprolotid1 + "' , '" + combocompcode.SelectedValue + "' ,'" + pp.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.orderqty + "','" + pp.process + "','" + pp.processcheck + "','" + pp.finyear + "')";
                            Utility.ExecuteNonQuery(ins1);
                        }
                        else
                        {
                            string up1 = "update  asptblprolotdet  set asptblpurdet1id='" + pp.asptblprolotid + "',asptblpurdetid='" + pp.asptblpurdetid + "',asptblpurid='" + pp.asptblpurid + "',asptblprolotid='" + pp.asptblprolotid + "' ,asptblprolotid1='" + pp.asptblprolotid1 + "', compcode='" + p.compcode + "' , pono='" + pp.pono + "' , colorname='" + pp.colorname + "',sizename='" + pp.sizename + "', orderqty='" + pp.orderqty + "' , processname='" + pp.process + "',processcheck='" + pp.processcheck + "', finyear='" + pp.finyear + "' where asptblprolotdetid='" + pp.asptblprolotdetid + "'";
                            Utility.ExecuteNonQuery(up1);
                        }
                    }
                }
                if (txtprodid.Text == "")
                {
                    MessageBox.Show("Record Saved Successfully " + txtprodid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                    tabControl1.SelectTab(tabPage2);
                  
                }
                else
                {
                    MessageBox.Show("Record Updated Successfully " + txtprodid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                    tabControl1.SelectTab(tabPage2);
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("colorname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }



       
        private void Production_FormClosed(object sender, FormClosedEventArgs e)
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
            txtprodid.Text = ""; txtprodno.Text = "";lblcount.Text = "";txtlotno.Text = "";
            txtprodid1.Text = "";txtbundle.Text = "";txtsize.Text = "";
            GlobalVariables.New_Flg = false;
            combostyle.Text = "";combopono.Text = "";combocompcode.Text = "";
            combobuyer.Text = ""; combopono.Text = "";
            txtprodid1.Text = "";
            checkactive.Checked = true;
            dataGridView1.Rows.Clear();

            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            dataGridView1.Rows.Clear();
            panel3.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;

            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            Class.Users.TableName = "asptblprolot";
            Class.Users.TableNameGrid = "asptblprolotdet";
            GlobalVariables.HideCols = new string[] { "asptblprolotdetid", "asptblprolotid", "asptblprolotid1", "compcode", "pono", "asptblpurdetid", "asptblpurid","ProcessName", "processcheck" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 0, 100, 0, 0, 300, 100, 50,0,0 };
            Class.Users.Query = "select a.asptblprolotdetid, a.asptblprolotid,a.asptblprolotid1,a.compcode, a.pono,a.asptblpurdet1id as QrCode,a.asptblpurdetid,a.asptblpurid,a.colorname,a.sizename,a.orderqty as qty,ProcessName,processcheck from asptblprolotdet a where a.asptblprolotdetid<0";
            CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols);
           
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
                string sel1 = "SELECT a.asptblprolotid,a.wdate, b.compcode,a.prodno, c.buyername,a.pono,d.stylename,a.active FROM  asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename   order by  a.asptblprolotid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                DataTable dt = ds.Tables["asptblprolot"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblprolotid"].ToString());
                        list.SubItems.Add(myRow["wdate"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["buyername"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());     
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["prodno"].ToString());
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
                    Cursor = Cursors.WaitCursor;
                    dataGridView1.Rows.Clear();
                    GlobalVariables.New_Flg = true;
                    txtprodid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "SELECT a.asptblprolotid,a.asptblprolotid1,b.compcode,a.prodno,a.wdate, c.buyercode,a.pono,d.stylename,a.lotno,a.bundle,a.size,a.active FROM  asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename  where a.asptblprolotid=" + txtprodid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                    DataTable dt = ds.Tables["asptblprolot"];
                    if (dt.Rows.Count > 0)
                    {
                        txtprodid.Text = Convert.ToString(dt.Rows[0]["asptblprolotid"].ToString());
                        txtprodid1.Text = Convert.ToString(dt.Rows[0]["asptblprolotid1"].ToString());
                        txtprodno.Text = Convert.ToString(dt.Rows[0]["prodno"].ToString());
                        dateTimePicker1.Text = dt.Rows[0]["wdate"].ToString();
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());                      
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                        combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                        txtbundle.Text = Convert.ToString(dt.Rows[0]["bundle"].ToString());
                        txtsize.Text = Convert.ToString(dt.Rows[0]["size"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        pono(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblprolotid"].ToString());
                    }
                    Cursor = Cursors.Default;
                }

            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show(ex.ToString());
            }
            Cursor = Cursors.Default;
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
                        if (item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text))
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
            //        string sel1 = "  SELECT  a.asptblprolotid,a.colorname,a.active from asptblprolot a  where a.colorname LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
            //        DataTable dt = ds.Tables["asptblprolot"];
            //        if (dt.Rows.Count > 0)
            //        {

            //            foreach (DataRow myRow in dt.Rows)
            //            {
            //                ListViewItem list = new ListViewItem();
            //                list.Text = iGLCount.ToString();
            //                list.SubItems.Add(myRow["asptblprolotid"].ToString());
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
            if (txtprodid.Text != "")
            {
                string sel1 = "select a.asptblprolotid from asptblprolot a join gtstatemast b on a.asptblprolotid=b.country where a.asptblprolotid='" + txtprodid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                DataTable dt = ds.Tables["asptblprolot"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtprodid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtprodid.Text != "")
                    {
                        string del = "delete from asptblprolot where asptblprolotid=" + txtprodid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptblprolotdet where asptblprolotdetid=" + txtprodid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + txtprodid1.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtprodid1.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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

                if (txtprodid.Text == "")
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
       
        public void pono(string s, string ss,string sss)
        {
            if (s != "" && ss != "" && sss == "" )
            {
                DataTable dt = new DataTable();
                dt = Utility.SQLQuery("select distinct  e.asptblsizgrpid,e.sizegroup,f.asptblbuymasid,f.buyercode,g.asptblstymasid, g.stylename from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblbuymas f on f.asptblbuymasid=a.buyer   join asptblstymas  g on g.asptblstymasid=a.stylename   where  b.compcode='" + s + "' and a.pono='" + ss + "'");
                if (dt.Rows.Count > 0)
                {
                    combostyle.DisplayMember = "sizegroup";
                    combostyle.ValueMember = "asptblsizgrpid";
                    combostyle.DataSource = dt;
                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.ValueMember = "asptblbuymasid";
                    combobuyer.DataSource = dt;

                    combostyle.DisplayMember = "stylenam";
                    combostyle.ValueMember = "asptblstymasid";  
                    combostyle.DataSource = dt;
                    dt = Utility.SQLQuery("select a.orderqty,a.lotno from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid    where  b.compcode='" + s + "' and a.pono='" + ss + "'");

                    txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                    txtlotno.Text = dt.Rows[0]["lotno"].ToString();

                }
                Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T'  order by 2  ", "asptblpromasid", "processname");

            }
            if (sss != "")
            {
                DataTable dt1 = new DataTable();
                dt1 = Utility.SQLQuery("select c.asptblprolotdetid,a.asptblprolotid,a.asptblprolotid1,b.compcode,a.pono,c.asptblpurdet1id,c.asptblpurdetid,c.asptblpurid, g.colorname,h.SIZENAME,c.orderqty,c.processname from asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblprolotdet c on c.asptblprolotid=a.asptblprolotid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblcolmas g on g.asptblcolmasid=c.colorname join asptblsizmas h on h.ASPTBLSIZMASID=c.sizename  where  b.compcode='" + s + "'  and a.pono='" + ss + "' and a.asptblprolotid='" + sss + "'");
                if (dt1.Rows.Count > 0)
                {
                    dataGridView1.Rows.Clear();

                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[i].Cells[0].Value = dt1.Rows[i]["asptblprolotdetid"].ToString();
                        dataGridView1.Rows[i].Cells[1].Value = dt1.Rows[i]["asptblprolotid"].ToString();
                        dataGridView1.Rows[i].Cells[2].Value = dt1.Rows[i]["asptblprolotid1"].ToString();
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
        void auto()
        {
            Models.CommonClass com = new Models.CommonClass();
            Class.Users.Finyear = System.DateTime.Now.Year.ToString();
            DataTable dt1 = com.autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName, "asptblprolot");
            if (dt1.Rows.Count > 0)
            {
                DataTable dt11 = com.shortcode(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName, "asptblprolot");
                if (dt11.Rows.Count < 0) { return; }
                else
                {
                    combocompname.Text = dt1.Rows[0]["COMPNAME"].ToString();
                    txtshortcode.Text = dt11.Rows[0]["shortcode"].ToString();
                    txtprodid1.Text = dt1.Rows[0]["id"].ToString();
                    txtprodno.Text = Class.Users.Finyear + "-" + txtshortcode.Text + "-" + txtprodid1.Text;
                    Utility.Load_Combo(combopono, "select a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid where  a.active='T' and b.compcode='" + combocompcode.Text + "' order by a.asptblpurid desc", "asptblpurid", "pono");
                }
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

            allip1.Items.Clear(); allip2.Items.Clear();
            if (combocompcode.SelectedIndex != -1)
            {
                if (combocompcode.Text != "" && txtprodid.Text == "")
                {
                    auto();
                }

            }
        }

        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopono.SelectedIndex != -1 && dataGridView1.Columns.Count > 0)
            {
                pono(combocompcode.Text, combopono.Text, txtprodid.Text);
                DataTable dt = CC.select("select MIN(A.asptblpurdet1ID) AS  MINID , MAX(A.asptblpurdet1ID) MAXID,count(a.asptblpurdet1id) cnt from asptblpurdet1 a where a.pono='" + combopono.Text+"'", "asptblpurdet1");
                lblbarcode.Refresh();
                lblbarcode.Text = "'" + dt.Rows[0]["MINID"].ToString() + "-" + dt.Rows[0]["MAXID"].ToString() + "'       Total Rows  :"+ dt.Rows[0]["cnt"].ToString();
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
            Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T'  order by 2  ", "asptblpromasid", "processname");

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

        DataTable dt1 = new DataTable();
        private void butadd_Click(object sender, EventArgs e)
        {
            string source = "";
            source=txtbarcode.Text.Trim();
            string data = getBetween(source, "'", "'");
            string[] data1 = data.Split('-');
            if (data1.Length == 2)
            {
                Cursor = Cursors.WaitCursor; 
                int a1 = Convert.ToInt32(data1[0]); dataGridView1.Rows.Clear();
                int a2 = Convert.ToInt32(data1[1]);
                rowcount = 0;
                for (int aa = a1; aa <= a2; aa++)
                {
                    DataTable dt = Utility.SQLQuery("select count(A.asptblprolotdetid) CNT from asptblprolotdet A WHERE a.compcode='" + combocompcode.SelectedValue + "' and A.pono='" + combopono.Text + "' AND A.asptblpurdet1id='" + aa.ToString() + "' AND A.processname='" + comboprocess.SelectedValue + "' and a.processcheck='T' ");
                    string s = dt.Rows[0]["CNT"].ToString();
                    if (Convert.ToInt32(s) >= 1) { lblcount.Refresh(); lblcount.Text = "Child Record Found  .  " + a1.ToString() + "   " + aa.ToString(); }
                    else
                    {
                        dt1 = Utility.SQLQuery("select f.asptblpurdet1id, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + combocompcode.Text + "'  and a.pono='" + combopono.Text + "' and f.asptblpurdet1id='" + aa + "'");
                        if (dt1.Rows.Count > 0)
                        {
                            bool chek = mas.checkduplicate1(5, dataGridView1, aa.ToString());
                            if (chek == true)
                            {

                                dataGridView1.Rows.Add();
                            
                                //rowcount = dataGridView1.Rows.Count - 1;
                                dataGridView1.Rows[rowcount].Cells[0].Value = "0";
                                dataGridView1.Rows[rowcount].Cells[1].Value = txtprodid.Text;
                                dataGridView1.Rows[rowcount].Cells[2].Value = txtprodid1.Text;
                                dataGridView1.Rows[rowcount].Cells[3].Value = combocompcode.Text;
                                dataGridView1.Rows[rowcount].Cells[4].Value = combopono.Text;
                                dataGridView1.Rows[rowcount].Cells[5].Value = dt1.Rows[0]["asptblpurdet1id"].ToString();
                                dataGridView1.Rows[rowcount].Cells[6].Value = dt1.Rows[0]["asptblpurdetid"].ToString();
                                dataGridView1.Rows[rowcount].Cells[7].Value = dt1.Rows[0]["asptblpurid"].ToString();
                                dataGridView1.Rows[rowcount].Cells[8].Value = dt1.Rows[0]["colorname"].ToString();
                                dataGridView1.Rows[rowcount].Cells[9].Value = dt1.Rows[0]["sizename"].ToString();
                                dataGridView1.Rows[rowcount].Cells[10].Value = 1;
                                dataGridView1.Rows[rowcount].Cells[11].Value = comboprocess.Text;
                                lblcount.Refresh();
                                lblcount.Text = rowcount.ToString();
                                rowcount++;
                                CommonFunctions.SetRowNumber(dataGridView1);
                            }
                            else
                            {
                                Cursor = Cursors.Default;
                                return;
                            }
                        }

                    }
                }
                Cursor = Cursors.Default;
            }
            else
            {
                DataTable dt = Utility.SQLQuery("select count(A.asptblprolotdetid) CNT from asptblprolotdet A WHERE a.compcode='" + combocompcode.SelectedValue + "' and A.pono='" + combopono.Text + "' AND A.asptblpurdet1id='" + txtbarcode.Text + "' AND A.processname='" + comboprocess.SelectedValue + "' and a.processcheck='T' ");
                string s = dt.Rows[0]["CNT"].ToString();
                if (Convert.ToInt32(s) >= 1) { lblcount.Refresh(); lblcount.Text = "Child Record Found  .  " + txtbarcode.Text; }
                else
                {
                    dt1 = Utility.SQLQuery("select f.asptblpurdet1id, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + combocompcode.Text + "'  and a.pono='" + combopono.Text + "' and f.asptblpurdet1id='" + txtbarcode.Text + "'");
                    if (dt1.Rows.Count > 0)
                    {
                        bool chek = mas.checkduplicate1(5, dataGridView1, txtbarcode.Text);
                        if (chek == true)
                        {
                            dataGridView1.Rows.Add();
                            rowcount = dataGridView1.Rows.Count - 1;
                            dataGridView1.Rows[rowcount].Cells[0].Value = "0";
                            dataGridView1.Rows[rowcount].Cells[1].Value = txtprodid.Text;
                            dataGridView1.Rows[rowcount].Cells[2].Value = txtprodid1.Text;
                            dataGridView1.Rows[rowcount].Cells[3].Value = combocompcode.Text;
                            dataGridView1.Rows[rowcount].Cells[4].Value = combopono.Text;
                            dataGridView1.Rows[rowcount].Cells[5].Value = dt1.Rows[0]["asptblpurdet1id"].ToString();
                            dataGridView1.Rows[rowcount].Cells[6].Value = dt1.Rows[0]["asptblpurdetid"].ToString();
                            dataGridView1.Rows[rowcount].Cells[7].Value = dt1.Rows[0]["asptblpurid"].ToString();
                            dataGridView1.Rows[rowcount].Cells[8].Value = dt1.Rows[0]["colorname"].ToString();
                            dataGridView1.Rows[rowcount].Cells[9].Value = dt1.Rows[0]["sizename"].ToString();
                            dataGridView1.Rows[rowcount].Cells[10].Value = 1;
                            dataGridView1.Rows[rowcount].Cells[11].Value = comboprocess.Text;
                            lblcount.Refresh();
                            lblcount.Text = rowcount.ToString();
                             CommonFunctions.SetRowNumber(dataGridView1);
                        }
                    }
                }
            }
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

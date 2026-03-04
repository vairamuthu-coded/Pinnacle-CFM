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



        private void autonumberload(string y, string com, string scr)
        {
            DataTable dt1 = new DataTable();
            string sel1 = "select max(a.asptblprolot1id)+1 as id,a.shortcode,B.COMPNAME from asptblprolot a join gtcompmast b on a.compcode = b.gtcompmastid  where  b.compcode='" + com + "' group by a.shortcode,B.COMPNAME";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
            dt1 = ds1.Tables["asptblprolot"];
           
          

            if (dt1.Rows.Count <= 0)
            {
                sel1 = "";
                sel1 = "select max(a.sequenceno)+1 as id,a.shortcode,B.COMPNAME from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptbluserrights c on c.userrightsid=a.screen where a.finyear='" + y + "' and b.compcode='" + com + "' AND C.MENUNAME='" + scr + "' group by a.shortcode,B.COMPNAME";
                ds1 = Utility.ExecuteSelectQuery(sel1, "asptblautogeneratemas");
                dt1 = ds1.Tables["asptblautogeneratemas"];
            
            }
            combocompname.Text = dt1.Rows[0]["COMPNAME"].ToString();
            txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
            txtprodid1.Text = dt1.Rows[0]["id"].ToString();
            txtprodno.Text = Class.Users.Finyear + "-" + txtshortcode.Text + "-" + txtprodid1.Text;

        }
        private void Production_Load(object sender, EventArgs e)
        {
          
            GridLoad();compload();
           
            combocompcode.Select();
        }
        string maxid = ""; 
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
        public void Saves()
        {
            try
            {

                Models.ProductionLot p = new Models.ProductionLot();
                p.asptblprolotid = Convert.ToInt64("0" + txtprodid.Text);
                p.asptblprolot1id = Convert.ToInt64("0" + txtprodid1.Text);
                p.finyear = Class.Users.Finyear;
                p.shortcode = Convert.ToString(txtshortcode.Text);
                p.wdate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                p.prodno = Convert.ToString(txtprodno.Text);
                p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                p.pono = Convert.ToString(combopono.Text);
                p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                p.lotno = Convert.ToString(txtlotno.Text);
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
                    string ins = "insert into asptblprolot(asptblprolot1id,finyear,shortcode,wdate,prodno,compcode,buyer,pono,orderqty,stylename,lotno,productioncancel,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.asptblprolot1id + "','" + p.finyear + "','" + p.shortcode + "','" + p.wdate + "','" + p.prodno + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.lotno + "','" + p.productioncancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "',date_format('" + p.modified + "','%Y-%m-%d'),'" + p.modifiedby + "','" + p.ipaddress + "');";
                    Utility.ExecuteNonQuery(ins);
                } 
                else
                {
                    string up = "update  asptblprolot   set  asptblprolot1id='" + p.asptblprolot1id + "',finyear='" + p.finyear + "' ,shortcode='"+p.shortcode+"',  wdate='" + p.wdate + "' , prodno ='" + p.prodno + "' ,compcode ='" + p.compcode + "' , buyer='" + p.buyer + "' , pono='" + p.pono + "' ,  orderqty ='" + p.orderqty + "' , stylename='" + p.stylename + "' ,lotno='" + p.lotno + "', productioncancel='" + p.productioncancel + "' , active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',modified=date_format('" + p.modified + "','%Y-%m-%d'), modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblprolotid='" + txtprodid.Text + "'";
                    Utility.ExecuteNonQuery(up);
                }
                Models.ProductionLotDet pp = new Models.ProductionLotDet();
                int i = 0;
                if (dataGridView1.Rows.Count >= 0)
                {
                    string sel2 = "select max(asptblprolotid) id    from  asptblprolot   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and prodno='" + txtprodno.Text + "'";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblprolot");
                    DataTable dt2 = ds2.Tables["asptblprolot"];
                    maxid = dt2.Rows[0]["id"].ToString();
                    for (i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        pp.asptblprolotdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                        p.asptblprolotid = Convert.ToInt64("0" + maxid);
                        p.asptblprolot1id = Convert.ToInt64(maxid);
                        p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                        pp.pono = Convert.ToString(combopono.Text);
                        colorid(dataGridView1.Rows[i].Cells[4].Value.ToString());
                        pp.colorname = coid;
                        sizeid(dataGridView1.Rows[i].Cells[5].Value.ToString());
                        pp.sizename = siid;
                        pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[6].Value.ToString());
                        pp.comqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[7].Value.ToString());
                        pp.lotqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[8].Value.ToString());
                        pp.balqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[9].Value.ToString());
                        pp.notes = Convert.ToString(dataGridView1.Rows[i].Cells[10].Value.ToString());
                        string sel1 = "select asptblprolotdetid   from  asptblprolotdet   where  asptblprolotid='" + p.asptblprolotid + "' and asptblprolot1id='" + p.asptblprolot1id + "' and compcode='" + p.compcode + "' and  pono='" + pp.pono + "'  and  colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "'  and comqty='" + pp.comqty + "' and lotqty='" + pp.lotqty + "' and  balqty ='" + pp.balqty + "' and  notes ='" + pp.notes + "'  ";
                        DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                        DataTable dt1 = ds1.Tables["asptblprolot"];
                        if (dt1.Rows.Count != 0)
                        {
                        }
                        else if (dt1.Rows.Count != 0 && pp.asptblprolotdetid == 0 || pp.asptblprolotdetid == 0)
                        {
                            string ins1 = "insert into asptblprolotdet(asptblprolotid,asptblprolot1id,compcode,pono,colorname,sizename,orderqty,comqty,lotqty,balqty,notes) values('" + maxid.ToString() + "' ,'" + p.asptblprolot1id + "' , '" + p.compcode + "' ,'" + pp.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.orderqty + "','" + pp.comqty + "','" + pp.lotqty + "','" + pp.balqty + "','" + pp.notes + "');";
                            Utility.ExecuteNonQuery(ins1);

                        }
                        else
                        {

                            string up1 = "update  asptblprolotdet  set asptblprolotid='" + p.asptblprolotid + "' ,asptblprolot1id='" + p.asptblprolot1id + "', compcode='" + p.compcode + "' , pono='" + pp.pono + "' , colorname='" + pp.colorname + "',sizename='" + pp.sizename + "', orderqty='" + pp.orderqty + "' , comqty='" + pp.comqty + "' , lotqty='" + pp.lotqty + "' ,  balqty ='" + pp.balqty + "' ,  notes ='" + pp.notes + "' where asptblprolotdetid='" + pp.asptblprolotdetid + "'";
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

           empty();
        }
        private void empty()
        {
            txtprodid.Text = ""; txtprodno.Text = "";
            txtprodid1.Text = "";
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
            panel6.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            Class.Users.TableNameGrid = "asptblprolotdet";
            GlobalVariables.HideCols = new string[] { "asptblprolotdetid", "asptblprolotid", "asptblprolot1id", "compcode", "pono" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 0, 300, 200, 100, 100, 100 };
            Class.Users.Query = "select * from asptblprolotdet a where a.asptblprolotdetid<0";
            CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblprolotdet");
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
                    dataGridView1.Rows.Clear();
                    GlobalVariables.New_Flg = true;
                    txtprodid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "SELECT a.asptblprolotid,a.asptblprolot1id,b.compcode,a.prodno,a.wdate, c.buyercode,a.pono,d.stylename,a.lotno,a.active FROM  asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename  where a.asptblprolotid=" + txtprodid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                    DataTable dt = ds.Tables["asptblprolot"];
                    if (dt.Rows.Count > 0)
                    {
                        txtprodid.Text = Convert.ToString(dt.Rows[0]["asptblprolotid"].ToString());
                        txtprodid1.Text = Convert.ToString(dt.Rows[0]["asptblprolot1id"].ToString());
                        txtprodno.Text = Convert.ToString(dt.Rows[0]["prodno"].ToString());
                        dateTimePicker1.Text = dt.Rows[0]["wdate"].ToString();
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());                      
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                        combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }

                        pono(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblprolotid"].ToString());
                     
                        
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
                dt = Utility.SQLQuery("select distinct  e.asptblsizgrpid,e.sizegroup,f.asptblbuymasid,f.buyercode,a.orderqty from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblbuymas f on f.asptblbuymasid=a.buyer    where  b.compcode='" + s + "' and a.pono='" + ss + "'");
                if (dt.Rows.Count > 0)
                {
                    combostyle.DisplayMember = "sizegroup";
                    combostyle.ValueMember = "asptblsizgrpid";
                    combostyle.DataSource = dt;
                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.ValueMember = "asptblbuymasid";
                    combobuyer.DataSource = dt;
                    txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                    DataTable dt1 = new DataTable();
                    dt1 = Utility.SQLQuery("select   g.colorname,h.SIZENAME,c.orderqty from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblcolmas g on g.asptblcolmasid=c.colorname join asptblsizmas h on h.ASPTBLSIZMASID=c.sizename  where  b.compcode='" + s + "'  and a.pono='" + ss + "'");
                    if (dt1.Rows.Count > 0)
                    {
                        dataGridView1.Rows.Clear();

                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[i].Cells[0].Value = "";
                            dataGridView1.Rows[i].Cells[1].Value = txtprodid.Text;
                            dataGridView1.Rows[i].Cells[2].Value = txtprodid1.Text;
                            dataGridView1.Rows[i].Cells[3].Value = combocompcode.Text;
                            dataGridView1.Rows[i].Cells[4].Value = combopono.Text;
                            dataGridView1.Rows[i].Cells[5].Value = dt1.Rows[i]["colorname"].ToString();
                            dataGridView1.Rows[i].Cells[6].Value = dt1.Rows[i]["sizename"].ToString();
                            dataGridView1.Rows[i].Cells[7].Value = dt1.Rows[i]["orderqty"].ToString();
                            dataGridView1.Rows[i].Cells[8].Value = "";
                            dataGridView1.Rows[i].Cells[9].Value = "";
                            dataGridView1.Rows[i].Cells[10].Value = "";
                            dataGridView1.Rows[i].Cells[11].Value = "";
                        }
                        CommonFunctions.SetRowNumber(dataGridView1);
                    }
                    else
                    {
                        dataGridView1.Rows.Clear();
                    }
                }
            }
            if(sss != "")
            {
                DataTable dt1 = new DataTable();
                dt1 = Utility.SQLQuery("select c.asptblprolotdetid,a.asptblprolotid,a.asptblprolot1id,b.compcode,a.pono, g.colorname,h.SIZENAME,c.orderqty,c.comqty,c.lotqty,c.balqty,c.notes from asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblprolotdet c on c.asptblprolotid=a.asptblprolotid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblcolmas g on g.asptblcolmasid=c.colorname join asptblsizmas h on h.ASPTBLSIZMASID=c.sizename  where  b.compcode='" + s + "'  and a.pono='" + ss + "' and a.asptblprolotid='" + sss + "'");
                if (dt1.Rows.Count > 0)
                {
                    dataGridView1.Rows.Clear();

                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[i].Cells[0].Value = dt1.Rows[i]["asptblprolotdetid"].ToString();
                        dataGridView1.Rows[i].Cells[1].Value = dt1.Rows[i]["asptblprolotid"].ToString();
                        dataGridView1.Rows[i].Cells[2].Value = dt1.Rows[i]["asptblprolot1id"].ToString();
                        dataGridView1.Rows[i].Cells[3].Value = dt1.Rows[i]["compcode"].ToString();
                        dataGridView1.Rows[i].Cells[4].Value = dt1.Rows[i]["pono"].ToString();
                        dataGridView1.Rows[i].Cells[5].Value = dt1.Rows[i]["colorname"].ToString();
                        dataGridView1.Rows[i].Cells[6].Value = dt1.Rows[i]["sizename"].ToString();
                        dataGridView1.Rows[i].Cells[7].Value = dt1.Rows[i]["orderqty"].ToString();
                        dataGridView1.Rows[i].Cells[8].Value = dt1.Rows[i]["comqty"].ToString();
                        dataGridView1.Rows[i].Cells[9].Value = dt1.Rows[i]["lotqty"].ToString();
                        dataGridView1.Rows[i].Cells[10].Value = dt1.Rows[i]["balqty"].ToString();
                        dataGridView1.Rows[i].Cells[11].Value = dt1.Rows[i]["notes"].ToString();
                    }
                    CommonFunctions.SetRowNumber(dataGridView1);
                }
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

       allip1.Items.Clear(); allip2.Items.Clear();
            if (combocompcode.SelectedIndex != -1)
            {

                    
                    DataTable dt = new DataTable();
                   // dt = Utility.SQLQuery("select A.asptblprolotid,a.prodno,B.COMPNAME  from asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + combocompcode.Text + "' order by 2");
                   // if (GlobalVariables.Dt.Rows.Count > 0)
                   // {
                    //txtwordordno.Text = dt.Rows[0]["prodno"].ToString();
                       // combocompname.Text = dt.Rows[0]["COMPNAME"].ToString();
                       
                        if (combocompcode.Text != "" && txtprodid.Text == "")
                        {
                         
                         autonumberload(Class.Users.Finyear, combocompcode.Text, Class.Users.ScreenName);
                    Utility.Load_Combo(combopono, "select a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid where  a.active='T' and b.compcode='" + combocompcode.Text+ "' order by a.asptblpurid desc", "asptblpurid", "pono");

                }
                //}

            }
        }

        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopono.SelectedIndex != -1 && dataGridView1.Columns.Count > 0)
            {
                pono(combocompcode.Text, combopono.Text,txtprodid.Text);

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
            //        //string sel1 = "select a.asptblprolotdetid from asptblprolotdet a join GTCOMPMAST b on a.compcode=b.gtcompmastid where a.asptblprolotdetid='" + item.SubItems[1].Text + "'";
            //        //DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolotdet");
            //        //DataTable dt = ds.Tables["asptblprolotdet"];
            //        //if (dt.Rows.Count > 0)
            //        //{
            //        //    MessageBox.Show("Child Record Found.Can Not Delete." + combopono.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            //        //}
            //        //else
            //        //{

            //            string del = "delete from asptblprolotdet where asptblprolotdetid='" + Convert.ToInt64("0" + item.SubItems[1].Text) + "'";
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

        private void butDelete_Click(object sender, EventArgs e)
        {

        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }
    }
}

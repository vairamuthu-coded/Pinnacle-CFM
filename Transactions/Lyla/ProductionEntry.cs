using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions.Lyla
{
    public partial class ProductionEntry : Form,ToolStripAccess
    {
        public ProductionEntry()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
          
        }
        private string screen = "ProductionEntry";
        private static ProductionEntry _instance; string coid = "", siid = "", fabid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
       
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        ListView allip = new ListView();
        ListView allip1 = new ListView();
        ListView allip2 = new ListView();
        ListView allip3delete = new ListView();
        public static ProductionEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ProductionEntry();
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
        Models.ProductionLot p = new Models.ProductionLot();
        public void Saves()
        {
            try
            {

                if (combonotes.Text == "")
                {
                    MessageBox.Show("Pls Select Notes", " No Data Found ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    return;
                }
                if (dataGridView1.Rows.Count > 0 && comboprocess.SelectedValue != null)
                {
                    Cursor.Current = Cursors.WaitCursor;                    
                    p.asptblprolotid = Convert.ToInt64("0" + txtprodid.Text);
                    p.asptblprolot1id = Convert.ToInt64("0" + txtprod1id.Text);
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
                    p.processname = Convert.ToInt64("0"+comboprocess.SelectedValue);
                    p.processtype = Convert.ToString(comboprocesstype.Text);
                    p.issuetype= Convert.ToString(comboprocesstype.Text);
                    p.panelmistake = Convert.ToString(comboprocesstype.Text);
                    if (comboprocesstype.Text == "INWARD") { p.inward1 = "INWARD"; p.stitching="T";p.inward = "T"; p.issuetype = "INWARD"; p.rechecking = "F"; } else { p.inward = "F"; }
                    if (comboprocesstype.Text == "DELIVERY") { p.delivery1 = "DELIVERY"; p.stitching = "T"; p.delivery = "T"; p.issuetype = "DELIVERY"; p.rechecking = "F"; } else { p.delivery = "F";}
                    if (comboprocesstype.Text == "REWORK") { p.rework1 = "REWORK"; p.inward = "T"; p.stitching = "T"; p.restitching = "T"; p.issuetype = "REWORK"; p.delivery = "T"; p.rechecking = "F"; } else { p.restitching = "F"; }
                    p.size = Convert.ToString(txtsize.Text.ToUpper());
                    p.notes = combonotes.Text;
                    p.compcode1 = Class.Users.COMPCODE;
                    p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
                    p.username = Class.Users.USERID;
                    p.createdby = Convert.ToString(Class.Users.HUserName);
                    p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                    // p.modified = Convert.ToString(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    p.modified = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                    p.modifiedby = Class.Users.HUserName;
                    p.ipaddress = Class.Users.IPADDRESS;
                    if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; }
                    if (checkcancel.Checked == true) { p.productioncancel = "T"; } else { p.productioncancel = "F"; }                    
                   Class.Users.dt= p.select(p.asptblprolotid, p.asptblprolot1id,p.compcode, p.finyear, p.shortcode, p.wdate, p.prodno, p.buyer, p.pono, p.orderqty, p.stylename, p.lotno, p.bundle, p.size, p.processname, p.processtype, p.productioncancel, p.active,p.issuetype,p.restitching,p.rechecking,p.inward,p.delivery);
                    if (Class.Users.dt.Rows.Count != 0)
                    { }
                     else if (Class.Users.dt.Rows.Count != 0 && p.asptblprolotid==0 || p.asptblprolotid==0)
                    {
                        
                         auto();
                        string ins = "insert into asptblprolot(asptblprolot1id,finyear,shortcode,wdate,prodno,compcode,buyer,pono,orderqty,stylename,lotno,bundle,size,processname,processtype, productioncancel,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress,issuetype,restitching,rechecking,inward,delivery,notes)  VALUES('" + p.asptblprolot1id + "','" + p.finyear + "','" + p.shortcode + "','" + p.wdate + "','" + p.prodno + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.lotno + "','" + p.bundle + "','" + p.size + "','" + p.processname + "','" + p.processtype + "','" + p.productioncancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "',date_format('" + p.modified + "','%Y-%m-%d'),'" + p.modifiedby + "','" + p.ipaddress + "','" + p.issuetype + "','" + p.restitching + "','" + p.rechecking + "','" + p.inward + "','" + p.delivery + "','" + p.notes + "');";
                        Utility.ExecuteNonQuery(ins);
                    }
                   else
                    {
                        //string up = "update  asptblprolot   set wdate=date_format('" + p.modified + "', '%Y-%m-%d'), compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',modified=date_format('" + p.modified + "','%Y-%m-%d'), modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblprolotid='" + txtprodid.Text + "'";
                        //Utility.ExecuteNonQuery(up);
                    }
                    Models.ProductionLotDet pp = new Models.ProductionLotDet();
                    int i = 0,j=1; int cnt = dataGridView1.Rows.Count;
                    if (cnt >= 0)
                    {
                        string sel2 = "select max(asptblprolotid) id    from  asptblprolot   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' ";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblprolot");
                        DataTable dt2 = ds2.Tables["asptblprolot"];
                        maxid = dt2.Rows[0]["id"].ToString();

                        for (i = 0; i < cnt; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value == null)
                            {
                                dataGridView1.Rows[i].Cells[0].Value = 0;
                            }
                            else
                            {
                                pp.asptblprolotdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                            }
                            p.asptblprolotid = Convert.ToInt64("0" + maxid);
                            p.asptblprolot1id = Convert.ToInt64("0"+maxid);
                            p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                            pp.pono = Convert.ToString(combopono.Text);
                            pp.barcode = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[5].Value.ToString());
                            pp.asptblpurdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[6].Value.ToString());
                            pp.asptblpurid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[7].Value.ToString());
                            colorid(dataGridView1.Rows[i].Cells[8].Value.ToString());
                            pp.colorname = coid;
                            sizeid(dataGridView1.Rows[i].Cells[9].Value.ToString());
                            pp.sizename = siid;
                            pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[10].Value.ToString());
                            pp.process = Convert.ToString(comboprocess.SelectedValue);
                            pp.processcheck = "T";
                            string sel1 = "select asptblprolotdetid   from  asptblprolotdet   where   asptblpurdet1id='" + pp.barcode + "' and barcode='" + pp.barcode + "' and asptblpurdetid='" + pp.asptblpurdetid + "' and asptblpurid='" + pp.asptblpurid + "' and asptblprolotid='" + p.asptblprolotid + "' and asptblprolot1id='" + p.asptblprolot1id + "' and compcode='" + combocompcode.SelectedValue + "' and  pono='" + p.pono + "'  and  colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "' and processtype='" + p.processtype + "' and processname='" + pp.process + "' and  processcheck='" + pp.processcheck + "' and  finyear='" + p.finyear + "' and    issuetype='" + p.issuetype + "' and  restitching ='" + p.restitching + "' and rechecking ='" + p.rechecking + "' and inward='" + p.inward + "' and delivery='" + p.delivery + "'  and panelmistake='" + p.panelmistake + "' ";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                            DataTable dt1 = ds1.Tables["asptblprolot"];
                            if (dt1.Rows.Count != 0 && pp.asptblprolotdetid == 0 || pp.asptblprolotdetid == 0)
                            {
                                string ins1 = "insert into asptblprolotdet(asptblpurdet1id,barcode,asptblpurdetid,asptblpurid,asptblprolotid,asptblprolot1id,compcode,pono,colorname,sizename,orderqty,processname,processcheck,finyear,processtype,issuetype,restitching,rechecking,inward,delivery,panelmistake,inward1,rework1,delivery1,modified,notes) values('" + pp.barcode + "','" + pp.barcode + "','" + pp.asptblpurdetid + "','" + pp.asptblpurid + "','" + maxid.ToString() + "' ,'" + p.asptblprolot1id + "' , '" + combocompcode.SelectedValue + "' ,'" + p.pono + "' , '" + pp.colorname + "','" + pp.sizename + "','" + pp.orderqty + "','" + pp.process + "','" + pp.processcheck + "','" + p.finyear + "','" + p.processtype + "','" + p.issuetype + "','" + p.restitching + "','" + p.rechecking + "','" + p.inward + "','" + p.delivery + "','" + p.panelmistake + "','" + p.inward1 + "','" + p.rework1 + "','" + p.delivery1 + "',date_format('" + p.modified + "','%Y-%m-%d'),'" + p.notes + "')";
                                Utility.ExecuteNonQuery(ins1);
                                                            
                                string up1 = "update asptblpurdet1 set panelmistake='" + p.panelmistake + "',issuetype='" + p.issuetype + "',RESTITCHING= '" + p.restitching + "',STITCHING= '" + p.stitching + "',rechecking='" + p.rechecking + "',inward='" + p.inward + "',delivery='" + p.delivery + "' , processcheck='T' where barcode='" + pp.barcode + "' AND PONO='" + p.pono + "' AND COMPCODE='" + p.compcode + "' AND FINYEAR='" + p.finyear + "'";
                                Utility.ExecuteNonQuery(up1);
                                
                                lblcount.Refresh();
                                lblcount.Text = "Inserted " + cnt + " of " + j.ToString();
                                j++;
                            }
                            else
                            {
                                //string up = "update asptblprolotdet set modified=date_format('" + p.modified + "','%Y-%m-%d') where asptblprolotdetid='" + pp.asptblprolotdetid + "' ";
                                //Utility.ExecuteNonQuery(up);
                            }
                        }
                    }
                    if (txtprodid.Text == "")
                    {
                        Cursor.Current = Cursors.Default;
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
                else
                {
                    MessageBox.Show("Invalid Grid Record " + dataGridView1.Rows.Count.ToString(), " No Data Found ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                ////string sel2 = "select max(asptblprolotid) id    from  asptblprolot   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "'";
                ////DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblprolot");
                ////DataTable dt2 = ds2.Tables["asptblprolot"];
                ////string del = "delete from asptblprolot a where a.asptblprolotid=" + dt2.Rows[0]["id"].ToString();
                ////Utility.ExecuteNonQuery(del);
                ////string comit = "commit;";
                ////    Utility.ExecuteNonQuery(comit);
                MessageBox.Show(ex.Message.ToString());
            }
            Cursor.Current = Cursors.Default;
        }



       
        private void Production_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       
        public void News()
        {
            GridLoad(); compload();

            empty();
            combocompcode.Focus();
        }
        private void empty()
        {
            txtsearch.Text = "";
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows.Clear();
                dataGridView1.AllowUserToAddRows = false;
            }
            Class.Users.UserTime = 0; txtorderno.Text = ""; txtstyleref.Text = "";
            combonotes.SelectedIndex = -1; Class.Users.UserTime = 0;
            txtprodid.Text = ""; txtprodno.Text = ""; lblcount.Text = ""; txtlotno.Text = ""; txtbarcode.Text = "";
            txtprod1id.Text = ""; txtbundle.Text = ""; txtsize.Text = ""; txtorderqty.Text = ""; lblbarcode.Text = "";
            GlobalVariables.New_Flg = false; GlobalVariables.Saves.Visible = true; txtbarcode.Text = "";
            combostyle.Text = ""; combopono.Text = ""; combocompcode.Text = ""; combocompname.Text = "";
            combobuyer.Text = ""; combopono.Text = ""; combobuyer.Text = ""; combostyle.Text = ""; combopono.Text = "";
            txtprod1id.Text = ""; comboprocess.Text = ""; comboprocesstype.Text = "";
            checkactive.Checked = true; dateTimePicker1.Value = Convert.ToDateTime(System.DateTime.Now.ToString());
            dataGridView1.Rows.Clear(); pictureBox1.Image = null;
            lblcount.ForeColor = Color.White; dataGridView1.Columns.Clear();
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
            GlobalVariables.HideCols = new string[] { "asptblprolotdetid", "asptblprolotid", "asptblprolot1id", "compcode", "pono", "asptblpurdetid", "asptblpurid", "ProcessName", "processcheck" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 0, 100, 0, 0, 300, 100, 50, 0, 0 };
            Class.Users.Query = "select a.asptblprolotdetid, a.asptblprolotid,a.asptblprolot1id,a.compcode, a.pono,a.asptblpurdet1id as QrCode,a.asptblpurdetid,a.asptblpurid,a.colorname,a.sizename,a.orderqty as qty,ProcessName,processcheck from asptblprolotdet a where a.asptblprolotdetid<0";
            CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblprolotdet");
            //if (dataGridView1.Rows.Count >= 1)
            //{
            //    do
            //    {
            //        int i = 0;
            //        for (i = 0; i < dataGridView1.Rows.Count; i++)
            //        {
            //            try
            //            {
            //                dataGridView1.Rows.RemoveAt(i);
            //            }
            //            catch(Exception ex) { }
            //        }
            //    }
            //    while (dataGridView1.Rows.Count > 1);
            //}
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
                string sel1 = "SELECT a.asptblprolotid,a.wdate, b.compcode,a.prodno, c.buyername,a.pono,d.stylename,a.lotno,a.orderqty,a.processtype, a.active FROM  asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename   order by  a.asptblprolotid desc";
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
                        list.SubItems.Add(myRow["wdate"].ToString().Substring(0,10));
                        list.SubItems.Add(myRow["prodno"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["buyername"].ToString());
                        list.SubItems.Add(myRow["lotno"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());                      
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["processtype"].ToString());
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
                    Cursor.Current = Cursors.WaitCursor;
                    dataGridView1.Rows.Clear();
            
                   
                    txtprodid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "SELECT a.asptblprolotid,a.asptblprolot1id,b.compcode,b.compname,a.prodno,a.wdate, c.buyercode,a.pono,d.stylename,a.lotno,a.bundle,a.size,a.active,a.processtype,A.processname,a.notes FROM  asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename  where a.asptblprolotid=" + txtprodid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                    DataTable dt = ds.Tables["asptblprolot"];
                    if (dt.Rows.Count > 0)
                    {
                        txtprodid.Text = Convert.ToString(dt.Rows[0]["asptblprolotid"].ToString());
                        txtprod1id.Text = Convert.ToString(dt.Rows[0]["asptblprolot1id"].ToString());
                        txtprodno.Text = Convert.ToString(dt.Rows[0]["prodno"].ToString());
                        dateTimePicker1.Text = dt.Rows[0]["wdate"].ToString();
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                        combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                        txtbundle.Text = Convert.ToString(dt.Rows[0]["bundle"].ToString());
                        txtsize.Text = Convert.ToString(dt.Rows[0]["size"].ToString());
                        combonotes.Text = Convert.ToString(dt.Rows[0]["Notes"].ToString());

                        string po = p.HideButton(combopono.Text, "asptblcutpanret");
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
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        comboprocess.SelectedValue = dt.Rows[0]["processname"].ToString();
                
                        comboprocesstype.Text = dt.Rows[0]["processtype"].ToString();

                        pono(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblprolotid"].ToString());
                        ponoGrid(dt.Rows[0]["compcode"].ToString(), dt.Rows[0]["pono"].ToString(), dt.Rows[0]["asptblprolotid"].ToString());

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
                if (txtsearch.Text.Length >= 3 && checkBox1.Checked == false)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[6].ToString().Contains(txtsearch.Text) || item.SubItems[8].ToString().Contains(txtsearch.Text) || item.SubItems[10].ToString().Contains(txtsearch.Text) || item.SubItems[11].ToString().Contains(txtsearch.Text))
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
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
               
                if (txtsearch.Text.Length>=5 && checkBox1.Checked == true)
                {
                    try
                    {
                        listView1.Items.Clear(); 
                        string sel1 = "SELECT a.asptblprolotid,a.wdate, b.compcode,a.prodno, c.buyername,a.pono,d.stylename,a.lotno,a.orderqty,a.processtype, a.active FROM  asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename  join asptblprolotdet e on e.asptblprolotid=a.asptblprolotid where e.BARCODE like'%" + txtsearch.Text + "%'    order by  a.asptblprolotid desc";
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
                                list.SubItems.Add(myRow["wdate"].ToString().Substring(0, 10));
                                list.SubItems.Add(myRow["prodno"].ToString());
                                list.SubItems.Add(myRow["compcode"].ToString());
                                list.SubItems.Add(myRow["buyername"].ToString());
                                list.SubItems.Add(myRow["lotno"].ToString());
                                list.SubItems.Add(myRow["pono"].ToString());
                                list.SubItems.Add(myRow["orderqty"].ToString());
                                list.SubItems.Add(myRow["stylename"].ToString());
                                list.SubItems.Add(myRow["processtype"].ToString());
                                list.SubItems.Add(myRow["active"].ToString());
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
                if (txtsearch.Text == "")
                {

                    ListView ll = new ListView();
                    checkBox1.Checked = false;
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
            if (txtprodid.Text != "")
            {
                string sel1 = "select a.asptblprolotid from asptblprolot a join gtstatemast b on a.asptblprolotid=b.country where a.asptblprolotid='" + txtprodid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblprolot");
                DataTable dt = ds.Tables["asptblprolot"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtprod1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtprodid.Text != "")
                    {
                        string del = "delete from asptblprolot where asptblprolotid=" + txtprodid.Text;
                        Utility.ExecuteNonQuery(del);
                        string del1 = "delete from asptblprolotdet where asptblprolotdetid=" + txtprodid.Text;
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully " + txtprod1id.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtprod1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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
        public void ponoGrid(string s, string ss, string sss)
        {
     
            if (sss != "")
            {
                DataTable dt1 = new DataTable();
                dt1 = Utility.SQLQuery("select c.asptblprolotdetid,a.asptblprolotid,a.asptblprolot1id,b.compcode,a.pono,c.asptblpurdet1id,c.asptblpurdetid,c.asptblpurid, g.colorname,h.SIZENAME,c.orderqty,c.processname from asptblprolot a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblprolotdet c on c.asptblprolotid=a.asptblprolotid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblcolmas g on g.asptblcolmasid=c.colorname join asptblsizmas h on h.ASPTBLSIZMASID=c.sizename  where  b.compcode='" + s + "'  and a.pono='" + ss + "' and a.asptblprolotid='" + sss + "'");
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
                    lblcount.Refresh();lblcount.Text ="Total Rows :"+ dataGridView1.Rows.Count.ToString();
                }
            }
        }
        public void pono(string s, string ss,string sss)
        {
            if (s != "" && ss != "")
            {
                DataTable dt = new DataTable();
            
                if (comboprocesstype.Text == "INWARD")
                {                
                     dt = Utility.SQLQuery("select distinct  e.asptblsizgrpid,e.sizegroup,f.asptblbuymasid,f.buyercode,g.asptblstymasid, g.stylename,a.orderqty,a.lotno,a.bundle,a.size,i.processname,i.asptblpromasid,a.garmentimage,a.orderno,a.styleref from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblbuymas f on f.asptblbuymasid=a.buyer   join asptblstymas  g on g.asptblstymasid=a.stylename  join asptblpromas i on i.asptblpromasid=a.processname  where  b.compcode='" + s + "' and a.pono='" + ss + "'");
                    if (dt.Rows.Count > 0)
                    {
                        p.stylename = Convert.ToInt64("0"+dt.Rows[0]["asptblstymasid"].ToString());
                        p.buyer = Convert.ToInt64("0" + dt.Rows[0]["asptblbuymasid"].ToString());
                        p.orderqty = Convert.ToInt64("0" + dt.Rows[0]["orderqty"].ToString());
                        p.processname = Convert.ToInt64("0" + dt.Rows[0]["orderqty"].ToString());
                       txtorderno.Text= Convert.ToString(dt.Rows[0]["orderno"].ToString());
                        txtstyleref.Text = Convert.ToString(dt.Rows[0]["styleref"].ToString());
                        combobuyer.DisplayMember = "buyercode";
                        combobuyer.ValueMember = "asptblbuymasid";
                        combobuyer.DataSource = dt;
                        combostyle.DisplayMember = "stylename";
                        combostyle.ValueMember = "asptblstymasid";
                        combostyle.DataSource = dt;
                        txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                        txtlotno.Text = dt.Rows[0]["lotno"].ToString();
                        txtbundle.Text = dt.Rows[0]["bundle"].ToString();
                        txtsize.Text = dt.Rows[0]["size"].ToString();
                    }
                }
                if (comboprocesstype.Text != "INWARD")
                {
                            dt = Utility.SQLQuery("select distinct  e.asptblsizgrpid,e.sizegroup,f.asptblbuymasid,f.buyercode,g.asptblstymasid, g.stylename,h.orderqty,h.lotno,h.bundle,h.size,i.processname,i.asptblpromasid,a.garmentimage,a.orderno,a.styleref from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid and c.compcode=a.compcode and c.compcode=b.gtcompmastid join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblbuymas f on f.asptblbuymasid=a.buyer   join asptblstymas  g on g.asptblstymasid=a.stylename JOIN asptblprolot h on h.pono=a.pono  join asptblpromas i on i.asptblpromasid=a.processname  where  b.compcode='" + s + "' and a.pono='" + ss + "'");
                        if (dt.Rows.Count > 0)
                    {
                        //comboprocess.DisplayMember = "processname";
                        //comboprocess.ValueMember = "asptblpromasid";
                        //comboprocess.DataSource = dt;

                        txtorderno.Text = Convert.ToString(dt.Rows[0]["orderno"].ToString());
                        txtstyleref.Text = Convert.ToString(dt.Rows[0]["styleref"].ToString());
                        combobuyer.DisplayMember = "buyercode";
                        combobuyer.ValueMember = "asptblbuymasid";
                        combobuyer.DataSource = dt;
                        combostyle.DisplayMember = "stylename";
                        combostyle.ValueMember = "asptblstymasid";
                        combostyle.DataSource = dt;
                        txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
                        txtlotno.Text = dt.Rows[0]["lotno"].ToString();
                        txtbundle.Text = dt.Rows[0]["bundle"].ToString();
                        txtsize.Text = dt.Rows[0]["size"].ToString();
                        txtorderno.Text = Convert.ToString(dt.Rows[0]["orderno"].ToString());
                        txtstyleref.Text = Convert.ToString(dt.Rows[0]["styleref"].ToString());
                    }
                }
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
        void auto()
        {
            if (combocompcode.Text != "")
            {


                DataTable dt1 = CC.autonumberload(Class.Users.Finyear, combocompcode.Text, screen, "asptblprolot");
                if (dt1.Rows.Count > 0)
                {
                    DataTable dt11 = CC.shortcode(Class.Users.Finyear, combocompcode.Text, screen, "asptblprolot");
                    if (dt11.Rows.Count < 0) { return; }
                    else
                    {
                        combocompname.Text = dt11.Rows[0]["COMPNAME"].ToString();
                        txtshortcode.Text = dt11.Rows[0]["shortcode"].ToString();
                        txtprod1id.Text = dt1.Rows[0]["id"].ToString();
                        txtprodno.Text = Class.Users.Finyear + "-" + txtshortcode.Text + "-" + txtprod1id.Text;
                        p.asptblprolot1id = Convert.ToInt64("0"+txtprod1id.Text);
                        p.prodno = Convert.ToString(txtprodno.Text);
                        Class.Users.UserTime = 0;
                    }
                }
            }
            else
            {
                MessageBox.Show("pls Select Compcode ");
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
        int  totalcount=0;
        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combopono.SelectedIndex != -1 && dataGridView1.Columns.Count > 0)
            {
                txtbarcode.Text = ""; combonotes.SelectedIndex = 0;
                pono(combocompcode.Text, combopono.Text, txtprodid.Text);
                DataTable dt = CC.select("select distinct MIN(A.barcode) AS  MINID , MAX(A.barcode) MAXID,count(a.barcode) cnt from asptblpurdet1 a where a.pono='" + combopono.Text + "'", "asptblpurdet1");
                if (dt.Rows[0]["MINID"].ToString() == "") { lblcount.Refresh(); lblcount.Text = "No Data Found.."; lblbarcode.Refresh(); lblbarcode.Text = "0"; }
                else
                {
                    lblbarcode.Refresh(); totalcount=0; totalcount =Convert.ToInt32(dt.Rows[0]["cnt"].ToString());
                    lblbarcode.Text = "'" + dt.Rows[0]["MINID"].ToString() + "-" + dt.Rows[0]["MAXID"].ToString() + "'";
                }
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
            try { 
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {

                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            }
            catch { Exception ex; }
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


        private void BindGrid(DataGridView grid, DataTable dt1, string barcode,int col)
        {
            bool chek = mas.checkduplicate1(col, grid, barcode);          
            if (chek == true)
            {
                grid.Rows.Add();
                rowcount = dataGridView1.Rows.Count - 1;
                grid.Rows[rowcount].Cells[0].Value = "0";
                grid.Rows[rowcount].Cells[1].Value = txtprodid.Text;
                grid.Rows[rowcount].Cells[2].Value = txtprod1id.Text;
                grid.Rows[rowcount].Cells[3].Value = combocompcode.Text;
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
                return;
            }
            grid.Sort(grid.Columns[col], ListSortDirection.Descending);
            grid.AllowUserToAddRows = false;
        }
        private DataTable barcode(string type, string com, string pono, string bar)
        {
            //if (type == "PANEL MISTAKE")
            //{
            //    Class.Users.dt2 = Utility.SQLQuery("select f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND F.INWARD='F' AND F.ISSUETYPE='PANEL MISTAKE'");
            //}
            if (type == "INWARD")
            {
                Class.Users.dt2 = Utility.SQLQuery("select f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND F.CUTTING='F'  AND F.STITCHING= 'F' AND F.CHECKING='F' ORDER BY 1");
            }
            if (type == "DELIVERY")
            {
                Class.Users.dt2 = Utility.SQLQuery("select DISTINCT f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid JOIN asptblprolot g on g.pono=a.pono and g.pono=f.pono JOIN asptblprolotdet h on h.asptblprolotid=g.asptblprolotid and h.pono=g.pono and h.barcode=f.barcode  where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND  F.STITCHING= 'T' and f.issuetype='INWARD' AND F.PANELMISTAKE='INWARD' ORDER BY 1");
            }
            if (type == "REWORK")
            {
                Class.Users.dt2 = Utility.SQLQuery("select  DISTINCT  f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid   where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND f.issuetype='CHECKING MISTAKE' AND f.REMARKS='STITCHING MISTAKE' and F.RECHECKING= 'T' union all select  DISTINCT  f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid   where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND f.issuetype='STITCHING MISTAKE' AND f.REMARKS='STITCHING MISTAKE' and F.RESTITCHING= 'F' union all select  DISTINCT  f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid   where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND f.issuetype='CHECKING MISTAKE' AND f.REMARKS='REWORK' and F.RESTITCHING= 'F' union all select  DISTINCT  f.barcode, c.asptblpurdetid,a.asptblpurid,a.pono,d.colorname,e.SIZENAME,f.orderqty from  asptblpur a  join gtcompmast b on a.compcode=b.gtcompmastid   join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptblcolmas d on d.asptblcolmasid=c.colorname  join asptblsizmas e on e.ASPTBLSIZMASID=c.sizename join asptblpurdet1 f on f.asptblpurdetid=c.asptblpurdetid and f.asptblpurid=a.asptblpurid   where  b.compcode='" + com + "'  and a.pono='" + pono + "' and f.barcode='" + bar + "' AND f.issuetype='REWORK' AND f.REMARKS='REWORK' and F.RESTITCHING= 'T' AND NOT F.PANELMISTAKE='CHECKING REWORK' AND  NOT F.PANELMISTAKE='CHECKING DELIVERY' ORDER BY 1");
            }
            return Class.Users.dt2;
        }
        private void butadd_Click(object sender, EventArgs e)
        {
            if (txtbarcode.Text.Length >= 5)
            {
                string source = ""; Class.Users.UserTime = 0;
                source = txtbarcode.Text.Trim();
                string data = getBetween(source, "'", "'"); rowcount = 0;
                string[] data1 = data.Split('-'); int aa = 0, cnt2 = 1;
                if (data1.Length == 2)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    int a1 = Convert.ToInt32(data1[0]); dataGridView1.Rows.Clear();
                    int a2 = Convert.ToInt32(data1[1]);
                    int a3 = 1;
                    a3 += a2 - a1;
                    for (aa = a1; aa <= a2; aa++)
                    {
                        try
                        {

                            DataTable dt = Utility.SQLQuery("select count(asptblprolotdetid) CNT from asptblprolotdet  where  barcode='" + aa.ToString() + "'  and compcode='" + combocompcode.SelectedValue + "' and  pono='" + combopono.Text + "'  and  processcheck='T' and  issuetype='" + comboprocesstype.Text + "' and  restitching ='T' and rechecking ='F' AND INWARD='T'  and delivery='F' ");
                            string s = dt.Rows[0]["CNT"].ToString();
                            if (Convert.ToInt32(s) >= 1)
                            {
                                lblcount.Refresh(); lblcount.Text = "Child Record Found  .  " + a2.ToString(); lblcount.Text = a3 + " of " + cnt2.ToString(); cnt2++;
                            }
                            else
                            {
                                Class.Users.dt1 = null;

                                Class.Users.dt1 = barcode(comboprocesstype.Text, combocompcode.Text, combopono.Text, aa.ToString());
                                if (Class.Users.dt1.Rows.Count > 0)
                                {
                                    BindGrid(dataGridView1, Class.Users.dt1, aa.ToString(),5);
                                    lblcount.Refresh();
                                    lblcount.Text = a3 + " of " + cnt2.ToString();
                                    rowcount++; cnt2++;


                                }
                                else
                                {
                                    lblcount.Refresh(); lblcount.Text = a3 + " of " + cnt2.ToString();
                                    rowcount++; cnt2++;
                                    if (a3 == cnt2)
                                    {
                                        MessageBox.Show("Invalid Barcode");
                                        lblcount.Refresh(); lblcount.Text = " Invalid BarCode " + dataGridView1.Rows.Count.ToString();
                                        txtbarcode.Text = ""; Cursor.Current = Cursors.Default;
                                    }
                                }

                            }
                        }
                        catch (Exception ex) { MessageBox.Show(ex.Message.ToString()); }
                    }
                    lblcount.Text = "Total Rows " + dataGridView1.Rows.Count.ToString();
                    Cursor.Current = Cursors.Default;

                }
                else
                {
                    if (txtbarcode.Text.Length >= 5)
                    {
                        try
                        {
                            DataTable dt = Utility.SQLQuery("select count(asptblprolotdetid) CNT from asptblprolotdet  where  barcode='" + txtbarcode.Text + "'  and compcode='" + combocompcode.SelectedValue + "' and  pono='" + combopono.Text + "'  and  processcheck='T' and  issuetype='" + comboprocesstype.Text + "' and  restitching ='T' and rechecking ='F' AND INWARD='T'  and delivery='F' ");
                            string s = dt.Rows[0]["CNT"].ToString();
                            if (Convert.ToInt32(s) >= 1)
                            {
                                MessageBox.Show("Child Record Found");
                                MessageBox.Show("Child Record Found ", "Alert", MessageBoxButtons.OK); lblcount.Refresh(); lblcount.Text = "Child Record Found  .  " + data1[0].ToString(); Cursor.Current = Cursors.Default; return;
                            }
                            else
                            {
                                Class.Users.dt1 = null;
                                Class.Users.dt1 = barcode(comboprocesstype.Text, combocompcode.Text, combopono.Text, txtbarcode.Text);

                                if (Class.Users.dt1.Rows.Count > 0)
                                {
                                    BindGrid(dataGridView1, Class.Users.dt1, txtbarcode.Text,5);
                                    lblcount.Refresh(); lblcount.Text = "Count :" + dataGridView1.Rows.Count.ToString() + " Of " + totalcount.ToString();
                                    Cursor.Current = Cursors.Default;
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
                        catch (Exception ex) { MessageBox.Show(ex.Message.ToString()); }
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
            if (txtprodid.Text != "")
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

        private void comboprocesstype_SelectedIndexChanged(object sender, EventArgs e)
        {

           
           
            if (comboprocesstype.Text == "INWARD")
            {
                combopono.DataSource = null;
                Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T' and not PROCESSNAME='CUTTING' and not PROCESSNAME='CHECKING' and not PROCESSNAME='PACKING' and not PROCESSNAME='PACKING' order by 2  ", "asptblpromasid", "processname");
                    Utility.Load_Combo(combopono, "select DISTINCT a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where   E.PONO IS NULL and  a.active='T' and b.compcode='" + combocompcode.Text + "'  order by 1 desc", "asptblpurid", "pono");

            }            
            if (comboprocesstype.Text == "DELIVERY")
            {
                combopono.DataSource = null;
                Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T' and not PROCESSNAME='CUTTING' and not PROCESSNAME='STITCHING'   and not PROCESSNAME='PACKING' order by 2  ", "asptblpromasid", "processname");
                Utility.Load_Combo(combopono, "select DISTINCT a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid JOIN asptblprolotdet C ON A.PONO=C.PONO AND A.COMPCODE=C.COMPCODE AND C.inward1='INWARD' LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where   E.PONO IS NULL and  a.active='T' and b.compcode='" + combocompcode.Text + "'  order by 1 desc", "asptblpurid", "pono");

            }
            if (comboprocesstype.Text == "REWORK")
            {
                combopono.DataSource = null;
                Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T' and not PROCESSNAME='CUTTING'  and not PROCESSNAME='PACKING' and not PROCESSNAME='STITCHING' and not PROCESSNAME='PACKING' order by 2  ", "asptblpromasid", "processname");
                Utility.Load_Combo(combopono, "select DISTINCT a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid JOIN asptblcutpanretdet C ON A.PONO=C.PONO AND A.COMPCODE=C.COMPCODE AND C.issuetype='CHECKING MISTAKE' AND C.remarks='STITCHING MISTAKE' OR   C.issuetype='STITCHING MISTAKE' AND C.remarks='STITCHING MISTAKE'  join asptblprolot d on d.compcode=a.compcode and d.pono=a.pono and d.pono=c.pono LEFT JOIN  asptblordclomas E ON E.PONO=A.PONO AND E.ACTIVE='T' where   E.PONO IS NULL and  a.active='T' and b.compcode='" + combocompcode.Text + "'  order by 1 desc", "asptblpurid", "pono");

            }
            pono(combocompcode.Text, combopono.Text, txtprodid.Text);
            dataGridView1.Rows.Clear();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (deleteToolStripMenuItem.Text == "Delete")
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    if (txtprodid.Text != "")
                    {
                        string up1 = "update asptblpurdet1 set panelmistake='F',issuetype='CUTTING',RESTITCHING= 'F',STITCHING= 'F',rechecking='F',inward='F',delivery='F' , processcheck='F',REMOVEROW='DELETE' where barcode='" + row.Cells[5].Value + "' AND PONO='" + p.pono + "' AND COMPCODE='" + p.compcode + "' AND FINYEAR='" + p.finyear + "'";
                        Utility.ExecuteNonQuery(up1);
                        string del = "delete from asptblprolotdet  where asptblprolotdetid='" + row.Cells[0].Value + "' AND PONO='" + combopono.Text + "' AND COMPCODE='" + combocompcode.SelectedValue + "'";
                        Utility.ExecuteNonQuery(del);
                    }
                    dataGridView1.Rows.RemoveAt(row.Index);
                    dataGridView1.Refresh();
                }
            }
          
        }

        private void deleteAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (deleteAllToolStripMenuItem.Text == "Delete All")
            {
                if (txtprodid.Text != "")
                {
                    //foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    //{
                       
                    //        string up1 = "update asptblpurdet1 set panelmistake='F',issuetype='CUTTING',RESTITCHING= 'F',STITCHING= 'F',rechecking='F',inward='F',delivery='F' , processcheck='F',REMOVEROW='DELETE' where barcode='" + row.Cells[5].Value + "' AND PONO='" + p.pono + "' AND COMPCODE='" + p.compcode + "' AND FINYEAR='" + p.finyear + "'";
                    //        Utility.ExecuteNonQuery(up1);
                    //    string del = "delete from asptblprolotdet  where asptblprolotdetid='" + row.Cells[0].Value + "' AND PONO='" + combopono.Text + "' AND COMPCODE='" + combocompcode.SelectedValue + "'";
                    //    Utility.ExecuteNonQuery(del);
                    //    dataGridView1.Rows.RemoveAt(row.Index);
                    //    dataGridView1.Refresh();
                    //}
                }
                else
                {
                    dataGridView1.Rows.Clear();
                    this.dataGridView1.AllowUserToAddRows = false;
                   
                }
            }
        }

        private void refreshToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            auto();
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
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            txtsearch.Text = "";
        }

        private void dataGridView1_Sorted(object sender, EventArgs e)
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

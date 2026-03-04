using Pinnacle.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Reflection.Emit;
using System.Text;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media;
using System.Xml.Linq;
using MySql.Data.MySqlClient;

namespace Pinnacle.Transactions.Lyla
{
    public partial class BarcodeGenerate : Form, ToolStripAccess
    {
        public BarcodeGenerate()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());


        }
        private string screen = "BarcodeGenerate";
        private static BarcodeGenerate _instance; string coid = "", siid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        Models.CommonClass com = new Models.CommonClass();
        Models.PurchasesModel p = new Models.PurchasesModel();
        public static BarcodeGenerate Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new BarcodeGenerate();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void usercheck(string s, string ss, string sss)
        {


        }
       
        void auto()
        {
            try
            {
                    if (combocompcode.Text != "" && txtasptblpurid.Text == "")
                    {
                

                                DataTable dt1 = com.autonumberload1(Class.Users.Finyear, combocompcode.Text,screen, "asptblpur");
                        if (dt1.Rows.Count > 0)
                    {
                        DataTable dt2 = com.shortcode1(Class.Users.Finyear, combocompcode.Text, screen, "asptblpur");
                        txtasppur1id.Text = dt1.Rows[0]["id"].ToString();
                        txtshortcode.Text = dt2.Rows[0]["shortcode"].ToString();
                        txtpono.Text = Class.Users.Finyear + "-" + dt2.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
                        if (dt1.Rows[0]["barcode"].ToString() != "")
                        {
                            txtautono.Text = dt1.Rows[0]["barcode"].ToString();
                        }
                        else { txtautono.Text = dt2.Rows[0]["barcode"].ToString(); }
                        p.pono = Convert.ToString(txtpono.Text);
                        p.asptblpur1id = Convert.ToInt64("0" + txtasppur1id.Text);
                        Class.Users.UserTime = 0;
                    }
                    else
                    {
                        txtasppur1id.Text = dt1.Rows[0]["id"].ToString();
                        txtshortcode.Text = dt1.Rows[0]["shortcode"].ToString();
                        txtpono.Text = Class.Users.Finyear + "-" + dt1.Rows[0]["shortcode"].ToString() + "-" + dt1.Rows[0]["id"].ToString();
                        txtautono.Text = dt1.Rows[0]["barcode"].ToString();
                        p.pono = Convert.ToString(txtpono.Text);
                        p.asptblpur1id = Convert.ToInt64("0"+txtasppur1id.Text);
                        Class.Users.UserTime = 0;
                    }
                }
            }
            catch(Exception EX) { throw new Exception(EX.Message); }
          
        }
           
      
        private void BarcodeGenerate_Load(object sender, EventArgs e)
        {

          
            combocompcode.Focus();
            
        }
        private void combostyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            sizeload(combosizegroup.Text);
            allip1.Items.Clear(); allip2.Items.Clear();
        }
        string[] col = { "COLORNAME", "P-COUNT" };
        void load1()
        {
            Utility.Load_Combo(comboprocess, "select asptblpromasid,processname from asptblpromas WHERE active='T' and processname='CUTTING' order by 2  ", "asptblpromasid", "processname");
            Utility.Load_Combo(combobuyer, "select asptblbuymasid,buyercode from asptblbuymas where  active='T' order by 1", "asptblbuymasid", "buyercode");
            Utility.Load_Combo(combosizegroup, "select asptblsizgrpid,sizegroup from asptblsizgrp where  active='T' order by 1", "asptblsizgrpid", "sizegroup");
            Utility.Load_Combo(combostyle, "select asptblstymasid,stylename from asptblstymas where  active='T' order by 1", "asptblstymasid", "stylename");

        }
        public void News()
        {

            empty();
            GridLoad(); 
            compload(); colorload();
            if (allip1.Items.Count <= 0 && allip2.Items.Count <= 0)
            {
                int cnt = 0;
                cnt = col.Length;
                foreach (string str in col)
                {
                    mas.SizeIndex.Add(str);
                }
                GlobalVariables.WidthCols = new Int32[] { 220, 80 };
                if (mas.ColIndex.Count <= 0)
                {
                    CommonFunctions.AddColumn(dataGridView1, mas.GridHeader.ToArray(), mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(), GlobalVariables.WidthCols, cnt);
                }

                mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
            }
            combocompcode.Focus();

        }

        private void empty()
        {
            txtorderno.Enabled = true;
            txtstyleref.Enabled = true;
            txtorderno.Text = "";
            txtstyleref.Text = "";txtexcesqty.Text = "";
            combotype.Enabled = true; combocompcode.Enabled = true;
            combosizegroup.Enabled = true;
            combostyle.Enabled = true;
            combobuyer.Enabled = true;
            txtorderqty.ReadOnly = false;butclose.Enabled = true;
            txtasptblpurid.Text = ""; GlobalVariables.Saves.Visible = true;pictureBox1.Image = null;
            combocompcode.Text = ""; combocompname.Text = "";combotype.Text = "";
            txtasppur1id.Text = ""; txtpono.Text = ""; GlobalVariables.New_Flg = false;
            combobuyer.Text = ""; combosizegroup.Text = ""; txtlotno.Text = "";
            lblpono.Visible = false;combopono1.Visible= false;label15.Text = "";lblgridtotal.Text = "";label15.Text = "";
            combobuyer.Text = ""; combostyle.Text = ""; 
            comboprocess.Text = ""; txtorderqty.Text = "";
            mas.ColIndex.Clear(); mas.SizeIndex.Clear();
            allip2.Items.Clear(); allip1.Items.Clear();
            dataGridView1.Rows.Clear(); dataGridView1.Columns.Clear();
            checksizeall.Checked = false;
            txtasppur1id.Text = "";
            checkactive.Checked = true; txtorderqty.Text = "";
            Class.Users.UserTime = 0;
            butheader.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            panel6.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            listView3.Font = Class.Users.FontName;
            foreach (ListViewItem l2 in listView2.Items)
                l2.Checked = false;
            foreach (ListViewItem l3 in listView3.Items)
                l3.Checked = false;
         
            tabControl2.SelectTab(tabPage3);

        }
        public void Saves()
        {
            if (combocompcode.Text == "")
            {
                MessageBox.Show("Pls Select CompCode is empty " + "CompCode", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combocompcode.Focus();
                return;
            }
            if (combotype.Text == "")
            {
                MessageBox.Show("Pls Select Type is empty " + "Type", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combotype.Focus();
                return;
            }
            //if (pictureBox1.Image==null)
            //{
            //    MessageBox.Show("Pls Upload Garment Image is empty " + " Garment Image ", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}
            if (txtasppur1id.Text == "" && dataGridView1.Rows.Count < 0)
            {
                MessageBox.Show("colorname Name is empty " + " Alert " + txtasppur1id.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (combotype.Text == "ORDER")
            {
                if (Convert.ToInt32("0" + label15.Text) != Convert.ToInt32("0" + txtorderqty.Text))
                {
                    MessageBox.Show("Order Qty Mismatch in Grid Row " + " Alert " + lblgridtotal.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtorderqty.Focus();
                    return;
                }
            }
            else
            {
                if (Convert.ToInt32("0" + label15.Text) != Convert.ToInt32("0" + txtexcesqty.Text))
                {
                    MessageBox.Show("Order Qty Mismatch in Grid Row " + " Alert " + lblgridtotal.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtexcesqty.Focus();
                    return;
                }
            }
            if (Convert.ToInt64(combobuyer.SelectedValue) <= 0)
            {
                MessageBox.Show("BuyerName is empty " + " Alert " + combobuyer.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combobuyer.Focus();
                return;
            }
            if (Convert.ToInt64(combostyle.SelectedValue) <= 0)
            {
                MessageBox.Show("StyleName is empty " + " Alert " + combostyle.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combostyle.Focus();
                return;
            }

            if (txtorderno.Text =="")
            {
                MessageBox.Show("OrderNo is empty " + " Alert " + txtorderno.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtorderno.Focus();
                return;
            }
            if (txtstyleref.Text == "")
            {
                MessageBox.Show("Style Reference is empty " + " Alert " + txtstyleref.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtstyleref.Focus();
                return;
            }
            if (Convert.ToInt64(combosizegroup.SelectedValue) <= 0)
            {
                MessageBox.Show("SizeGroup Name is empty " + " Alert " + combosizegroup.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combosizegroup.Focus();
                return;
            }
            if (txtorderqty.Text == "")
            {
                MessageBox.Show("OrderQty is empty " + " Alert " + txtorderqty.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtorderqty.Focus();
                return;
            }

            if (txtasppur1id.Text != "" && dataGridView1.Rows.Count > 0)
            {
                string maxid = ""; 
                Cursor.Current = Cursors.WaitCursor;
                p.asptblpurid = Convert.ToInt64("0" + txtasptblpurid.Text);
                p.asptblpur1id = Convert.ToInt64("0" + txtasppur1id.Text);
                p.shortcode = Convert.ToString(txtshortcode.Text);
                p.finyear = Class.Users.Finyear;
                p.podate = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                p.pono = Convert.ToString(txtpono.Text);
                p.lotno = Convert.ToString(txtlotno.Text);
                p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
                p.excessqty = Convert.ToInt64("0" + txtexcesqty.Text);
                p.sizegroup = Convert.ToInt64("0" + combosizegroup.SelectedValue);
                p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                p.processname = Convert.ToInt64("0" + comboprocess.SelectedValue);
                p.processtype = Convert.ToString(combotype.Text);
                p.issuetype = "PANEL MISTAKE";
                p.panelmistake = "F";
                p.compcode1 = Class.Users.COMPCODE;
                p.username = Class.Users.USERID;
                p.createdby = Convert.ToString(Class.Users.HUserName);
                p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                p.modified = Convert.ToString(dateTimePicker1.Value.ToString());
                p.modified1 = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                p.modifiedby = Class.Users.HUserName;
                p.ipaddress = Class.Users.IPADDRESS;
                p.barcode =Convert.ToInt64("0"+txtautono.Text);
                p.orderno = Convert.ToString(txtorderno.Text.ToUpper());
                p.styleref = Convert.ToString(txtstyleref.Text.ToUpper());
                if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; }
                if (checkpo.Checked == true) { p.pocancel = "T"; } else { p.pocancel = "F"; }
                string sel0 = "select asptblpurid    from  asptblpur   where  barcode='"+p.barcode+"' and asptblpur1id='" + p.asptblpur1id+"' and shortcode='" + p.shortcode+"' and finyear ='" + p.finyear+"' and podate='" + p.podate+"'  and stylename='" + p.stylename+"' and orderqty ='" + p.orderqty+"' and  buyer='" + p.buyer+"'  and processname ='" + p.processname+"' and processtype='" + p.processtype+"' and lotno ='" + p.lotno+"' and pocancel ='" + p.pocancel+"' and active='" + p.active+ "' and excessqty='" + p.excessqty+ "' and orderno='" + p.orderno + "' and styleref='" + p.styleref + "'";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "asptblpur");
                DataTable dt0 = ds0.Tables["asptblpur"];
                if (txtasptblpurid.Text=="") { p.asptblpurid = 0; }
                else
                {
                    p.asptblpurid = Convert.ToInt64("0" + txtasptblpurid.Text);
                }
                if (dt0.Rows.Count != 0) { }
                else if (dt0.Rows.Count != 0 && p.asptblpurid == 0 || p.asptblpurid == 0)
                {
                    auto();
                    string ins = "insert into asptblpur(asptblpur1id,shortcode,finyear,podate,compcode,stylename,orderqty,pono,buyer,sizegroup,processname,processtype,lotno,pocancel,active,compcode1,username,createdby,modifiedby,modified,ipaddress,issuetype,excessqty,barcode,orderno,styleref,modified1)  VALUES('" + p.asptblpur1id + "','" + p.shortcode + "','" + p.finyear + "','" + p.podate + "','" + p.compcode + "','" + p.stylename + "','" + p.orderqty + "','" + p.pono + "','" + p.buyer + "','" + p.sizegroup + "','" + p.processname + "','" + p.processtype + "','" + p.lotno + "','" + p.pocancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + p.modifiedby + "','" + p.modified + "','" + p.ipaddress + "','" + p.issuetype + "','" + p.excessqty + "','" + p.barcode + "','" + p.orderno + "','" + p.styleref + "',date_format('" + p.modified1 + "','%Y-%m-%d'))";
                    Utility.ExecuteNonQuery(ins);
                    string sel2 = "select max(asptblpur1id) id    from  asptblpur   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and pono='" + txtpono.Text + "' ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
                    DataTable dt2 = ds2.Tables["asptblpur"];
                    maxid = dt2.Rows[0]["id"].ToString();
                    if (dt2.Rows.Count > 0 && stdbytes != null)
                    {
                        MySqlCommand cmd;
                        string ins1 = "UPDATE  asptblpur SET imagebytes='" + stdbytes.Length.ToString() + "',garmentimage=@garmentimage where  asptblpur1id='" + dt2.Rows[0]["ID"].ToString() + "'";
                        cmd = new MySqlCommand(ins1, Utility.Connect());
                        cmd.Parameters.Add(new MySqlParameter("@garmentimage", stdbytes));
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    string up = "update  asptblpur  set   compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "',orderno='" + p.orderno + "',styleref='" + p.styleref + "' ,modified1=date_format('" + p.modified1.ToString().Substring(0, 10) + "','%Y-%m-%d') where asptblpurid='" + p.asptblpurid + "'";
                    Utility.ExecuteNonQuery(up);
                    maxid = txtasptblpurid.Text;
                    MySqlCommand cmd;
                    if (stdbytes != null)
                    {
                        string ins1 = "UPDATE  asptblpur SET imagebytes='" + stdbytes.Length.ToString() + "',garmentimage=@garmentimage where  asptblpur1id='" + txtasppur1id.Text + "'";
                        cmd = new MySqlCommand(ins1, Utility.Connect());
                        cmd.Parameters.Add(new MySqlParameter("@garmentimage", stdbytes));
                        cmd.ExecuteNonQuery();
                    }
                }
                
                int i = 0, j = 0;
                if (dataGridView1.Rows.Count >= 0 && txtasptblpurid.Text =="")
                {
                    Models.PurchasesModeldetail pp = new Models.PurchasesModeldetail();
                    string sel2 = "select max(asptblpurid) id ,max(barcode) barcode   from  asptblpur   where  compcode='" + p.compcode + "'  and finyear='" + Class.Users.Finyear + "' and pono='" + txtpono.Text + "'";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
                    DataTable dt2 = ds2.Tables["asptblpur"];
                    maxid = dt2.Rows[0]["id"].ToString();
                 
                     bar1 =Convert.ToInt64("0"+dt2.Rows[0]["barcode"].ToString());
                    for (i = 0; i < dataGridView1.Columns.Count - 2; i++)
                    {
                        for (j = 0; j < dataGridView1.Rows.Count; j++)
                        {
                            pp.barcode = Convert.ToInt64(bar1)+ Convert.ToInt64(1);
                            colorid(dataGridView1.Rows[j].Cells[0].Value.ToString());
                            pp.colorname = coid;
                            if (Convert.ToInt64("0" + dataGridView1.Rows[j].Cells[1].Value.ToString()) <= 0)
                            {
                                pp.portion = 1;
                            }
                            else
                            {
                                pp.portion = Convert.ToInt64("0" + dataGridView1.Rows[j].Cells[1].Value.ToString());
                            }
                                sizeid(dataGridView1.Columns[i + 2].HeaderText.ToString());
                            pp.sizename = siid;
                            pp.orderqty = Convert.ToInt64("0" + dataGridView1.Rows[j].Cells[i + 2].Value.ToString());
                            pp.asptblpurdetid = Convert.ToInt64("0" + dataGridView1.Rows[j].Cells[1].Value.ToString());
                            pp.asptblpurid = Convert.ToInt64("0" + maxid);
                            pp.asptblpur1id = Convert.ToInt64(txtasppur1id.Text);
                            pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                            pp.pono = Convert.ToString(txtpono.Text);
                            string sel1 = "select asptblpurdetid   from  asptblpurdet   where   asptblpurid='" + pp.asptblpurid + "' AND   asptblpur1id='" + pp.asptblpur1id + "' and compcode='" + p.compcode + "' and  pono='" + pp.pono + "'  and  colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "'  ";// and portion = '" + pp.portion + "' and orderqty = '" + pp.orderqty + "' 
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                            DataTable dt1 = ds1.Tables["asptblpur"];
                            if (dt1.Rows.Count <= 0) { pp.asptblpurdetid = 0; }
                            else
                            {
                                pp.asptblpurdetid = Convert.ToInt64("0" + dt1.Rows[0]["asptblpurdetid"].ToString());
                            }
                            if (dt1.Rows.Count != 0 || pp.asptblpurdetid > 0)
                            {
                                string up1 = "update  asptblpurdet  set barcode='" + pp.barcode + "' ,asptblpurid='" + pp.asptblpurid + "' ,asptblpur1id='" + pp.asptblpur1id + "', compcode='" + p.compcode + "' , pono='" + pp.pono + "' , colorname='" + pp.colorname + "', portion='" + pp.portion + "', sizename='" + pp.sizename + "', orderqty='" + pp.orderqty + "' where asptblpurid='" + pp.asptblpurid + "' AND  compcode='" + p.compcode + "'  AND pono='" + pp.pono + "'  AND colorname='" + pp.colorname + "' AND sizename='" + pp.sizename + "'";
                                Utility.ExecuteNonQuery(up1);
                              
                            }
                            else
                            {
                                string ins1 = "insert into asptblpurdet(barcode,asptblpurid,asptblpur1id,compcode,pono,colorname,portion,sizename,orderqty) values('" + pp.barcode+ "' ,'" + maxid.ToString() + "' ,'" + txtasppur1id.Text + "' , '" + combocompcode.SelectedValue + "' ,'" + pp.pono + "' , '" + pp.colorname + "','" + pp.portion + "','" + pp.sizename + "','" + pp.orderqty + "');";
                                Utility.ExecuteNonQuery(ins1);
                               
                            }
                            
                        }
                    }
                    barcodeload();
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
        Int64 bar1 = 0;
        private void coladd(DataGridView grid)
        {
            foreach (DataGridViewColumn col1 in grid.Columns)
            {
                Class.Users.QrCodeArray.Columns.Add(col1.HeaderText);
            }
        }
        private void rowadd(int inx,int cel,DataGridView grid)
        {
            for (inx = 0; inx < grid.Rows.Count; inx++)
            {
                Class.Users.QrCodeArray.Rows.Add(); cel = 0;
                foreach (DataGridViewCell cell in grid.Rows[inx].Cells)
                {
                    Class.Users.QrCodeArray.Rows[Class.Users.QrCodeArray.Rows.Count - 1][grid.Columns[cel].HeaderText] = grid.Rows[inx].Cells[cel].Value.ToString();
                    cel++;
                }
            }
        }
        private void barcodeload()
        {

            int i = 0;
            if (Class.Users.QrCodeArray.Rows.Count > 0)
            {
                Class.Users.QrCodeArray.Rows.Clear();
                Class.Users.QrCodeArray.Columns.Clear();
            }
            int rowindex = 0; int cel = 0;           
            coladd(dataGridView1);
            rowadd(rowindex,cel,dataGridView1);   
            Class.Users.Query = "select asptblpurdet1id,asptblpurdetid,asptblpurid,asptblpur1id,compcode,pono,colorname,sizename,orderqty,PORTION,colorname1,qrcode from asptblpurdet1  where asptblpurdet1id<0" + ":'" + combocompcode.SelectedValue + "':'" + Class.Users.Finyear + "':'" + txtpono.Text + "':'asptblpurdet1'";
            string[] sarray = Class.Users.Query.Split(':');
            Class.Users.QrCode = com.select(sarray[0], sarray[4]);
            int n = 2;
            int cnt = 0, k = 0, cnt1 = 0, cnt2 = 1, row = 0, col = 0, rowcount = 0, tot = 0;
            col = Class.Users.QrCodeArray.Columns.Count - 2;
            int totcount = 0;
                rowcount = Class.Users.QrCodeArray.Rows.Count;
            int cc = 0;  
            for (int i1 = 0; i1 < rowcount; i1++)
            {
                tot = 0; row = 0; cc++;
                cnt1 = 0; row = 2;int cnt5 = 0;
                
                for (i = 0; i < col; i++)
                {
                    if (col != tot)
                    {
                        cnt2 = cc - 1; cnt1 = i + row; cnt = 0; int cnt4 = 0; 
                        int cnt3 = Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2]["P-COUNT"].ToString());
                        cnt = Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2]["P-COUNT"].ToString());/*Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString());*/
                        cnt5 = Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString());
                        if (cnt >= 1)
                        {
                            k++;
                            int id1 = 0, id2 = 0, id3 = 0; string id4 = "";
                            Class.Users.Query = "select a.asptblpurdetid ,a.asptblpurid,b.asptblpur1id,b.pono   from  asptblpurdet a join asptblpur b on a.asptblpurid=b.asptblpurid join asptblsizmas  c on c.asptblsizmasid=a.sizename  where  b.compcode=" + sarray[1] + "  and b.finyear=" + sarray[2] + " and b.pono=" + sarray[3] + " and c.sizename='" + Class.Users.QrCodeArray.Columns[cnt1].ToString() + "' ";
                            Class.Users.dt = com.select(Class.Users.Query, "asptblpurdet");
                            if (Class.Users.dt.Rows.Count > 0)
                            {

                                id1 = Convert.ToInt32(Class.Users.dt.Rows[i1]["asptblpurdetid"].ToString());
                                id2 = Convert.ToInt32(Class.Users.dt.Rows[i1]["asptblpurid"].ToString());
                                id3 = Convert.ToInt32(Class.Users.dt.Rows[i1]["asptblpur1id"].ToString());
                                id4 = Convert.ToString(Class.Users.dt.Rows[i1]["pono"].ToString());
                            }
                            int mi = 1; k = 1;
                            if (cnt3 == 1 || cnt5==1) { cnt4 = cnt5; } else { cnt4 = 0;cnt4=cnt5; }
                            for (int a = 0; a < cnt4; a++)
                            {
                                int totcount1 = 0; 
                                for (int b = 0; b < cnt3; b++)
                                {
                                    
                                    if (totcount1 < cnt3)
                                    {
                                        totcount++; totcount1++;
                                        Class.Users.QrCode.Rows.Add(0);
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][1] = id1;
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][2] = id2;
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][3] = id3;
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][4] = Class.Users.COMPCODE;
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][5] = id4;
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][6] = Class.Users.QrCodeArray.Rows[i1]["COLORNAME"].ToString(); //GenerateQrCode(Class.Users.QrCodeArray.Rows[i]["COLORNAME"].ToString());
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][7] = Class.Users.QrCodeArray.Columns[cnt1].ToString();
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][8] = Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString();
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][8] = Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString();
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][9] = Class.Users.QrCodeArray.Rows[cnt2]["P-COUNT"].ToString();
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][10] = Class.Users.QrCodeArray.Rows[i1]["COLORNAME"].ToString() + " " + i1 + " " + k.ToString() + " " + Class.Users.QrCodeArray.Columns[cnt1].ToString() + " " + Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString() + " " + totcount;
                                        Class.Users.QrCode.Rows[Class.Users.QrCode.Rows.Count - mi][11] = bar1;
                                        lblcount.Refresh();
                                        lblcount.Text = totcount.ToString();
                                    }
                                    
                                }
                                bar1++; k++;
                            }
                            tot++;
                           
                        }
                    }
                }
            }
            if (txtasppur1id.Text != "")
            {
                Models.PurchasesModel1detail p3 = new Models.PurchasesModel1detail();
                for (int m = 0; m < Class.Users.QrCode.Rows.Count; m++)
                {
                  
                    p3.sno = Convert.ToInt64("0" + m.ToString());
                    p3.asptblpurdetid = Convert.ToInt64("0" + Class.Users.QrCode.Rows[m]["asptblpurdetid"].ToString());
                    p3.asptblpurid = Convert.ToInt64("0" + Class.Users.QrCode.Rows[m]["asptblpurid"].ToString());
                    p3.asptblpur1id = Convert.ToInt64("0" + Class.Users.QrCode.Rows[m]["asptblpur1id"].ToString());
                    p3.compcode = Convert.ToInt64("0" + Class.Users.QrCode.Rows[m]["compcode"].ToString());
                    p3.pono = Class.Users.QrCode.Rows[m]["pono"].ToString();
                    p3.colorname = Class.Users.QrCode.Rows[m]["colorname"].ToString();
                    p3.sizename = Class.Users.QrCode.Rows[m]["sizename"].ToString();
                    p3.portion = Convert.ToInt64("0" + Class.Users.QrCode.Rows[m]["portion"].ToString());
                    p3.orderqty = Convert.ToInt64("0" + Class.Users.QrCode.Rows[m]["orderqty"].ToString());
                    p3.colorname1 = Convert.ToString(Class.Users.QrCode.Rows[m]["colorname1"].ToString());
                    p3.barcode = Convert.ToInt64("0" + Class.Users.QrCode.Rows[m]["qrcode"].ToString());
                    p3.finyear = Class.Users.Finyear;
                    p3.cutting = "F";
                    p3.stitching = "F";
                    p3.checking = "F";
                    p3.restitching = "F";
                    p3.rechecking = "F";
                    p3.inward = "F";
                    p3.delivery = "F";
                    p.panelmistake = "F";
                    p.modified = Convert.ToString(dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                    Class.Users.dt = p3.Select(p3.barcode, p3.asptblpurdetid, p3.asptblpurid, p3.asptblpur1id, p3.compcode, p3.pono, p3.colorname, p3.portion, p3.sizename, p3.sno, p3.orderqty);
                    if (Class.Users.dt.Rows.Count <= 0) { p3.asptblpurdet1id = 0; }
                    else
                    {
                        p3.asptblpurdet1id = Convert.ToInt64("0" + Class.Users.dt.Rows[0]["asptblpurdet1id"].ToString());
                    }
                    if (Class.Users.dt.Rows.Count != 0) { }
                    else if (Class.Users.dt.Rows.Count != 0 && p3.asptblpurdet1id == 0 || p3.asptblpurdet1id == 0)
                    {
                        p3 = new PurchasesModel1detail(p3.barcode, p3.asptblpurdetid, p3.asptblpurid, p3.asptblpur1id, p3.compcode, p3.pono, p3.colorname, p3.portion, p3.sizename, p3.orderqty, p3.colorname1, p3.finyear, p3.cutting, p3.stitching, p3.checking, p3.restitching, p3.rechecking, p3.sno, p3.inward, p3.delivery, p.panelmistake, p.modified);
                        lblcount.Refresh();
                        lblcount.Text = "Inserted " + m;
                    }
                    else
                    {
                        p3 = new PurchasesModel1detail(p3.barcode, p3.asptblpurid, p3.compcode, p3.pono, p3.colorname, p3.portion, p3.sizename, p3.orderqty, p3.restitching, p3.rechecking, p3.inward, p3.delivery, p.panelmistake, p3.modified);
                        lblcount.Refresh();
                        lblcount.Text = "Updated " + m;
                    }
                }
                
            }
            Class.Users.Query = "select max(a.asptblpurid) as asptblpurid from asptblpur a where a.compcode='" + combocompcode.SelectedValue + "' and a.finyear='" + Class.Users.Finyear + "' and a.pono='" + txtpono.Text + "' ";
            DataTable dt = com.select(Class.Users.Query, "asptblpurid");
            string s = "select max(a.barcode) as barcode,max(a.asptblpurid) as asptblpurid from asptblpurdet1 a where a.compcode='" + combocompcode.SelectedValue + "' and a.finyear='" + Class.Users.Finyear + "'";
            DataTable dt0 = com.select(s, "asptblpurdet1");
            string up = "update  asptblpur  set  barcode='" + dt0.Rows[0]["barcode"].ToString() + "'  where compcode='" + p.compcode + "' and finyear='" + p.finyear + "' and asptblpurid='" + dt0.Rows[0]["asptblpurid"].ToString() + "'";
            Utility.ExecuteNonQuery(up);
            Class.Users.Query = "select cast(a.barcode as char) AS QRCODE,a.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode,a.pono,a.colorname,a.sizename,a.orderqty,a.PORTION,a.colorname1 from asptblpurdet1 a where a.compcode='" + combocompcode.SelectedValue + "'  and a.pono='" + txtpono.Text + "' AND a.asptblpurid='" + dt.Rows[0]["asptblpurid"].ToString() + "'";

            Class.Users.QrCode = null;
            Class.Users.QrCode = com.select(Class.Users.Query, "asptblpurdet1");
            Cursor.Current = Cursors.Default;
            Pinnacle.ReportFormate.Lyla.LylaReport ly = new ReportFormate.Lyla.LylaReport();
            ly.Show();
        }

       

        private void BarcodeGenerate_FormClosed(object sender, FormClosedEventArgs e)
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
                string sel1 = " SELECT c.asptblcolmasid,c.colorname  FROM   asptblcolmas c   where  c.active='T' order by 1 ";
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
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView3.Items.Add(list);
                            i++;
                        }
                        lbltotalcolor.Refresh();
                        lbltotalcolor.Text = "Total Count    :" + listView3.Items.Count;
                    }
                    //panelcolorsize.Visible = true;
                    //panelcolorsize.Width = 450;
                    //panelcolorsize.Height = 300;
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
                string sel1 = "SELECT a.asptblpurid,b.compcode,a.podate,a.pono,f.stylename,e.sizegroup,a.orderqty, a.active FROM  asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblstymas f on f.asptblstymasid=a.stylename   order by  a.asptblpurid desc";
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
                        list.SubItems.Add(myRow["podate"].ToString().Substring(0,10));                       
                        list.SubItems.Add(myRow["pono"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["sizegroup"].ToString());                      
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
                  
                    Cursor.Current = Cursors.WaitCursor; Class.Users.Query = "";
                    txtasptblpurid.Text = listView1.SelectedItems[0].SubItems[2].Text;                    
                    Class.Users.Query = "select a.asptblpurid,a.asptblpur1id,a.finyear,a.podate,b.compcode,f.asptblstymasid,f.stylename,a.pono,b.compname,C.asptblbuymasid,c.buyercode,a.orderqty,a.excessqty,d.sizegroup,e.processname,a.processtype,a.lotno,a.pocancel,a.active,a.garmentimage,a.orderno,a.styleref  from  asptblpur a join gtcompmast b on b.gtcompmastid=a.compcode join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblsizgrp d on d.asptblsizgrpid=a.sizegroup join asptblpromas e on e.asptblpromasid=a.processname  JOIN asptblstymas  F on f.asptblstymasid=a.stylename   where a.asptblpurid=" + txtasptblpurid.Text;
                    Class.Users.dt = com.select(Class.Users.Query, "asptblpur");
                    if (Class.Users.dt.Rows.Count > 0)
                    {
                        txtasppur1id.Text = Convert.ToString(Class.Users.dt.Rows[0]["asptblpur1id"].ToString());
                        combocompcode.Text = Convert.ToString(Class.Users.dt.Rows[0]["compcode"].ToString());
                        txtpono.Text = Convert.ToString(Class.Users.dt.Rows[0]["pono"].ToString());
                        combopono1.Text= Convert.ToString(Class.Users.dt.Rows[0]["pono"].ToString());
                        combobuyer.SelectedValue = Convert.ToString(Class.Users.dt.Rows[0]["asptblbuymasid"].ToString());
                        combostyle.DisplayMember = "stylename";
                        combostyle.ValueMember = "asptblstymasid";
                        combostyle.DataSource = Class.Users.dt;
                        combobuyer.DisplayMember = "buyercode";
                        combobuyer.ValueMember = "asptblbuymasid";
                        combobuyer.DataSource = Class.Users.dt;

                        txtorderno.Text = Convert.ToString(Class.Users.dt.Rows[0]["orderno"].ToString());
                        txtstyleref.Text = Convert.ToString(Class.Users.dt.Rows[0]["styleref"].ToString());
                      dateTimePicker1.Value= Convert.ToDateTime(Class.Users.dt.Rows[0]["podate"].ToString());
                        int c = Class.Users.dt.Rows.Count;
                        if (c >= 1)
                        {
                            if (c >= 1 && Class.Users.dt.Rows[0]["garmentimage"].ToString() != "")
                            {
                                pictureBox1.Image = null; stdbytes = null;
                                stdbytes = (byte[])Class.Users.dt.Rows[0]["garmentimage"];
                                Image img = Models.Device.ByteArrayToImage(stdbytes);
                                pictureBox1.Image = img;
                            }
                            else
                            {
                                pictureBox1.Image = null; stdbytes = null;
                            }
                        }

                        combotype.Text = Class.Users.dt.Rows[0]["processtype"].ToString();
                        if (combotype.Text == "ORDER")
                        {
                            string sel1 = "select a.asptblpurid,a.asptblpur1id,a.finyear,a.podate,b.compcode,f.stylename,a.pono,b.compname,c.buyercode,a.orderqty,d.sizegroup,e.processname,a.lotno,a.pocancel,a.active  from  asptblpur a join gtcompmast b on b.gtcompmastid=a.compcode join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblsizgrp d on d.asptblsizgrpid=a.sizegroup join asptblpromas e on e.asptblpromasid=a.processname   JOIN asptblstymas  F on f.asptblstymasid=a.stylename  where  a.asptblpurid=" + txtasptblpurid.Text;
                            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                            DataTable dt = ds.Tables["asptblpur"];
                            if (dt.Rows.Count > 0)
                            {

                                combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                                combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());                             
                                combosizegroup.Text = Convert.ToString(dt.Rows[0]["sizegroup"].ToString());
                                sizeload(combosizegroup.Text);
                                comboprocess.Text = Convert.ToString(dt.Rows[0]["processname"].ToString());


                            }
                        }
                        else
                        {
                            string sel1 = "select a.asptblpurid,a.asptblpur1id,a.finyear,a.podate,b.compcode,f.stylename,a.pono,b.compname,c.buyercode,a.orderqty,d.asptblsizgrpid,d.sizegroup,e.asptblpromasid,e.processname,a.lotno,a.pocancel,a.active  from  asptblpur a join gtcompmast b on b.gtcompmastid=a.compcode join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblsizgrp d on d.asptblsizgrpid=a.sizegroup join asptblpromas e on e.asptblpromasid=a.processname   JOIN asptblstymas  F on f.asptblstymasid=a.stylename  where  a.asptblpurid=" + txtasptblpurid.Text;
                            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                            DataTable dt = ds.Tables["asptblpur"];
                            if (dt.Rows.Count > 0)
                            {

                                combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                                combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                                combopono1.SelectedValue = Convert.ToString(dt.Rows[0]["pono"].ToString());
                                txtpono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                               // combostyle.SelectedValue = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                               // combobuyer.SelectedValue = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                                combosizegroup.Text = Convert.ToString(dt.Rows[0]["sizegroup"].ToString());
                                combosizegroup.DisplayMember = "sizegroup";
                                combosizegroup.ValueMember = "asptblsizgrpid";
                                combosizegroup.DataSource = dt;
                                sizeload(combosizegroup.Text);
                                comboprocess.DisplayMember = "processname";
                                comboprocess.ValueMember = "asptblpromasid";
                                comboprocess.DataSource = dt;
                                //comboprocess.Text = Convert.ToString(dt.Rows[0]["processname"].ToString());
                            }
                        }
                        txtlotno.Text = Convert.ToString(Class.Users.dt.Rows[0]["lotno"].ToString());
                        txtorderqty.Text = Convert.ToString(Class.Users.dt.Rows[0]["orderqty"].ToString());
                        txtexcesqty.Text = Convert.ToString(Class.Users.dt.Rows[0]["excessqty"].ToString());
                        txtasptblpurid.Text = Convert.ToString(Class.Users.dt.Rows[0]["asptblpurid"].ToString());
                        if (Class.Users.dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        string sel2 = "select distinct d.colorname,a.asptblpurdetid from asptblpurdet a join asptblpur b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizmas e on e.asptblsizmasid=a.sizename  where c.compcode='" + combocompcode.Text + "' and a.asptblpurid='" + Class.Users.dt.Rows[0]["asptblpurid"].ToString() + "' order by 2";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpur");
                        DataTable dt21 = ds2.Tables["asptblpur"];
                        foreach (DataRow row1 in dt21.Rows)
                        {
                            foreach (ListViewItem item0 in listView2.Items)
                            {
                                if (item0.SubItems[3].Text == row1["colorname"].ToString())
                                {
                                    item0.Checked = true;
                                }
                            }
                        }
                        string sel4 = "select distinct d.ASPTBLSIZMASID,d.SIZENAME,b.asptblpurdetid from asptblpur a join asptblpurdet b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=a.compcode join asptblsizmas d on d.ASPTBLSIZMASID=b.sizename join asptblcolmas e on e.asptblcolmasid=b.colorname  where c.compcode='" + combocompcode.Text + "' AND A.asptblpurid='" + Class.Users.dt.Rows[0]["asptblpurid"].ToString() + "' order by 3";
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
                        int i = 0, j = 0;
                        int k = 1;
                        int col = dataGridView1.Columns.Count - 2;
                        for (i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            string sel5 = "select  b.portion,d.ASPTBLSIZMASID,  d.SIZENAME,b.orderqty,b.asptblpurdetid from asptblpur a join asptblpurdet b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=a.compcode join asptblsizmas d on d.ASPTBLSIZMASID=b.sizename join asptblcolmas e on e.asptblcolmasid=b.colorname  where c.compcode='" + combocompcode.Text + "' AND A.pono='" + txtpono.Text + "' AND A.asptblpurid='" + Class.Users.dt.Rows[0]["asptblpurid"].ToString() + "' and e.colorname='" + dataGridView1.Rows[i].Cells[0].EditedFormattedValue.ToString() + "'  order by 5";//and d.sizename='" + dataGridView1.Columns[i + 1].HeaderText.ToString() + "'
                            DataSet ds5 = Utility.ExecuteSelectQuery(sel5, "asptblpur");
                            DataTable dt5 = ds5.Tables["asptblpur"];
                            if (col <= i)
                            {
                                dataGridView1.Rows[i].Cells[k].Value = dt5.Rows[0]["portion"].ToString();
                            }
                            else
                            {
                                dataGridView1.Rows[i].Cells[k].Value = dt5.Rows[i]["portion"].ToString();
                            }
                            int ci = 2;
                            for (j = 0; j < dt5.Rows.Count; j++)
                            {
                                dataGridView1.Rows[i].Cells[j + 2].Value = dt5.Rows[j]["orderqty"].ToString();
                                GridEdit(ci, dataGridView1);
                                ci++;
                            }
                        }
                    }
                    else
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Invalid");
                    }
                }
                else
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("No Data Found in ListView");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            Models.CommonClass CC = new Models.CommonClass();
            string po = CC.HideButton(txtpono.Text, "asptblprolot");
            if (po != "")
            {
                if (Class.Users.HUserName == "ADMIN" || Class.Users.HUserName == "VAIRAM")
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
            combotype.Enabled = false;combocompcode.Enabled = false;
            combosizegroup.Enabled = false; 
            combostyle.Enabled = false; combobuyer.Enabled = false;
            txtorderqty.ReadOnly = true; butclose.Enabled = false;
            tabControl1.SelectTab(tabPage1);
            combocompcode.Select(); Cursor.Current = Cursors.Default;
        }

        private void listView2_ItemChecked1(object sender, ItemCheckedEventArgs e)
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
        private void butAdd_Click(object sender, EventArgs e)
        {
           
            GlobalVariables.WidthCols = new Int32[] {220, 80};
            if (allip1.Items.Count > 0 && allip2.Items.Count > 0)
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
                if (mas.SizeIndex.Count > 0)
                {
                    CommonFunctions.AddColumn(dataGridView1, mas.GridHeader.ToArray(), mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(), GlobalVariables.WidthCols, cnt);
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
                        if (item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[7].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[8].ToString().Contains(txtsearch.Text))
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
                    listView1.Items.Clear(); //listView1.BackColor = System.Drawing.Color.LightGray;
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
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtasppur1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {
                    if (txtasptblpurid.Text != "")
                    {
                        com.query = "delete from asptblpur where asptblpurid=" + txtasptblpurid.Text;
                        Utility.ExecuteNonQuery(com.query);
                        com.query = "delete from asptblpurdet where asptblpurid=" + txtasptblpurid.Text;
                        Utility.ExecuteNonQuery(com.query);
                        com.query = "delete from asptblpurdet1 where asptblpurid=" + txtasptblpurid.Text;
                        Utility.ExecuteNonQuery(com.query);
                        MessageBox.Show("Record Deleted Successfully " + txtasppur1id.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        MessageBox.Show("Invalid." + txtasppur1id.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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
                    listView2.Items.Clear();
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
            combocompname.DataSource = dt; Class.Users.UserTime = 0;     
        }
      
        public void Prints()
        {            
            Class.Users.Query = "select cast(a.BARCODE as char) AS QRCODE,a.asptblpurdetid,a.asptblpurid,a.asptblpur1id,a.compcode,a.pono,a.colorname,a.sizename,a.orderqty,a.PORTION,a.colorname1 from asptblpurdet1 a where a.compcode='" + combocompcode.SelectedValue + "'  AND a.asptblpurid='" + txtasptblpurid.Text + "'";
            Class.Users.QrCode = null;
            Class.Users.QrCode = com.select(Class.Users.Query, "asptblpurdet1");
            Pinnacle.ReportFormate.Lyla.LylaReport ly = new ReportFormate.Lyla.LylaReport();
            ly.Show();
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
            throw new NotImplementedException();
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            News();
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
            //CommonFunctions.SetRowNumber(dataGridView1);
        }

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();
        }


        //private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        //{
        //    e.Control.KeyPress -= new KeyPressEventHandler(Column1_KeyPress);
        //    if (dataGridView1.CurrentCell.ColumnIndex > 0) //Desired Column
        //    {
        //        TextBox tb = e.Control as TextBox;
        //        if (tb != null)
        //        {
        //            tb.KeyPress += new KeyPressEventHandler(Column1_KeyPress);
        //        }
        //    }
        //    Class.Users.UserTime = 0;
        //}
        private void Column1_KeyPress(object sender, KeyPressEventArgs e)
        {            
            //if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) || (sender as TextBox).Text.Length >= 8)
            //{
            //    e.Handled = true;
            //}
            
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void butshow_Click(object sender, EventArgs e)
        {
           
        }

        private void butclose_Click(object sender, EventArgs e)
        {
          
            Class.Users.UserTime = 0;
            butAdd_Click(sender,e);
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void comboprocess_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkactive_CheckedChanged(object sender, EventArgs e)
        {

        }
        void FieldDisable()
        {
            txtpono.Text = "";
            txtexcesqty.Enabled = true;
            combobuyer.Enabled = false;
            combostyle.Enabled = false;
            combostyle.Enabled = false;
            txtorderqty.Enabled = false;
            combosizegroup.Enabled = false;
            comboprocess.Enabled=false;
             txtlotno.Enabled = false;
        }
        void FieldEnable()
        {
           
            combobuyer.Enabled = true;
            //combostyle.Enabled = true;
            //combostyle.Enabled = true;
            txtorderqty.Enabled = true;
            combosizegroup.Enabled = true;
            comboprocess.Enabled = false;
            txtlotno.Enabled = true;
        }
        private void combotype_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (combotype.Text == "EXCESS QTY")
                {
                    txtexcesqty.Enabled = true;
                    lblpono.Visible = true; txtorderno.Enabled = false;
                    txtstyleref.Enabled = false;
                    combopono1.Visible = true; 
                    FieldDisable(); 
                   
                    Utility.Load_Combo(combopono1, "select a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  where  a.active='T' and b.compcode='" + Class.Users.HCompcode + "'  order by a.asptblpurid desc", "pono", "pono");

                }
                if (combotype.Text == "ORDER")
                {
                    txtorderno.Enabled = true;
                    txtstyleref.Enabled = true;
                    if (txtasptblpurid.Text == "")
                    {
                       
                        txtasppur1id.Text = ""; txtpono.Text = ""; GlobalVariables.New_Flg = false;
                        combobuyer.Text = ""; combosizegroup.Text = ""; txtlotno.Text = "";
                        lblpono.Visible = false; combopono1.Visible = false; label15.Text = ""; lblgridtotal.Text = ""; label15.Text = "";
                        combobuyer.Text = ""; //combostyle.Text = ""; comboprocess.Text = ""; txtorderqty.Text = "";
                        dataGridView1.Rows.Clear(); dataGridView1.Columns.Clear();
                        checksizeall.Checked = false;
                        txtasppur1id.Text = "";
                        checkactive.Checked = true; txtorderqty.Text = "";

                        if (allip1.Items.Count <= 0 && allip2.Items.Count <= 0)
                        {
                            int cnt = 0;
                            cnt = col.Length;
                            foreach (string str in col)
                            {
                                mas.SizeIndex.Add(str);
                            }
                            GlobalVariables.WidthCols = new Int32[] { 220, 80 };
                            if (mas.ColIndex.Count <= 0)
                            {
                                CommonFunctions.AddColumn(dataGridView1, mas.GridHeader.ToArray(), mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(), GlobalVariables.WidthCols, cnt);
                            }

                            mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
                        }
                    }
                    txtexcesqty.Enabled = false;
                    lblpono.Visible = false;
                    combopono1.Visible = false;
                    FieldEnable();
                    load1();
                    auto(); sizeload(combosizegroup.Text);
                }
                else
                {
                   
                    //combobuyer.Text = ""; combosizegroup.Text = ""; txtlotno.Text = ""; txtasptblpurid.Text = "";
                    //combobuyer.Text = ""; combostyle.Text = ""; comboprocess.Text = ""; txtorderqty.Text = "";
                }
            }
            catch (Exception ex) {  }
        }

        private void combopono1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            if (combopono1.Text != "" && combotype.Text== "EXCESS QTY")
            {
                mas.ColIndex.Clear(); mas.SizeIndex.Clear();
                allip2.Items.Clear(); allip1.Items.Clear();
                dataGridView1.Rows.Clear(); dataGridView1.Columns.Clear();
                compload(); colorload(); sizeload(combosizegroup.Text);
                DataTable dt = new DataTable();
                string sel1 = "select a.asptblpurid,a.asptblpur1id,a.finyear,a.podate,b.compcode,f.asptblstymasid,f.stylename,a.pono,b.compname,c.asptblbuymasid,c.buyercode,a.orderqty,d.asptblsizgrpid,d.sizegroup,e.asptblpromasid,e.processname,a.lotno,a.pocancel,a.active,a.orderno,a.styleref  from  asptblpur a join gtcompmast b on b.gtcompmastid=a.compcode join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblsizgrp d on d.asptblsizgrpid=a.sizegroup join asptblpromas e on e.asptblpromasid=a.processname  JOIN asptblstymas  F on f.asptblstymasid=a.stylename   where a.pono='" + combopono1.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpur");
                dt = ds.Tables["asptblpur"];
                if (dt.Rows.Count > 0)
                {
                    GlobalVariables.New_Flg = false;
                    combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                    auto();
                    combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                  //  combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                   // combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                    combosizegroup.Text = Convert.ToString(dt.Rows[0]["sizegroup"].ToString());
                    combostyle.DisplayMember = "stylename";
                    combostyle.ValueMember = "asptblstymasid";
                    combostyle.DataSource = dt;
                    combobuyer.DisplayMember = "buyercode";
                    combobuyer.ValueMember = "asptblbuymasid";
                    combobuyer.DataSource = dt;
                    combosizegroup.DisplayMember = "sizegroup";
                    combosizegroup.ValueMember = "asptblsizgrpid";
                    combosizegroup.DataSource = dt;
                    sizeload(combosizegroup.Text);
                    comboprocess.DisplayMember = "processname";
                    comboprocess.ValueMember = "asptblpromasid";
                    comboprocess.DataSource = dt;

                    comboprocess.Text = Convert.ToString(dt.Rows[0]["processname"].ToString());
                    txtlotno.Text = Convert.ToString(dt.Rows[0]["lotno"].ToString());
                    txtorderqty.Text = Convert.ToString(dt.Rows[0]["orderqty"].ToString());
                    txtorderno.Text = Convert.ToString(dt.Rows[0]["orderno"].ToString());
                    txtstyleref.Text = Convert.ToString(dt.Rows[0]["styleref"].ToString());
                    if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    string sel2 = "select distinct e.asptblsizmasid,  d.colorname from asptblpurdet a join asptblpur b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizmas e on e.asptblsizmasid=a.sizename  where c.compcode='" + combocompcode.Text + "' and a.asptblpurid='" + Convert.ToString(dt.Rows[0]["asptblpurid"].ToString()) + "' order by 1";
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

                    string sel4 = "select distinct d.ASPTBLSIZMASID,d.SIZENAME from asptblpur a join asptblpurdet b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=a.compcode join asptblsizmas d on d.ASPTBLSIZMASID=b.sizename join asptblcolmas e on e.asptblcolmasid=b.colorname  where c.compcode='" + combocompcode.Text + "' AND A.asptblpurid='" + Convert.ToString(dt.Rows[0]["asptblpurid"].ToString()) + "' order by 1";
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

                    int i = 0, j = 0;
                    int k = 1;
                    int col = dataGridView1.Columns.Count - 2;
                    for (i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        string sel5 = "select  b.portion,d.ASPTBLSIZMASID,  d.SIZENAME,b.orderqty,b.asptblpurdetid from asptblpur a join asptblpurdet b on a.asptblpurid=b.asptblpurid join gtcompmast c on c.gtcompmastid=a.compcode join asptblsizmas d on d.ASPTBLSIZMASID=b.sizename join asptblcolmas e on e.asptblcolmasid=b.colorname  where c.compcode='" + combocompcode.Text + "' AND A.asptblpurid='" + Convert.ToString(dt.Rows[0]["asptblpurid"].ToString()) + "' and e.colorname='" + dataGridView1.Rows[i].Cells[0].EditedFormattedValue.ToString() + "'  order by 1";//and d.sizename='" + dataGridView1.Columns[i + 1].HeaderText.ToString() + "'
                        DataSet ds5 = Utility.ExecuteSelectQuery(sel5, "asptblpur");
                        DataTable dt5 = ds5.Tables["asptblpur"];
                        if (dt5.Rows.Count > 0)
                        {
                            if (col <= i)
                            {

                                dataGridView1.Rows[i].Cells[k].Value = dt5.Rows[0]["portion"].ToString();
                            }
                            else
                            {
                                dataGridView1.Rows[i].Cells[k].Value = dt5.Rows[i]["portion"].ToString();
                            }
                        }
                    }
                    allip1.Items.Clear();
                }
            }
        
        }
        DataGridViewComboBoxCell combo;
        object item;
        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
               

                if (txtasppur1id.Text == "")
                {

                }
                combocompcode.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            {
                txtsearch.Select();
                empty();
            }
        }
        Int32 s = 0;
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex >= 1)
            //{

            //    int ss = 0;
            //    for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //    {
            //        if (e.RowIndex == 0)
            //        {
            //            ss = Convert.ToInt32("0" + dataGridView1.Rows[i].Cells[1].Value);
            //        }
            //        else
            //        {
            //            ss = Convert.ToInt32("0" + dataGridView1.Rows[i].Cells[1].Value);

            //        }
            //        if (ss <= 0)
            //        {
            //            dataGridView1.Rows[i].Cells[1].Value = 1;
            //        }
            //    }


            //}
        }

        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {

            //s = 0;
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{

            //    for (int j = 0; j < dataGridView1.Columns.Count - 2; j++)
            //    {
            //        s += Convert.ToInt32("0" + dataGridView1.Rows[i].Cells[j + 2].Value);
            //        if (Convert.ToInt32("0" + txtorderqty.Text)<s)
            //        {

            //            lblgridtotal.Refresh();
            //            label15.Text = s.ToString();
            //            lblgridtotal.Text = "Excess Qty" + s.ToString() + " OrderQty " + txtorderqty.Text + " Invalid";
            //            dataGridView1.Rows[i].Cells[j + 2].Value = null;
            //            dataGridView1.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.Red;
            //        }
            //        else
            //        {
            //        lblgridtotal.Refresh();
            //        label15.Text = s.ToString();
            //        lblgridtotal.Text = "Excess Qty" + s.ToString() + " OrderQty " + txtorderqty.Text + " Invalid";

            //        dataGridView1.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.White;
            //        }

            //    }

            //}
        }
        void GridEdit(int colindex,DataGridView grid)
        {
            if (colindex > 1)
            {
                if (combotype.Text == "ORDER")
                {
                    s = 0;
                    for (int i = 0; i < grid.Rows.Count; i++)
                    {
                        
                        for (int j = 0; j < grid.Columns.Count - 2; j++)
                        {
                            s += Convert.ToInt32("0" + grid.Rows[i].Cells[j + 2].Value);
                            if (Convert.ToInt32("0" + txtorderqty.Text) < s)
                            {

                                lblgridtotal.Refresh();
                                label15.Text = s.ToString();
                                lblgridtotal.Text = "Size wise Qty Mismatch " + s.ToString() + " OrderQty " + txtorderqty.Text + " Invalid";
                                grid.Rows[i].Cells[j + 2].Value = null;
                                grid.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.Red;
                            }
                            else
                            {
                                label15.Refresh();
                                label15.Text = s.ToString(); lblgridtotal.Text = "";
                                grid.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.White;
                            }

                        }
                       
                    }
                }
                if (combotype.Text == "EXCESS QTY")
                {
                    s = 0;
                    for (int i = 0; i < grid.Rows.Count; i++)
                    {
                       
                        for (int j = 0; j < grid.Columns.Count - 2; j++)
                        {
                            s += Convert.ToInt32("0" + grid.Rows[i].Cells[j + 2].Value);
                            if (Convert.ToInt32("0" + txtexcesqty.Text) < s)
                            {

                                lblgridtotal.Refresh();
                                label15.Text = s.ToString();
                                lblgridtotal.Text = "Size wise Qty Mismatch " + s.ToString() + " OrderQty " + txtexcesqty.Text + " Invalid";
                                grid.Rows[i].Cells[j + 2].Value = null;
                                grid.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.Red;
                            }
                            else
                            {
                                label15.Refresh();
                                label15.Text = s.ToString(); lblgridtotal.Text = "";
                                grid.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.White;
                            }

                        }

                    }
                }
            }
        }
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            GridEdit(e.ColumnIndex,dataGridView1);
            //if (e.ColumnIndex > 1)
            //{
            //    if (combotype.Text == "ORDER")
            //    {
            //        s = 0;
            //        for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //        {

            //            for (int j = 0; j < dataGridView1.Columns.Count - 2; j++)
            //            {
            //                s += Convert.ToInt32("0" + dataGridView1.Rows[i].Cells[j + 2].Value);
            //                if (Convert.ToInt32("0" + txtorderqty.Text) < s)
            //                {

            //                    lblgridtotal.Refresh();
            //                    label15.Text = s.ToString();
            //                    lblgridtotal.Text = "Size wise Qty Mismatch " + s.ToString() + " OrderQty " + txtorderqty.Text + " Invalid";
            //                    dataGridView1.Rows[i].Cells[j + 2].Value = null;
            //                    dataGridView1.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.Red;
            //                }
            //                else
            //                {
            //                    label15.Refresh();
            //                    label15.Text = s.ToString(); lblgridtotal.Text = "";
            //                    dataGridView1.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.White;
            //                }

            //            }

            //        }
            //    }
            //    if (combotype.Text == "EXCESS QTY")
            //    {
            //                s = 0;
            //        for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //        {

            //            for (int j = 0; j < dataGridView1.Columns.Count - 2; j++)
            //            {
            //                s += Convert.ToInt32("0" + dataGridView1.Rows[i].Cells[j + 2].Value);
            //                if (Convert.ToInt32("0" + txtexcesqty.Text) < s)
            //                {

            //                    lblgridtotal.Refresh();
            //                    label15.Text = s.ToString();
            //                    lblgridtotal.Text = "Size wise Qty Mismatch " + s.ToString() + " OrderQty " + txtexcesqty.Text + " Invalid";
            //                    dataGridView1.Rows[i].Cells[j + 2].Value = null;
            //                    dataGridView1.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.Red;
            //                }
            //                else
            //                {
            //                    label15.Refresh();
            //                    label15.Text = s.ToString(); lblgridtotal.Text = "";
            //                    dataGridView1.Rows[i].Cells[j + 2].Style.BackColor = System.Drawing.Color.White;
            //                }

            //            }

            //        }
            //    }
            //}
        }

        private void txtorderqty_TextChanged(object sender, EventArgs e)
        {
            lblgridtotal.Refresh();
            lblgridtotal.Text = "Size wise Qty Mismatch " + s.ToString() + " OrderQty " + txtorderqty.Text + " Invalid";

        }

        private void txtexcesqty_TextChanged(object sender, EventArgs e)
        {
            lblgridtotal.Refresh();
            lblgridtotal.Text = "Size wise Qty Mismatch " + s.ToString() + " OrderQty " + txtexcesqty.Text + " Invalid";

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void lblcount_Click(object sender, EventArgs e)
        {

        }
        byte[] stdbytes; Int64 std;
        OpenFileDialog open = new OpenFileDialog();
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                stdbytes = null;

                open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.png;*.jpg; *.jpeg; *.bmp";
                if (open.ShowDialog() == DialogResult.OK)
                {

                    pictureBox1.Image = new Bitmap(open.FileName);
                    stdbytes = Models.Device.ImageToByteArray(pictureBox1);
                    System.Text.Encoding enc = System.Text.Encoding.ASCII;
                    std = Convert.ToInt64("0" + stdbytes.Length);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void txtorderqty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtexcesqty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        //public static DataTable ToDataTable<T>(this IList<T> data)
        //{
        //    PropertyDescriptorCollection properties =
        //        TypeDescriptor.GetProperties(typeof(T));
        //    DataTable table = new DataTable();
        //    foreach (PropertyDescriptor prop in properties)
        //        table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
        //    foreach (T item in data)
        //    {
        //        DataRow row = table.NewRow();
        //        foreach (PropertyDescriptor prop in properties)
        //            row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
        //        table.Rows.Add(row);
        //    }
        //    return table;
        //}

        //private void butqrcode_Click(object sender, EventArgs e)
        //{
        //  //  int i = 0;
        //  //  if (Class.Users.QrCodeArray.Rows.Count > 0)
        //  //  {
        //  //      Class.Users.QrCodeArray.Rows.Clear(); Class.Users.QrCodeArray.Columns.Clear();                
        //  //  }            
        //  //Class.Users.Query = "select asptblpurdet1id,asptblpurdetid,asptblpurid,asptblpur1id,compcode,pono,colorname,sizename,orderqty,PORTION,colorname1,qrcode from asptblpurdet1  where asptblpurdet1id<0"+ ":'" + combocompcode.SelectedValue+ "':'" + Class.Users.Finyear + "':'" + txtpono.Text + "':'asptblpurdet1'";
        //  //  int rowindex = 0;
        //  //  foreach (DataGridViewColumn col1 in dataGridView1.Columns)
        //  //  {
        //  //      Class.Users.QrCodeArray.Columns.Add(col1.HeaderText);
        //  //  }
        //  //   foreach (DataGridViewRow row1 in dataGridView1.Rows)
        //  //  {
        //  //      Class.Users.QrCodeArray.Rows.Add();
        //  //      if (row1.Cells[rowindex].FormattedValue.ToString() != null)
        //  //      {
        //  //          foreach (DataGridViewCell cell in row1.Cells)
        //  //          {
        //  //              Class.Users.QrCodeArray.Rows[Class.Users.QrCodeArray.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
        //  //          }
        //  //      }
        //  //      else
        //  //      {
        //  //          return;
        //  //      }
        //  //      rowindex++;
        //  //  }
        //  //  DataTable dt = new DataTable();
        //  //  Models.CommonClass com = new Models.CommonClass();

        //  //  string[] sarray = Class.Users.Query.Split(':');
        //  //  dt = com.select(sarray[0], sarray[4]);

        //  //  int n = 2;
        //  //  int cnt = 0, k = 0, cnt1 = 0, cnt2 = 1, row = 0, col = 0, rowcount = 0, tot = 0;
        //  //  col = Class.Users.QrCodeArray.Columns.Count - 2;
        //  //  int totcount = 0;
        //  //  rowcount = Class.Users.QrCodeArray.Rows.Count;
        //  //  int cc = 0;
        //  //  for (int i1 = 0; i1 < rowcount; i1++)
        //  //  {
        //  //      tot = 0; row = 0; cc++;
        //  //      cnt1 = 0; row = 2;

        //  //      for (i = 0; i < col; i++)
        //  //      {
        //  //          if (col != tot)
        //  //          {
        //  //              cnt2 = cc - 1; cnt1 = i + row; cnt = 0;
        //  //              cnt = Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2]["P-COUNT"].ToString()) * Convert.ToInt32("0" + Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString());
        //  //              if (cnt >= 1)
        //  //              {
        //  //                  k++;
        //  //                  string sel2 = "select a.asptblpurdetid ,a.asptblpurid,b.asptblpur1id,b.pono   from  asptblpurdet a join asptblpur b on a.asptblpurid=b.asptblpurid join asptblsizmas  c on c.asptblsizmasid=a.sizename  where  b.compcode=" + sarray[1] + "  and b.finyear=" + sarray[2] + " and b.pono=" + sarray[3] + " and c.sizename='" + Class.Users.QrCodeArray.Columns[cnt1].ToString() + "' ";
        //  //                  DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblpurdet");
        //  //                  DataTable dt2 = ds2.Tables["asptblpurdet"];
        //  //                  int id1 = Convert.ToInt32(dt2.Rows[i1]["asptblpurdetid"].ToString());
        //  //                  int id2 = Convert.ToInt32(dt2.Rows[i1]["asptblpurid"].ToString());
        //  //                  int id3 = Convert.ToInt32(dt2.Rows[i1]["asptblpur1id"].ToString());
        //  //                  string id4 = Convert.ToString(dt2.Rows[i1]["pono"].ToString());
        //  //                  for (int a = 0; a < cnt; a++)
        //  //                  {
        //  //                      totcount++;
        //  //                      dt.Rows.Add(0);
        //  //                      dt.Rows[dt.Rows.Count - 1][1] = id1;
        //  //                      dt.Rows[dt.Rows.Count - 1][2] = id2;
        //  //                      dt.Rows[dt.Rows.Count - 1][3] = id3;
        //  //                      dt.Rows[dt.Rows.Count - 1][4] = Class.Users.COMPCODE;
        //  //                      dt.Rows[dt.Rows.Count - 1][5] = id4;
        //  //                      dt.Rows[dt.Rows.Count - 1][6] = Class.Users.QrCodeArray.Rows[i1]["COLORNAME"].ToString(); //GenerateQrCode(Class.Users.QrCodeArray.Rows[i]["COLORNAME"].ToString());
        //  //                      dt.Rows[dt.Rows.Count - 1][7] = Class.Users.QrCodeArray.Columns[cnt1].ToString();
        //  //                      dt.Rows[dt.Rows.Count - 1][8] = Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString();
        //  //                      dt.Rows[dt.Rows.Count - 1][9] = Class.Users.QrCodeArray.Rows[cnt2]["P-COUNT"].ToString();
        //  //                      dt.Rows[dt.Rows.Count - 1][10] = Class.Users.QrCodeArray.Rows[i1]["COLORNAME"].ToString() + " " + i1 + " " + k.ToString() + " " + Class.Users.QrCodeArray.Columns[cnt1].ToString() + " " + Class.Users.QrCodeArray.Rows[cnt2][Class.Users.QrCodeArray.Columns[cnt1].ToString()].ToString() + " " + totcount;
        //  //                      dt.Rows[dt.Rows.Count - 1][11] = "";
        //  //                  }

        //  //                  tot++;
        //  //              }
        //  //          }

        //  //      }
        //  //  }
        //    //Pinnacle.ReportFormate.Lyla.LylaReport ly = new ReportFormate.Lyla.LylaReport();
        //    //ly.Show();
        //}

        public void ReadOnlys()
        {
        }
    }
}

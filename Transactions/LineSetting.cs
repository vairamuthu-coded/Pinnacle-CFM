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
    public partial class LineSetting : Form,ToolStripAccess
    {
        public LineSetting()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            tabControl1.SelectTab(tabPage2);
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            panel6.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
        }

        private static LineSetting _instance; string coid = "", siid = "";
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfilter2 = new ListView();
        ListView listfilter4 = new ListView();
        public static LineSetting Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new LineSetting();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


   
        public void autonumberload()
        {
            try
            {
                //txtdeltid1.Text = ""; txtcertificateno.Text = "";
                //if(txtdeltid.Text != "") { }

                //string sel = "select max(a.asptbllayid1)+1 as id from asptbllay a join gtcompmast b on a.compcode = b.gtcompmastid  where  a.finyear='" + Class.Users.Finyear + "' and b.compcode ='" + combocompcode.Text + "'; ";
                //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllay");
                //DataTable dt = ds.Tables["asptbllay"];
                //int cnt = dt.Rows.Count;
                //if (dt.Rows[0]["id"].ToString() == "")
                //{


                //    txtlayno.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + 1;
                //    txtlayid1.Text = "1";
                //}
                //else
                //{

                //    txtlayno.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                //    txtlayid1.Text = dt.Rows[0]["id"].ToString();
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LayCutting_Load(object sender, EventArgs e)
        {
            empty(); txtsearch.Select();
        }

        public void Saves()
        {
            try
            {

                if (txtlayid1.Text == "" && listView4.Items.Count < 0)
                {
                    MessageBox.Show("colorname Name is empty " + " Alert " + txtlayid1.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtlayid1.Text != "" && listView4.Items.Count > 0)
                {
                    string maxid = "";
                    Models.LayCuttingModel p = new Models.LayCuttingModel();
                    p.asptbllayid = Convert.ToInt64("0" + txtlayid.Text);                   
                    p.finyear = Class.Users.Finyear;
                    p.laydate = Convert.ToString(dateTimePicker1.Value.ToShortDateString());
                    p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                    p.pono = Convert.ToString(combopono.Text);                   
                    p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);                   
                    p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                    p.tableno = Convert.ToInt64(txttableno.Text);
                    p.compcode1 = Class.Users.COMPCODE;
                    p.username = Class.Users.USERID;
                    p.createdby = Convert.ToString(Class.Users.HUserName);
                    p.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString()).ToString();
                    p.modified = Convert.ToString(System.DateTime.Now.ToString());
                    p.modifiedby = Class.Users.HUserName;
                    p.ipaddress = Class.Users.IPADDRESS;


                    if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; checkactive.Checked = false; }
                    if (checklaycancel.Checked == true) { p.laycancel = "T"; } else { p.laycancel = "F"; checklaycancel.Checked = false; }

                    string sel = "select asptbllayid    from  asptbllay    WHERE  finyear='" + p.finyear + "' and laydate='" + p.laydate + "' and compcode ='" + p.compcode + "' and buyer='" + p.buyer + "' and pono='" + p.pono + "' and orderqty ='" + p.orderqty + "'  and stylename='" + p.stylename + "' and tableno='" + p.tableno + "'  and laycancel='" + p.laycancel + "' and  active='" + p.active + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllay");
                    DataTable dt = ds.Tables["asptbllay"];
                    if (dt.Rows.Count != 0)
                    {
                        // MessageBox.Show("Child Record Found " + " Alert " + txtpono.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtlayid.Text) == 0 || Convert.ToInt32("0" + txtlayid.Text) == 0)
                    {
                        string ins = "insert into asptbllay(finyear,laydate,compcode,buyer,pono,orderqty,stylename,tableno,laycancel,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.finyear + "','" + p.laydate + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderqty + "','" + p.stylename + "','" + p.tableno + "','" + p.laycancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + Convert.ToDateTime(p.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + p.modified + "','" + p.modifiedby + "','" + p.ipaddress + "');";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptbllayid) id    from  asptbllay   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and pono='" + txtlayid1.Text + "' ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbllay");
                        DataTable dt2 = ds2.Tables["asptbllay"];
                        maxid = dt2.Rows[0]["id"].ToString();
                    }
                    else
                    {
                        string up = "update  asptbllay  set   finyear='" + p.finyear + "' ,  laydate='" + p.pono + "' , compcode ='" + p.compcode + "' , buyer='" + p.buyer + "' , pono='" + p.pono + "' ,  orderqty ='" + p.orderqty + "' , stylename='" + p.stylename + "' , tableno='" + p.tableno + "'  , laycancel='" + p.laycancel + "' , active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbllayid='" + txtlayid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        maxid = txtlayid.Text;
                    }
                    Models.LayCuttingModel.LayCuttingModeldet pp = new Models.LayCuttingModel.LayCuttingModeldet();
                    int i = 0;
                    if (listView4.Items.Count >= 0)
                    {
                        for (i = 0; i < listView4.Items.Count; i++)
                        {
                            pp.asptbllaydetid = Convert.ToInt64("0" + listView4.Items[i].SubItems[2].Text);
                            pp.asptbllayid = Convert.ToInt64("0" + txtlayid.Text);
                         
                            pp.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                            pp.pono = Convert.ToString(txtlayid.Text);                          
                            colorid(listView4.Items[i].SubItems[3].Text);
                            sizeid(listView4.Items[i].SubItems[4].Text);
                            pp.colorname = coid;
                            pp.sizename = siid;


                            string sel1 = "select asptbllaydetid    from  asptbllaydet   where   asptbllayid='" + txtlayid.Text + "'  and asptbllayid1='" + txtlayid1.Text + "' and colorname='" + pp.colorname + "' and sizename='" + pp.sizename + "' ";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbllaydet");
                            DataTable dt1 = ds1.Tables["asptbllaydet"];
                            if (dt1.Rows.Count != 0)
                            {

                                tabControl1.SelectTab(tabPage2);
                            }
                            else if (dt1.Rows.Count != 0 && pp.asptbllaydetid == 0 || pp.asptbllaydetid == 0)
                            {

                                string ins1 = "insert into asptbllaydet(asptbllayid,compcode,pono,orderno,colorname,sizename) values('" + maxid.ToString() + "' , '" + combocompcode.SelectedValue + "' ,'" + txtlayid1.Text + "' , '" + comboorderno.Text + "' ,'" + pp.colorname + "','" + pp.sizename + "' );";
                                Utility.ExecuteNonQuery(ins1);
                            }
                            else
                            {
                                string up1 = "update  asptbllaydet  set asptbllayid='" + txtlayid.Text + "' , compcode='" + combocompcode.SelectedValue + "' , pono='" + txtlayid1.Text + "' , orderno='" + comboorderno.Text + "' ,colorname='" + pp.colorname + "',sizename='" + pp.sizename + "' where asptbllaydetid='" + pp.asptbllayid + "';";
                                Utility.ExecuteNonQuery(up1);
                            }


                        }
                    }
                    if (txtlayid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtlayid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        autonumberload();
                        tabControl1.SelectTab(tabPage2);
                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtlayid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        autonumberload();
                        tabControl1.SelectTab(tabPage2);
                    }
                }
                else
                {
                    MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("colorname " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void LayCutting_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       
        public void News()
        {

           GridLoad();   compload(); empty();
        }
        private void empty()
        {
            txtlayid.Text = ""; txtorderqty.Text = "";
            txtlayid1.Text = ""; combopono.Text = ""; comboorderno.Text = "";
            txttableno.Text = ""; 
            combocompcode.Text = ""; combocompcode.SelectedIndex = -1;
            combostyle.Text = ""; combostyle.SelectedIndex = -1;
            combobuyer.Text = ""; combobuyer.SelectedIndex = -1; combocolor.Text = ""; combocolor.SelectedIndex = -1;
            txtlayid1.Text = "";
            checkactive.Checked = true;listView2.Items.Clear();listView3.Items.Clear(); listView4.Items.Clear();
            checklaycancel.Checked = false; listView4.Items.Clear();
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            panel6.BackColor = Class.Users.BackColors;
            panel7.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            listView3.Font = Class.Users.FontName;
            listView4.Font = Class.Users.FontName;

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
                string sel = "select asptblsizemasid from  asptblsizemas where sizename='" + s + "' ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsizemas");
                DataTable dt = ds.Tables["asptblsizemas"];
                siid = "";
                siid = Convert.ToString(dt.Rows[0]["asptblsizemasid"].ToString());

            }
            catch (Exception EX)
            { }
        }
        //public void stylenameload()
        //{
        //    try
        //    {
        //        string sel = "select asptblstylemasid,stylename from  asptblstylemas  order by 1 ;";
        //        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblstylemas");
        //        DataTable dt = ds.Tables["asptblstylemas"];

        //        combostyle.DisplayMember = "stylename";
        //        combostyle.ValueMember = "asptblstylemasid";
        //        combostyle.DataSource = dt;
        //        combostyle.Text = ""; combostyle.SelectedIndex = -1;
        //    }
        //    catch (Exception EX)
        //    { }
        //}

        public void compload()
        {
            try
            {

                DataTable dt = mas.comcode();

                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt; 
                combocompcode.Text = ""; combocompcode.SelectedIndex = -1;
               
            }
            catch (Exception EX)
            { }
        }
        //private void colorload()
        //{
        //    try
        //    {
        //        listView2.Items.Clear();
        //        string sel1 = "   SELECT a.colorname   FROM  asptblcolmas a   order by 1";
        //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcolmas");
        //        DataTable dt = ds.Tables["asptblcolmas"];
        //        if (dt.Rows.Count > 0)
        //        {
        //            int i = 1;
        //            foreach (DataRow myRow in dt.Rows)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                list.SubItems.Add(i.ToString());
        //                list.SubItems.Add(myRow["colorname"].ToString());
        //                list.SubItems.Add("");
        //                this.listfilter2.Items.Add((ListViewItem)list.Clone());
        //                listView2.Items.Add(list);
        //                i++;
        //            }
        //            lbltotalsize.Text = "Total Count    :" + listView2.Items.Count;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        public void GridLoad()
        {
            //try
            //{
            //    listView1.Items.Clear();
            //    string sel1 = "SELECT a.asptbllayid,b.compcode,c.buyername,a.pono,a.orderno ,a.orderqty,d.stylename,e.tableno,a.active FROM  asptbllay a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstylemas d on d.asptblstylemasid=a.stylename join asptbltableno e on e.asptbltablenoid=a.tableno   order by  a.asptbllayid desc;";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllay");
            //    DataTable dt = ds.Tables["asptbllay"];
            //    if (dt.Rows.Count > 0)
            //    {
            //        int i = 1;
            //        foreach (DataRow myRow in dt.Rows)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            list.SubItems.Add(i.ToString());
            //            list.SubItems.Add(myRow["asptbllayid"].ToString());
            //            list.SubItems.Add(myRow["compcode"].ToString());
            //            list.SubItems.Add(myRow["buyername"].ToString());
            //            list.SubItems.Add(myRow["pono"].ToString());
            //            list.SubItems.Add(myRow["orderno"].ToString());
            //            list.SubItems.Add(myRow["orderqty"].ToString());
            //            list.SubItems.Add(myRow["stylename"].ToString());
            //            list.SubItems.Add(myRow["tableno"].ToString());
            //            list.SubItems.Add(myRow["active"].ToString());
            //            this.listfilter.Items.Add((ListViewItem)list.Clone());
            //                                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;


            //            listView1.Items.Add(list);
            //            i++;
            //        }
            //        lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtlayid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptbllayid,a.asptbllayid1,a.finyear,a.laydate,b.compcode,e.buyername,a.pono,a.orderno,a.orderqty, a.shipdate,c.stylename,d.tableno,a.laycancel,a.active  from  asptbllay a join gtcompmast b on b.gtcompmastid=a.compcode  join asptblstylemas c on c.asptblstylemasid=a.stylename join asptbltableno d on d.asptbltablenoid=a.tableno  join asptblbuymas e on e.asptblbuymasid=a.buyer  where a.asptbllayid=" + txtlayid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllay");
                    DataTable dt = ds.Tables["asptbllay"];
                    if (dt.Rows.Count > 0)
                    {
                        txtlayid.Text = Convert.ToString(dt.Rows[0]["asptbllayid"].ToString());
                        txtlayid1.Text = Convert.ToString(dt.Rows[0]["asptbllayid1"].ToString());
                        dateTimePicker1.Value = Convert.ToDateTime(dt.Rows[0]["laydate"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyername"].ToString());
                        combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        comboorderno.Text = Convert.ToString(dt.Rows[0]["orderno"].ToString());
                        txtorderqty.Text = Convert.ToString(dt.Rows[0]["orderqty"].ToString());                       
                        combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        txttableno.Text = Convert.ToString(dt.Rows[0]["tableno"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        if (dt.Rows[0]["laycancel"].ToString() == "T") { checklaycancel.Checked = true; } else { checklaycancel.Checked = true; checklaycancel.Checked = false; }
                    }
                    string sel2 = "select d.colorname,e.sizename,a.asptbllaydetid from asptbllaydet a join asptbllay b on a.asptbllayid=b.asptbllayid join gtcompmast c on c.gtcompmastid=b.compcode join asptblcolmas d on d.asptblcolmasid=a.colorname  join asptblsizemas e on e.asptblsizemasid=a.sizename  where a.asptbllayid=" + txtlayid.Text;
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbllay");
                    DataTable dt2 = ds2.Tables["asptbllay"];
                    if (dt2.Rows.Count > 0)
                    {
                        int i = 0; listView4.Items.Clear(); listfilter4.Items.Clear();
                        foreach (DataRow myRow in dt2.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(myRow["asptbllaydetid"].ToString());
                            list.SubItems.Add(myRow["colorname"].ToString());
                            list.SubItems.Add(myRow["sizename"].ToString());

                            this.listfilter4.Items.Add((ListViewItem)list.Clone());
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;


                            listView4.Items.Add(list);

                            i++;
                        }
                        lbltotal.Text = "Total Count    :" + listView4.Items.Count;
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
                        if (listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[5].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[6].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[8].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
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

                        this.listView1.Items.Add((ListViewItem)item.Clone());


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
                string[] fab;listView4.Items.Clear();
                string[] col; string[] siz; string[] sqty;
                List<string> listfab = new List<string>();
                List<string> listcolor = new List<string>();
                List<string> listsqty = new List<string>();
                List<string> listsize = new List<string>();

                for (int j = 0; j < allip2.Items.Count; j++)
                {
                    if (allip2.Items[j].SubItems[2].Text == "True")
                    {
                        listfab.Add(allip2.Items[j].SubItems[1].Text);

                    }
                }
                fab = listfab.ToArray<string>();


                for (int j = 0; j < allip1.Items.Count; j++)
                {
                    if (allip1.Items[j].SubItems[4].Text == "True")
                    {
                        listcolor.Add(allip1.Items[j].SubItems[1].Text);
                       
                    }
                }
                col = listcolor.ToArray<string>();
                for (int j = 0; j < allip1.Items.Count; j++)
                {
                    if (allip1.Items[j].SubItems[4].Text == "True")
                    {
                       
                        listsize.Add(allip1.Items[j].SubItems[2].Text);
                        
                    }
                }
                siz = listsize.ToArray<string>();

                for (int j = 0; j < allip1.Items.Count; j++)
                {
                    if (allip1.Items[j].SubItems[4].Text == "True")
                    {
                       
                        listsqty.Add(allip1.Items[j].SubItems[3].Text);
                    }
                }
                sqty = listsqty.ToArray<string>();


                int item0 = 1;
                for (int i = 0; i < fab.Length; i++)
                {
                    for (int j = 0; j < col.Length; j++)
                    {

                        ListViewItem list = new ListViewItem();

                        list.SubItems.Add(item0.ToString());
                        list.SubItems.Add("");
                        list.SubItems.Add(fab[i].ToString());
                        list.SubItems.Add(col[j].ToString());
                        list.SubItems.Add(siz[j].ToString());
                        list.SubItems.Add(txtmarkerno.Text);
                        list.SubItems.Add(sqty[j].ToString());
                        this.listfilter4.Items.Add((ListViewItem)list.Clone());
                        if (item0 % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }

                        listView4.Items.Add(list);
                        item0++;

                    }
                }
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
            lbltotalall.Text = "Total Count: " + listView4.Items.Count;
        }

        private void checkall_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //foreach (ListViewItem item in listView4.Items)
                //{
                //    item.Checked = true;
                //}
                //foreach (ListViewItem item in listView3.Items)
                //{
                //    item.Checked = false;
                //}
                if (checkall.Checked == true)
                {

                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        int i = 0;
                        allip1.Items.Clear(); allip2.Items.Clear();
                        listView4.Items.Clear();
                        checkall.Checked = false;
                        checksizeall.Checked = false;
                        //listView2.Items.Clear();
                        //listView3.Items.Clear();
                        //colorload();
                        //sizeload(combotableno.Text);
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
        private void listView2_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {


                ListViewItem it = new ListViewItem();
                if (e.Item.Checked == true)
                {

                    e.Item.SubItems[5].Text = "Connected";
                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.SubItems[3].Text);
                    it.SubItems.Add(e.Item.SubItems[4].Text);
                   // it.SubItems.Add(e.Item.SubItems[5].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip1.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[5].Text == "Connected")
                {
                    e.Item.SubItems[5].Text = "DisConnected";
                    e.Item.Checked = false;
                    for (int c = 0; c < allip1.Items.Count; c++)
                    {
                        if (listView2.SelectedItems[0].SubItems[3].Text == allip1.Items[c].SubItems[2].Text)
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

                    e.Item.SubItems[3].Text = "Connected";
                    it.SubItems.Add(e.Item.SubItems[2].Text);                  
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip2.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[3].Text == "Connected")
                {
                    e.Item.SubItems[3].Text = "DisConnected";
                    for (int c = 0; c < allip2.Items.Count; c++)
                    {
                        if (listView3.SelectedItems[0].SubItems[2].Text == allip2.Items[c].SubItems[1].Text)
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
                        foreach (ListViewItem eachItem in listView4.SelectedItems)
                        {
                            string ID = eachItem.SubItems[2].Text;
                            if (Convert.ToInt64("0" + ID) > 1)
                            {
                                string del = "delete from asptbllaydet where asptbllaydetid=" + ID;
                                Utility.ExecuteNonQuery(del);
                            }
                            listView4.Items.Remove(eachItem);
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
            try
            {
                int item0 = 0; listView4.Items.Clear();
                if (txtpursearch.Text.Length >= 1)
                {

                    foreach (ListViewItem item in listfilter4.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter4.Items[item0].SubItems[2].ToString().Contains(txtpursearch.Text) || listfilter4.Items[item0].SubItems[3].ToString().Contains(txtpursearch.Text))
                        {


                            list.Text = listfilter4.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter4.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter4.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter4.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter4.Items[item0].SubItems[4].Text);
                            if (item0 % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }

                            listView4.Items.Add(list);
                        }
                        item0++;
                    }
                    lbltotalall.Text = "Total Count: " + listView4.Items.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    item0 = listfilter.Items.Count;
                    listView4.Items.Clear();
                    foreach (ListViewItem item in listfilter4.Items)
                    {
                        if (item0 % 2 == 0) { item.BackColor = Color.WhiteSmoke; } else { item.BackColor = Color.White; }

                        this.listView4.Items.Add((ListViewItem)item.Clone());

                        item0++;
                    }

                    lbltotalall.Text = "Total Count: " + listView4.Items.Count;
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
                    autonumberload();
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


        public void orderno(Int64 s)
        {
            try
            {
                if (s >0)
                {
                    string sel1 = "SELECT a.asptblordentryid, a.orderno   FROM  asptblordentry a  join gtcompmast b on a.compcode=b.gtcompmastid join asptblpur c on c.orderno=a.orderno where a.compcode='" + s + "'  order by a.asptblordentryid desc";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordentry");
                    DataTable dt = ds.Tables["asptblordentry"];
                    comboorderno.DataSource = null;

                    comboorderno.DisplayMember = "orderno";
                    comboorderno.ValueMember = "asptblordentryid";
                    comboorderno.DataSource = dt;

                  
                }
            }
            catch (Exception EX)
            { }
        }


        public void color(string ord)
        {
            try
            {
                string sel1 = "SELECT distinct e.asptblcolmasid,e.colorname  FROM  asptblordentry a  join gtcompmast b on a.compcode=b.gtcompmastid join asptblpur c on c.orderno=a.orderno and c.compcode=b.gtcompmastid join asptblpurdet d on d.asptblpurid=c.asptblpurid join asptblcolmas e on e.asptblcolmasid=d.colorname where a.orderno='" + ord + "'  order by a.asptblordentryid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordentry");
                DataTable dt = ds.Tables["asptblordentry"];
                combocolor.DataSource = null;

                combocolor.DisplayMember = "colorname";
                combocolor.ValueMember = "asptblcolmasid";
                combocolor.DataSource = dt;
            }
            catch (Exception EX)
            { }
        }
        public void pono(string s)
        {
            try
            {
                string sel1 = "SELECT c.asptblpurid,c.pono,d.asptblstylemasid,d.stylename,e.asptblbuymasid,e.buyername,a.orderqty  FROM  asptblordentry a  join gtcompmast b on a.compcode=b.gtcompmastid join asptblpur c on c.orderno=a.orderno  join asptblstylemas d on d.asptblstylemasid=c.stylename join asptblbuymas e on e.asptblbuymasid=c.buyer where a.orderno='" + s + "'  order by c.asptblpurid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordentry");
                DataTable dt = ds.Tables["asptblordentry"];
                combopono.DataSource = null;

                combopono.DisplayMember = "pono";
                combopono.ValueMember = "asptblpurid";
                combopono.DataSource = dt;

                combobuyer.DisplayMember = "buyername";
                combobuyer.ValueMember = "asptblbuymasid";
                combobuyer.DataSource = dt;

                combostyle.DisplayMember = "stylename";
                combostyle.ValueMember = "asptblstylemasid";
                combostyle.DataSource = dt;
                txtorderqty.Text = dt.Rows[0]["orderqty"].ToString();
            }
            catch (Exception EX)
            { }
        }
        private void comboorderno_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboorderno.SelectedValue != null)
            {
                pono(comboorderno.Text);
                color(comboorderno.Text);
            }
        }

        private void combopono_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void combocolor_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboorderno.SelectedValue != null && combocolor.SelectedValue != null)
            {
                colorsizeqty(comboorderno.Text, combocolor.Text);
                fabricsizeqty(comboorderno.Text, combocolor.Text); listView4.Items.Clear(); allip1.Items.Clear(); allip2.Items.Clear();
            }
        }
        private void colorsizeqty(string ord,string col)
        {
            try
            {
                if (ord != "" && col != "")
                {
                    listView2.Items.Clear();
                    string sel1 = "select f.colorname,g.sizename,c.shipqty  from asptblordentry a join asptblordentrydet b on a.asptblordentryid=b.asptblordentryid join asptblordentrysubdet c on c.asptblordentryid=a.asptblordentryid and c.asptblordentryid=b.asptblordentryid AND b.indexno=c.indexno join gtcompmast e on e.gtcompmastid=a.compcode and e.gtcompmastid=b.compcode join asptblcolmas f on f.asptblcolmasid=b.colorname join asptblsizemas g on g.asptblsizemasid=c.sizename where  a.orderno='" + ord + "' and f.colorname='" + col + "' order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizemas");
                    DataTable dt = ds.Tables["asptblsizemas"];
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
                            if (i % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }
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
                    string sel1 = "select distinct h.fabric  from asptblordentry a  join asptblordentrydet b on a.asptblordentryid=b.asptblordentryid   join asptblordentrysubdet c on c.asptblordentryid=a.asptblordentryid and c.asptblordentryid=b.asptblordentryid  AND b.indexno=c.indexno  join gtcompmast e on e.gtcompmastid=a.compcode and e.gtcompmastid=b.compcode join asptblcolmas f on f.asptblcolmasid=C.colorname AND f.asptblcolmasid=b.colorname join asptblsizemas g on g.asptblsizemasid=c.sizename join asptblfabricmas h on h.asptblfabricmasid=b.fabric and h.asptblfabricmasid=c.fabric  where  a.orderno='" + ord + "' and f.colorname='" + col + "' order by 1;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizemas");
                    DataTable dt = ds.Tables["asptblsizemas"];
                    if (dt.Rows.Count > 0)
                    {
                        int i = 1;
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(i.ToString());
                            list.SubItems.Add(myRow["fabric"].ToString());
                            list.SubItems.Add("");
                            if (i % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }

                            listView3.Items.Add(list);
                            i++;
                        }
                        lbltotalfabric.Text = "Total Count    :" + listView3.Items.Count;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string sel = "select asptblbuymasid,buyercode from  asptblbuymas a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + combocompcode.Text + "'   order by 1 ;";
            //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbuymas");
            //DataTable dt = ds.Tables["asptblbuymas"];
            //combobuyer.DisplayMember = "buyercode";
            //combobuyer.ValueMember = "asptblbuymasid";
            //combobuyer.DataSource = dt;
            // autonumberload();
            listView4.Items.Clear(); allip1.Items.Clear(); allip2.Items.Clear();
            if (combocompcode.SelectedValue != null)
            {
                Int64 c = Convert.ToInt64(combocompcode.SelectedValue.ToString());
                orderno(c);
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

        private void butDelete_Click(object sender, EventArgs e)
        {

        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }
    }
}

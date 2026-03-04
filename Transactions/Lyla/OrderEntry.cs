using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions.Lyla
{
    public partial class OrderEntry : Form, ToolStripAccess
    {
        public OrderEntry()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            tabControl1.SelectTab(tabPage2); tabControl2.SelectTab(tabPage4);

        }

        private static OrderEntry _instance; string coid = "", siid = "", fabid = ""; int griddelrow;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        ListView listfiltergrid = new ListView();
        ListView listfilter4 = new ListView();
        ListView listfilter2search = new ListView(); int totalqty = 0; int tot = 0; int tot1 = 0; int shiptot1 = 0; int shptot = 0; int ordqty1 = 0;
        public static OrderEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new OrderEntry();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {


                    }
                }


            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        public void autonumberload()
        {
            try
            {

                string sel = "select count(a.asptblordentryid1)+1 as id from asptblordentry a join gtcompmast b on a.compcode = b.gtcompmastid  where  a.finyear='" + Class.Users.Finyear + "' and b.compcode ='" + combocompcode.Text + "'; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblordentry");
                DataTable dt = ds.Tables["asptblordentry"];
                int cnt = dt.Rows.Count;
                if (dt.Rows[0]["id"].ToString() == "")
                {
                    txtorderno.Text = "ORD" + "/" + combocompcode.Text + "/" + Class.Users.Finyear + "/" + 1;
                    txtorderid1.Text = "1";
                }
                else
                {
                    txtorderno.Text = "ORD" + "/" + combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                    txtorderid1.Text = dt.Rows[0]["id"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void OrderEntry_Load(object sender, EventArgs e)
        {
            News(); txtsearch.Select(); //datagridviewrow();
        }

        public void News()
        {

            empty(); GridLoad(); colorload(); Fabricload(); sizegroupload(); compload(); stylenameload(); employeeload(); currencyload(); listView2.Items.Clear();
        }
        private void empty()
        {
            dataGridView1.Rows.Clear(); tot1 = 0; tot = 0; totalqty = 0; shiptot1 = 0; shptot = 0; ordqty1 = 0; txtorderid.Text = "";
            lbltotalall.Text = ""; Class.Users.UserTime = 0;
            lblorderqty.Text = ""; griddelrow = -1;
            lblshipqty.Text = ""; listView2.Items.Clear();

            GlobalVariables.HideCols = new string[] { "SNo", "asptblordentrydetid", "compcode", "orderno" };
            CommonFunctions.RemoveColumn(dataGridView2, GlobalVariables.HideCols);
            CommonFunctions.SetRowNumber(dataGridView2);
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel5.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            listView2.Font = Class.Users.FontName;
            listView3.Font = Class.Users.FontName;

            //do
            //{
            //    for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            //}
            //while (dataGridView1.Rows.Count > 1);

            do
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++) { try { dataGridView2.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView2.Rows.Count > 1);

            allip1.Items.Clear(); allip2.Items.Clear(); allip3.Items.Clear(); allip4.Items.Clear();
            if (listView1.Items.Count > 1)
            {
                listView1.Items[0].Selected = true;
            }
        }
        bool isvalid = false;
        private bool checkvalidate()
        {
            int i = 0, j = 1;
            int cnt = dataGridView2.Rows.Count - 1;
            if (Convert.ToInt64(combocompcode.SelectedValue) < 0)
            {
                MessageBox.Show("CompCode is empty " + " Alert " + combocompcode.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combocompcode.Select();
                return false;
            }
            if (txtorderqty.Text == "")
            {
                MessageBox.Show("Order Qty is empty " + " Alert " + txtorderqty.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtorderqty.Select();
                return false;
            }
            if (Convert.ToInt64(combobuyagent.SelectedValue) <= 0)
            {
                MessageBox.Show("bUyeing Agent is empty " + " Alert " + combobuyagent.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combobuyagent.Select();
                return false;
            }
            if (Convert.ToInt64(combosizegroup.SelectedValue) < 0)
            {
                MessageBox.Show("Size Group is empty " + " Alert " + combosizegroup.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                combosizegroup.Select();
                return false;
            }
            for (i = 0; i < cnt; i++)
            {

                int v4 = Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value);
                int v5 = Convert.ToInt32(dataGridView2.Rows[i].Cells[5].Value);
                int v6 = Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value);
                int v7 = Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value);

                if (v4 < 0)
                {
                    MessageBox.Show("Fabric is empty " + " Alert " + i.ToString(), "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    dataGridView2.Rows[i].Cells[4].Selected = true;
                    return false;

                }
                if (v5 < 0)
                {
                    MessageBox.Show("ColorName is empty " + " Alert " + i.ToString(), "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    dataGridView2.Rows[i].Cells[5].Selected = true;
                    return false;

                }
                if (v6 < 0)
                {
                    MessageBox.Show("Qty Name is empty " + " Alert " + i.ToString(), "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    dataGridView2.Rows[i].Cells["6"].Selected = true;
                    return false;

                }
                if (v6 != v7)
                {
                    MessageBox.Show("Qty mismactch is empty " + " Alert " + i.ToString(), "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    dataGridView2.Rows[i].Cells[7].Selected = true;
                    return false;

                }
                if (v7 < 0)
                {
                    MessageBox.Show("pls Enter Subgrid Value " + " Alert " + i.ToString(), "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    dataGridView2.Rows[i].Cells[7].Selected = true;
                    return false;

                }
                if (j == cnt)
                {
                    return true;
                }
                j++;

            }
            return isvalid;
        }
        public void Saves()
        {
            try
            {
                if (checkvalidate() == true)
                {

                    string maxid = ""; string maxdetid = "";
                    Models.OrderEntryModel p = new Models.OrderEntryModel();

                    p.asptblordentryid = Convert.ToInt64("0" + txtorderid.Text);
                    p.asptblordentryid1 = Convert.ToString(txtorderid1.Text);
                    p.finyear = Class.Users.Finyear;
                    p.orderdate = Convert.ToString(orderdate.Value.ToShortDateString());
                    p.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    p.compname = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    p.buyer = Convert.ToInt64("0" + combobuyer.SelectedValue);
                    p.buyingagent = Convert.ToInt64("0" + combobuyagent.SelectedValue);
                    p.orderno = Convert.ToString(txtorderno.Text);
                    p.orderqty = Convert.ToInt64("0" + txtorderqty.Text);
                    p.pono = Convert.ToString(txtpono.Text);
                    p.stylerefno = Convert.ToString("0" + txtstyrefno.Text);
                    p.stylename = Convert.ToInt64("0" + combostyle.SelectedValue);
                    p.sizegroup = Convert.ToInt64("0" + combosizegroup.SelectedValue);
                    p.merchandiser = Convert.ToInt64("0" + combomerch.SelectedValue);
                    p.currency = Convert.ToInt64("0" + combocurrency.SelectedValue);
                    p.currencyvalue = Convert.ToInt64("0" + txtconvertvalue.Text);
                    p.compcode1 = Class.Users.COMPCODE;
                    p.username = Class.Users.USERID;
                    p.createdby = Convert.ToString(Class.Users.HUserName);
                    p.createdon = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    p.modified = Convert.ToString(System.DateTime.Now.ToString());
                    p.modifiedby = Class.Users.HUserName;
                    p.ipaddress = Class.Users.IPADDRESS;
                    if (checkactive.Checked == true) { p.active = "T"; } else { p.active = "F"; checkactive.Checked = false; }
                    if (checkordercancel.Checked == true) { p.ordercancel = "T"; } else { p.ordercancel = "F"; checkordercancel.Checked = false; }
                    string sel = "select asptblordentryid    from  asptblordentry    WHERE asptblordentryid1='" + p.asptblordentryid1 + "' and finyear='" + p.finyear + "' and orderdate='" + p.orderdate + "' and compcode ='" + p.compcode + "' and buyer='" + p.buyer + "' and pono='" + p.pono + "' and orderno='" + p.orderno + "' and orderqty ='" + p.orderqty + "' and stylerefno='" + p.stylerefno + "' and stylename='" + p.stylename + "' and sizegroup='" + p.sizegroup + "'  and merchandiser='" + p.merchandiser + "' and currency='" + p.currency + "' and currencyvalue='" + p.currencyvalue + "' and ordercancel='" + p.ordercancel + "' and  active='" + p.active + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblordentry");
                    DataTable dt = ds.Tables["asptblordentry"];
                    if (dt.Rows.Count != 0)
                    {
                        // MessageBox.Show("Child Record Found " + " Alert " + txtpono.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        //  maxid = txtorderid1.Text;
                        maxid = txtorderid.Text;
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtorderid.Text) == 0 || Convert.ToInt32("0" + txtorderid.Text) == 0)
                    {
                        string ins = "insert into asptblordentry(asptblordentryid1,finyear,orderdate,compcode,buyer,pono,orderno,orderqty,stylerefno,stylename,sizegroup,merchandiser,currency,currencyvalue,ordercancel,active,compcode1,username,createdby,createdon,modified,modifiedby,ipaddress)  VALUES('" + p.asptblordentryid1 + "','" + p.finyear + "','" + p.orderdate + "','" + p.compcode + "','" + p.buyer + "','" + p.pono + "','" + p.orderno + "','" + p.orderqty + "','" + p.stylerefno + "','" + p.stylename + "','" + p.sizegroup + "','" + p.merchandiser + "','" + p.currency + "','" + p.currencyvalue + "','" + p.ordercancel + "','" + p.active + "','" + p.compcode1 + "','" + p.username + "','" + p.createdby + "','" + p.createdon + "','" + p.modified + "','" + p.modifiedby + "','" + p.ipaddress + "');";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptblordentryid) id    from  asptblordentry   where  compcode='" + combocompcode.SelectedValue + "'  and finyear='" + Class.Users.Finyear + "' and orderno='" + txtorderno.Text + "' ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblordentry");
                        DataTable dt2 = ds2.Tables["asptblordentry"];
                        maxid = dt2.Rows[0]["id"].ToString();
                    }
                    else
                    {
                        string up = "update  asptblordentry  set   asptblordentryid1='" + p.asptblordentryid1 + "' ,finyear='" + p.finyear + "' ,  orderdate='" + p.orderdate + "' , compcode ='" + p.compcode + "' , buyer='" + p.buyer + "' , pono='" + p.pono + "' , orderno='" + p.orderno + "' , orderqty ='" + p.orderqty + "' , stylerefno='" + p.stylerefno + "' , stylename='" + p.stylename + "' , sizegroup='" + p.sizegroup + "'  , merchandiser='" + p.merchandiser + "',currency='" + p.currency + "',currencyvalue='" + p.currencyvalue + "' ,ordercancel='" + p.ordercancel + "' , active='" + p.active + "' ,compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', createdon='" + p.createdon + "',modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblordentryid='" + txtorderid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        maxid = txtorderid.Text;
                    }
                    Models.OrderEntryModeldetail pp = new Models.OrderEntryModeldetail();

                    int i = 0, j = 0; int rowcount = dataGridView2.Rows.Count - 1;
                    if (rowcount >= 1)
                    {

                        for (i = 0; i < rowcount; i++)
                        {
                            if (Convert.ToInt64(dataGridView2.Rows[i].Cells["fabric"].Value.ToString()) > 0 && Convert.ToInt64(dataGridView2.Rows[i].Cells["colorname"].Value.ToString()) > 0 && Convert.ToInt64(dataGridView2.Rows[i].Cells["qty"].Value.ToString()) > 0)
                            {
                                pp.asptblordentrydetid = Convert.ToInt64("0" + dataGridView2.Rows[i].Cells["asptblordentrydetid"].Value);
                                pp.asptblordentryid = Convert.ToInt64("0" + maxid);
                                pp.asptblordentryid1 = Convert.ToString(p.asptblordentryid1);
                                pp.compcode = Convert.ToInt64("0" + p.compcode);
                                pp.orderno = Convert.ToString(p.orderno);
                                fabricid(dataGridView2.Rows[i].Cells["fabric"].FormattedValue.ToString());
                                colorid(dataGridView2.Rows[i].Cells["colorname"].FormattedValue.ToString());
                                pp.fabric = fabid;
                                pp.colorname = coid;
                                pp.qty = Convert.ToInt64("0" + dataGridView2.Rows[i].Cells["qty"].Value.ToString());
                                pp.notes = Convert.ToString(dataGridView2.Rows[i].Cells["notes"].Value.ToString());
                                pp.indexno = Convert.ToInt64(dataGridView2.Rows[i].Cells["sno"].FormattedValue.ToString());
                                string sel1 = "select asptblordentrydetid    from  asptblordentrydet   where   asptblordentryid='" + pp.asptblordentryid + "'  and asptblordentryid1='" + pp.asptblordentryid1 + "' and fabric='" + pp.fabric + "' and colorname='" + pp.colorname + "' and qty='" + pp.qty + "' and  notes='" + pp.notes + "'";
                                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblordentrydet");
                                DataTable dt1 = ds1.Tables["asptblordentrydet"];
                                if (dt1.Rows.Count != 0)
                                {
                                }
                                else if (dt1.Rows.Count != 0 && pp.asptblordentrydetid == 0 || pp.asptblordentrydetid == 0)
                                {
                                    string ins1 = "insert into asptblordentrydet(asptblordentryid,asptblordentryid1,compcode,orderno,fabric,colorname,qty,notes,indexno) values('" + pp.asptblordentryid + "' ,'" + pp.asptblordentryid1 + "' , '" + pp.compcode + "' ,'" + pp.orderno + "' ,'" + pp.fabric + "','" + pp.colorname + "','" + pp.qty + "' ,'" + pp.notes + "','" + pp.indexno + "' );";
                                    Utility.ExecuteNonQuery(ins1);

                                    string sel2 = "select max(a.asptblordentrydetid) as  maxdetid    from  asptblordentrydet a join asptblordentry b on a.asptblordentryid=b.asptblordentryid   where  b.compcode='" + combocompcode.SelectedValue + "'  and b.finyear='" + Class.Users.Finyear + "' and b.orderno='" + txtorderno.Text + "' ;";
                                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblordentrydet");
                                    DataTable dt2 = ds2.Tables["asptblordentrydet"];
                                    maxdetid = dt2.Rows[0]["maxdetid"].ToString();

                                    for (j = 0; j < listView2.Items.Count; j++)
                                    {
                                        if (listView2.Items[j].SubItems[9].Text == i.ToString())
                                        {
                                            Models.OrderEntryModelsubdetail s = new Models.OrderEntryModelsubdetail();

                                            s.asptblordentrysubdetid = Convert.ToInt64("0" + listView2.Items[j].SubItems[2].Text);
                                            s.asptblordentrydetid = Convert.ToInt64("0" + maxdetid);
                                            s.asptblordentryid = Convert.ToInt64("0" + pp.asptblordentryid);
                                            s.asptblordentryid1 = Convert.ToString(pp.asptblordentryid1);
                                            s.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                                            s.orderno = Convert.ToString("0" + txtorderno.Text);
                                            ////fabricid(listView2.Items[j].SubItems[3].Text);
                                            ////colorid(listView2.Items[j].SubItems[4].Text);
                                            sizeid(listView2.Items[j].SubItems[5].Text);
                                            s.fabric = pp.fabric;
                                            s.colorname = pp.colorname;
                                            s.sizename = siid;
                                            s.qty = Convert.ToInt64("0" + listView2.Items[j].SubItems[6].Text);
                                            s.excessqty = Convert.ToInt64("0" + listView2.Items[j].SubItems[7].Text);
                                            s.shipqty = Convert.ToInt64("0" + listView2.Items[j].SubItems[8].Text);
                                            s.indexno = Convert.ToInt64("0" + listView2.Items[j].SubItems[9].Text);

                                            string sel3 = "select asptblordentrysubdetid    from  asptblordentrysubdet   where  asptblordentrydetid='" + s.asptblordentrydetid + "'  and  asptblordentryid='" + s.asptblordentryid + "'  and asptblordentryid1='" + s.asptblordentryid1 + "' and fabric='" + s.fabric + "' and colorname='" + s.colorname + "' and sizename='" + s.sizename + "'and qty='" + s.qty + "' and  excessqty='" + s.excessqty + "' and  shipqty='" + s.shipqty + "' ";
                                            DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblordentrysubdet");
                                            DataTable dt3 = ds3.Tables["asptblordentrysubdet"];
                                            if (dt3.Rows.Count != 0)
                                            {

                                            }
                                            else if (dt3.Rows.Count != 0 && s.asptblordentrysubdetid == 0 || s.asptblordentrysubdetid == 0)
                                            {
                                                string ins3 = "insert into asptblordentrysubdet(asptblordentrydetid ,asptblordentryid ,asptblordentryid1,compcode,orderno,fabric ,colorname,sizename ,qty ,excessqty ,shipqty,indexno) values('" + s.asptblordentrydetid + "','" + s.asptblordentryid + "' ,'" + s.asptblordentryid1 + "' , '" + s.compcode + "' ,'" + s.orderno + "' ,'" + s.fabric + "','" + s.colorname + "','" + s.sizename + "','" + s.qty + "' ,'" + s.excessqty + "','" + s.shipqty + "' ,'" + s.indexno + "' );";
                                                Utility.ExecuteNonQuery(ins3);
                                            }
                                            else
                                            {
                                                string up3 = "update  asptblordentrysubdet  set  asptblordentrydetid='" + s.asptblordentrydetid + "' , asptblordentryid='" + s.asptblordentryid + "' ,asptblordentryid1='" + s.asptblordentryid1 + "', compcode='" + s.compcode + "' , orderno='" + s.orderno + "' , fabric='" + s.fabric + "', colorname='" + s.colorname + "', sizename='" + s.sizename + "',  qty='" + s.qty + "', excessqty='" + s.excessqty + "', shipqty='" + s.shipqty + "',indexno='" + s.indexno + "' where asptblordentrysubdetid='" + s.asptblordentrysubdetid + "';";
                                                Utility.ExecuteNonQuery(up3);
                                            }

                                        }
                                    }
                                }
                                else
                                {

                                    string up1 = "update  asptblordentrydet  set asptblordentryid='" + txtorderid.Text + "' ,asptblordentryid1='" + txtorderid1.Text + "', compcode='" + combocompcode.SelectedValue + "' , orderno='" + txtorderno.Text + "' , fabric='" + pp.fabric + "', colorname='" + pp.colorname + "', qty='" + pp.qty + "', notes='" + pp.notes + "', indexno='" + pp.indexno + "' where asptblordentrydetid='" + pp.asptblordentrydetid + "';";
                                    Utility.ExecuteNonQuery(up1);

                                    for (j = 0; j < listView2.Items.Count; j++)
                                    {
                                        if (listView2.Items[j].SubItems[9].Text == i.ToString())
                                        {
                                            Models.OrderEntryModelsubdetail s = new Models.OrderEntryModelsubdetail();
                                            // MessageBox.Show(listView2.Items[j].SubItems[1].Text + "==" + listView2.Items[j].SubItems[2].Text);
                                            s.asptblordentrysubdetid = Convert.ToInt64("0" + listView2.Items[j].SubItems[2].Text);
                                            s.asptblordentrydetid = Convert.ToInt64("0" + pp.asptblordentrydetid);
                                            s.asptblordentryid = Convert.ToInt64("0" + pp.asptblordentryid);
                                            s.asptblordentryid1 = Convert.ToString(txtorderid1.Text);
                                            s.compcode = Convert.ToInt64("0" + combocompcode.SelectedValue);
                                            s.orderno = Convert.ToString(txtorderno.Text);
                                            //fabricid(listView2.Items[j].SubItems[3].Text);
                                            //colorid(listView2.Items[j].SubItems[4].Text);
                                            sizeid(listView2.Items[j].SubItems[5].Text);
                                            s.fabric = pp.fabric;// fabid;
                                            s.colorname = pp.colorname;// coid;
                                            s.sizename = siid;
                                            s.qty = Convert.ToInt64("0" + listView2.Items[j].SubItems[6].Text);
                                            s.excessqty = Convert.ToInt64("0" + listView2.Items[j].SubItems[7].Text);
                                            s.shipqty = Convert.ToInt64("0" + listView2.Items[j].SubItems[8].Text);
                                            s.indexno = Convert.ToInt64(listView2.Items[j].SubItems[9].Text);

                                            string sel3 = "select asptblordentrysubdetid    from  asptblordentrysubdet   where  asptblordentrydetid='" + s.asptblordentrydetid + "'  and  asptblordentryid='" + s.asptblordentryid + "'  and asptblordentryid1='" + s.asptblordentryid1 + "' and fabric='" + s.fabric + "' and colorname='" + s.colorname + "' and sizename='" + s.sizename + "'and qty='" + s.qty + "' and  excessqty='" + s.excessqty + "' and  shipqty='" + s.shipqty + "' ";
                                            DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblordentrysubdet");
                                            DataTable dt3 = ds3.Tables["asptblordentrysubdet"];
                                            if (dt3.Rows.Count != 0)
                                            {

                                            }
                                            else if (dt3.Rows.Count != 0 && s.asptblordentrysubdetid == 0 || s.asptblordentrysubdetid == 0)
                                            {
                                                string ins3 = "insert into asptblordentrysubdet(asptblordentrydetid ,asptblordentryid ,asptblordentryid1,compcode,orderno,fabric ,colorname,sizename ,qty ,excessqty ,shipqty,indexno) values('" + s.asptblordentrydetid + "','" + s.asptblordentryid + "' ,'" + s.asptblordentryid1 + "' , '" + s.compcode + "' ,'" + s.orderno + "' ,'" + s.fabric + "','" + s.colorname + "','" + s.sizename + "','" + s.qty + "' ,'" + s.excessqty + "','" + s.shipqty + "' ,'" + s.indexno + "' );";
                                                Utility.ExecuteNonQuery(ins3);
                                            }
                                            else
                                            {
                                                string up3 = "update  asptblordentrysubdet  set  asptblordentrydetid='" + s.asptblordentrydetid + "' , asptblordentryid='" + s.asptblordentryid + "' ,asptblordentryid1='" + s.asptblordentryid1 + "', compcode='" + s.compcode + "' , orderno='" + s.orderno + "' , fabric='" + s.fabric + "', colorname='" + s.colorname + "', sizename='" + s.sizename + "',  qty='" + s.qty + "', excessqty='" + s.excessqty + "', shipqty='" + s.shipqty + "',indexno='" + s.indexno + "' where asptblordentrysubdetid='" + s.asptblordentrysubdetid + "';";
                                                Utility.ExecuteNonQuery(up3);
                                            }

                                        }
                                    }
                                }


                            }
                        }
                    }
                    if (txtorderid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtorderid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        autonumberload();
                        tabControl1.SelectTab(tabPage2);
                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtorderid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                        autonumberload();
                        tabControl1.SelectTab(tabPage2);
                    }

                }
                else
                {
                    // MessageBox.Show("pls Enter Subgrid Value ", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Saves_Click " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void OrderEntry_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }




        public void sizegroupload()
        {
            try
            {
                string sel = "select asptblsizgrpid,sizegroup from  asptblsizgrp  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsizgrp");
                DataTable dt = ds.Tables["asptblsizgrp"];

                combosizegroup.DisplayMember = "sizegroup";
                combosizegroup.ValueMember = "asptblsizgrpid";
                combosizegroup.DataSource = dt;
                // combosizegroup.Text = ""; combosizegroup.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void buyerload(string s)
        {
            try
            {
                string sel = "select a.asptblbuymasid,a.buyername from  asptblbuymas a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + s + "'  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbuymas");
                DataTable dt = ds.Tables["asptblbuymas"];

                combobuyer.DisplayMember = "buyername";
                combobuyer.ValueMember = "asptblbuymasid";
                combobuyer.DataSource = dt;

                combobuyagent.DisplayMember = "buyername";
                combobuyagent.ValueMember = "asptblbuymasid";
                combobuyagent.DataSource = dt;
                // combosizegroup.Text = ""; combosizegroup.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void employeeload()
        {
            try
            {
                string sel = "select asptblempmasid,empname from  asptblempmas  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblempmas");
                DataTable dt = ds.Tables["asptblempmas"];

                combomerch.DisplayMember = "empname";
                combomerch.ValueMember = "asptblempmasid";
                combomerch.DataSource = dt;
                // combosizegroup.Text = ""; combosizegroup.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void currencyload()
        {
            try
            {
                string sel = "select asptblcurmasid,currency from  asptblcurmas  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblcurmas");
                DataTable dt = ds.Tables["asptblcurmas"];

                combocurrency.DisplayMember = "currency";
                combocurrency.ValueMember = "asptblcurmasid";
                combocurrency.DataSource = dt;
                // combosizegroup.Text = ""; combosizegroup.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
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
        public void fabricid(string s)
        {
            try
            {
                string sel = "select asptblfabmasid from  asptblfabmas where fabric='" + s + "' ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblfabmas");
                DataTable dt = ds.Tables["asptblfabmas"];
                fabid = "";
                fabid = Convert.ToString(dt.Rows[0]["asptblfabmasid"].ToString());


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
        public void stylenameload()
        {
            try
            {
                string sel = "select asptblstymasid,stylename from  asptblstymas  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblstymas");
                DataTable dt = ds.Tables["asptblstymas"];

                combostyle.DisplayMember = "stylename";
                combostyle.ValueMember = "asptblstymasid";
                combostyle.DataSource = dt;
                combostyle.Text = ""; combostyle.SelectedIndex = -1;
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
                combocompcode.Text = ""; combocompcode.SelectedIndex = -1;

            }
            catch (Exception EX)
            { }
        }

        private void Fabricload()
        {
            try
            {

                string sel1 = "SELECT  A.ASPTBLFABMASID,A.FABRIC  FROM  asptblfabmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblfabmas");
                DataTable dt = ds.Tables["asptblfabmas"];

                Fabric.DisplayMember = "fabric";
                Fabric.ValueMember = "asptblfabmasid";
                Fabric.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void colorload()
        {
            try
            {
                string sel1 = "   SELECT a.asptblcolmasid,a.colorname   FROM  asptblcolmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblcolmas");
                DataTable dt = ds.Tables["asptblcolmas"];

                colorname.DisplayMember = "colorname";
                colorname.ValueMember = "asptblcolmasid";
                colorname.DataSource = dt;
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
                listView3.Items.Clear(); lblsizecount.Text = "";
                string sel1 = "SELECT c.asptblsizmasid, c.sizename   FROM  asptblsizgrpDet a join asptblsizgrp b on a.asptblsizgrpid=b.asptblsizgrpid join asptblsizmas c on c.asptblsizmasid=a.sizename  where b.sizegroup='" + sizgroup + "'  order by 1;";
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
                    lblsizecount.Text = dt.Rows.Count.ToString();
                }
                //SizeName.DisplayMember = "sizename";
                //SizeName.ValueMember = "sizename";
                //SizeName.DataSource = dt;
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
                string sel1 = "SELECT a.asptblordentryid,b.compcode,c.buyername,a.pono,a.orderno ,a.orderqty,d.stylename,e.sizegroup,a.active FROM  asptblordentry a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup   order by  a.asptblordentryid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordentry");
                DataTable dt = ds.Tables["asptblordentry"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblordentryid"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["buyername"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());
                        list.SubItems.Add(myRow["orderno"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["stylename"].ToString());
                        list.SubItems.Add(myRow["sizegroup"].ToString());
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
                    int i, j = 0;
                    txtorderid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblordentryid,a.asptblordentryid1,a.finyear,a.orderdate,b.compcode,c.buyername,c.buyingagent, a.orderno,a.orderqty,a.pono,a.stylerefno, d.stylename, e.sizegroup,f.empname,g.currency,a.currencyvalue, a.ordercancel,a.active  from asptblordentry a join gtcompmast b on a.compcode=b.gtcompmastid join asptblbuymas c on c.asptblbuymasid=a.buyer join asptblstymas d on d.asptblstymasid=a.stylename join asptblsizgrp e on e.asptblsizgrpid=a.sizegroup join asptblempmas f on f.asptblempmasid=a.merchandiser join asptblcurmas g on g.asptblcurmasid=a.currency  where a.asptblordentryid=" + txtorderid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordentry");
                    DataTable dt = ds.Tables["asptblordentry"];
                    if (dt.Rows.Count > 0)
                    {
                        txtorderid.Text = Convert.ToString(dt.Rows[0]["asptblordentryid"].ToString());
                        txtorderid1.Text = Convert.ToString(dt.Rows[0]["asptblordentryid1"].ToString());
                        orderdate.Value = Convert.ToDateTime(dt.Rows[0]["orderdate"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyername"].ToString());
                        combobuyagent.Text = Convert.ToString(dt.Rows[0]["buyingagent"].ToString());
                        txtorderno.Text = Convert.ToString(dt.Rows[0]["orderno"].ToString());
                        txtorderqty.Text = Convert.ToString(dt.Rows[0]["orderqty"].ToString());
                        txtpono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        txtpono.Text = Convert.ToString(dt.Rows[0]["stylerefno"].ToString());
                        combostyle.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        combosizegroup.Text = Convert.ToString(dt.Rows[0]["sizegroup"].ToString());
                        combomerch.Text = Convert.ToString(dt.Rows[0]["empname"].ToString());
                        combocurrency.Text = Convert.ToString(dt.Rows[0]["currency"].ToString());
                        txtconvertvalue.Text = Convert.ToString(dt.Rows[0]["currencyvalue"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                        if (dt.Rows[0]["ordercancel"].ToString() == "T") { checkordercancel.Checked = true; } else { checkordercancel.Checked = true; checkordercancel.Checked = false; }
                    }
                    string sel2 = "select a.asptblordentrydetid,a.compcode,a.orderno, d.asptblfabmasid as fabric, e.asptblcolmasid as colorname,a.qty,a.notes,a.indexno from asptblordentrydet a join asptblordentry b on a.asptblordentryid=b.asptblordentryid join gtcompmast c on c.gtcompmastid=b.compcode join  asptblfabmas d on d.asptblfabmasid=a.fabric join asptblcolmas e on e.asptblcolmasid=a.colorname   where a.asptblordentryid=" + txtorderid.Text;
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblordentry");
                    DataTable dt2 = ds2.Tables["asptblordentry"];
                    if (dt2.Rows.Count > 0)
                    {
                        listfilter4.Items.Clear();
                        dataGridView2.DataSource = dt2;

                        for (j = 0; j < dt2.Rows.Count; j++)
                        {


                            if (Convert.ToInt64(dataGridView2.Rows[j].Cells[1].Value) > 0)
                            {
                                dataGridView2.Rows[j].Cells[1].Value = Convert.ToInt32("0" + dt2.Rows[j]["asptblordentrydetid"].ToString());
                                dataGridView2.Rows[j].Cells[2].Value = Convert.ToInt32("0" + dt2.Rows[j]["compcode"].ToString());
                                dataGridView2.Rows[j].Cells[3].Value = Convert.ToString(dt2.Rows[j]["orderno"].ToString());
                                dataGridView2.Rows[j].Cells[4].Value = Convert.ToInt32("0" + dt2.Rows[j]["fabric"].ToString());
                                dataGridView2.Rows[j].Cells[5].Value = Convert.ToInt64("0" + dt2.Rows[j]["colorname"].ToString());
                                dataGridView2.Rows[j].Cells[6].Value = Convert.ToInt32("0" + dt2.Rows[j]["qty"].ToString());
                                dataGridView2.Rows[j].Cells[7].Value = Convert.ToString(dt2.Rows[j]["notes"].ToString());
                                dataGridView2.BackColor = j % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;


                            }
                        }

                        lbltotal.Text = "Total Count    :" + dataGridView2.Rows.Count;
                    }

                    listView2.Items.Clear();
                    string sel3 = "select a.asptblordentrysubdetid,e.fabric, f.colorname,g.sizename, a.qty,a.excessqty ,a.shipqty,a.indexno from asptblordentrysubdet a   join asptblordentrydet b on a.asptblordentrydetid=b.asptblordentrydetid  join asptblordentry c on c.asptblordentryid=a.asptblordentryid join gtcompmast d on d.gtcompmastid=a.compcode join  asptblfabmas e on e.asptblfabmasid=a.fabric join asptblcolmas f on f.asptblcolmasid=a.colorname  join asptblsizmas g on g.asptblsizmasid=a.sizename   where a.asptblordentryid=" + txtorderid.Text;
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "asptblordentry");
                    DataTable dt3 = ds3.Tables["asptblordentry"];
                    i = 1;
                    foreach (DataRow myRow in dt3.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblordentrysubdetid"].ToString());
                        list.SubItems.Add(myRow["fabric"].ToString());
                        list.SubItems.Add(myRow["colorname"].ToString());
                        list.SubItems.Add(myRow["sizename"].ToString());
                        list.SubItems.Add(myRow["qty"].ToString());
                        list.SubItems.Add(myRow["excessqty"].ToString());
                        list.SubItems.Add(myRow["shipqty"].ToString());
                        list.SubItems.Add(myRow["indexno"].ToString());
                        if (i % 2 == 0)
                        {
                            list.BackColor = System.Drawing.Color.White;

                        }
                        else
                        {
                            list.BackColor = System.Drawing.Color.WhiteSmoke;
                        }
                        listView2.Items.Add(list);
                        i++;
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            tabControl1.SelectTab(tabPage1); combocompcode.Select(); dataGridView1.Rows.Clear();
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {


                int item0 = 0; listView1.Items.Clear(); int i = listfilter.Items.Count;
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
                            if (i % 2 == 0)
                            {
                                list.BackColor = System.Drawing.Color.White;

                            }
                            else
                            {
                                list.BackColor = System.Drawing.Color.WhiteSmoke;
                            }
                            listView1.Items.Add(list);


                        }
                        item0++; i++;
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

        }

        public void Deletes()
        {
            try
            {
                if (listView2.Items.Count > 0)
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        int i = 0;
                        for (i = 0; i < listView2.Items.Count; i++)
                        {
                            if (txtorderid.Text != "")
                            {
                                if (listView2.Items[i].SubItems[7].Text == griddelrow.ToString())
                                {
                                    if (Convert.ToInt64("0" + listView2.Items[i].SubItems[2].Text) > 1)
                                    {
                                        string del1 = "delete from ASPTBLNHDAYDET where asptblholdetid='" + listView2.Items[i].SubItems[2].Text + "';";
                                        Utility.ExecuteNonQuery(del1);
                                    }
                                    MessageBox.Show("Index of Row:   " + listView2.Items[i].SubItems[7].Text + "      Name:  " + listView2.Items[i].SubItems[2].Text + "Deleted Successfully", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);


                                    listView2.Items[i].Remove();
                                    i--;
                                }
                            }
                            else
                            {
                                dataGridView1.Rows.Clear(); tot1 = 0; tot = 0; totalqty = 0; shiptot1 = 0; shptot = 0;
                                lbltotalall.Text = "";
                                lblorderqty.Text = "";
                                lblshipqty.Text = "";
                            }

                        }


                    }
                }
                else
                {
                    MessageBox.Show("If you want to Remove,Double Click a Specific Row in ListView.", "error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
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


        private void ContextMenuStrip3_Click(object sender, EventArgs e)
        {

        }


        ListView allip = new ListView();
        ListView allip1 = new ListView();
        ListView allip2 = new ListView();
        ListView allip3 = new ListView();
        ListView allip4 = new ListView();
        private void listView2_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {


                ListViewItem it = new ListViewItem();
                if (e.Item.Checked == true)
                {

                    e.Item.SubItems[3].Text = "Connected";


                    e.Item.BackColor = SystemColors.MenuHighlight;
                    e.Item.ForeColor = System.Drawing.Color.White;
                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip2.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[3].Text == "Connected")
                {
                    e.Item.SubItems[3].Text = "DisConnected";
                    e.Item.Checked = false;
                    e.Item.BackColor = SystemColors.ControlLight;
                    e.Item.ForeColor = System.Drawing.Color.Black;
                    for (int c = 0; c < allip2.Items.Count; c++)
                    {
                        if (allip2.Items[c].SubItems[2].Text == "True")
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

                    e.Item.SubItems[4].Text = "Connected";
                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.SubItems[3].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip3.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[4].Text == "Connected")
                {
                    e.Item.SubItems[4].Text = "DisConnected";
                    for (int c = 0; c < allip3.Items.Count; c++)
                    {
                        if (allip3.Items[c].SubItems[3].Text == "true")
                        {
                            allip3.Items[c].Remove();
                            c--;
                        }
                    }
                    Cursor = Cursors.Default;
                }

            }
            catch (Exception ex) { 
               // MessageBox.Show("       ---  " + e.Item.ToString() + "             ======================" + ex.ToString()); 
            }
        }
        private void listView4_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {

                ListViewItem it = new ListViewItem();
                if (e.Item.Checked == true)
                {

                    e.Item.SubItems[3].Text = "Connected";
                    it.SubItems.Add(e.Item.SubItems[2].Text);
                    it.SubItems.Add(e.Item.Checked.ToString());
                    allip4.Items.Add(it);


                }
                if (e.Item.Checked == false && e.Item.SubItems[3].Text == "Connected")
                {
                    e.Item.SubItems[3].Text = "DisConnected";
                    for (int c = 0; c < allip4.Items.Count; c++)
                    {
                        if (allip4.Items[c].SubItems[2].Text == "true")
                        {
                            allip4.Items[c].Remove();
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




        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void refeshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            empty(); GridLoad(); colorload(); sizegroupload(); compload(); stylenameload();
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {

                if (txtorderid.Text == "")
                {
                    empty();
                    autonumberload();
                }
                combocompcode.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])//your specific tabname
            {
                txtsearch.Select();

            }
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }

        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView2.SelectedCells)
                {
                    if (oneCell.Selected)
                    {
                        dataGridView2.Rows.RemoveAt(oneCell.RowIndex);
                        if (txtorderid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow >= 1)
                                {
                                    string sel = "select * from  asptblordentrysubdet a join asptblpur b on a.orderno=b.orderno  Where  b.asptblordentrydetid='" + griddelrow + "';; ";
                                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblordentrysubdet");
                                    DataTable dt = ds.Tables["asptblordentrysubdet"];
                                    if (dt.Rows.Count >= 1)
                                    {
                                        MessageBox.Show("Child Record Found.Can Can not Delete.");
                                    }
                                    else
                                    {

                                        string del1 = "DELETE   FROM  asptblordentrysubdet     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  asptblordentrydetid='" + griddelrow + "'";
                                        Utility.ExecuteNonQuery(del1);
                                        string del2 = "DELETE   FROM  asptblordentrydet     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  asptblordentrydetid='" + griddelrow + "'";
                                        Utility.ExecuteNonQuery(del1);
                                        griddelrow = 0;
                                    }
                                }
                            }
                        }
                    }


                }
            }
            catch (Exception EX)
            {
                // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            }
        }

        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 8)
                {
                    int[] qty = new int[dataGridView1.Rows.Count];
                    int[] shipqty = new int[dataGridView1.Rows.Count];
                    if (dataGridView1.Rows[e.RowIndex].Cells[4].Value == null)
                    {

                    }
                    if (Convert.ToInt32("0" + dataGridView1.Rows[e.RowIndex].Cells[5].Value) > 1 && Convert.ToInt32("0" + dataGridView1.Rows[e.RowIndex].Cells[6].Value) > 1)
                    {
                        qty[e.RowIndex] = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[5].Value);
                        tot1 += qty[e.RowIndex];
                        if (totalqty < tot1 || totalqty == tot1)
                        {
                            if (totalqty != tot)
                            {
                                if (Convert.ToInt32(lblsizecount.Text) == qty.Count())
                                {
                                    shptot = Convert.ToInt32(Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[5].Value) + (Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[5].Value) * Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[6].Value) / 100));
                                    dataGridView1.Rows[e.RowIndex].Cells[7].Value = shptot.ToString();
                                    shiptot1 += Convert.ToInt32("0" + dataGridView1.Rows[e.RowIndex].Cells[7].Value);
                                    lblshipqty.Text = "S-Qty : " + shiptot1.ToString();
                                    return;
                                }
                            }
                        }
                        else
                        {
                            tot = totalqty - tot1;
                            shptot = Convert.ToInt32(Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[5].Value) + (Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[5].Value) * Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[6].Value) / 100));
                            dataGridView1.Rows[e.RowIndex].Cells[7].Value = shptot.ToString();
                            dataGridView1.Rows[e.RowIndex + 1].Cells[5].Value = tot.ToString(); ;
                        }
                    }
                    else
                    {
                        if (Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[5].Value) == "")
                        {
                            this.dataGridView1.Rows[e.RowIndex].Cells[5].Selected = true;

                            MessageBox.Show("pls Enter Excess Qty Field " + " Alert " + e.RowIndex, "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                        }
                    }
                    shiptot1 = 0; ordqty1 = 0;
                    for (int k = 0; k < dataGridView1.Rows.Count + 1; k++)
                    {
                        if (k < Convert.ToInt32(lblsizecount.Text))
                        {
                            ordqty1 += Convert.ToInt32("0" + dataGridView1.Rows[k].Cells[5].Value);
                            lblorderqty.Text = "O-Qty : " + ordqty1.ToString();
                            shiptot1 += Convert.ToInt32("0" + dataGridView1.Rows[k].Cells[7].Value);
                            lblshipqty.Text = "S-Qty : " + shiptot1.ToString();
                        }
                    }
                }
            }
            catch (Exception ex) { }
        }


        //private void refreshToolStripMenuItem1_Click_1(object sender, EventArgs e)
        //{
        //    for (int i = 0; i < dataGridView1.Rows.Count; i++)
        //    {
        //        if (i == 0)
        //        {
        //            dataGridView1.Rows[i].Cells[4].Value = totalqty.ToString();
        //        }

        //    }
        //}
        private void txtorderqty_KeyPress(object sender, KeyPressEventArgs e)
        {
            txtbox_KeyPress(sender, e);
        }
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox txtbox = e.Control as TextBox;
            if (txtbox != null)
            {
                txtbox.KeyPress += new KeyPressEventHandler(txtbox_KeyPress);
            }
        }

        void txtbox_KeyPress(object sender, KeyPressEventArgs e)
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

        private void datagridviewrow()
        {

            dataGridView1.ColumnCount = 9;
            dataGridView1.Columns[0].Name = "S.No";
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[0].Width = 35;
            dataGridView1.Columns[1].Name = "ID";
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[2].Name = "Fabric Name";
            dataGridView1.Columns[2].Width = 120;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].Name = "Color";
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[3].ReadOnly = true;
            dataGridView1.Columns[4].Name = "Size";
            dataGridView1.Columns[4].Width = 40;
            dataGridView1.Columns[4].ReadOnly = true;
            dataGridView1.Columns[5].Name = "Qty *";
            dataGridView1.Columns[5].Width = 70;
            dataGridView1.Columns[5].ContextMenuStrip = contextMenuStrip3;
            dataGridView1.Columns[6].Name = "Ex-Qty *";
            dataGridView1.Columns[6].Width = 70;
            dataGridView1.Columns[7].Name = "Ship Qty";
            dataGridView1.Columns[7].Width = 70;
            dataGridView1.Columns[7].ReadOnly = true;
            dataGridView1.Columns[8].Name = "Index";
            dataGridView1.Columns[8].Width = 1;
            dataGridView1.Columns[8].ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            CommonFunctions.SetRowNumber(dataGridView1);
            //dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.MenuHighlight;
            //dataGridView1.EnableHeadersVisualStyles = false;
            //dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex == 8)
            {

                butclear_Click(sender, e);

                var senderGrid = (DataGridView)sender;

                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewLinkColumn && e.RowIndex >= 0 && dataGridView2.Rows[e.RowIndex].Cells[4].Value != null && dataGridView2.Rows[e.RowIndex].Cells[5].Value != null && dataGridView2.Rows[e.RowIndex].Cells[6].Value != null)
                {
                    griddelrow = e.RowIndex; bool valid = false;
                    List<string> list4 = new List<string>(); dataGridView1.Rows.Clear();
                    List<string> list2 = new List<string>();
                    int rowcount = 0; totalqty = 0; Int64 totqty = 0, totshipqty = 0;
                    rowcount = dataGridView2.Rows.Count - 1;
                    string[] col; string[] fab;
                    string[] siz; int i, j, k, l; int oqty = 0, sqty = 0;
                    totalqty = Convert.ToInt32("0" + dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString());
                    datagridviewrow();
                    for (i = 0; i < rowcount; i++)
                    {

                        if (listView2.Items.Count < 1)
                        {
                            list4.Add(dataGridView2.Rows[i].Cells[4].FormattedValue.ToString());
                            list2.Add(dataGridView2.Rows[i].Cells[5].FormattedValue.ToString());
                        }
                        if (listView2.Items.Count > 1)
                        {

                            if (this.dataGridView2.Rows[i].Cells[0].Value.ToString() == e.RowIndex.ToString())
                            {
                                for (l = 0; l < listView2.Items.Count; l++)
                                {

                                    if (listView2.Items[l].SubItems[9].Text == e.RowIndex.ToString() && listView2.Items[l].SubItems[3].Text == this.dataGridView2.Rows[i].Cells[4].FormattedValue.ToString() && listView2.Items[l].SubItems[4].Text == this.dataGridView2.Rows[i].Cells[5].FormattedValue.ToString())
                                    {
                                        valid = true;

                                        dataGridView1.Rows.Add("", listView2.Items[l].SubItems[2].Text, listView2.Items[l].SubItems[3].Text, listView2.Items[l].SubItems[4].Text, listView2.Items[l].SubItems[5].Text, listView2.Items[l].SubItems[6].Text, listView2.Items[l].SubItems[7].Text, listView2.Items[l].SubItems[8].Text, listView2.Items[l].SubItems[9].Text);

                                        totqty += Convert.ToInt64("0" + listView2.Items[l].SubItems[6].Text);
                                        lblorderqty.Text = totqty.ToString();
                                        totshipqty += Convert.ToInt64("0" + listView2.Items[l].SubItems[8].Text);
                                        lblshipqty.Text = totshipqty.ToString();


                                    }

                                }
                                if (listView2.Items[i].SubItems[9].Text != e.RowIndex.ToString() && valid == false)
                                {

                                    list4.Add(dataGridView2.Rows[i].Cells[4].FormattedValue.ToString());
                                    list2.Add(dataGridView2.Rows[i].Cells[5].FormattedValue.ToString());
                                }

                            }


                        }

                    }
                    fab = list4.ToArray<string>();
                    col = list2.ToArray<string>();
                    List<string> list3 = new List<string>();
                    int item0 = 1;
                    foreach (ListViewItem item in listView3.Items)
                    {
                        list3.Add(item.SubItems[2].Text);

                        item0++;
                    }

                    siz = list3.ToArray<string>();

                    for (i = 0; i < fab.Length; i++)
                    {

                        for (k = 0; k < siz.Length; k++)
                        {
                            if (i == 0 && k == 0)
                            {
                                dataGridView1.Rows.Add("", "-1", fab[i].ToString(), col[i].ToString(), siz[k].ToString(), totalqty, "", "", e.RowIndex);

                                dataGridView1.Rows[k].Cells[5].Selected = true;
                            }
                            else
                            {

                                dataGridView1.Rows.Add("", "-1", fab[i].ToString(), col[i].ToString(), siz[k].ToString(), "", "", "", e.RowIndex);

                            }

                        }

                    }
                    allip3.Items.Clear();
                    lbltotalall.Text = "T-Rows :" + dataGridView1.Rows.Count;

                }
                else
                {
                    MessageBox.Show("Invalid.Pls enter qty");
                }
            }
        }
        private void butok_Click(object sender, EventArgs e)
        {
            lblorderqty.Text = ""; lbltotal.Text = ""; lblshipqty.Text = "";
            int rowcnt2 = dataGridView2.Rows.Count - 1;
            int rowcnt = rowcnt2 - griddelrow;
            int rowcnt1 = 0; int toval = 0;
            if (rowcnt >= 1)
            {
                if (txtorderid.Text != "")
                {
                    for (int l = 0; l < rowcnt; l++)
                    {
                        if (dataGridView1.Rows.Count > 0)
                        {

                            if (Convert.ToInt32(dataGridView1.Rows[l].Cells[1].Value.ToString()) < 1)
                            {
                                for (int k = 0; k < listView2.Items.Count; k++)
                                {
                                    if (listView2.Items[k].SubItems[9].Text == griddelrow.ToString())
                                    {
                                        listView2.Items[k].Remove();
                                        k--;
                                    }
                                }

                                if (dataGridView1.Rows[l].Cells[1].Value.ToString() == "-1" || dataGridView1.Rows[l].Cells[1].Value.ToString() == "0")
                                {
                                    toval = 0;
                                    for (int k = 0; k < dataGridView1.Rows.Count; k++)
                                    {
                                        if (Convert.ToString(dataGridView1.Rows[k].Cells[8].Value.ToString()) == griddelrow.ToString())
                                        {
                                            ListViewItem list = new ListViewItem();
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[0].Value.ToString());
                                            list.SubItems.Add("0");
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[2].Value.ToString());
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[3].Value.ToString());
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[4].Value.ToString());
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[5].Value.ToString());
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[6].Value.ToString());
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[7].Value.ToString());
                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[8].Value.ToString());
                                            listView2.Items.Add(list);
                                            if (k % 2 == 0)
                                            {
                                                list.BackColor = System.Drawing.Color.White;

                                            }
                                            else
                                            {
                                                list.BackColor = System.Drawing.Color.WhiteSmoke;
                                            }
                                            toval += Convert.ToInt32(dataGridView1.Rows[k].Cells[5].Value.ToString());

                                        }

                                    }

                                }
                            }

                            if (Convert.ToInt32(dataGridView1.Rows[l].Cells[1].Value.ToString()) > 1)
                            {
                                bool ss = false;


                                for (int k = 0; k < listView2.Items.Count; k++)
                                {
                                    if (ss == false)
                                    {
                                        if (listView2.Items[k].SubItems[9].Text == griddelrow.ToString())
                                        {
                                            ss = true;
                                            for (int p = 0; p < listView2.Items.Count; p++)
                                            {
                                                if (listView2.Items[p].SubItems[9].Text == griddelrow.ToString())
                                                {
                                                    listView2.Items[p].Remove();
                                                    p--;
                                                }
                                            }
                                            toval = 0;

                                            for (int t = 0; t < dataGridView1.Rows.Count; t++)
                                            {
                                                ListViewItem list = new ListViewItem();
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[0].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[1].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[2].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[3].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[4].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[5].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[6].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[7].Value.ToString());
                                                list.SubItems.Add(dataGridView1.Rows[t].Cells[8].Value.ToString());
                                                listView2.Items.Add(list);
                                                toval += Convert.ToInt32(dataGridView1.Rows[t].Cells[5].Value.ToString());

                                            }
                                            //dataGridView2.Rows[l].Cells[7].Value = toval.ToString();
                                            //MessageBox.Show(dataGridView2.Rows[l].Cells[7].Value.ToString()+"=="+l.ToString() +" ==="+ dataGridView2.Columns[7].Index.ToString());
                                        }
                                    }

                                }

                            }

                            dataGridView2.Rows[griddelrow].Cells[7].Value = toval.ToString();
                        }
                        else
                        {
                            MessageBox.Show("No data found in subgridview");
                            return;
                        }
                    }

                }
                else
                {
                    toval = 0;
                    if (listView2.Items.Count <= 0)
                    {

                        // if (rowcnt != rowcnt1)
                        // {

                        for (int l = griddelrow; l < dataGridView1.Rows.Count; l++)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[0].Value.ToString());
                            list.SubItems.Add("0");
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[2].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[3].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[4].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[5].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[6].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[7].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[8].Value.ToString());
                            // list.SubItems.Add(dataGridView1.Rows[l].Cells[9].Value.ToString());1093101104944;

                            listView2.Items.Add(list);
                            toval += Convert.ToInt32("0" + dataGridView1.Rows[l].Cells[5].Value.ToString());
                            if (l % 2 == 0)
                            {
                                list.BackColor = System.Drawing.Color.White;

                            }
                            else
                            {
                                list.BackColor = System.Drawing.Color.WhiteSmoke;
                            }
                        }
                        dataGridView2.Rows[griddelrow].Cells[7].Value = toval.ToString();
                        rowcnt1++;
                        // }
                        return;
                    }

                    if (listView2.Items.Count > 1)
                    {

                        for (int l = 0; l < listView2.Items.Count; l++)
                        {
                            if (listView2.Items[l].SubItems[9].Text == griddelrow.ToString())
                            {
                                listView2.Items[l].Remove();
                                l--;
                            }
                        }

                        for (int l = 0; l < dataGridView1.Rows.Count; l++)
                        {
                            ListViewItem list = new ListViewItem();
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[0].Value.ToString());
                            list.SubItems.Add("0");
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[2].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[3].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[4].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[5].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[6].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[7].Value.ToString());
                            list.SubItems.Add(dataGridView1.Rows[l].Cells[8].Value.ToString());
                            // list.SubItems.Add(dataGridView1.Rows[l].Cells[9].Value.ToString());1093101104944;
                            if (l % 2 == 0)
                            {
                                list.BackColor = System.Drawing.Color.White;

                            }
                            else
                            {
                                list.BackColor = System.Drawing.Color.WhiteSmoke;
                            }
                            listView2.Items.Add(list);
                            toval += Convert.ToInt32("0" + dataGridView1.Rows[l].Cells[5].Value.ToString());

                        }
                    }
                    dataGridView2.Rows[griddelrow].Cells[7].Value = toval.ToString();
                }



                dataGridView1.Rows.Clear();
            }
            else
            {
                MessageBox.Show("indalid");
                // datagridviewrow();
            }

            foreach (DataGridViewRow row2 in dataGridView2.Rows)
            {
                if (row2.Cells[4].Value == null)
                {
                    dataGridView2.ClearSelection();
                    int nRowIndex = dataGridView2.Rows.Count - 1;
                    dataGridView2.Rows[nRowIndex].Cells[4].Selected = true;

                }

            }

        }

        //private void butok_Click(object sender, EventArgs e)
        //{
        //    lblorderqty.Text = ""; lbltotal.Text = ""; lblshipqty.Text = "";
        //    int rowcnt2 = dataGridView2.Rows.Count - 1;
        //    int rowcnt = rowcnt2 - griddelrow;
        //    int rowcnt1 = 0;
        //    if (rowcnt >= 1)
        //    {
        //        if (txtorderid.Text != "")
        //        {
        //            bool sss = false; int toval = 0;
        //            for (int l = 0; l < rowcnt2; l++)
        //            {
        //                for (int z = 0; z < listView2.Items.Count; z++)
        //                {

        //                    if (sss == false)
        //                    {
        //                        if (griddelrow.ToString() == listView2.Items[z].SubItems[9].Text.ToString())
        //                        {
        //                            sss = true;
        //                            if (Convert.ToInt32(dataGridView1.Rows[l].Cells[1].Value.ToString()) < 1)
        //                            {
        //                                for (int k = 0; k < listView2.Items.Count; k++)
        //                                {
        //                                    if (listView2.Items[k].SubItems[9].Text == griddelrow.ToString())
        //                                    {
        //                                        listView2.Items[k].Remove();
        //                                        k--;
        //                                    }
        //                                }
        //                                if (dataGridView1.Rows[l].Cells[1].Value.ToString() == "-1" || dataGridView1.Rows[l].Cells[1].Value.ToString() == "0")
        //                                {
        //                                    toval = 0;
        //                                    for (int k = 0; k < dataGridView1.Rows.Count; k++)
        //                                    {
        //                                        if (Convert.ToString(dataGridView1.Rows[k].Cells[8].Value.ToString()) == griddelrow.ToString())
        //                                        {
        //                                            ListViewItem list = new ListViewItem();
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[0].Value.ToString());
        //                                            list.SubItems.Add("0");
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[2].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[3].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[4].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[5].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[6].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[7].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[k].Cells[8].Value.ToString());
        //                                            listView2.Items.Add(list);
        //                                            toval += Convert.ToInt32(dataGridView1.Rows[k].Cells[5].Value.ToString());
        //                                        }

        //                                    }

        //                                }
        //                            }
        //                            if (Convert.ToInt32(dataGridView1.Rows[l].Cells[1].Value.ToString()) > 1)
        //                            {
        //                                bool ss = false;
        //                                sss = true;


        //                                if (ss == false)
        //                                {
        //                                    if (listView2.Items[z].SubItems[9].Text == griddelrow.ToString())
        //                                    {
        //                                        ss = true; toval = 0;
        //                                        for (int p = 0; p < listView2.Items.Count; p++)
        //                                        {
        //                                            if (listView2.Items[p].SubItems[9].Text == griddelrow.ToString())
        //                                            {
        //                                                listView2.Items[p].Remove();
        //                                                p--;
        //                                            }
        //                                        }

        //                                        for (int t = 0; t < dataGridView1.Rows.Count; t++)
        //                                        {
        //                                            ListViewItem list = new ListViewItem();
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[0].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[1].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[2].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[3].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[4].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[5].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[6].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[7].Value.ToString());
        //                                            list.SubItems.Add(dataGridView1.Rows[t].Cells[8].Value.ToString());
        //                                            listView2.Items.Add(list);
        //                                            toval += Convert.ToInt32(dataGridView1.Rows[t].Cells[5].Value.ToString());

        //                                        }

        //                                    }
        //                                }



        //                            }
        //                            dataGridView2.Rows[griddelrow].Cells[7].Value = toval.ToString();
        //                        }
        //                    }
        //                }
        //            }
        //        }              
        //        else
        //        {
        //            if (listView2.Items.Count <= 0)
        //            {

        //                if (rowcnt != rowcnt1)
        //                {

        //                    for (int l = 0; l < dataGridView1.Rows.Count; l++)
        //                    {
        //                        ListViewItem list = new ListViewItem();
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[0].Value.ToString());
        //                        list.SubItems.Add("0");
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[2].Value.ToString());
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[3].Value.ToString());
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[4].Value.ToString());
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[5].Value.ToString());
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[6].Value.ToString());
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[7].Value.ToString());
        //                        list.SubItems.Add(dataGridView1.Rows[l].Cells[8].Value.ToString());
        //                        // list.SubItems.Add(dataGridView1.Rows[l].Cells[9].Value.ToString());1093101104944;

        //                        listView2.Items.Add(list);

        //                    }

        //                    rowcnt1++;
        //                }
        //            }

        //            if (listView2.Items.Count > 1)
        //            {

        //                for (int l = 0; l < listView2.Items.Count; l++)
        //                {
        //                    if (listView2.Items[l].SubItems[9].Text == griddelrow.ToString())
        //                    {
        //                        listView2.Items[l].Remove();
        //                        l--;
        //                    }
        //                }

        //                for (int l = 0; l < dataGridView1.Rows.Count; l++)
        //                {
        //                    ListViewItem list = new ListViewItem();
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[0].Value.ToString());
        //                    list.SubItems.Add("0");
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[2].Value.ToString());
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[3].Value.ToString());
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[4].Value.ToString());
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[5].Value.ToString());
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[6].Value.ToString());
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[7].Value.ToString());
        //                    list.SubItems.Add(dataGridView1.Rows[l].Cells[8].Value.ToString());
        //                    // list.SubItems.Add(dataGridView1.Rows[l].Cells[9].Value.ToString());1093101104944;

        //                    listView2.Items.Add(list);


        //                }
        //            }
        //        }


        //        //for (int i = 0; i < rowcnt2; i++)
        //        //{
        //        //    //if (row2.Cells[4].Value == null)
        //        //    //{
        //        //    //    dataGridView2.ClearSelection();
        //        //    //    int nRowIndex = dataGridView2.Rows.Count - 1;                      
        //        //    //    dataGridView2.Rows[nRowIndex].Cells[5].Selected = true;

        //        //    //}
        //        //    int toval = 0;
        //        //    for (int j = 0; j < dataGridView1.Rows.Count; j++)
        //        //    {
        //        //        if (Convert.ToString(dataGridView2.Rows[i].Cells[0].Value) == griddelrow.ToString())
        //        //        {
        //        //            toval += Convert.ToInt32(dataGridView1.Rows[i].Cells[5].Value.ToString());


        //        //        }
        //        //    }
        //        //    dataGridView2.Rows[i].Cells[7].Value = toval.ToString();
        //        //}
        //        dataGridView1.Rows.Clear();
        //    }
        //    else
        //    {
        //        MessageBox.Show("indalid");
        //       // datagridviewrow();
        //    }
        //}




        //private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        //{


        //    //if (listView2.Items[l].SubItems[7].Text != e.RowIndex.ToString())
        //    //{
        //    //    if (listView2.Items[l].SubItems[7].Text == e.RowIndex.ToString() && listView2.Items[l].SubItems[2].Text == dataGridView2.Rows[e.RowIndex].Cells[5].FormattedValue.ToString() && listView2.Items[l].SubItems[3].Text == dataGridView2.Rows[e.RowIndex].Cells[6].FormattedValue.ToString())
        //    //    {
        //    //        listView2.Items[l].Remove();
        //    //        l--;
        //    //    }

        //    //}
        //}

        private void butclear_Click(object sender, EventArgs e)
        {

            dataGridView1.Rows.Clear(); tot1 = 0; tot = 0; totalqty = 0; shiptot1 = 0; shptot = 0; ordqty1 = 0;
            lbltotalall.Text = "";
            lblorderqty.Text = "";
            lblshipqty.Text = "";
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    if (i == 0)
            //    {
            //        dataGridView1.Rows[i].Cells[4].Value = totalqty.ToString();
            //    }

            //}
        }

        private void Searchs_Click(object sender, EventArgs e)
        {

        }

        private void refresh2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            empty(); GridLoad(); colorload(); sizegroupload(); compload(); stylenameload();
        }

        private void combocompcode_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            autonumberload();
            buyerload(combocompcode.Text);
        }

        private void butSave_Click(object sender, EventArgs e)
        {
            Saves();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                dataGridView1.Rows.Clear(); tot1 = 0; tot = 0; totalqty = 0; shiptot1 = 0; shptot = 0; ordqty1 = 0;
                lbltotalall.Text = "";
                lblorderqty.Text = ""; griddelrow = -1;
                lblshipqty.Text = "";
                //SubGrid sub = new SubGrid();

                //sub.Show();


            }
            try
            {
                if (txtorderid.Text != "")
                {
                    griddelrow = 0;
                    griddelrow = Convert.ToInt32("0" + dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
                }
                // dataGridView1.BeginEdit(true);
            }
            catch (Exception ex)
            {

            }

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // listView2.Items.Clear();
            // empty();
        }

        private void txtconvertvalue_KeyPress(object sender, KeyPressEventArgs e)
        {
            txtbox_KeyPress(sender, e);
        }

        private void dataGridView2_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void dataGridView2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox txtbox = e.Control as TextBox;
            if (txtbox != null)
            {
                txtbox.KeyPress += new KeyPressEventHandler(txtbox_KeyPress);
            }
        }

        private void tabControl3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        List<string> sizedupli = new List<string>();
        private void dataGridView2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //if (e.ColumnIndex == 5)
            //{
            //    int cc = 0;
            //    foreach (DataGridViewCell oneCell in dataGridView2.SelectedCells)
            //    {
            //        if (oneCell.Selected)
            //        {

            //            string s = dataGridView2.Rows[oneCell.RowIndex].Cells[4].EditedFormattedValue.ToString() + "-" + dataGridView2.Rows[oneCell.RowIndex].Cells[5].EditedFormattedValue.ToString(); ;
            //            if (!(sizedupli.Contains(s)) && s != null)
            //            {
            //                sizedupli.Add(s);

            //            }
            //            else
            //            {
            //                MessageBox.Show("Duplicate" + s.ToString());

            //            }
            //        }
            //    }

            //}
        }

        private void refresh3ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            butclear_Click(sender, e);

        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sel = "select asptblbuymasid,buyercode from  asptblbuymas a join gtcompmast b on a.compcode=b.gtcompmastid where b.compcode='" + combocompcode.Text + "'   order by 1 ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbuymas");
            DataTable dt = ds.Tables["asptblbuymas"];
            combobuyer.DisplayMember = "buyercode";
            combobuyer.ValueMember = "asptblbuymasid";
            combobuyer.DataSource = dt;
            autonumberload();
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

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void butNew_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            CommonFunctions.SetRowNumber(dataGridView2);
        }

        private void combobuyer_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class StyleMaster : Form,ToolStripAccess
    {
        private static StyleMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        public static StyleMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new StyleMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public StyleMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());

        }
    
        private void StyleMaster_Load(object sender, EventArgs e)
        {
          
            txtstylename.Select();

        }
        private bool ValidateContact(TextBox s)
        {
            if (!Regex.Match(s.Text, "^[A-Z][a-zA-Z]*$").Success)
            {

                s.Text.Remove(s.Text.Length - 1);

            }
            return true;
        }

        public void Saves()
        {
            try
            {
                if (combocompcode.Text == "")
                {
                    MessageBox.Show("pls Enter CompCode ", "CompCode", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    combocompcode.Focus();
                    return;
                }
                if (combobuyer.Text == "")
                {
                    MessageBox.Show("pls Enter BuyerName ", "BuyerName", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    combobuyer.Focus();
                    return;
                }
                if (txtstylename.Text == "")
                {
                    MessageBox.Show("pls Enter StyleName ", "StyleName", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtstylename.Focus();
                    return;
                }
               
                if (txtstylename.Text != "")
                {
                   // Models.Validate va = new Models.Validate();
                    

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select ASPTBLSTYMASID    from  ASPTBLSTYMAS    WHERE compcode='"+combocompcode.SelectedValue+"' and  buyercode='"+combobuyer.SelectedValue+"' and  styleno='"+txtstyleno.Text+"' and  stylename='"+txtstylename.Text+"' and  description='"+ txtdescription.Text + "' and  stylecategory='"+ combostylecategory.SelectedValue + "' and  packtype='"+ combostylepacktype.SelectedValue + "'  and  currency='"+combocurrecny.SelectedValue+ "'  and season='" + combosession.SelectedValue + "'  and  active='"+chk+"' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLSTYMAS");
                        DataTable dt = ds.Tables["ASPTBLSTYMAS"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtstylename.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtstyleid.Text) == 0 || Convert.ToInt32("0" + txtstyleid.Text) == 0)
                        {
                            string ins = "insert into asptblstymas(compcode,  buyercode ,  styleno,  stylename,  description,  stylecategory,  packtype,  currency,season, active,  username,  createdby,  createdon,  modifiedby,  modified,  ipaddress)  VALUES('"+combocompcode.SelectedValue+ "',  '" + combobuyer.SelectedValue + "' ,  '" + txtstyleno.Text + "',  '" + txtstylename.Text + "',  '" + txtdescription.Text + "',  '" + combostylecategory.SelectedValue + "',  '" + combostylepacktype.SelectedValue + "', '" + combocurrecny.SelectedValue + "', '" + combosession.SelectedValue + "', '" + chk + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + System.DateTime.Now.ToString("yyyy-MM-dd") + "','" + Class.Users.HUserName + "','" + System.DateTime.Now.ToString("yyyy-MM-dd") + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtstylename.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  ASPTBLSTYMAS  set   compcode='" + combocompcode.SelectedValue + "' ,  buyercode='" + combobuyer.SelectedValue + "' ,  styleno='" + txtstyleno.Text + "' ,  stylename='" + txtstylename.Text + "' ,  description='" + txtdescription.Text + "' ,  stylecategory='" + combostylecategory.SelectedValue + "' , season='" + combosession.SelectedValue + "' , packtype='" + combostylepacktype.SelectedValue + "' , currency='" + combocurrecny.SelectedValue + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where ASPTBLSTYMASID='" + txtstyleid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtstylename.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("" + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void StyleMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        public void News()
        {
            GridLoad();
            empty();
        }
        private void empty()
        {
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            txtstyleid.Text = "";
            txtstylename.Text = "";
            combocompcode.Text = "";
            combobuyer.Text = "";
            txtstyleno.Text = "";
            txtstylename.Text = "";
            txtdescription.Text = "";
            combostylecategory.Text = "";
            combostylepacktype.Text = "";
        
            combocurrecny.Text = "";
            combosession.Text = "";
            
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            Utility.Load_Combo(combocompcode, "select gtcompmastid,compcode from gtcompmast  where active='T' order by 2", "gtcompmastid", "compcode");
            Utility.Load_Combo(combosession, "select 0 asptblseasonmasid,''season from dual union all select asptblseasonmasid,season from asptblseasonmas  where active='T' order by 2", "asptblseasonmasid", "season");
            Utility.Load_Combo(combostylecategory, "select 0 asptblstycatmasid,''stylecategory from dual union all select asptblstycatmasid,stylecategory from asptblstycatmas  where active='T' order by 2", "asptblstycatmasid", "stylecategory");
            Utility.Load_Combo(combobuyer, "select 0 asptblbuymasid,''buyercode from dual union all select asptblbuymasid,buyercode from asptblbuymas  where active='T' order by 2", "asptblbuymasid", "buyercode");
            Utility.Load_Combo(combocurrecny, "select 0 asptblcurmasid,''currency from dual union all  select asptblcurmasid,currency from asptblcurmas  where active='T' order by 2", "asptblcurmasid", "currency");
            Utility.Load_Combo(combostylepacktype, "select 0 asptblordpacmasid,''packtype from dual union all select asptblordpacmasid,packtype from asptblordpacmas  where active='T' order by 2", "asptblordpacmasid", "packtype");

            txtstylename.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select a.asptblstymasid, b.compcode,  c.buyercode ,  a.styleno,  a.stylename,  a.description,  d.stylecategory,  A.packtype,  a.styleyear, g.season,  f.currency,  a.orderqty,   a.active from asptblstymas a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblbuymas c on c.asptblbuymasid=a.buyercode  join asptblstycatmas d on d.asptblstycatmasid=a.stylecategory  join asptblcurmas f on asptblcurmasid=a.currency join asptblseasonmas g on g.asptblseasonmasid=a.season  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTYMAS");
                DataTable dt = ds.Tables["ASPTBLSTYMAS"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLSTYMASID"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["STYLENAME"].ToString());
                        list.SubItems.Add(myRow["styleno"].ToString());
                        list.SubItems.Add(myRow["BuyerCode"].ToString());
                        list.SubItems.Add(myRow["stylecategory"].ToString());
                        list.SubItems.Add(myRow["description"].ToString());
                        list.SubItems.Add(myRow["Season"].ToString());
                        list.SubItems.Add(myRow["orderqty"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());

                        listView1.Items.Add(list);
                        listfilter.Items.Add((ListViewItem)list.Clone());
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

                    txtstyleid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select a.ASPTBLSTYMASID,b.compcode,  c.buyercode ,  a.styleno,  a.stylename, " +
                        " a.description,  d.stylecategory,  A.packtype,  a.styleyear,   f.currency,g.season,  a.orderqty,   a.active from asptblstymas a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblbuymas c on c.asptblbuymasid=a.buyercode  join asptblstycatmas d on d.asptblstycatmasid=a.stylecategory   join asptblcurmas f on asptblcurmasid=a.currency join asptblseasonmas g on g.asptblseasonmasid=a.season  where a.ASPTBLSTYMASID=" + txtstyleid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTYMAS");
                    DataTable dt = ds.Tables["ASPTBLSTYMAS"];

                    if (dt.Rows.Count > 0)
                    {
                        txtstyleid.Text = Convert.ToString(dt.Rows[0]["ASPTBLSTYMASID"].ToString());
                        combocompcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combobuyer.Text = Convert.ToString(dt.Rows[0]["buyercode"].ToString());
                        txtstyleno.Text = Convert.ToString(dt.Rows[0]["styleno"].ToString());
                        txtstylename.Text = Convert.ToString(dt.Rows[0]["stylename"].ToString());
                        txtdescription.Text = Convert.ToString(dt.Rows[0]["description"].ToString());
                        combostylecategory.Text = Convert.ToString(dt.Rows[0]["stylecategory"].ToString());
                        combostylepacktype.SelectedValue = Convert.ToString(dt.Rows[0]["packtype"].ToString());
                     
                        combocurrecny.Text = Convert.ToString(dt.Rows[0]["currency"].ToString());
                        combosession.Text = Convert.ToString(dt.Rows[0]["season"].ToString());
                  
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                    tabControl1.SelectTab(tabPage1);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Searchs()
        {
            try
            {
                if (txtsearch.Text.ToUpper() != "")
                {
                    listView1.Items.Clear(); int iGLCount = 1;
                    string sel1 = "  SELECT  a.ASPTBLSTYMASID,a.STYLENAME ,a.active from ASPTBLSTYMAS a  where a.STYLENAME LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTYMAS");
                    DataTable dt = ds.Tables["ASPTBLSTYMAS"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["ASPTBLSTYMASID"].ToString());
                            list.SubItems.Add(myRow["STYLENAME"].ToString());
                            list.SubItems.Add(myRow["active"].ToString());
                            listView1.Items.Add(list);
                            iGLCount++;
                        }
                        lbltotal.Text = "Total Count    :" + listView1.Items.Count;
                    }
                    else
                    {
                        listView1.Items.Clear();
                    }
                }
                else
                {

                    listView1.Items.Clear();
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Deletes()
        {
            if (txtstyleid.Text != "")
            {
                string sel1 = "select a.ASPTBLSTYMASID from asptblstymas a join asptblstygrpdet b on a.asptblstymasid=b.stylename where a.ASPTBLSTYMASID='" + txtstyleid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLSTYMAS");
                DataTable dt = ds.Tables["ASPTBLSTYMAS"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtstylename.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from ASPTBLSTYMAS where ASPTBLSTYMASID='" + Convert.ToInt64("0" + txtstyleid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtstylename.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        public void Prints()
        {
            throw new NotImplementedException();
        }

       


        public void ReadOnlys()
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

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[4].ToString().Contains(txtsearch.Text) || item.SubItems[5].ToString().Contains(txtsearch.Text) || item.SubItems[6].ToString().Contains(txtsearch.Text) || item.SubItems[8].ToString().Contains(txtsearch.Text) || item.SubItems[9].ToString().Contains(txtsearch.Text))
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
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {
                    try
                    {
                        listView1.Items.Clear(); item0 = 0;
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
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);
                          
                            item0++;
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("---" + ex.ToString());
            };
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void txtbuyname_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}

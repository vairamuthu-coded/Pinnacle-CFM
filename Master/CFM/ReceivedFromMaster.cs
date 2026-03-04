using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master.CFM
{
    public partial class ReceivedFromMaster : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        public ReceivedFromMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
           
        }

        private static ReceivedFromMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        bool tabcheck = false;
        public static ReceivedFromMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ReceivedFromMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
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


        public void compid(string s)
    {
        string sel = "select max(a.asptblpartymasid1) as id  from  asptblpartymas a  where a.partyname='" + s.ToUpper() + "' ;";
        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
        DataTable dt = ds.Tables["asptblpartymas"];
        Int64 partycount = 1;
        if (dt.Rows.Count >= 1)
        {
            txtpartyid1.Text = "";
            Int64 cc = Convert.ToInt64("0" + dt.Rows[0]["id"].ToString()) + partycount;
            txtpartyid1.Text = cc.ToString();
        }
        else
        {
            txtpartyid1.Text = "1";
        }
    }
  
 
        private void ReceivedFrom_Load(object sender, EventArgs e)
        {
           
        }
        
    public void Saves()
    {
        try
        {
              
                Models.Validate va = new Models.Validate();
           
            if (va.IsStringSpacehypenpracket(txtpartyname.Text) == false)
            {
                MessageBox.Show("'SupplierName  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtpartyname.Focus();
                return;
            }
          
            if (txtaddress.Text == "")
            {
                MessageBox.Show("'Address  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtaddress.Focus(); return;
            }
            else
            {
                    if (txtpartyid.Text == "")
                    {
                        compid(txtpartyname.Text.Substring(1, 3));
                    }
                    if (va.IsStringSpacehypenpracket(txtpartyname.Text) == true)
                    {
                        string chk = "", chkcustomer = "";
                        txtPartycode.Text = txtpartyname.Text.Substring(0, 3);
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        if (checkcustomer.Checked == true) { chkcustomer = "CUSTOMER"; } else { chkcustomer = ""; checkcustomer.Checked = false; }
                        string sel = "select asptblpartymasid    from  asptblpartymas   WHERE partyname='" + txtpartyname.Text + "' and city='" + combocity.Text + "' and state='" + combostate.Text + "' and country='" + combocountry.Text + "' and address='" + txtaddress.Text + "' and pincode='" + txtpincode.Text + "' and panno='" + txtpanno.Text + "' and tinno='" + txttinno.Text + "' and gstno='" + txtgstno.Text + "' and gstdate='" + txtgstdate.Text + "' and phoneno='" + txtphoneno.Text + "' and email='" + txtemail.Text + "' and website='" + txtwebsite.Text + "'  and contactname='" + txtcontact.Text + "' and customer='" + chkcustomer + "' and active='" + chk + "' and modifiedby='" + Class.Users.HUserName + "' and ipaddress='" + Class.Users.IPADDRESS + "' and asptblpartymasid='" + txtpartyid.Text + "';";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                        DataTable dt = ds.Tables["asptblpartymas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtpartyid.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt64("0" + txtpartyid.Text) == 0 || Convert.ToInt64("0" + txtpartyid.Text) == 0)
                        {
                            string ins = "insert into asptblpartymas(asptblpartymasid1,partycode,partyname,city,state,country,address,pincode,panno,tinno,gstno,gstdate,phoneno,email,website,contactname,active,customer,createdby,modifiedby,ipaddress)  VALUES('" + txtpartyid1.Text + "','" + txtPartycode.Text.ToUpper() + "','" + txtpartyname.Text.ToUpper() + "', '" + combocity.Text + "', '" + combostate.Text + "','" + combocountry.Text + "' ,'" + txtaddress.Text.ToUpper() + "','" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "','" + txtpanno.Text + "','" + txttinno.Text.ToUpper() + "','" + txtgstno.Text.ToUpper() + "','" + txtgstdate.Text + "','" + txtphoneno.Text + "','" + txtemail.Text.ToUpper() + "','" + txtwebsite.Text.ToUpper() + "','" + txtphoneno.Text.ToUpper() + "','" + chk + "', '" + chkcustomer + "' ,'" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "');";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtpartyid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty(); tabControl1.SelectTab(tabPage3);
                        }
                        else
                        {
                            string up = "update  asptblpartymas  set  asptblpartymasid1='" + txtpartyid.Text + "' , partycode='" + txtPartycode.Text.ToUpper() + "',partyname='" + txtpartyname.Text.ToUpper() + "', city='" + combocity.Text + "', state='" + combostate.Text + "',country='" + combocountry.Text + "' ,address='" + txtaddress.Text.ToUpper() + "',pincode='" + Convert.ToInt64("0" + txtpincode.Text).ToString() + "',panno='" + txtpanno.Text.ToUpper() + "',tinno='" + txttinno.Text.ToUpper() + "',gstno='" + txtgstno.Text.ToUpper() + "',gstdate='" + txtgstdate.Text + "',phoneno='" + txtphoneno.Text + "',email='" + txtemail.Text.ToUpper() + "',website='" + txtwebsite.Text.ToUpper() + "',contactname='" + txtphoneno.Text.ToUpper() + "',customer='" + chkcustomer + "',active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where  asptblpartymasid='" + txtpartyid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtpartyid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                            tabControl1.SelectTab(tabPage3);
                        }

                    }
                    //else
                    //{
                    //    MessageBox.Show("Invalid Data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    return;

                    //}
                }

        }
        catch (Exception ex)
        {

            MessageBox.Show("Error " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }
    }

    private void ReceivedFrom_FormClosed(object sender, FormClosedEventArgs e)
    {
        _instance = null;
    }


        public void News()
    {
            empty();
    }
    private void empty()
        {
            txtpartyid.Text = "";
            txtpartyid1.Text = ""; txtpartyname.Select(); pictureBox1.Width = 150; pictureBox1.Height = 150;
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            Class.Users.UserTime = 0;

            combocity.Text = ""; combocity.Text = "";
            combostate.Text = ""; combostate.Text = "";
            combocountry.Text = ""; combocountry.Text = ""; txtsearch.Text = "";
            checkactive.Checked = true;
            checkcustomer.Checked = true;

            txtPartycode.Text = "";
            txtpartyname.Text = "";
            txtaddress.Text = "";
            txtpincode.Text = "";
            txtpanno.Text = "";
            txttinno.Text = "";
            txtgstno.Text = "";
            txtgstdate.Text = "";
            txtphoneno.Text = "";
            txtemail.Text = "";
            txtwebsite.Text = "";
            txtphoneno.Text = "";
            this.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            GridLoad();
            if (tabcheck == false)
            {
                tabControl1.SelectTab(tabPage3);
                tabcheck = true;
            }
            else
            {
                tabControl1.SelectTab(tabPage1);
                tabcheck = true;
            }
            txtsearch.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.asptblpartymasid,a.PartyName,a.city,a.state,a.country,a.address,  a.pincode,a.panno,a.tinno,  a.gstno,'' as gstdate,  a.phoneno, a.email,  a.website,  a.active from asptblpartymas a  order by A.asptblpartymasid DESC;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                DataTable dt = ds.Tables["asptblpartymas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblpartymasid"].ToString());
                        list.SubItems.Add(myRow["PartyName"].ToString());
                        list.SubItems.Add(myRow["city"].ToString());
                        list.SubItems.Add(myRow["state"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        list.SubItems.Add(myRow["pincode"].ToString());
                        list.SubItems.Add(myRow["panno"].ToString());
                        list.SubItems.Add(myRow["tinno"].ToString());
                        list.SubItems.Add(myRow["gstno"].ToString());
                        list.SubItems.Add(myRow["phoneno"].ToString());
                        listView1.Items.Add(list);
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
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
                if (listView1.SelectedItems.Count == 0)
                    return;

                txtpartyid.Text = listView1.SelectedItems[0].SubItems[2].Text;

                string sel1 = "select a.asptblpartymasid,a.PartyCode,a.PartyName,a.city,a.state,a.country,a.address,  a.pincode,a.panno,a.tinno,  a.gstno,'' as gstdate,  a.phoneno, a.email,  a.website,  a.active,a.customer from asptblpartymas a  where a.asptblpartymasid='" + txtpartyid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                DataTable dt = ds.Tables["asptblpartymas"];

                if (dt.Rows.Count == 0)
                    return;

                DataRow row = dt.Rows[0];

                txtpartyid.Text = row["asptblpartymasid"].ToString();
                txtpartyid1.Text = row["asptblpartymasid"].ToString();

                txtPartycode.Text = row["PartyCode"].ToString();
                txtpartyname.Text = row["PartyName"].ToString();
                combocity.Text = row["city"].ToString();
                combostate.Text = row["state"].ToString();
                combocountry.Text = row["country"].ToString();
                txtaddress.Text = row["address"].ToString();
                txtpincode.Text = row["pincode"].ToString();
                txtpanno.Text = row["panno"].ToString();
                txttinno.Text = row["tinno"].ToString();
                txtgstno.Text = row["gstno"].ToString();
                txtgstdate.Text = row["gstdate"].ToString();
                txtphoneno.Text = row["phoneno"].ToString();
                txtemail.Text = row["email"].ToString();
                txtwebsite.Text = row["website"].ToString();

                checkactive.Checked = row["active"].ToString() == "T";
                checkcustomer.Checked = row["customer"].ToString() == "CUSTOMER";

                tabControl1.SelectTab(tabPage1);
                txtpartyname.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //try
            //{
            //    if (listView1.Items.Count > 0)
            //    {

            //        txtpartyid.Text = listView1.SelectedItems[0].SubItems[2].Text;
            //        //txtcompid1.Text = listView1.SelectedItems[0].SubItems[2].Text;

            //        string sel1 = "select a.asptblpartymasid,a.PartyCode,a.PartyName,a.city,a.state,a.country,a.address,  a.pincode,a.panno,a.tinno,  a.gstno,'' as gstdate,  a.phoneno, a.email,  a.website,  a.active,a.customer from asptblpartymas a  where a.asptblpartymasid='" + txtpartyid.Text + "'";
            //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
            //        DataTable dt = ds.Tables["asptblpartymas"];

            //        if (dt.Rows.Count > 0)
            //        {
            //            // txtcompid.Text = "";
            //            txtpartyid.Text = Convert.ToString(dt.Rows[0]["asptblpartymasid"].ToString());
            //            txtpartyid1.Text = Convert.ToString(dt.Rows[0]["asptblpartymasid"].ToString());
            //            txtPartycode.Text = Convert.ToString(dt.Rows[0]["PartyCode"].ToString());
            //            txtpartyname.Text = Convert.ToString(dt.Rows[0]["PartyName"].ToString());
            //            combocity.Text = Convert.ToString(dt.Rows[0]["city"].ToString());
            //            combostate.Text = Convert.ToString(dt.Rows[0]["state"].ToString());
            //            combocountry.Text = Convert.ToString(dt.Rows[0]["country"].ToString());
            //            txtaddress.Text = Convert.ToString(dt.Rows[0]["address"].ToString());
            //            txtpincode.Text = Convert.ToString(dt.Rows[0]["pincode"].ToString());
            //            txtpanno.Text = Convert.ToString(dt.Rows[0]["panno"].ToString());
            //            txttinno.Text = Convert.ToString(dt.Rows[0]["tinno"].ToString());
            //            txtgstno.Text = Convert.ToString(dt.Rows[0]["gstno"].ToString());
            //            txtgstdate.Text = Convert.ToString(dt.Rows[0]["gstdate"].ToString());
            //            txtphoneno.Text = Convert.ToString(dt.Rows[0]["phoneno"].ToString());
            //            txtemail.Text = Convert.ToString(dt.Rows[0]["email"].ToString());
            //            txtwebsite.Text = Convert.ToString(dt.Rows[0]["website"].ToString());

            //            if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
            //            if (dt.Rows[0]["customer"].ToString() == "CUSTOMER") { checkcustomer.Checked = true; } else { checkcustomer.Checked = false; }


            //            //  Combocity_SelectedIndexChanged(sender, e);
            //            // Combostate_SelectedIndexChanged(sender, e);
            //            tabControl1.SelectTab(tabPage1);txtpartyname.Select();
            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                listView1.BeginUpdate();
                listView1.Items.Clear();

                string searchText = txtsearch.Text.Trim().ToUpper();
                int rowIndex = 0;

                if (searchText.Length > 1)
                {
                    foreach (ListViewItem src in listfilter.Items)
                    {
                        if (src.SubItems[3].Text.ToUpper().Contains(searchText))
                        {
                            ListViewItem item = (ListViewItem)src.Clone();

                            item.BackColor = rowIndex % 2 == 0
                                ? Class.Users.Color1
                                : Class.Users.Color2;

                            listView1.Items.Add(item);
                            rowIndex++;
                        }
                    }
                }
                else
                {
                    GridLoad();
                }

                lbltotal.Text = $"Total {listView1.Items.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                listView1.EndUpdate();
            }

            //try
            //{
            //    int item0 = 0;
            //    if (txtsearch.Text.Length > 1)
            //    {
            //        listView1.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (item.SubItems[3].ToString().Contains(txtsearch.Text.ToUpper()))
            //            {
            //                list.Text = item.SubItems[0].Text;
            //                list.SubItems.Add(item.SubItems[1].Text);
            //                list.SubItems.Add(item.SubItems[2].Text);
            //                list.SubItems.Add(item.SubItems[3].Text);                           
            //                list.SubItems.Add(item.SubItems[4].Text);
            //                list.SubItems.Add(item.SubItems[5].Text);
            //                list.SubItems.Add(item.SubItems[6].Text);
            //                list.SubItems.Add(item.SubItems[7].Text);
            //                list.SubItems.Add(item.SubItems[8].Text);
            //                list.SubItems.Add(item.SubItems[9].Text);
            //                list.SubItems.Add(item.SubItems[10].Text);
            //                list.SubItems.Add(item.SubItems[11].Text);
            //                list.SubItems.Add(item.SubItems[12].Text);
            //                list.SubItems.Add(item.SubItems[13].Text);

            //                list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

            //                listView1.Items.Add(list);
            //            }
            //            item0++;
            //        }
            //    }
            //    else
            //    {
            //        listView1.Items.Clear();
            //        GridLoad();
            //    }
            //    lbltotal.Refresh();
            //    lbltotal.Text = "Total " + listView1.Items.Count.ToString();

            //    //if (txtsearch.Text.ToUpper() != "")
            //    //{
            //    //    listView1.Items.Clear(); int iGLCount = 1;
            //    //    string sel1 = "select a.asptblpartymasid,a.partyname, a.city,a.state, a.active,a.pincode,a.panno,a.tinno,  a.gstno,  a.phoneno," +
            //    //        "   from  asptblpartymas a    where a.PartyCode LIKE'%" + txtsearch.Text.ToUpper() + "%'  || a.city LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.state LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.active LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.ptransaction LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.pincode LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.panno LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.gstno LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.phoneno LIKE'%" + txtsearch.Text.ToUpper() + "%' || a.contactname LIKE'%" + txtsearch.Text.ToUpper() + "%';";
            //    //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
            //    //    DataTable dt = ds.Tables["asptblpartymas"];
            //    //    if (dt.Rows.Count > 0)
            //    //    {

            //    //        foreach (DataRow myRow in dt.Rows)
            //    //        {
            //    //            ListViewItem list = new ListViewItem();
            //    //            list.SubItems.Add(iGLCount.ToString());
            //    //            list.SubItems.Add(myRow["asptblpartymasid"].ToString());
            //    //            list.SubItems.Add(myRow["PartyName"].ToString());
            //    //            list.SubItems.Add(myRow["city"].ToString());
            //    //            list.SubItems.Add(myRow["state"].ToString());
            //    //            list.SubItems.Add(myRow["active"].ToString());
            //    //            // list.SubItems.Add(myRow["customer"].ToString());
            //    //            list.SubItems.Add(myRow["pincode"].ToString());
            //    //            list.SubItems.Add(myRow["panno"].ToString());
            //    //            list.SubItems.Add(myRow["tinno"].ToString());
            //    //            list.SubItems.Add(myRow["gstno"].ToString());
            //    //            list.SubItems.Add(myRow["phoneno"].ToString());
            //    //            listView1.Items.Add(list);
            //    //            listfilter.Items.Add((ListViewItem)list.Clone());
            //    //            list.BackColor = iGLCount % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
            //    //            iGLCount++;
            //    //        }
            //    //        lbltotal.Text = "Total Count: " + listView1.Items.Count;
            //    //    }
            //    //    else
            //    //    {
            //    //        listView1.Items.Clear();
            //    //    }
            //    //}
            //    //else
            //    //{

            //    //    listView1.Items.Clear();
            //    //    GridLoad();
            //    //}
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }


        public void Deletes()
        {
            if (txtpartyid.Text != "")
            {
                var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmation == DialogResult.Yes)
                {
                    string sel1 = "select a.asptblpartymasid from asptblpartymas a join asptblrawmaterial b on a.asptblpartymasid=b.receivedfrom where a.asptblpartymasid='" + txtpartyid.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblpartymas");
                    DataTable dt = ds.Tables["asptblpartymas"];
                    if (dt.Rows.Count > 0)
                    {
                      
                        MessageBox.Show("Child Record Found.Can Not Delete." + txtpartyname.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                    }
                    else
                    {

                        string del = "delete from asptblpartymas where asptblpartymasid='" + Convert.ToInt64("0" + txtpartyid.Text) + "'";
                        Utility.ExecuteNonQuery(del);
                        MessageBox.Show("Record Deleted Successfully " + txtpartyname.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty(); tabControl1.SelectTab(tabPage3);
                    }
                }
            }
        }



        private void txtPartycode_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void txtpartyname_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combodivision.Select();
            
            }
        }

        private void combodivision_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combocity.Select();
         
            }
        }

        private void combocity_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combostate.Select();
               
            }
        }

        private void combostate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                combocountry.Select();
               
            }
        }
        private void combocountry_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtaddress.Select();
            
            }
        }
        private void txtaddress_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtpincode.Select();
               
            }
        }

        private void txtpincode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtpanno.Select();
               
            }
        }

        private void txtpanno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txttinno.Select();
                
            }
        }

        private void txttinno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtgstno.Select();
               
            }
        }

        private void txtgstno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtgstdate.Select();
              
            }
        }

        private void txtgstdate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtphoneno.Select();
               
            }
        }

        private void txtphoneno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtemail.Select();
               
            }
        }
        private void txtwebsite_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtphoneno.Select();
               
            }
        }
        private void txtemail_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtwebsite.Select();
               
            }
        }

        private void txtcontact_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                  buttsave.Select();
               
            }
        }

       

       
       

        private void buttsave_Click(object sender, EventArgs e)
        {
            Saves();
          
        }

        private void buttsearch_Click(object sender, EventArgs e)
        {
            Searchs_Click(sender, e);
        }

        private void Searchs_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPage3);
        }

        private void buttdelete_Click(object sender, EventArgs e)
        {
            Deletes();
        }

     

        private void combodivision_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtphoneno_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])//your specific tabname
            {
                txtpartyname.Select();

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
            {
                txtsearch.Select();
            }
        }

        private void txtpartyname_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || char.IsWhiteSpace(e.KeyChar) || e.KeyChar == (char)Keys.Back || e.KeyChar == '-' || e.KeyChar == '(' || e.KeyChar == ')');
        }

        private void buttcancel_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(tabPage3);
        }
        byte[] bytes;
        public static byte[] ImageToByteArray(PictureBox imageIn)
        {
            var ms = new MemoryStream();
            imageIn.Image.Save(ms, imageIn.Image.RawFormat);
            return ms.ToArray();
        }
        public Image ByteArrayToImage1(byte[] byteArrayIn)
        {
            System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
            Image img = (Image)converter.ConvertFrom(byteArrayIn);

            return img;
        }
        public static Image ByteArrayToImage(byte[] byteArrayIn)
        {
            var ms = new MemoryStream(byteArrayIn);
            Image returnImage = Image.FromStream(ms);
            return returnImage;
        }
        OpenFileDialog open = new OpenFileDialog();
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                bytes = null;
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {
                        p.Image = new Bitmap(open.FileName);
                        bytes = ImageToByteArray(p);
                        //  System.Text.Encoding enc = System.Text.Encoding.ASCII;


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Prints()
        {
            throw new NotImplementedException();
        }

        public void Searchs()
        {
            throw new NotImplementedException();
        }

        public void Searchs(int EditID)
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
    }
}

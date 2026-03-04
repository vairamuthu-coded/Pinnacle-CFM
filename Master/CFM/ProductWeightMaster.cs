using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master.CFM
{
    public partial class ProductWeightMaster : Form,ToolStripAccess
    {
        ListView listfilter = new ListView();
        public ProductWeightMaster()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToShortTimeString().ToString());
        }

        private static ProductWeightMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.Validate va = new Models.Validate();
       
        public static ProductWeightMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ProductWeightMaster();
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



        private void ProductWeightMaster_Load(object sender, EventArgs e)
        {
           
        }

        public void loadproduct()
        {
            string sel = "SELECT  a.asptblproductmasid,a.productname from asptblproductmas a where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproductmas");
            DataTable dt = ds.Tables["asptblproductmas"];
           
            comboproductname.DisplayMember = "productname";
            comboproductname.ValueMember = "asptblproductmasid";
            comboproductname.DataSource = dt;
           
        }
       
        public void Saves()
        {
            try
            {
                Models.Validate va = new Models.Validate();

                // -------- Validation --------
                if (comboproductname.SelectedValue == null)
                {
                    MessageBox.Show("Product Name is required", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    comboproductname.Focus();
                    return;
                }

                if (!va.IsDecimal(txtnetweight.Text))
                {
                    MessageBox.Show("Net Weight is invalid", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtnetweight.Focus();
                    return;
                }

                if (!va.IsDecimal(txtgrossweight.Text))
                {
                    MessageBox.Show("Gross Weight is invalid", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtgrossweight.Focus();
                    return;
                }

                if (!va.IsInteger(comboproductname.SelectedValue.ToString()))
                {
                    MessageBox.Show("Invalid Product", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string activeFlag = checkactive.Checked ? "T" : "F";
                int productWeightId = Convert.ToInt32("0" + txtproductweightid.Text);

                // -------- Duplicate Check --------

                string sel = "select a.asptblproductweightmasid    from  asptblproductweightmas a join asptblproductmas b on a.productname=b.asptblproductmasid   WHERE a.productname='" + comboproductname.SelectedValue + "' and a.netweight='" + txtnetweight.Text + "'and a.grossweight='" + txtgrossweight.Text + "' and a.active='" + activeFlag + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproductweightmas");         
                bool recordExists = ds.Tables[0].Rows.Count > 0;
                if (recordExists)
                {
                    MessageBox.Show("Record already exists",
                        "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // -------- INSERT --------
                else if (productWeightId == 0)
                {
                   
                    string ins = "insert into asptblproductweightmas(productname, netweight,grossweight,active,productname1,createdby,modifiedby,ipaddress)  VALUES('" + comboproductname.SelectedValue + "','" + txtnetweight.Text + "','" + txtgrossweight.Text + "','" + activeFlag + "','" + txtproductname1.Text + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                    Utility.ExecuteNonQuery(ins);
                    MessageBox.Show("Record Saved Successfully","Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                // -------- UPDATE --------
                else
                {


                    string up = "update  asptblproductweightmas  set productname='" + comboproductname.SelectedValue + "', netweight='" + txtnetweight.Text + "',grossweight='" + txtgrossweight.Text + "' , active='" + activeFlag + "' , productname1='" + txtproductname1.Text + "',modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblproductweightmasid='" + txtproductweightid.Text + "';";
                    Utility.ExecuteNonQuery(up);



                    MessageBox.Show("Record Updated Successfully",
                        "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                GridLoad();
                empty();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //try
            //{
            //    Models.Validate va = new Models.Validate();
            //    if (comboproductname.SelectedValue == null)
            //    {
            //        MessageBox.Show("'Product Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.comboproductname.Focus();
            //        return;
            //    }
            //    if (va.IsDecimal(txtnetweight.Text) == false)
            //    {
            //        MessageBox.Show("'Netweight Weight  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtnetweight.Focus();
            //        return;
            //    }
            //    if (va.IsDecimal(txtgrossweight.Text) == false)
            //    {

            //        MessageBox.Show("'Gross Weight  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtgrossweight.Focus(); return;
            //    }
            //    else
            //    {

            //        if (va.IsInteger(comboproductname.SelectedValue.ToString()) == true && va.IsDecimal(txtnetweight.Text) == true && va.IsDecimal(txtgrossweight.Text) == true)
            //        {

            //            string chk = "";
            //            if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
            //            string sel = "select a.asptblproductweightmasid    from  asptblproductweightmas a join asptblproductmas b on a.productname=b.asptblproductmasid   WHERE a.productname='" + comboproductname.SelectedValue + "' and a.netweight='" + txtnetweight.Text + "'and a.grossweight='" + txtgrossweight.Text + "' and a.active='" + chk + "';";
            //            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblproductweightmas");
            //            DataTable dt = ds.Tables["asptblproductweightmas"];
            //            if (dt.Rows.Count != 0)
            //            {
            //                MessageBox.Show("Child Record Found " + " Alert " + comboproductname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
            //            }
            //            else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtproductweightid.Text) == 0 || Convert.ToInt32("0" + txtproductweightid.Text) == 0)
            //            {
            //                string ins = "insert into asptblproductweightmas(productname, netweight,grossweight,active,productname1,createdby,modifiedby,ipaddress)  VALUES('" + comboproductname.SelectedValue + "','" + txtnetweight.Text + "','" + txtgrossweight.Text + "','" + chk + "','" + txtproductname1.Text + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
            //                Utility.ExecuteNonQuery(ins);
            //                MessageBox.Show("Record Saved Successfully " + comboproductname.SelectedValue, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                GridLoad(); empty();
            //            }
            //            else
            //            {
            //                string up = "update  asptblproductweightmas  set productname='" + comboproductname.SelectedValue + "', netweight='" + txtnetweight.Text + "',grossweight='" + txtgrossweight.Text + "' , active='" + chk + "' , productname1='"+txtproductname1.Text+"',modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblproductweightmasid='" + txtproductweightid.Text + "';";
            //                Utility.ExecuteNonQuery(up);
            //                MessageBox.Show("Record Updated Successfully " + comboproductname.SelectedValue, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                GridLoad();
            //                empty();
            //            }

            //        }
            //        else
            //        {
            //            MessageBox.Show("Invalid Data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //            return;

            //        }
            //    }

            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show("grossweight " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //}
        }

        private void ProductWeightMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }


        public void News()
        {

            GridLoad(); loadproduct(); empty();
        }
        private void empty()
        {
            txtproductweightid.Text = ""; GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;

            comboproductname.SelectedIndex = -1;  comboproductname.Select();
            txtnetweight.Text = ""; ; txtgrossweight.Text = "";
            txtgrossweight.Text = ""; txtsearch.Text = "";txtproductname1.Text = "";
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                listfilter.Items.Clear();
                string sel1 = " select a.asptblproductweightmasid,a.productname1, a.netweight, a.grossweight , a.active    from  asptblproductweightmas a  join asptblproductmas b on a.productname=b.asptblproductmasid    order by  a.asptblproductweightmasid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproductweightmas");
                DataTable dt = ds.Tables["asptblproductweightmas"];
                if (dt.Rows.Count == 0)
                    return;

                int rowIndex = 1;

                foreach (DataRow row in dt.Rows)
                {
                    ListViewItem item = new ListViewItem();
                    item.SubItems.Add(rowIndex.ToString());
                    item.SubItems.Add(row["asptblproductweightmasid"].ToString());
                    item.SubItems.Add(row["productname1"].ToString());
                    item.SubItems.Add(row["netweight"].ToString());
                    item.SubItems.Add(row["grossweight"].ToString());
                    item.SubItems.Add(row["active"].ToString());

                    item.BackColor = rowIndex % 2 == 0
                        ? Class.Users.Color1
                        : Class.Users.Color2;

                    listView1.Items.Add(item);
                    listfilter.Items.Add((ListViewItem)item.Clone());

                    rowIndex++;
                }

                lbltotal.Text = $"Total Count: {listView1.Items.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.SelectedItems.Count == 0)
                    return;

                txtproductweightid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                string sel1 = "select a.asptblproductweightmasid,b.productname, a.netweight, a.grossweight , a.active,a.productname1    from  asptblproductweightmas a  join asptblproductmas b on a.productname=b.asptblproductmasid where a.asptblproductweightmasid=" + txtproductweightid.Text;
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproductweightmas");
                DataTable dt = ds.Tables["asptblproductweightmas"];

                if (dt.Rows.Count == 0)
                    return;

                DataRow row = dt.Rows[0];

                txtproductweightid.Text = row["asptblproductweightmasid"].ToString();
                comboproductname.Text = row["productname"].ToString();
                txtnetweight.Text = row["netweight"].ToString();
                txtgrossweight.Text = row["grossweight"].ToString();

                checkactive.Checked = row["active"].ToString() == "T";

                txtproductname1.Text = $"{comboproductname.Text}-{txtnetweight.Text}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                listView1.BeginUpdate();
                listView1.Items.Clear();

                if (txtsearch.Text.Length > 1)
                {
                    string searchText = txtsearch.Text.ToUpperInvariant();
                    int rowIndex = 0;

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        // Always use .Text (not .ToString())
                        if (item.SubItems[3].Text.ToUpperInvariant().Contains(searchText) ||
                            item.SubItems[4].Text.ToUpperInvariant().Contains(searchText))
                        {
                            ListViewItem list = new ListViewItem(item.SubItems[0].Text);

                            for (int i = 1; i < item.SubItems.Count; i++)
                                list.SubItems.Add(item.SubItems[i].Text);

                            list.BackColor = (rowIndex % 2 == 0)
                                ? Class.Users.Color1
                                : Class.Users.Color2;

                            listView1.Items.Add(list);
                            rowIndex++;
                        }
                    }
                }
                else
                {
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                listView1.EndUpdate();
            }

            lbltotal.Text = $"Total {listView1.Items.Count}";

            //try
            //{
            //    int item0 = 0;
            //    if (txtsearch.Text.Length > 1)
            //    {
            //        listView1.Items.Clear();
            //        foreach (ListViewItem item in listfilter.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (item.SubItems[3].ToString().Contains(txtsearch.Text.ToUpper()) || item.SubItems[4].ToString().Contains(txtsearch.Text.ToUpper()))
            //            {
            //                list.Text = item.SubItems[0].Text;
            //                list.SubItems.Add(item.SubItems[1].Text);
            //                list.SubItems.Add(item.SubItems[2].Text);
            //                list.SubItems.Add(item.SubItems[3].Text);
            //                list.SubItems.Add(item.SubItems[4].Text);
            //                list.SubItems.Add(item.SubItems[5].Text);
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
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            //lbltotal.Refresh();
            //lbltotal.Text = "Total "+listView1.Items.Count.ToString();
        }


        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            loadproduct(); GridLoad();
        }

        private void Txtgrossweight_TextChanged(object sender, EventArgs e)
        {
          
            if (va.IsIntegerdot(txtgrossweight.Text) == false)
            {
               // txtgrossweight.Text = "";
            }
            if (va.IsIntegerdot(txtgrossweight.Text) == true)
            {
                txtproductname1.Text = "";
                txtproductname1.Text = comboproductname.Text + " " + txtnetweight.Text + " " + "KG"; ; 
            }
        }
        private void Txtnetweight_TextChanged(object sender, EventArgs e)
        {
            txtproductname1.Text = ""; 
         //   txtproductname1.Text = comboproductname.Text + " " + txtnetweight.Text;
            txtproductname1.Text = comboproductname.Text + " " + txtnetweight.Text + " " + "KG"; ;
        }

        public void Deletes()
        {
            if (txtproductweightid.Text != "")
            {
                string sel1 = "select a.asptblproductweightmasid from asptblproductweightmas a join asptbldeliverydet b on a.asptblproductweightmasid=b.productname where a.asptblproductweightmasid='" + txtproductweightid.Text+"';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblproductweightmas");
                DataTable dt = ds.Tables["asptblproductweightmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + comboproductname.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblproductweightmas where asptblproductweightmasid='" + Convert.ToInt64("0" + txtproductweightid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + comboproductname.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void comboproductname_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Enter)
            {
                txtnetweight.Focus();
               
            }
        }

        private void txtnetweight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtgrossweight.Focus();
            }
        }

        private void txtnetweight_KeyPress(object sender, KeyPressEventArgs e)
        {
          
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
         //   e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtgrossweight_KeyPress(object sender, KeyPressEventArgs e)
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

        private void comboproductname_MouseClick(object sender, MouseEventArgs e)
        {
            this.comboproductname.BackColor = Color.Yellow;
        }

        private void txtgrossweight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                
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

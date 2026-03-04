using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Transactions.Lyla
{
    public partial class OrderCloseEntry : Form,ToolStripAccess
    {
        public OrderCloseEntry()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();

        }
        private static OrderCloseEntry _instance;
        public static OrderCloseEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new OrderCloseEntry();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        Models.CommonClass CC = new Models.CommonClass();
        ListView listfilter = new ListView();
        void empty()
        {
            combofinyear.Text = Class.Users.Finyear;          
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            Class.Users.dt = CC.select("select distinct b.gtcompmastid,b.compcode from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid where  a.active='T' and b.compcode='" + Class.Users.HCompcode + "' order by b.gtcompmastid desc", "ASPTBLPUR");
            if (Class.Users.dt != null)
            {
                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = Class.Users.dt;
            }
            Class.Users.dt = CC.select("select distinct a.asptblpurid,a.pono from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  where  a.active='T' and b.compcode='" + Class.Users.HCompcode + "' order by a.asptblpurid desc", "ASPTBLPUR");
            if (Class.Users.dt != null)
            {
                combopono.DisplayMember = "pono";
                combopono.ValueMember = "pono";
                combopono.DataSource = Class.Users.dt;
            }
           
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
   combopono.Focus();
            txtsearch.Text = "";
            txtasptblordcloid.Text = "";
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
        public void News()
        {
            GridLoad(); empty();
        }
        public void Saves()
        {
            try
            {
                if (combopono.Text == "")
                {
                    MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combopono.Focus();
                    return;
                }

                else
                {
                    if (combopono.Text != "")
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select a.asptblordclomasID    from  asptblordclomas a     WHERE a.finyear='" + combofinyear.Text + "' and a.pono='" + combopono.Text + "' and a.active='" + chk + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblordclomas");
                        DataTable dt = ds.Tables["asptblordclomas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + combopono.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtasptblordcloid.Text) == 0 || Convert.ToInt32("0" + txtasptblordcloid.Text) <= 0)
                        {
                            string ins = "insert into asptblordclomas(finyear,pono,active,compcode, createdby,modifiedby,ipaddress)  VALUES('" + combofinyear.Text + "','" + combopono.Text.ToUpper() + "','" + chk + "','" + combocompcode.SelectedValue + "','" + Class.Users.HUserName + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + combopono.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblordclomas  set finyear='" + combofinyear.Text + "' ,pono='" + combopono.Text + "',  active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblordclomasID='" + txtasptblordcloid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + combopono.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid Data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;

                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("pono " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = " select b.asptblpurid,a.pono , a.active,a.asptblordclomasid ,c.compcode   from  asptblordclomas a  join asptblpur b on a.pono=b.pono join gtcompmast c on c.gtcompmastid=a.compcode  where c.compcode='"+Class.Users.HCompcode+"'   order by b.asptblpurid desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordclomas");
                DataTable dt = ds.Tables["asptblordclomas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblordclomasID"].ToString());
                        list.SubItems.Add(myRow["pono"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        listView1.Items.Add(list);
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
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
                Class.Users.UserTime = 0;
                if (listView1.Items.Count > 0)
                {

                    txtasptblordcloid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select b.asptblpurid,a.pono , a.active,a.asptblordclomasid ,c.compcode   from  asptblordclomas a  join asptblpur b on a.pono=b.pono join gtcompmast c on c.gtcompmastid=a.compcode   where a.asptblordclomasID=" + txtasptblordcloid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblordclomas");
                    DataTable dt = ds.Tables["asptblordclomas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtasptblordcloid.Text = Convert.ToString(dt.Rows[0]["asptblordclomasID"].ToString());
                        combopono.Text = Convert.ToString(dt.Rows[0]["pono"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }


                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
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
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text))
                        {
                            list.SubItems.Add(item.SubItems[0].Text);
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
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
                            list.SubItems.Add(item.SubItems[0].Text);
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
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
            }
        }

        public void ChangePasswords()
        {
            
        }

        public void ChangeSkins()
        {
            
        }

        public void Deletes()
        {
            
        }

        public void DownLoads()
        {
            throw new NotImplementedException();
        }

     

        public void GlobalSearchs()
        {
        }

      
        public void Imports()
        {
        }

        public void Logins()
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
        public void Pdfs()
        {
            
        }

        public void Prints()
        {
        }

        public void ReadOnlys()
        {
        }

   

        public void Searchs()
        {
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void TreeButtons()
        {
            throw new NotImplementedException();
        }

        private void OrderCloseEntry_FormClosing(object sender, FormClosingEventArgs e)
        {
           
        }

        private void OrderCloseEntry_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }
    }
}

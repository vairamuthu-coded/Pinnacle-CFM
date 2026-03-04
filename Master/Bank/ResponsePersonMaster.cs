using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master.Bank
{
    public partial class ResponsePersonMaster : Form, ToolStripAccess
    {
        public ResponsePersonMaster()
        {
            InitializeComponent();

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());

        }
        ListView listfilter = new ListView();
        private static ResponsePersonMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();

        public static ResponsePersonMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ResponsePersonMaster();
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

        private void ResponsePersonMaster_Load(object sender, EventArgs e)
        {

        }
        public void Saves()
        {
            try
            {
                if (txtpersonid.Text == "" && txtperson.Text.Length >= 2)
                {
                    string sel = "select asptblresmasid    from  asptblresmas    WHERE resonseperson='" + txtperson.Text.ToUpper().Trim() + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblresmas");
                    DataTable dt = ds.Tables["asptblresmas"];
                    if (dt.Rows.Count > 0)
                    {
                        txtperson.Text = "";
                        MessageBox.Show("Child Recrod Found");
                        return;
                    }
                }
                if (txtperson.Text == "")
                {
                    MessageBox.Show("Person Name is empty " + " Alert " + txtperson.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtperson.Text != "")
                {
                    Models.Validate va = new Models.Validate();
                    if (va.IsStringNumbericslace(txtperson.Text) == true)
                    {

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select asptblresmasid    from  asptblresmas    WHERE resonseperson='" + txtperson.Text + "' and active='" + chk + "' ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblresmas");
                        DataTable dt = ds.Tables["asptblresmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtperson.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtpersonid.Text) == 0 || Convert.ToInt32("0" + txtpersonid.Text) == 0)
                        {
                            string ins = "insert into asptblresmas(resonseperson,active,createdby,modifiedby,ipaddress)  VALUES('" + txtperson.Text.ToUpper() + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtperson.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  asptblresmas  set   resonseperson='" + txtperson.Text.ToUpper() + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblresmasid='" + txtpersonid.Text + "';";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtperson.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    }
                    else
                    {
                        MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("resonseperson " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void ResponsePersonMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        public void News()
        {

            empty(); GridLoad();
        }
        private void empty()
        {
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            txtpersonid.Text = "";
            txtperson.Text = ""; txtperson.Select();
            checkactive.Checked = false;
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "   SELECT A.asptblresmasid, A.resonseperson , a.active  FROM  asptblresmas a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblresmas");
                DataTable dt = ds.Tables["asptblresmas"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblresmasid"].ToString());
                        list.SubItems.Add(myRow["resonseperson"].ToString());
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

                    txtpersonid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblresmasid, a.resonseperson , a.active    from  asptblresmas a    where a.asptblresmasid=" + txtpersonid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblresmas");
                    DataTable dt = ds.Tables["asptblresmas"];

                    if (dt.Rows.Count > 0)
                    {
                        txtpersonid.Text = Convert.ToString(dt.Rows[0]["asptblresmasid"].ToString());
                        txtperson.Text = Convert.ToString(dt.Rows[0]["resonseperson"].ToString());
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
                if (txtsearch.Text != "")
                {

                    int item0 = 0; listView1.Items.Clear();


                    foreach (ListViewItem item in listfilter.Items)
                    {

                        if (item.SubItems[3].ToString().Contains(txtsearch.Text))
                        {
                            ListViewItem list = new ListViewItem();

                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            if (item0 % 2 == 0)
                            {
                                list.BackColor = System.Drawing.Color.White;

                            }
                            else
                            {
                                list.BackColor = System.Drawing.Color.WhiteSmoke;
                            }
                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;



                }
                else
                {

                    GridLoad();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
        }

        public void Deletes()
        {
            if (txtpersonid.Text != "")
            {
                string sel1 = "select a.gtcountrymastid from gtcountrymast a join gtstatemast b on a.gtcountrymastid=b.country where a.gtcountrymastid='" + txtpersonid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtcountrymast");
                DataTable dt = ds.Tables["gtcountrymast"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtperson.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from gtcountrymast where gtcountrymastid='" + Convert.ToInt64("0" + txtpersonid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtperson.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void txtperson_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetter(e.KeyChar) || (char.IsWhiteSpace(e.KeyChar)) || e.KeyChar=='.' || e.KeyChar == (char)Keys.Back);

        }

        private void txtperson_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}

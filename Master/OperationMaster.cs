using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class OperationMaster : Form,ToolStripAccess
    {
        public OperationMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
        }

        private static OperationMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights(); ListView listfilter = new ListView();
        public static OperationMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new OperationMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


       

        private void OperationMaster_Load(object sender, EventArgs e)
        {
            News();GridLoad();
        }
        public void Saves()
        {
            try
            {
                if (txtoperation.Text == "")
                {
                    MessageBox.Show("operation Name is empty " + " Alert " + txtoperation.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtshortname.Text == "")
                {
                    MessageBox.Show("ShoftName is empty " + " Alert " + txtshortname.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtoperation.Text != "" && txtshortname.Text != "")
                {


                    string chk = "",sew = "",qc = "";
                   
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblopemasid    from  asptblopemas    WHERE operation='" + txtoperation.Text + "' and shortname='" + txtshortname.Text + "' and  groupname='" + combogroup.SelectedValue + "'  and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblopemas");
                    DataTable dt = ds.Tables["asptblopemas"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtoperation.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtoperationid.Text) == 0 || Convert.ToInt32("0" + txtoperationid.Text) == 0)
                    {
                        string ins = "insert into asptblopemas(operation,shortname,groupname,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtoperation.Text.ToUpper() + "','" + txtshortname.Text.ToUpper() + "', '" + combogroup.SelectedValue + "', '" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtoperation.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblopemas  set   operation='" + txtoperation.Text.ToUpper() + "' , shortname='" + txtshortname.Text.ToUpper() + "', groupname='" + combogroup.SelectedValue + "' , active='" + chk + "' , modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblopemasid='" + txtoperationid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtoperation.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty();
                    }

                }
                else
                {
                    MessageBox.Show("'Invalid Data'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("operation " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void operationMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

      

        public void News()
        {

            empty(); GridLoad();
            Utility.Load_Combo(combogroup, "select asptblgromasID, groupname    from  asptblgromas  where active='T'   order by 2", "asptblgromasID", "groupname");
        }
        private void empty()
        {
            txtoperationid.Text = "";
            txtoperation.Text = "";
            txtshortname.Text = "";
            combogroup.Text = "";
            checkactive.Checked = true;
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;

            txtoperation.Select();
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "   SELECT a.asptblopemasid, a.operation ,a.shortname,b.groupname ,a.active  FROM  asptblopemas a  join asptblgromas b on a.groupname=b.asptblgromasid   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblopemas");
                DataTable dt = ds.Tables["asptblopemas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblopemasid"].ToString());
                        list.SubItems.Add(myRow["operation"].ToString());
                        list.SubItems.Add(myRow["shortname"].ToString());                      
                        list.SubItems.Add(myRow["groupname"].ToString());
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

                    txtoperationid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblopemasid, a.operation ,a.shortname,b.groupname, a.active    from  asptblopemas a  join asptblgromas b on a.groupname=b.asptblgromasid   where a.asptblopemasid=" + txtoperationid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblopemas");
                    DataTable dt = ds.Tables["asptblopemas"];
                     
                    if (dt.Rows.Count > 0)
                    {
                        txtoperationid.Text = Convert.ToString(dt.Rows[0]["asptblopemasid"].ToString());
                        txtoperation.Text = Convert.ToString(dt.Rows[0]["operation"].ToString());
                        txtshortname.Text = Convert.ToString(dt.Rows[0]["shortname"].ToString());
                        combogroup.Text = Convert.ToString(dt.Rows[0]["groupname"].ToString());

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
                if (txtsearch.Text.ToUpper() != "")
                {

                    int item0 = 0; listView1.Items.Clear();
                    if (txtsearch.Text.Length > 1)
                    {

                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text))
                            {


                                list.Text = listfilter.Items[item0].SubItems[0].Text;
                                list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);               
                                list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
               
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

                        listView1.Items.Clear(); item0 = listfilter.Items.Count;
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            


                                list.Text = listfilter.Items[item0].SubItems[0].Text;
                                list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                                list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                                if (item0 % 2 == 0) { list.BackColor = Color.WhiteSmoke; } else { list.BackColor = Color.White; }

                                listView1.Items.Add(list);


                          
                            item0++;
                        }
                        lbltotal.Text = "Total Count: " + listView1.Items.Count;
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
        }

        public void Deletes()
        {
            if (txtoperationid.Text != "")
            {
                string sel1 = "select a.asptblopemasid from asptblopemas a join gtstatemast b on a.asptblopemasid=b.country where a.asptblopemasid='" + txtoperationid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblopemas");
                DataTable dt = ds.Tables["asptblopemas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtoperation.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblopemas where asptblopemasid='" + Convert.ToInt64("0" + txtoperationid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtoperation.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void refeshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();
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
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void OperationMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            _instance = null;
        }
    }
}

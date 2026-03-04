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
    public partial class ShiftTypeMaster : Form,ToolStripAccess
    {
        public ShiftTypeMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }

        private static ShiftTypeMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();

        public static ShiftTypeMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ShiftTypeMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }


        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                

            }
            else
            {
                MessageBox.Show("Invalid");
            }

        }
        private void ShiftTypeMaster_Load(object sender, EventArgs e)
        {
            News();
        }

        public void Saves()
        {
            try
            {
                if (txtshifttype.Text == "")
                {
                    MessageBox.Show("shifttype Name is empty " + " Alert " + txtshifttype.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (txtshifttype.Text != "")
                {


                    string chk = "";
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    string sel = "select asptblshitypeid    from  asptblshitype    WHERE shifttype='" + txtshifttype.Text + "' and shiftno='" + txtshiftno.Text + "' and active='" + chk + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblshitype");
                    DataTable dt = ds.Tables["asptblshitype"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtshifttype.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtshifttypeid.Text) == 0 || Convert.ToInt32("0" + txtshifttypeid.Text) == 0)
                    {
                        string ins = "insert into asptblshitype(shifttype,shiftno,active,compcode,username,createdby,modifiedby,ipaddress)  VALUES('" + txtshifttype.Text.ToUpper() + "','" + txtshiftno.Text + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtshifttype.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  asptblshitype  set   shifttype='" + txtshifttype.Text.ToUpper() + "' , shiftno='" + txtshiftno.Text + "', active='" + chk + "' ,compcode='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "', modifiedby='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblshitypeid='" + txtshifttypeid.Text + "';";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtshifttype.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

                MessageBox.Show("shifttype " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void shifttypeMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        
        public void News()
        {

            empty(); GridLoad();
        }
        private void empty()
        {
            txtshifttypeid.Text = "";
            txtshifttype.Text = "";
            txtshiftno.Text = "";
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;

            checkactive.Checked = false;
            txtshifttype.Select();
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                string sel1 = "SELECT asptblshitypeid, shifttype ,shiftno, active  FROM  asptblshitype    order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblshitype");
                DataTable dt = ds.Tables["asptblshitype"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblshitypeid"].ToString());
                        list.SubItems.Add(myRow["shifttype"].ToString());
                        list.SubItems.Add(myRow["shiftno"].ToString());
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

                    txtshifttypeid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.asptblshitypeid, a.shifttype ,a.shiftno, a.active    from  asptblshitype a    where a.asptblshitypeid=" + txtshifttypeid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblshitype");
                    DataTable dt = ds.Tables["asptblshitype"];

                    if (dt.Rows.Count > 0)
                    {
                        txtshifttypeid.Text = Convert.ToString(dt.Rows[0]["asptblshitypeid"].ToString());
                        txtshifttype.Text = Convert.ToString(dt.Rows[0]["shifttype"].ToString());
                        txtshiftno.Text = Convert.ToString(dt.Rows[0]["shiftno"].ToString());
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


                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length > 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            this.listfilter.Items.Add((ListViewItem)list.Clone());
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    listView1.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {

                    ListView ll = new ListView();
                    item0 = listfilter.Items.Count;
                   
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();



                        list.Text = item.SubItems[0].Text;
                        list.SubItems.Add(item.SubItems[1].Text);
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        if (item0 % 2 == 0) { list.BackColor = Color.White; } else { list.BackColor = Color.WhiteSmoke; }

                        listView1.Items.Add(list);



                        item0++;
                    }
                    listView1.Text = "Total Count: " + listView1.Items.Count;
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
        }

        public void Deletes()
        {
            if (txtshifttypeid.Text != "")
            {
                string sel1 = "select a.asptblshitypeid from asptblshitype a join gtstatemast b on a.asptblshitypeid=b.country where a.asptblshitypeid='" + txtshifttypeid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblshitype");
                DataTable dt = ds.Tables["asptblshitype"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtshifttype.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblshitype where asptblshitypeid='" + Convert.ToInt64("0" + txtshifttypeid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtshifttype.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void txtshiftno_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            //{
            //    e.Handled = false; //Do not reject the input
            //}
            //else
            //{
            //    e.Handled = true; //Reject the input
            //}
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void txtsearch_VisibleChanged(object sender, EventArgs e)
        {

        }
    }
}

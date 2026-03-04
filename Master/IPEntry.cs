using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Pinnacle.Models;
namespace Pinnacle.Master
{
    public partial class IPEntry : Form, ToolStripAccess
    {
        private static IPEntry _instance;
        public IPEntry()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;

        }

      

        public static IPEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new IPEntry();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        private void IPEntry_Load(object sender, EventArgs e)
        {
            GridLoad(); GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            DataTable dt = mas.findcomcode(Class.Users.HCompcode, Class.Users.HUserName);
            if (dt.Rows.Count > 0)
            {
                combo_compcode.DisplayMember = "COMPCODE";
                combo_compcode.ValueMember = "GTCOMPMASTID";
                combo_compcode.DataSource = dt;


            }
            empty();
        }
       
        private bool ValidateContact(TextBox s)
        {
            if (!Regex.Match(s.Text, "^[0-9.][0-9.]*$").Success)
            {

                s.Text.Remove(s.Text.Length - 1);

            }
            return true;
        }

       public void Saves() { 
            try
            {
                if (txtMACIP.Text != "")
                {
                   
                   

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; } 
                        string sel = "select ASPTBLMACIPID    from  ASPTBLMACIP    WHERE MACIP='" + txtMACIP.Text + "' and compcode='" + combo_compcode.SelectedValue + "' and username='" + combouser.Text + "'  and active='" + chk + "' AND  MACNO='" + txtmachineno.Text + "' AND MTYPE='" + combomactype.Text + "' AND MTYPE2='" + combomactype2.Text + "'  ";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLMACIP");
                        DataTable dt = ds.Tables["ASPTBLMACIP"];
                    if (dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found " + " Alert " + txtMACIP.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtMACIPid.Text) == 0 || Convert.ToInt32("0" + txtMACIPid.Text) == 0)
                    {
                        string ins = "insert into ASPTBLMACIP(MACIP,compcode,active,username,createdon,modifiedon,ipaddress,MACNO,MTYPE,MTYPE2)  VALUES('" + txtMACIP.Text.ToUpper() + "','" + combo_compcode.SelectedValue + "','" + chk + "','" + combouser.Text + "','" + Class.Users.HUserName + "','" + System.DateTime.Now.ToString() + "','" + Class.Users.IPADDRESS + "','" + txtmachineno.Text + "','" + combomactype.Text + "','" + combomactype2.Text + "' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtMACIP.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                    else
                    {
                        string up = "update  ASPTBLMACIP  set   MACIP='" + txtMACIP.Text + "' ,  compcode='" + combo_compcode.SelectedValue + "', active='" + chk + "' ,username='" + combouser.Text + "', modifiedon='" + System.DateTime.Now.ToString() + "',ipaddress='" + Class.Users.IPADDRESS + "', MACNO='" + txtmachineno.Text + "',MTYPE='" + combomactype.Text + "',MTYPE2='" + combomactype2.Text + "' where ASPTBLMACIPID=" + txtMACIPid.Text;
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtMACIP.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad();
                        empty();
                    }

                    
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("MACIP " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void MachineIPEntry_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }



        public void News()
        {

            empty(); GridLoad();
        }
        private void empty()
        {
            txtMACIPid.Text = "";
            txtMACIP.Text = "";combouser.Text = "";
            checkactive.Checked = false;combomactype.Text = "";combomactype2.Text = "";txtmachineno.Text = "";
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            butfooter.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font= Class.Users.FontName;
        }
        public void GridLoad()
        {
           
            try
            {
                listView1.Items.Clear();
                string sel1 = "SELECT  A.ASPTBLMACIPID, A.MACIP , B.COMPCODE,D.USERNAME,  A.ACTIVE    FROM  ASPTBLMACIP   A JOIN GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN ASPTBLUSERMAS  D ON D.COMPCODE=B.GTCOMPMASTID AND D.COMPCODE=A.COMPCODE AND D.USERNAME=A.USERNAME where B.COMPCODE='" + Class.Users.HCompcode + "'  AND D.username='" + Class.Users.HUserName + "'  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACIP");
                DataTable dt = ds.Tables["ASPTBLMACIP"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["ASPTBLMACIPid"].ToString());
                        list.SubItems.Add(myRow["MACIP"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
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

                    txtMACIPid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " SELECT  A.ASPTBLMACIPID, A.MACIP , B.COMPCODE,D.USERNAME,A.MACNO,A.MTYPE,A.MTYPE2 , A.ACTIVE    FROM  ASPTBLMACIP   A JOIN GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE  JOIN ASPTBLUSERMAS  D ON D.COMPCODE=B.GTCOMPMASTID AND D.COMPCODE=A.COMPCODE AND D.USERNAME=A.USERNAME  where a.ASPTBLMACIPid=" + txtMACIPid.Text+";";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACIP");
                    DataTable dt = ds.Tables["ASPTBLMACIP"];

                    if (dt.Rows.Count > 0)
                    {
                        txtMACIPid.Text = Convert.ToString(dt.Rows[0]["ASPTBLMACIPID"].ToString());
                        txtMACIP.Text = Convert.ToString(dt.Rows[0]["MACIP"].ToString());
                        combo_compcode.Text = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combouser.Text = Convert.ToString(dt.Rows[0]["username"].ToString());
                        txtmachineno.Text = Convert.ToString(dt.Rows[0]["MACNO"].ToString());
                        combomactype.Text = Convert.ToString(dt.Rows[0]["MTYPE"].ToString());
                        combomactype2.Text = Convert.ToString(dt.Rows[0]["MTYPE2"].ToString());
                 
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }


                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            txtMACIP.Focus();
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtsearch.Text.ToUpper() != "")
                {
                    listView1.Items.Clear(); int iGLCount = 1;
                    string sel1 = "  SELECT  a.ASPTBLMACIPID,a.MACIP ,a.active from ASPTBLMACIP a  where a.MACIP LIKE'%" + txtsearch.Text.ToUpper() + "%' ;";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACIP");
                    DataTable dt = ds.Tables["ASPTBLMACIP"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["ASPTBLMACIPID"].ToString());
                            list.SubItems.Add(myRow["MACIP"].ToString());
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


        public void Prints()
        {
           // 
        }

        public void Searchs()
        {
            //
        }

        public void Deletes()
        {

            if (txtMACIPid.Text != "")
            {
                string sel1 = "select a.ASPTBLMACIPID from ASPTBLMACIP a  where a.ASPTBLMACIPID='" + txtMACIPid.Text + "';";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACIP");
                DataTable dt = ds.Tables["ASPTBLMACIP"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtMACIP.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    string del = "delete from ASPTBLMACIP where ASPTBLMACIPID='" + Convert.ToInt64("0" + txtMACIPid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtMACIP.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        public void Imports()
        {
            
        }

        public void Pdfs()
        {
           // 
        }

        public void ChangePasswords()
        {
           // 
        }

        public void DownLoads()
        {
           // 
        }

        public void ChangeSkins()
        {
          //  
        }

        public void Logins()
        {
           // 
        }

        public void GlobalSearchs()
        {
           // 
        }

        public void TreeButtons()
        {
          //  
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        private void combo_compcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (combo_compcode.SelectedIndex >= 0)
                {

                    Int64 s = Convert.ToInt64(combo_compcode.SelectedValue);
                    DataTable dt1 = mas.comcode1(s);
                    if (dt1.Rows.Count > 0)
                    {
                        combouser.DisplayMember = "USERNAME";
                        combouser.ValueMember = "userid";
                        combouser.DataSource = dt1;

                    }
                    combouser.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void txtmachineno_TextChanged(object sender, EventArgs e)
        {

        }

        private void combomactype_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

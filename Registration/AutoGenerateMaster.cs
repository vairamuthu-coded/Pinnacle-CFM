using Pinnacle.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Registration
{
    public partial class AutoGenerateMaster : Form,ToolStripAccess
    {
        public AutoGenerateMaster()
        {
            InitializeComponent();
         
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            
        }

        ListView listfilter = new ListView();
        private static AutoGenerateMaster _instance;
        Models.Master mas = new Models.Master();
      //  Models.UserRights sm = new Models.UserRights();
     
        public static AutoGenerateMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new AutoGenerateMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

        CommonClass CC=new CommonClass();
        public DataTable autonumberload(string year,string com,string scode,Int64 sc)
        {
            try
            {

                CC.query = "select max(a.asptblautogeneratemasid) as  id from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid  where  a.finyear='" + year + "' and b.compcode ='" + com + "' and a.shortcode ='" + scode + "' and a.screen ='" + sc + "' ";
                CC.ds = Utility.ExecuteSelectQuery(CC.query, "asptblautogeneratemas");
                CC.dt = CC.ds.Tables["asptblautogeneratemas"];
                int cnt = CC.dt.Rows.Count;
                if (CC.dt.Rows[0]["id"].ToString() == "")
                {
                    txtshortcode.Text = ""; txtseqid.Text = ""; 
                    txtseqid.Text = combofinyear.Text + "-" + scode + "-" + 1;
                    txtshortcode.Text = scode;
                }
                else
                {
                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("autonumberload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return CC.dt;
        }
        public void News()
        {
            finyear();
            DataTable dt1 = mas.aspcomcode();
            if (dt1.Rows.Count >= 0)
            {


                combocompcode.DisplayMember = "COMPCODE";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt1;


            }

            Utility.Load_Combo(comboscreen, "select DISTINCT A.menuid, a.menuname from asptblnavigation a join gtcompmast b on a.compcode=b.gtcompmastid where a.active='T' and b.compcode='" + combocompcode.Text+"' and a.parentmenuid> 1 order by 1 ", "MENUID", "menuname");
            Utility.Load_Combo(combocompcode1, "SELECT distinct b.gtcompmastid, B.compcode FROM asptblautogeneratemas A JOIN GTCOMPMAST B ON A.compcode=B.GTCOMPMASTID order by 2 ", "gtcompmastid", "compcode");
            GridLoad();
            empty();
        }
        public void ReadOnlys()
        {

        }
        public void Saves()
        {
            try
            {
                if (combofinyear.Text != null && combocompcode.Text != null)
                {
                  
                    string chk = ""; Class.Users.Finyear = combofinyear.Text;
                    if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                    GlobalVariables.Dt = Utility.SQLQuery("select * from  asptblautogeneratemas  where    finyear='" + combofinyear.Text + "' and   compcode='" + combocompcode.SelectedValue + "' and    screen='" + comboscreen.SelectedValue + "' and shortcode='" + txtshortcode.Text + "'  and   sequenceno='" + txtseqno.Text + "' and  active='" + chk + "' and  barcode='" + txtbarcode.Text + "'");
                    if (GlobalVariables.Dt.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found. " + txtautoid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (GlobalVariables.Dt.Rows.Count != 0 && Convert.ToInt64("0" + txtautoid.Text) == 0 || Convert.ToInt64("0" + txtautoid.Text) == 0)
                    {
                        autonumberload(combofinyear.Text, combocompcode.Text, txtshortcode.Text, Convert.ToInt64(comboscreen.SelectedValue.ToString()));

                        string ins = "insert into asptblautogeneratemas(sequenceid,  finyear,   compcode,    screen,  shortcode,  sequenceno,  active,  compcode1,  username,  createdby,   modifiedby, ipaddress,barcode)  VALUES('" + txtseqid.Text + "',  '" + combofinyear.Text + "',   '" + combocompcode.SelectedValue + "',    '" + comboscreen.SelectedValue + "',   '" + txtshortcode.Text + "', '" + txtseqno.Text + "',  '" + chk + "',  '" + Class.Users.COMPCODE + "',  '" + Class.Users.USERID + "',  '" + System.DateTime.Now.ToString() + "', '" + Class.Users.HUserName + "', '" + Class.Users.IPADDRESS + "','"+txtbarcode.Text+"' )";
                        Utility.ExecuteNonQuery(ins);
                        MessageBox.Show("Record Saved Successfully " + txtautoid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(combocompcode.Text); empty();
                    }
                    else
                    {

                        string up = "update  asptblautogeneratemas  set sequenceid='" + txtseqid.Text + "',finyear='" +combofinyear.Text+ "',compcode='" + combocompcode.SelectedValue + "',screen='" + comboscreen.SelectedValue + "',shortcode='" + txtshortcode.Text + "',sequenceno='" + txtseqno.Text + "',active='" + chk + "',barcode='" + txtbarcode.Text + "',compcode1='" + Class.Users.COMPCODE + "',username='" + Class.Users.USERID + "',createdby='" + System.DateTime.Now.ToString() + "',modifiedby='" + Class.Users.HUserName + "', ipaddress='" + Class.Users.IPADDRESS + "' where asptblautogeneratemasid='" + txtautoid.Text + "'";
                        Utility.ExecuteNonQuery(up);
                        MessageBox.Show("Record Updated Successfully " + txtautoid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(combocompcode.Text);
                        empty();
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Grade " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        void empty()
        {
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            txtbarcode.Text = "";txtautoid.Text = "";
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName; GlobalVariables.New_Flg = false;
            comboscreen.Text = ""; txtseqid.Text = ""; txtseqno.Text = "";
            txtsearch.Text = ""; comboscreen.Select(); txtshortcode.Text = "";

        }
        public void GridLoad(string s)
        {
            if (s != "")
            {
              
                listfilter.Items.Clear();
                string sel = "SELECT A.asptblautogeneratemasID, a.sequenceid,a.FINYEAR,B.COMPCODE,d.menuname,a.shortcode,A.sequenceno,A.ACTIVE  FROM  asptblautogeneratemas A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID  join ASPTBLNAVIGATION d on d.MENUID=a.screen where b.compcode='" + s + "' ORDER BY a.asptblautogeneratemasID desc";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblautogeneratemas");
                DataTable dt = ds.Tables["asptblautogeneratemas"];
                listView1.Items.Clear();
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblautogeneratemasID"].ToString());
                        list.SubItems.Add(myRow["sequenceid"].ToString());
                        list.SubItems.Add(myRow["FINYEAR"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["menuname"].ToString());
                        list.SubItems.Add(myRow["shortcode"].ToString());
                        list.SubItems.Add(myRow["sequenceno"].ToString());
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count    :" + listView1.Items.Count;
                }
            }
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            if (listView1.Items.Count > 0)
            {

                txtautoid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                string sel = "SELECT A.asptblautogeneratemasID, a.sequenceid,a.FINYEAR,B.COMPCODE,d.menuname ,a.shortcode,A.sequenceno,A.ACTIVE,A.BARCODE  FROM  asptblautogeneratemas A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID  join asptblnavigation d on d.MENUID=a.screen WHERE A.asptblautogeneratemasID='" + txtautoid.Text + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblautogeneratemas");
                DataTable dt = ds.Tables["asptblautogeneratemas"];
                if (dt.Rows.Count > 0)
                {
                    GlobalVariables.New_Flg = true;
                    txtautoid.Text = Convert.ToString(dt.Rows[0]["asptblautogeneratemasID"].ToString());
                    txtseqid.Text = Convert.ToString(dt.Rows[0]["sequenceid"].ToString());
                    combofinyear.Text = Convert.ToString(dt.Rows[0]["FINYEAR"].ToString());
                    combocompcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                    comboscreen.Text = Convert.ToString(dt.Rows[0]["menuname"].ToString().ToUpper());
                    txtshortcode.Text = Convert.ToString(dt.Rows[0]["shortcode"].ToString().ToUpper());
                    txtseqno.Text = Convert.ToString(dt.Rows[0]["sequenceno"].ToString());
                    txtbarcode.Text = Convert.ToString(dt.Rows[0]["BARCODE"].ToString());
                    if (dt.Rows[0]["ACTIVE"].ToString() == "T") { checkactive.Checked = true; checkactive.ForeColor = Color.Red; } else { checkactive.Checked = false; checkactive.ForeColor = Color.Black; }


                    tabControl1.SelectTab(tabPage1);
                }
            }
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; int i = 1;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();

                        if (item.SubItems[6].ToString().Contains(txtsearch.Text) || item.SubItems[7].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = i.ToString();
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.SubItems.Add(item.SubItems[7].Text);
                            list.SubItems.Add(item.SubItems[8].Text);
                            list.SubItems.Add(item.SubItems[9].Text);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                        }
                        i++;
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
                            list.Text = i.ToString();                         
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);

                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);

                            list.SubItems.Add(item.SubItems[7].Text);
                            list.SubItems.Add(item.SubItems[8].Text);
                            list.SubItems.Add(item.SubItems[9].Text);

                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                            item0++; i++;
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
        void finyear()
        {
          
           Utility.Load_Combo(combofinyear, "SELECT SUBSTRING(now(),1,4) AS FINYEAR ", "finyear", "finyear");
           
        }
        private void AutoGenerateMaster_Load(object sender, EventArgs e)
        {
            News();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad(combocompcode1.Text);
        }

        public void Prints()
        {
           
        }

        public void Searchs()
        {
           
        }

        public void Deletes()
        {
            Class.Users.Query= "DELETE FROM asptblautogeneratemas A WHERE A.asptblautogeneratemasID="+txtautoid.Text;
            Utility.ExecuteNonQuery(Class.Users.Query);
            MessageBox.Show("Record Deleted...");
            GridLoad(combocompcode1.Text);
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
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }

        private void checktrue_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        private void comboscreen_SelectedIndexChanged(object sender, EventArgs e)
        {
         
        }

        private void txtseqno_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void combocompcode1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridLoad(combocompcode1.Text);
        }

        public void GridLoad()
        {
            GridLoad(combocompcode1.Text);
        }

        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

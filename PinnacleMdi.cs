using Pinnacle.Master;
using Pinnacle.Transactions.Lyla;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle
{
    public partial class PinnacleMdi : Form, ToolStripAccess
    {

        int mid = 0; int mid1 = 0; string systemuser = ""; int sessioncount = 0;
        DataTable dtall = new DataTable();
        public PinnacleMdi()
        {
            InitializeComponent();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            systemuser = Environment.UserName;

            tabControl1.TabPages.Remove(Masters);
            tabControl1.TabPages.Remove(Transactions);
            tabControl1.TabPages.Remove(Registration);
            tabControl1.TabPages.Remove(TreeView);
            tabControl1.TabPages.Remove(Reports);
            tabControl1.TabPages.Remove(Approval);
            tabControl1.TabPages.Remove(Approval2);
            mdipanel.Show();
            lblheader.Text = "";

        }
        Image img;
        Models.UserRights sm = new Models.UserRights();
        List<Models.UserRights> usr = new List<Models.UserRights>();
        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        ToolStripMenuItem MnuStripItem = new ToolStripMenuItem();
        TreeNode mainNode = new TreeNode();
        DataTable dtlist = new DataTable();


        List<string> listfilter = new List<string>(); List<string> uniqueList = new List<string>();
        void companylogo(string c)
        {
            string sel = "select a.companylogo from gtcompmast a where a.compcode='" + c + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            if (dt != null)
            {
                int cc = dt.Rows.Count;
          
                byte[] stdbytes; Int64 std;
                if (cc >= 1 && dt.Rows[0]["companylogo"].ToString() != "")
                {
                    img = null; stdbytes = null;
                    stdbytes = (byte[])dt.Rows[0]["companylogo"];
                    img = Models.Device.ByteArrayToImage(stdbytes);

                }
                else
                {
                    img = null; stdbytes = null;
                }
            }
        }

        private void PinnacleMdi_Load(object sender, EventArgs e)
        {
            try
            {
                Class.Users.screen = false; tabControl1.Visible = true;
                paneltree.Visible = false; toolStripStatusLabel3.Text = Class.Users.UserTime.ToString();
                PinnacleMdi.ActiveForm.Text = Class.Users.HCompName.ToString() + "  UserName: " + Class.Users.HUserName.ToString() + "  ProjectName  :" + Class.Users.ProjectID + " - " + Class.Users.HostelName;
                treeload(); systemuser = Environment.UserName;
                combosearchload(); MnuStrip.Items.Clear();
                lblMarquee1.Text = "Welcome to " + Class.Users.HCompName.ToString();
                timer2.Enabled = true;
                Class.Users.Query = "select a.userrightsid, a.menuid,e.menuname,a.parentmenuid  from  asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username join asptblnavigation d on d.menuid=a.menuid join asptblmenuname  e on e.menunameid=d.menunameid where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.active='T'  and a.parentmenuid = 1 order by 1";// and a.parentmenuid = 1
                DataSet ds0 = Utility.ExecuteSelectQuery(Class.Users.Query, "asptbluserrights");
                DataTable dt0 = ds0.Tables["asptbluserrights"];
                this.butnext.ForeColor = System.Drawing.Color.White;
                foreach (DataRow dr in dt0.Rows)
                {
                    if (dr["menuname"].ToString() == "Approval" || dr["menuname"].ToString() == "Approval2") { }
                    else
                    {
                        MnuStripItem = new ToolStripMenuItem(dr["menuname"].ToString());
                        mid = Convert.ToInt32("0" + dr["menuid"].ToString());
                        SubMenu(MnuStripItem, mid);
                        MnuStrip.Items.Add(MnuStripItem);
                    }
                }
                
            }
            catch (Exception ex)
            {

            }
            timer1.Enabled = true;
            uniqueList.Clear();
            this.Font = Class.Users.FontName;
            GlobalVariables.MenuStrip1 = this.MnuStrip;
            GlobalVariables.Toolstrip1 = this.toolStrip1;
            GlobalVariables.News = this.News;
            GlobalVariables.Saves = this.Saves;
            GlobalVariables.Prints = this.Prints;
            GlobalVariables.Searchs = this.Searchs;
            GlobalVariables.Deletes = this.Deletes;
            GlobalVariables.ReadOnlys = this.ReadOnlys;
            GlobalVariables.MdiPanel = this.mdipanel;
            GlobalVariables.Imports = this.Imports;
            GlobalVariables.Pdfs = this.Pdfs;
            GlobalVariables.DownLoads = this.DownLoads;
            GlobalVariables.ChangeSkins = this.ChangeSkins;
            GlobalVariables.ChangePasswords = this.ChangePasswords;
            GlobalVariables.Logins = this.Logins;
            GlobalVariables.GlobalSearchs = this.GlobalSearchs;
            GlobalVariables.TreeButtons = this.TreeButtons;
            GlobalVariables.HeaderName = this.lblheader;
            GlobalVariables.Exit = this.Exit;
            GlobalVariables.MasterForm = this;
            GlobalVariables.TabCtrl = this.TabCtrl;
            combosearch.Select();
            GlobalVariables.Toolstrip1.Font = Class.Users.FontName;
            GlobalVariables.Toolstrip1.Visible = false;
            GlobalVariables.MenuStrip1.Font = Class.Users.FontName;
            labelsearch.Font = Class.Users.FontName;
            this.Font = Class.Users.FontName;
            TabCtrl.Visible = false;

            LoginForm log1 = new LoginForm();
            log1.FormClosed += Log1_FormClosed;
            dtlist = ConvertListToDataTable(usr);
            tabControl1.Visible = true;
            Backcolor3_Click(sender, e);
            companylogo(Class.Users.HCompcode);
            string sel = "select a.userrightsid, a.menuid,e.menuname,a.parentmenuid  from  asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username join asptblnavigation d on d.menuid=a.menuid join asptblmenuname  e on e.menunameid=d.menunameid where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.active='T'  and a.parentmenuid = 1 order by 1";// and a.parentmenuid = 1
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblmenuname");
            DataTable dt = ds.Tables["asptblmenuname"];
           
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if ("Masters" == dt.Rows[i]["menuname"].ToString())
                {

                    DataTable dt1 = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, dt.Rows[i]["menuname"].ToString());

                    if (dt1.Rows.Count > 0)
                    {
                        
                        Class.Users.Description = "";
                        pop1(2, flowLayoutPanel1);
                        tabControl1.TabPages.Add(Masters);
                    }

                }

                if ("Transactions" == dt.Rows[i]["menuname"].ToString())
                {
                    DataTable dt1 = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, dt.Rows[i]["menuname"].ToString());

                    if (dt1.Rows.Count > 0)
                    {
                        tabControl1.TabPages.Add(Transactions);
                        pop1(3, flowLayoutPanel2);
                        tabControl1.SelectTab(Transactions);
                    }

                }
                if ("Reports" == dt.Rows[i]["menuname"].ToString())
                {
                    DataTable dt1 = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, dt.Rows[i]["menuname"].ToString());

                    if (dt1.Rows.Count > 0)
                    {
                        tabControl1.TabPages.Add(Reports);
                    }
                    

                }
                if ("TreeView" == dt.Rows[i]["menuname"].ToString())
                {
                    DataTable dt1 = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, dt.Rows[i]["menuname"].ToString());

                    if (dt1.Rows.Count > 0)
                    {
                        tabControl1.TabPages.Add(TreeView);
                    }

                }
                if ("Registration" == dt.Rows[i]["menuname"].ToString())
                {
                    DataTable dt1 = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, dt.Rows[i]["menuname"].ToString());

                    if (dt1.Rows.Count > 0)
                    {
                        tabControl1.TabPages.Add(Registration);
                    }

                }
                //if ("Approval" == dt.Rows[i]["menuname"].ToString())
                //{

                //    DataTable dt1 = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, dt.Rows[i]["menuname"].ToString());

                //    if (dt1.Rows.Count > 0)
                //    {
                //        Class.Users.Description = "Approval";
                //        pop2(0, dataGridView1);
                //        tabControl1.TabPages.Add(Approval);
                //        tabControl1.SelectTab(Approval);
                //    }
                //    else
                //    {
                //        Class.Users.Description = "";
                        
                //            pop1(3, flowLayoutPanel2);
                //            tabControl1.SelectTab(Transactions);
                       
                //    }

                //}
                //if ("Approval2" == dt.Rows[i]["menuname"].ToString())
                //{


                //    DataTable dt1 = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, dt.Rows[i]["menuname"].ToString());

                //    if (dt1.Rows.Count > 0)
                //    {
                //        Class.Users.Description = "Approval2";
                //        pop2(0, dataGridView2);
                //        tabControl1.TabPages.Add(Approval2);
                //        tabControl1.SelectTab(Approval2);
                //    }
                //    else
                //    {
                //        Class.Users.Description = "";
                //        pop1(3, flowLayoutPanel2);
                //        tabControl1.SelectTab(Transactions);
                //    }

                //}
                //if (dt.Rows.Count == 1)
                //{
                //    if (dt.Rows[0]["menuname"].ToString() == "Approval2") {
                        
                //        tabControl1.SelectTab(Approval2);
                //    }
                //    else
                //    {
                //        pop1(2, flowLayoutPanel1);
                //        tabControl1.SelectTab(Masters);
                //    }
                   
                //}
            }


           


        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0; 
            if (tabControl1.SelectedTab == tabControl1.TabPages["Masters"])
            {
                Class.Users.Description = "";
                pop1(2, flowLayoutPanel1);

            }


            if (tabControl1.SelectedTab == tabControl1.TabPages["Transactions"])
            {
                Class.Users.Description = "";
                pop1(3, flowLayoutPanel2);
            }

            if (tabControl1.SelectedTab == tabControl1.TabPages["TreeView"])
            {
                Class.Users.Description = "";
                pop1(4, flowLayoutPanel3);
            }

            if (tabControl1.SelectedTab == tabControl1.TabPages["Registration"])
            {
                Class.Users.Description = "";
                pop1(5, flowLayoutPanel4);
            }


            if (tabControl1.SelectedTab == tabControl1.TabPages["Reports"])
            {
                Class.Users.Description = "";
                pop1(6, flowLayoutPanel5);
            }

            if (tabControl1.SelectedTab == tabControl1.TabPages["Approval"])
            {
                Class.Users.Description = "Approval";
                pop2(0, dataGridView1);
                CommonFunctions.SetRowNumber(dataGridView1);
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["Approval2"])
            {
                Class.Users.Description = "Approval2";
                pop2(0, dataGridView2);
                CommonFunctions.SetRowNumber(dataGridView2);
            }


        }

        private void pop1(int param, FlowLayoutPanel lay)
        {

            Class.Users.Description = "";
            DataTable dt = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, param);
            CustomControl[] items = new CustomControl[dt.Rows.Count];
            lay.Controls.Clear(); pictureBox1.Image = img;
            foreach (DataRow myRow in dt.Rows)
            {
                items[i] = new CustomControl();
                items[i].menuname.Text = Convert.ToString(myRow["menuname"].ToString());
                items[i].userimage = img;
                lay.Controls.Add(items[i]);
                items[i].menuname.Click += Menuname_Click;
                items[i].iconbackground.BackColor = Class.Users.BackColors;
                items[i].menuname.ForeColor = Class.Users.BackColors;
                items[i].Border.BackColor = Class.Users.BackColors;

            }



        }


        private void pop2(int param, DataGridView lay)
        {
            try
            {
                this.BackColor = Class.Users.BackColors;
                this.Font = Class.Users.FontName; Class.Users.UserTime = 0;
                DataTable dt = new DataTable();
                if (Class.Users.Description == "Approval")
                {
                    string f = "select  a.asptbladvpaymasid, a.dateofpayment,d.department, b.partyname,a.orderno,a.itemdesc, a.deductionamt,a.advanceterms,a.advanceamount,a.modeofpayment as paymenttype,e.resonseperson,'' as Attach,'' as ManagerRemarks,A.MDStatus from  asptbladvpaymas a join  asptblpartymas b on a.partyname=b.asptblpartymasid join asptblbanmas c on c.asptblbanmasid=b.bankname join asptbldeptmas d on d.asptbldeptmasID=a.department join asptblresmas e on e.asptblresmasid=a.responseperson where a.approval='F' and a.mdapproval='F' and a.conformed='F' and not a.approval='R'   order by a.asptbladvpaymasid desc ";
                    DataSet ds0 = Utility.ExecuteSelectQuery(f, "asptbladvpaymas");
                    dt = ds0.Tables[0];
                }
                if (Class.Users.Description == "Approval2")
                {
                    dt = null;
                    string f = "select  a.asptbladvpaymasid, a.dateofpayment,d.department, b.partyname,a.orderno,a.itemdesc,a.deductionamt,a.advanceterms,a.advanceamount,a.modeofpayment as paymenttype,e.resonseperson,'' as Attach1,'' as MDRemarks,a.ManagerRemarks,a.MDStatus,a.Notes from  asptbladvpaymas a join  asptblpartymas b on a.partyname=b.asptblpartymasid join asptblbanmas c on c.asptblbanmasid=b.bankname join asptbldeptmas d on d.asptbldeptmasID=a.department join asptblresmas e on e.asptblresmasid=a.responseperson where a.approval='T' and a.mdapproval='F' and a.conformed='F' and not a.approval='R' order by a.asptbladvpaymasid desc ";
                    DataSet ds0 = Utility.ExecuteSelectQuery(f, "asptbladvpaymas");
                    dt = ds0.Tables[0];
                }
                lay.Controls.Clear();
                if (dt.Rows.Count > 0)
                {
                    butwaitapproval.Text = "Waiting for Approval";
                    butmdapproval.Text = "Waiting for Approval";
                    Class.Users.Description = "";
                    pictureBox1.Image = img;                   
                    lay.Font = Class.Users.FontName;                  
                    CommonFunctions.SetRowNumber(lay);
                    lay.DataSource = dt;
                }
                else
                {
                    butwaitapproval.Text = "No Content";
                    butmdapproval.Text = "No Content";
                    do
                    {
                        for (int i = 0; i < lay.Rows.Count; i++) { try { lay.Rows.RemoveAt(i); } catch (Exception) { } }
                    }
                    while (lay.Rows.Count > 1);

                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message) ; }
        }



   
        int listcount = 0;
        private void Log1_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Close();
        }
        static DataTable ConvertListToDataTable(List<Models.UserRights> list)
        {

            DataTable table = new DataTable();
            table.Columns.Add("MenuName", typeof(string));
            foreach (var dvd in list)
            {
                var row = table.NewRow();
                row["MenuName"] = dvd.MenuName;
                table.Rows.Add(row);
            }
            return table;
        }
        public void SubMenu(ToolStripMenuItem mnu, int midd)
        {

            Class.Users.Query = "  select a.menuid, e.aliasname as menuname ,a.parentmenuid  from   asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username  join asptblnavigation d on d.menuid = a.menuid join asptblmenuname  e on e.menunameid=d.menunameid where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.parentmenuid='" + midd + "' and  a.active='T' order by 1";
            DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "asptbluserrights");
            DataTable dtchild = ds.Tables["asptbluserrights"];
            for (int j = 0; j < dtchild.Rows.Count; j++)
            {
                ToolStripMenuItem SSMenu = new ToolStripMenuItem(dtchild.Rows[j]["menuname"].ToString(), null, new EventHandler(ChildClick));
                mnu.DropDownItems.Add(SSMenu);
                Class.Users.Query = "select a.menuid, e.aliasname as menuname ,a.parentmenuid  from   asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username  join asptblnavigation d on d.menuid = a.menuid join asptblmenuname  e on e.menunameid=d.menunameid  where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.parentmenuid='" + dtchild.Rows[j]["menuid"].ToString() + "' and  a.active='T' order by 1";
                DataSet ds1 = Utility.ExecuteSelectQuery(Class.Users.Query, "asptbluserrights");
                DataTable dtchild1 = ds1.Tables["asptbluserrights"];
                if (dtchild1.Rows.Count>0)
                {
                    for (int k = 0; k < dtchild1.Rows.Count; k++)
                    {
                      
                        ToolStripMenuItem SSMenu1 = new ToolStripMenuItem(dtchild1.Rows[k]["menuname"].ToString(), null, new EventHandler(ChildClick));
                        SSMenu.DropDownItems.Add(SSMenu1);

                        Class.Users.Query = "select a.menuid, e.aliasname as menuname ,a.parentmenuid  from   asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username  join asptblnavigation d on d.menuid = a.menuid  join asptblmenuname  e on e.menunameid=d.menunameid  where b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.parentmenuid='" + dtchild1.Rows[k]["menuid"].ToString() + "' and  a.active='T' order by 1";
                        DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "asptbluserrights");
                        DataTable dtchild2 = ds2.Tables["asptbluserrights"];
                        if (dtchild2.Rows.Count > 0)
                        {
                            for (int l = 0; l < dtchild2.Rows.Count; l++)
                            {
                                ToolStripMenuItem SSMenu2 = new ToolStripMenuItem(dtchild2.Rows[l]["menuname"].ToString(), null, new EventHandler(ChildClick));
                                SSMenu1.DropDownItems.Add(SSMenu2);
                                Class.Users.Query = "select a.menuid,e.aliasname as menuname,a.parentmenuid  from   asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username  join asptblnavigation d on d.menuid = a.menuid join asptblmenuname  e on e.menunameid=d.menunameid  where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.parentmenuid='" + dtchild2.Rows[l]["menuid"].ToString() + "' and  a.active='T' order by 1";
                                DataSet ds3 = Utility.ExecuteSelectQuery(Class.Users.Query, "asptbluserrights");
                                DataTable dtchild3 = ds3.Tables["asptbluserrights"];
                                if (dtchild3.Rows.Count > 0)
                                {
                                    for (int m = 0; m < dtchild3.Rows.Count; m++)
                                    {

                                        ToolStripMenuItem SSMenu3 = new ToolStripMenuItem(dtchild3.Rows[m]["menuname"].ToString(), null, new EventHandler(ChildClick));
                                        SSMenu2.DropDownItems.Add(SSMenu3);
                                    }
                                }
                            }
                        }

                    }
                }
              

            }

        }
        


        int i = 0;

        private void combosearchload()
        {
          
            DataTable dt1 = sm.headerdropdowns();
            combosearch.DisplayMember = "menuname";
            combosearch.ValueMember = "menuname";
            combosearch.DataSource = dt1;

            combosearch.Text = ""; combosearch.SelectedIndex = -1;
        }
        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            int cnt = dt1.Rows.Count;
            if (cnt >= 1)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {
                        if (dt1.Rows[r]["NEWS"].ToString() == "T") { GlobalVariables.News.Visible = true; } else { GlobalVariables.News.Visible = false; }
                        if (dt1.Rows[r]["SAVES"].ToString() == "T") { GlobalVariables.Saves.Visible = true; } else { GlobalVariables.Saves.Visible = false; }
                        if (dt1.Rows[r]["PRINTS"].ToString() == "T") { GlobalVariables.Prints.Visible = true; } else { GlobalVariables.Prints.Visible = false; }
                        if (dt1.Rows[r]["READONLY"].ToString() == "T") { GlobalVariables.ReadOnlys.Visible = false; this.Enabled = true; } else { this.Enabled = false; }
                        if (dt1.Rows[r]["SEARCH"].ToString() == "T") { GlobalVariables.Searchs.Visible = true; } else { GlobalVariables.Searchs.Visible = false; }
                        if (dt1.Rows[r]["DELETES"].ToString() == "T") { GlobalVariables.Deletes.Visible = true;  } else { GlobalVariables.Deletes.Visible = false; }
                        if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { GlobalVariables.TreeButtons.Visible = false; } else { GlobalVariables.TreeButtons.Visible = false; }
                        if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { GlobalVariables.GlobalSearchs.Visible = true; } else { GlobalVariables.GlobalSearchs.Visible = false; }
                        if (dt1.Rows[r]["LOGIN"].ToString() == "T") { GlobalVariables.Logins.Visible = true; } else { GlobalVariables.Logins.Visible = false; }
                        if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { GlobalVariables.ChangePasswords.Visible = true; } else { GlobalVariables.ChangePasswords.Visible = false; }
                        if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { GlobalVariables.ChangeSkins.Visible = true; } else { GlobalVariables.ChangeSkins.Visible = false; }
                        if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { GlobalVariables.DownLoads.Visible = true; } else { GlobalVariables.DownLoads.Visible = false; }
                        if (dt1.Rows[r]["Pdf"].ToString() == "T") { GlobalVariables.Pdfs.Visible = true; } else { GlobalVariables.Pdfs.Visible = false; }
                        if (dt1.Rows[r]["Imports"].ToString() == "T") { GlobalVariables.Imports.Visible = true; } else { GlobalVariables.Imports.Visible = false; }

                    }
                }   
            }
            else
            {

                MessageBox.Show("No Screen Rights Defined .Pls Contact Your Administrator...");

                this.Hide();


            }


        }

      

        string[] s;
        private void Menuname_Click(object sender, EventArgs e)
        {
            Class.Users.ScreenName = ""; Class.Users.UserTime = 0;
            s = sender.ToString().Split(',');
            sender = s[1].Substring(7).TrimEnd();
            ChildClick(sender, e);
        }
        //private void AppCheckBox_CheckedChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        Class.Users.ScreenName = ""; Class.Users.UserTime = 0;
        //        s = ((CheckBox)sender).Text.Split(':');
        //        string ss = s[0].TrimEnd();
        //        string update = "update asptbladvpaymas set approval='T'  where asptbladvpaymasid='" + ss + "'";
        //        Utility.ExecuteNonQuery(update);
        //        Models.Master mas = new Models.Master();
        //        mas.pop("Approved", "", "");


        //    }
        //    catch (Exception ex) { }
        //}


   
        private void Username_Click(object sender, EventArgs e)
        {
            Class.Users.ScreenName = "";
            
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            mdipanel.Show();

        }
        Int64 parentID = 0; TreeNode parentNode = null;
        private void treeload()
        {
            string sel = "select a.userrightsid, a.menuid,  e.aliasname as  menuname ,a.parentmenuid  from  asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username    join asptblnavigation d on d.menuid=a.menuid   join asptblmenuname  e on e.menunameid=d.menunameid where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.active='T' and a.parentmenuid = 1 order by 1";// and a.parentmenuid = 1
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            DataTable dt = ds.Tables["asptbluserrights"];
            foreach (DataRow dr in dt.Rows)
            {
                parentNode = treeView1.Nodes.Add(dr["menuname"].ToString());
                PopulateTreeView(Convert.ToInt64(dr["menuid"].ToString()), parentNode);

            }
        }

        private void PopulateTreeView(Int64 parentId, TreeNode parentNode)
        {
          
             DataTable dtchildc = sm.headerdropdowns(Class.Users.HCompcode, Class.Users.HUserName, parentId);

            TreeNode childNode;
            foreach (DataRow dr in dtchildc.Rows)
            {
                Models.UserRights sm1 = new Models.UserRights();
                if (parentNode == null)
                {
                    childNode = treeView1.Nodes.Add(dr["menuname"].ToString());

                }
                else
                {
                    childNode = parentNode.Nodes.Add(dr["menuname"].ToString());

                }
                PopulateTreeView(Convert.ToInt32("0" + dr["menuid"].ToString()), childNode);
                sm1.MenuName = dr["menuname"].ToString();
                usr.Add(sm1);

            }
        }
       
       


        public void SubMenu1(TreeNode mnu, int midd)
        {
            DataTable dtchild = sm.TreeView(Class.Users.HCompcode, Class.Users.HUserName, midd);
         
            for (int j = 0; j < dtchild.Rows.Count; j++)
            {               
                TreeNode SSMenu = new TreeNode(dtchild.Rows[j]["menuname"].ToString());
                mnu.Nodes.Add(SSMenu);

            }
        }
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {

            sender = e.Node.Text;
            ChildClick(sender, e);
        }
       
        

        private DataTable ConvertListToDataTable()
        {

            List<string> uniqueList = listfilter.Distinct().ToList();
            dtlist.Rows.Clear();
            dtlist.Columns.Clear();
            if (dtlist.Rows.Count == 0)
            {
                dtlist.Columns.Add("MenuName", typeof(string));
            }

            foreach (string str in uniqueList)
            {
                DataRow row = dtlist.NewRow();
                row["MenuName"] = str;
                dtlist.Rows.Add(row);
            }


            return dtlist;
        }
        
        string screen1 = "";
        private void button5_Click(object sender, EventArgs e)
        {

            butprevious_Click(sender,e);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            butnext_Click(sender,e);
        }


        private void ChildClick(object sender, EventArgs e)
        {
            if (sender.ToString() != "")
            {
                Class.Users.ScreenName = "";
                Class.Users.ScreenName = sender.ToString();
                GlobalVariables.MenuStrip1.Visible = true;
                GlobalVariables.TabCtrl.Visible = true;
                usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
                uniqueList.Add(sender.ToString().Trim());
                GlobalVariables.HeaderName.Text = sender.ToString();
                GlobalVariables.TabCtrl.Font = Class.Users.FontName;
                GlobalVariables.Toolstrip1.Visible = true;
                switchclass(sender.ToString().Trim());
            }
        }


        void switchclass(string scr)
        {
            
            switch (scr)
            {
            

                case "Godown Master":
                    CommonFunctions.ShowPopUpForm(Master.CFM.GodownMaster.Instance, this); button1.Show(); return;
                case "Product Master":
                    CommonFunctions.ShowPopUpForm(Master.CFM.ProductMaster.Instance, this); button1.Show(); return;
                case "Port Master":
                    CommonFunctions.ShowPopUpForm(Master.CFM.PortMaster.Instance, this); button1.Show(); return;
                case "Product Weight Master":
                    CommonFunctions.ShowPopUpForm(Master.CFM.ProductWeightMaster.Instance, this); button1.Show(); return;
                case "Received From Master":
                    CommonFunctions.ShowPopUpForm(Master.CFM.ReceivedFromMaster.Instance, this); button1.Show(); return;
                case "Wheat Variety Master":
                    CommonFunctions.ShowPopUpForm(Master.CFM.WheatVarietyMaster.Instance, this); button1.Show(); return;
                case "Delivery Weight Certificate":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.CFM.DeliveryWeightCertificate.Instance, this); button1.Show(); return;
                case "Raw Material Weight Certificate":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.CFM.RawMaterialWeightCertificate.Instance, this); button1.Show(); return;
                case "Raw Material Report":
                    CommonFunctions.ShowPopUpForm(Report.CFM.RawMaterialReport.Instance, this); button1.Show(); return;
                case "Delivery Report":
                    CommonFunctions.ShowPopUpForm(Report.CFM.DeliveryCertificateReport.Instance, this); button1.Show(); return;
                case "Audit Report":
                    CommonFunctions.ShowPopUpForm(Report.CFM.AuditReport.Instance, this); button1.Show(); return;
                case "Manage Users":
                    CommonFunctions.ShowPopUpForm(Master.CFM.User.Instance, this); button1.Show(); return;

                case "Create AutoGenerate":
                    //showform(Registration.AboutMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Registration.AutoGenerateMaster.Instance, this); button1.Show(); return;
                case "Create About":
                    //showform(Registration.AboutMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Registration.AboutMaster.Instance, this); button1.Show(); return;

                case "Create Generate":
                    //showform(Registration.GenerateMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Registration.GenerateMaster.Instance, this); button1.Show(); return;

                case "Create Register":
                    //showform(Registration.RegisterMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Registration.RegisterMaster.Instance, this); button1.Show(); return;

                case "CountryMaster":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CountryMaster.Instance, this); button1.Show(); return;
                case "Create Country":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CountryMaster.Instance, this); button1.Show(); return;
                case "Create State":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.StateMaster.Instance, this); button1.Show(); return;
                case "Create City":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CityMaster.Instance, this); button1.Show(); return;
                case "Create Company":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CompanyMaster.Instance, this); button1.Show(); return;

                case "Create IfsCode":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.Bank.IFSCMaster.Instance, this); button1.Show(); return;

                case "Create Payment":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Bank.Payment.Instance, this); button1.Show(); return;
                case "Create UTR":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Bank.UTR.Instance, this); button1.Show(); return;

                case "Create Bank":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.Bank.BankMaster.Instance, this); button1.Show(); return;
                case "Create Advance Payment":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Bank.AdvancePayment.Instance, this); button1.Show(); return;
                case "Create Response Person":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.Bank.ResponsePersonMaster.Instance, this); button1.Show(); return;
                case "Create Party":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.Bank.PartyMaster.Instance, this); button1.Show(); return;
                case "Create Department":
                    CommonFunctions.ShowPopUpForm(Master.DepartmentMaster.Instance, this); button1.Show(); return;

                case "Create Payment Report":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Report.Bank.PaymentDetails.Instance, this); button1.Show(); return;

                //case "StateMaster":
                //    CommonFunctions.ShowPopUpForm(Pinnacle.Master.StateMaster.Instance, this); button1.Show(); return;
               
                //case "CityMaster":
                //    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CityMaster.Instance, this); button1.Show(); return;
              
                //case "CompanyMaster":
                //    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CompanyMaster.Instance, this); button1.Show(); return;
               case "Create Transaction Compenstate":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Tally.TransactionCompenstate.Instance, this); button1.Show(); return;
                case "Create ProductionStatusReport":
                    CommonFunctions.ShowPopUpForm(Pinnacle.ReportFormate.Lyla.ProductionStatusReport.Instance, this); button1.Show(); return;

                case "Create OrderCloseEntry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Lyla.OrderCloseEntry.Instance, this);
                    button1.Show(); return;
                case "Create Defect":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.DefectMaster.Instance, this);
                    button1.Show(); return;
               
               
                case "Create SewingLine":
                    //showform(Master.PartyMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.SewingLineMaster.Instance, this); button1.Show(); return;

                case "Create Category":
                    //showform(Master.EmployeeMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CategoryMaster.Instance, this); button1.Show(); return;
                case "Create Employee":
                    //showform(Master.EmployeeMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.PIN.EmployeeMaster.Instance, this); button1.Show(); return;

                case "Create Shiftr":
                    //showform(Master.DashBoard.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.ShiftMaster.Instance, this); button1.Show(); return;

                case "Create FinYear":
                    //showform(Master.FinYearMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Master.FinYearMaster.Instance, this); button1.Show(); return;

                case "Create Holiday":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.HolidayMaster.Instance, this); button1.Show(); return;

                case "Create Location":
                    //   showform(Master.SeasonMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.LocationMaster.Instance, this); button1.Show(); return;

                case "Create Brand":
                    // showform(Master.BrandMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.BrandMaster.Instance, this); button1.Show(); return;
                case "Create Blood Group":
                    CommonFunctions.ShowPopUpForm(Master.BloodGroupMaster.Instance, this); button1.Show(); return;
                case "Create Relation":
                    CommonFunctions.ShowPopUpForm(Master.RelationMaster.Instance, this); button1.Show(); return;
                case "Create Proof":
                    // showform(Master.SizeMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.ProofMaster.Instance, this); button1.Show(); return;
                case "Create Size":
                    // showform(Master.SizeMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.SizeMaster.Instance, this); button1.Show(); return;
                case "Create SizeGroup":
                    CommonFunctions.ShowPopUpForm(Master.SizeGroupMaster.Instance, this); button1.Show(); return;
                case "Create Style":
                    //showform(Master.StyleMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.StyleMaster.Instance, this); button1.Show(); return;

                case "Create Gsm":
                    //showform(Master.GsmMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.GsmMaster.Instance, this); button1.Show(); return;

                case "Create ShiftType":
                    //showform(Master.SampleDepartmentMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.ShiftTypeMaster.Instance, this); button1.Show(); return;

                case "Create Fabric":
                    //showform(Master.FabricMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.FabricMaster.Instance, this); button1.Show(); return;

                case "Create SubStyle":
                    //  showform(Master.SubStyleMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.SubStyleMaster.Instance, this); button1.Show(); return;

                case "Create Counts":
                    // showform(Master.CountsMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.CountsMaster.Instance, this); button1.Show(); return;

                case "Create Gauge":
                    // showform(Master.GaugeMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.GaugeMaster.Instance, this); button1.Show(); return;

                case "Create Buyer":
                    // showform(Master.BuyerMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.BuyerMaster.Instance, this); button1.Show(); return;

                case "Create Color":
                    //showform(Master.ColorMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.ColorMaster.Instance, this); button1.Show(); return;

                case "Create Currency":
                    // showform(Master.CurrencyMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.CurrencyMaster.Instance, this); button1.Show(); return;


                case "Create FabricType":
                    //showform(Master.SamplePackTypeMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.FabricTypeMaster.Instance, this); button1.Show(); return;

                case "Create Rack":
                    //  showform(Master.RackMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.RackMaster.Instance, this); button1.Show(); return;

                case "Create Bin":
                    //showform(Master.BinMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.BinMaster.Instance, this); button1.Show(); return;

                case "Create SampleType":
                    //showform(Master.SampleCollection.SampleTypeMaster.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.SampleCollection.SampleTypeMaster.Instance, this); button1.Show(); return;
               
                case "Create User":
                    //   showform(TreeView.UserMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.TreeView.UserMaster.Instance, this); button1.Show(); return;

                case "Create TreeView":
                    // showform(TreeView.TreeViewMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.TreeView.TreeViewMaster.Instance, this); button1.Show(); return;

                case "Create MenuName":
                    // showform(TreeView.MenuNameMaster.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.TreeView.MenuNameMaster.Instance, this); button1.Show(); return;

                case "Create UserRights":
                    // showform(TreeView.UserRights.Instance);  break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.TreeView.UserRights.Instance, this); button1.Show(); return;

                case "Create Navigation":
                    CommonFunctions.ShowPopUpForm(Pinnacle.TreeView.NavigationMaster.Instance, this); button1.Show(); return;



                case "Create Grade":
                    //showform(ReportFormate.SampleCollectionReport.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.GradeMaster.Instance, this); button1.Show(); return;

                case "Create Season":
                    // showform(Report.DateTimeReport.Instance); break;
                    CommonFunctions.ShowPopUpForm(Master.SeasonMaster.Instance, this); button1.Show(); return;
                case "Create DateTime Report":
                    // showform(Report..Instance); break;
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.PIN.DateTimeReport.Instance, this); button1.Show(); return;

                case "Create Batch":
                    CommonFunctions.ShowPopUpForm(Master.BatchMaster.Instance, this); button1.Show(); return;
                case "Create School":
                    CommonFunctions.ShowPopUpForm(Master.School.ChaitanyaSchool.SchoolMaster.Instance, this); return;
                case "Create Student":
                    CommonFunctions.ShowPopUpForm(Master.School.ChaitanyaSchool.StudentMaster.Instance, this); return;
                case "Create Reader":
                    CommonFunctions.ShowPopUpForm(Master.ReaderMaster.Instance, this); button1.Show(); return;
                case "Create Break":
                    CommonFunctions.ShowPopUpForm(Master.BreakMaster.Instance, this); button1.Show(); return;
                case "Create Standard":
                    CommonFunctions.ShowPopUpForm(Master.School.ChaitanyaSchool.StandardMaster.Instance, this); return;
                case "Create Section":
                    CommonFunctions.ShowPopUpForm(Master.SectionMaster.Instance, this); button1.Show(); return;
                case "Create Machine":
                    CommonFunctions.ShowPopUpForm(Master.MachineMaster.Instance, this); button1.Show(); return;
                case "Create Block":
                    CommonFunctions.ShowPopUpForm(Master.School.ChaitanyaSchool.BlockMaster.Instance, this); return;
                case "Create OrderPack":
                    CommonFunctions.ShowPopUpForm(Master.OrderPackMaster.Instance, this); button1.Show(); return;
                case "Create Voting":
                    CommonFunctions.ShowPopUpForm(Master.School.ChaitanyaSchool.VotingMaster.Instance, this); return;
                case "Create StyleCategory":
                    CommonFunctions.ShowPopUpForm(Master.StyleCategoryMaster.Instance, this); button1.Show(); return;
                case "Create ElectionPost":
                    CommonFunctions.ShowPopUpForm(Master.School.ChaitanyaSchool.ElectionPostMaster.Instance, this); button1.Show(); return;
                case "Create Laying":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Laying.Instance, this); button1.Show(); return;

                case "Create Patient":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.PatientMaster.Instance, this); button1.Show(); break;
                case "Create LabTest":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.LabTestMaster.Instance, this); button1.Show(); break;
                case "Create Hospital":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.HospitalMaster.Instance, this); button1.Show(); break;

                case "Create LabTestItem":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.LabTestItemMaster.Instance, this); button1.Show(); break;
                case "Create Group":
                    CommonFunctions.ShowPopUpForm(Master.GroupMaster.Instance, this); button1.Show(); return;
                case "Create Prescription Entry":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.PrescriptionEntry.Instance, this); button1.Show(); break;
                case "Create OP Registration":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.OPRegistration.Instance, this); button1.Show(); break;
                case "Create Operation":
                    CommonFunctions.ShowPopUpForm(Master.OperationMaster.Instance, this); button1.Show(); return;
                case "Create Process":
                    CommonFunctions.ShowPopUpForm(Master.ProcessMaster.Instance, this); button1.Show(); return;

                case "Create ProcessGroup":
                    CommonFunctions.ShowPopUpForm(Master.ProcessGroupMaster.Instance, this); button1.Show(); return;

                case "Create Doctor":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.DoctorMaster.Instance, this); button1.Show(); break;
                case "DepartmentMaster":
                    CommonFunctions.ShowPopUpForm(Master.DepartmentMaster.Instance, this); button1.Show(); return;
                case "Create Designation":
                    CommonFunctions.ShowPopUpForm(Master.DesignationMaster.Instance, this); button1.Show(); return;
                case "Create Hospital Department":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.HospitalDepartmentMaster.Instance, this); button1.Show(); return;
                case "Create Medicine":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.MedicineMaster.Instance, this); button1.Show(); return;

                // showform(Master.Hospital.DepartmentMaster.Instance); button1.Show(); break;
                case "Create LineTarget":
                    CommonFunctions.ShowPopUpForm(Master.LineTargetMaster.Instance, this); button1.Show(); return;
                case "Create Display":
                    CommonFunctions.ShowPopUpForm(Master.DisplayMaster.Instance, this); button1.Show(); return;
                case "Create Line":
                    CommonFunctions.ShowPopUpForm(Master.LineMaster.Instance, this); button1.Show(); return;
                case "Create LineGroup":
                    CommonFunctions.ShowPopUpForm(Master.LineGroupMaster.Instance, this); button1.Show(); return;
                case "Create HSN":
                    CommonFunctions.ShowPopUpForm(Master.HSNMaster.Instance, this); button1.Show(); return;
                case "Create CuttingTable":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.CuttingTableMaster.Instance, this); button1.Show(); return;
                case "Create ListView":
                    CommonFunctions.ShowPopUpForm(Master.Hospital.ListViewMaster.Instance, this); button1.Show(); return;
                case "Create StyleGroup":
                    CommonFunctions.ShowPopUpForm(Master.StyleGroupMaster.Instance, this); button1.Show(); return;
                case "Create BarcodeGenerate":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Lyla.BarcodeGenerate.Instance, this); button1.Show(); return;
                case "Create ProductionInOut":
                    CommonFunctions.ShowPopUpForm(Pinnacle.ReportFormate.Lyla.ProductionInOut .Instance, this); button1.Show(); return;
                case "Create CheckingInOut":
                    CommonFunctions.ShowPopUpForm(Pinnacle.ReportFormate.Lyla.CheckingInOut.Instance, this); button1.Show(); return;
                case "Create DefectInOut":
                    CommonFunctions.ShowPopUpForm(Pinnacle.ReportFormate.Lyla.DefectInOut.Instance, this); button1.Show(); return;

                case "Create CheckingEntry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Lyla.CheckingEntry.Instance, this); button1.Show(); return;

                case "Create Cutting":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Cutting.Instance, this); button1.Show(); return;
                case "Create Production":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Production.Instance, this); button1.Show(); return;
                case "Create ProductionEntry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Lyla.ProductionEntry.Instance, this); button1.Show(); return;
                case "Create DefectEntry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Lyla.DefectEntry.Instance, this); button1.Show(); return;
                case "Create Remarks":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Master.Lyla.RemarksMaster.Instance, this); button1.Show(); return;
                case "Create OrderEntry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.Lyla.OrderEntry.Instance, this); button1.Show(); return;
                case "Create PortTransaction":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.PortTransaction.Instance, this); button1.Show(); return;
                case "Create DeviceCommunication":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.DeviceCommunication.Instance, this); button1.Show(); return;
                case "Create Attendance":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.PIN.Attendance.Instance, this); button1.Show(); return;
                case "Create IPMaster":
                    CommonFunctions.ShowPopUpForm(Master.IPMaster.Instance, this); button1.Show(); return;
                case "Create IPEntry":
                    CommonFunctions.ShowPopUpForm(Master.IPEntry.Instance, this); button1.Show(); return;
                case "BundleGeneration":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.BundleGeneration.Instance, this); button1.Show(); return;
                //case "HolidayCategory":
                //    showform(Master.HolidayCategory.Instance); button1.Show(); return;
                case "Create HRPay Details":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.SKL.HRPayDetails.Instance, this); button1.Show(); return;
                case "Create NHDay Entry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.SKL.NHDayEntry.Instance, this); button1.Show(); return;
                case "Create Day Shuffle Entry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.SKL.DayShuffleEntry.Instance, this); button1.Show(); return;
                case "Create Buyer Shuffle Entry":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.SKL.BuyerShuffleEntry.Instance, this); button1.Show(); return;
                case "Create TimeCard Process":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.SKL.TimeCardProcessEntry.Instance, this); button1.Show(); return;
                case "Create TimeCard":
                    CommonFunctions.ShowPopUpForm(Pinnacle.Transactions.SKL.TimeCard.Instance, this); button1.Show(); return;

            }
        }

        private void MnuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.Text == "Masters" || e.ClickedItem.Text == "Transactions" || e.ClickedItem.Text == "TreeView" || e.ClickedItem.Text == "Reports" || e.ClickedItem.Text == "Registration")
            {
                return;
            }
            else
            {
                Class.Users.UserTime = 0;
                Class.Users.screen = true;Class.Users.ValidCheck = false;
                if (e.ClickedItem.Text.Length > 7)
                {
                     Class.Users.ScreenName = "";
                     sender = e.ClickedItem.Text;
                    Class.Users.ScreenName = e.ClickedItem.Text;
                    ChildClick(sender, e);
                    
                }
            }

        }

        private void PinnacleMdi_MouseEnter(object sender, EventArgs e)
        {
            this.BackColor = Color.Silver;
        }

        private void PinnacleMdi_MouseLeave(object sender, EventArgs e)
        {
            this.BackColor = Color.White;
          
        }
        int co = 0;

        private void Timer1_Tick(object sender, EventArgs e)
        {

            toolStripStatusLabel1.Text = "  Date  :" + System.DateTime.Now.ToShortDateString() + "        Time  : " + System.DateTime.Now.ToLongTimeString();
            
            if (co == 0)
            {
                lblMarquee1.ForeColor = Color.White;
                pictureBox1.Show(); Class.Users.UserTime += 1;
                co = 1;
            }
            else
            {
                lblMarquee1.ForeColor = Color.Silver;
                pictureBox1.Hide(); Class.Users.UserTime += 1;
                co = 0;
            }
            toolStripStatusLabel3.Text = Class.Users.UserTime.ToString();
            if (Class.Users.UserTime > Class.Users.LoginTime)
            {
                //string name = "";
                //name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
                //string up = "delete from asptblsessionmas  WHERE  osuser='" + name + "'";
                //Utility.ExecuteNonQuery(up);
                Application.Exit();
            }

        }
        private void PinnacleMdi_FormClosed(object sender, FormClosedEventArgs e)
        {
            //string name = "";
            //name = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1];
            //string up = "delete from asptblsessionmas  WHERE  osuser='" + name + "'";
            //Utility.ExecuteNonQuery(up);
            Application.Exit();         
        }


        private void Button4_Click(object sender, EventArgs e)
        {
            Application.Exit();
            button1.Hide();
        }


        private void RefreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            tabControl1.TabPages.Remove(Masters);
            tabControl1.TabPages.Remove(Transactions);
            tabControl1.TabPages.Remove(Registration);
            tabControl1.TabPages.Remove(TreeView);
            tabControl1.TabPages.Remove(Reports);
            tabControl1.TabPages.Remove(Approval);
            tabControl1.TabPages.Remove(Approval2);
            PinnacleMdi_Load(sender,e);
            Backcolor3_Click(sender, e);
            companylogo(Class.Users.HCompcode);
    
        }




        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void buntreepanel_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            if (paneltree.Visible == true)
            {
                paneltree.Visible = false;
            }
            else
            {
                paneltree.Visible = true;
            }
        }

        private void btnpanel_Click(object sender, EventArgs e)
        {
            paneltree.Visible = false;
        }

    

        private void btnrefresh_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            tabControl1.TabPages.Remove(Masters);
            tabControl1.TabPages.Remove(Transactions);
            tabControl1.TabPages.Remove(Registration);
            tabControl1.TabPages.Remove(TreeView);
            tabControl1.TabPages.Remove(Reports);
            tabControl1.TabPages.Remove(Approval);
            tabControl1.TabPages.Remove(Approval2);
            PinnacleMdi_Load(sender, e);
        }

        private void buttreeclose_Click(object sender, EventArgs e)
        {
            paneltree.Visible = false; 
            
        }

        private void toolStriptreeopen_Click(object sender, EventArgs e)
        {
            paneltree.Visible = true; 
            
                


           
           
        }

       
        private bool NodeFiltering(TreeNode Nodo, string Texto)
        {
            bool resultado = false;

            if (Nodo.Nodes.Count == 0)
            {
                if (Nodo.Text.Substring(0,1).ToUpper().Contains(Texto.Substring(0,1).ToUpper()))
                {
                    resultado = true;
                }
                else
                {
                    Nodo.Remove();
                }
            }
            else
            {
                for (int i = Nodo.Nodes.Count; i > 0; i--)
                {
                    if (NodeFiltering(Nodo.Nodes[i - 1], Texto))
                        resultado = true;
                }

                if (!resultado)
                    Nodo.Remove();
            }

            return resultado;
        }
        internal class ListItem
        {
        }

        private void buttreesearch_Click(object sender, EventArgs e)
        {

        }

        private void txtpanelsearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
           
        }
       
        private void butprevious_Click(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;

            try
            {
                listfilter = uniqueList.Distinct().ToList();
               // if (listfilter.Count > 1)
                //{
                    dtlist = ConvertListToDataTable();
                    DataTable dt = dtlist.Copy();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["MenuName"].ToString() == screen1)
                        {
                            listcount++;
                            break;
                        }
                    }
                    if (dt.Rows.Count > listcount)
                    {
                        sender = dt.Rows[listcount]["MenuName"].ToString();
                        screen1 = dt.Rows[listcount]["MenuName"].ToString();                       
                        ChildClick(sender, e);
                    }
                    else
                    {
                        MessageBox.Show("No Screen Available in the Index  :" + listcount.ToString(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void butnext_Click(object sender, EventArgs e)
        {
            try
            {
                listfilter = uniqueList.Distinct().ToList();
                // if (listfilter.Count > 1)
                // {
                Class.Users.UserTime = 0;
                dtlist = ConvertListToDataTable();
                    DataTable dt = dtlist.Copy();
                    if (listcount >= 1)
                    {
                        listcount--;
                        sender = dt.Rows[listcount]["MenuName"].ToString();
                        screen1 = dt.Rows[listcount]["MenuName"].ToString();
                       
                        ChildClick(sender, e);
                    }
                    else
                    {
                        MessageBox.Show("No Screen Available in the Index  :" + listcount.ToString(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void News_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                ((ToolStripAccess)GlobalVariables.CurrentForm).News();
                Class.Users.UserTime = 0;
            }
        }

        private void Saves_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Saves();
            }
        }

        private void Deletes_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Deletes();
            }
        }

        private void Searchs_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Searchs();
            }
        }

        private void Prints_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Prints();
            }
        }

        private void Imports_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Imports();
            }
        }

        private void Pdfs_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Pdfs();
            }
        }

        private void DownLoads_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).DownLoads();
            }
        }

        private void ChangeSkins_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).ChangeSkins();
            }
        }

        private void Logins_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Logins();
            }
        }

        private void ChangePasswords_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).ChangePasswords();
            }
        }

        private void GlobalSearchs_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).GlobalSearchs();
            }
        }

        private void TreeButtons_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).TreeButtons();
            }
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).Exit();
                if (GlobalVariables.TabCtrl.TabPages.Count <= 0)
                {

                    GlobalVariables.TabCtrl.Visible = false;
                    GlobalVariables.Toolstrip1.Visible = false;
                }
                else
                {
                    GlobalVariables.TabCtrl.Visible = true;
                    GlobalVariables.Toolstrip1.Visible = true;
                    ((ToolStripAccess)GlobalVariables.CurrentForm).News();
                }

            }
            //if (Class.Users.ColorID == "Backcolor1_Click")
            //{
            //    Backcolor1_Click(sender, e);
            //}
            //if (Class.Users.ColorID == "Backcolor2_Click")
            //{
            //    Backcolor2_Click(sender, e);
            //}
            //if (Class.Users.ColorID == "Backcolor3_Click")
            //{
            //    Backcolor3_Click(sender, e);
            //}
            //if (Class.Users.ColorID == "Backcolor4_Click")
            //{
            //    Backcolor4_Click(sender, e);
            //}
            //if (Class.Users.ColorID == "Backcolor5_Click")
            //{
            //    Backcolor5_Click(sender, e);
            //}
        }

      

        private void ReadOnlys_Click(object sender, EventArgs e)
        {
            if (GlobalVariables.CurrentForm != null)
            {
                Class.Users.UserTime = 0;
                ((ToolStripAccess)GlobalVariables.CurrentForm).ReadOnlys();
            }
        }

        private void combosearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combosearch.SelectedIndex >=1 )
            {
                Class.Users.UserTime = 0;
                sender = combosearch.Text;
             
                ChildClick(sender, e);
                combosearch.Select();
                combosearch.SelectedIndex = -1;
                
                
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (paneltree.Visible == true) { paneltree.Visible = false;
                paneltree.Width = 0; }
            else
            {
                paneltree.Width = 294;
                paneltree.Visible = true;
            }
            Class.Users.UserTime = 0;
        }

        void ToolStripAccess.News()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.Saves()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.Prints()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.Searchs()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.Deletes()
        {
           // throw new NotImplementedException();
        }

        void ToolStripAccess.ReadOnlys()
        {
           // throw new NotImplementedException();
        }

        void ToolStripAccess.Imports()
        {
           // throw new NotImplementedException();
        }

        void ToolStripAccess.Pdfs()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.ChangePasswords()
        {
          //  throw new NotImplementedException();
        }

        void ToolStripAccess.DownLoads()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.ChangeSkins()
        {
           // throw new NotImplementedException();
        }

        void ToolStripAccess.Logins()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.GlobalSearchs()
        {
           // throw new NotImplementedException();
        }

        void ToolStripAccess.TreeButtons()
        {
            //throw new NotImplementedException();
        }

        void ToolStripAccess.Exit()
        {
            Class.Users.UserTime = 0;
        }

        public void GridLoad()
        {
            //throw new NotImplementedException();
        }

        private void Gbuttonok_Click(object sender, EventArgs e)
        {
           
            //string sel = "select distinct gtcompmastid,  compcode,compname,username,userid,gatename,PASWORD from (  select distinct A.gtcompmastid,  A.compcode,A.compname,c.username,C.USERID,C.GATENAME,C.PASWORD from  gtcompmast A   join asptblusermas c on c.compcode = a.gtcompmastid  where c.active='T'  and A.COMPCODE='AGF'    UNION ALL    select distinct B.gtcompmastid,  b.compcode,b.compname,c.username,C.USERID,C.GATENAME,C.PASWORD from  gtcompmast b    join asptblusermas c on c.compcode = b.gtcompmastid  where c.active='T'  and      B.COMPCODE='AGFM'     UNION ALL       select distinct A.gtcompmastid,  A.compcode,A.compname,c.username,C.USERID,C.GATENAME,C.PASWORD from  gtcompmast A    join asptblusermas c on c.compcode = a.gtcompmastid  where c.active='T'  and A.COMPCODE='AGFMGII'       UNION ALL        select distinct B.gtcompmastid,  b.compcode,b.compname,c.username,C.USERID,C.GATENAME,C.PASWORD from  gtcompmast b    join asptblusermas c on c.compcode = b.gtcompmastid  where c.active='T'  and B.COMPCODE='FLF'        UNION ALL           select distinct B.gtcompmastid,  b.compcode,b.compname,c.username,C.USERID,C.GATENAME,C.PASWORD from  gtcompmast b    join asptblusermas c on c.compcode = b.gtcompmastid  where c.active='T'  and B.COMPCODE='FLFD'       ) DUAL where compcode='" + GlobalCompcode1.Text + "' AND USERNAME='" + Class.Users.HUserName + "' AND PASWORD='" + Class.Users.PWORD + "'";
            //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
            //DataTable dt = ds.Tables["asptblusermas"];

            //if (dt.Rows.Count > 0)
            //{
            //    Class.Users.HUserName = dt.Rows[0]["username"].ToString();
            //    Class.Users.USERID = Convert.ToInt64(dt.Rows[0]["userid"].ToString());
            //    Class.Users.HGateName = System.DateTime.Now.Year + "/" + dt.Rows[0]["gatename"].ToString();
            //    Class.Users.HCompName = dt.Rows[0]["compname"].ToString();
            //    Class.Users.COMPCODE = Convert.ToInt64(dt.Rows[0]["gtcompmastid"].ToString());




            //    Globalpanel3.Visible = false;
            //    PinnacleMdi_Load(sender, e);
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new PinnacleMdi();
            form2.Closed += (s, args) => this.Close();
            LoginForm lf = new LoginForm();
            lf.Show();
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void GlobalCompcode1_SelectedIndexChanged(object sender, EventArgs e)
        {
           // Class.Users.HCompcode = GlobalCompcode1.Text;
        }

   

        private void colors1_Load(object sender, EventArgs e)
        {

            colors1.backcolor1.Click += Backcolor1_Click;
            colors1.backcolor2.Click += Backcolor2_Click;
            colors1.backcolor3.Click += Backcolor3_Click;
            colors1.backcolor4.Click += Backcolor4_Click;
            colors1.backcolor5.Click += Backcolor5_Click;
            Class.Users.UserTime = 0;
            
        }       
     

        private void Backcolor1_Click(object sender, EventArgs e)
        {
          
            combosearch.Font= colors1.backcolor1.Font;
            GlobalVariables.MenuStrip1.Font= colors1.backcolor1.Font;
            GlobalVariables.Toolstrip1.Font = colors1.backcolor1.Font;
            GlobalVariables.TabCtrl.Font = colors1.backcolor1.Font;
            butrightborder.BackColor = colors1.backcolor1.BackColor;
            Class.Users.FontName = colors1.backcolor1.Font;
            this.panelheader.BackColor = colors1.backcolor1.BackColor;
             this.statusStrip.BackColor = colors1.backcolor1.BackColor;
            Class.Users.ForeColors = colors1.backcolor1.ForeColor;
            Class.Users.BackColors = colors1.backcolor1.BackColor;
            this.butfooter.BackColor = colors1.backcolor1.BackColor;
            toolStrip1.Font = new System.Drawing.Font(colors1.backcolor1.Font.Name, colors1.backcolor1.Font.Size, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            lblMarquee1.ForeColor = colors1.backcolor1.ForeColor;
            lbldashboard.ForeColor = Class.Users.Color1;
            butwaitapproval.ForeColor = Class.Users.Color1;
            butwaitapproval.BackColor = Class.Users.BackColors;
            butmdapproval.ForeColor = Class.Users.Color1;
            butmdapproval.BackColor = Class.Users.BackColors;
            lblheader.ForeColor = colors1.backcolor1.ForeColor;
            toolStrip1.ForeColor = colors1.backcolor1.ForeColor;
            Class.Users.ColorID = "Backcolor1_Click";           
           
            News_Click(sender, e);

        }

        private void Backcolor2_Click(object sender, EventArgs e)
        {
         
            combosearch.Font = colors1.backcolor2.Font;
            GlobalVariables.MenuStrip1.Font = colors1.backcolor2.Font;
            GlobalVariables.Toolstrip1.Font = colors1.backcolor2.Font;
            GlobalVariables.TabCtrl.Font = colors1.backcolor2.Font;
            Class.Users.FontName = colors1.backcolor2.Font;
            butrightborder.BackColor = colors1.backcolor2.BackColor;
            this.panelheader.BackColor = colors1.backcolor2.BackColor;
            this.statusStrip.BackColor = colors1.backcolor2.BackColor;
            Class.Users.ForeColors = colors1.backcolor2.ForeColor;
            Class.Users.BackColors= colors1.backcolor2.BackColor;
            this.butfooter.BackColor = colors1.backcolor2.BackColor; 
            toolStrip1.Font = new System.Drawing.Font(colors1.backcolor2.Font.Name, colors1.backcolor2.Font.Size, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));


            toolStrip1.ForeColor = colors1.backcolor2.ForeColor;
            lblMarquee1.ForeColor = Class.Users.Color1;
            lblheader.ForeColor = colors1.backcolor2.ForeColor;
            lbldashboard.ForeColor = colors1.backcolor2.ForeColor;
            butwaitapproval.ForeColor = Class.Users.Color1;
            butwaitapproval.BackColor = Class.Users.BackColors;
            butmdapproval.ForeColor = Class.Users.Color1;
            butmdapproval.BackColor = Class.Users.BackColors;
            Class.Users.ColorID = "Backcolor2_Click";
           
            News_Click(sender, e);
        }
        private void Backcolor3_Click(object sender, EventArgs e)
        {
           
            combosearch.Font = colors1.backcolor3.Font;
            GlobalVariables.MenuStrip1.Font = colors1.backcolor3.Font;
            GlobalVariables.Toolstrip1.Font = colors1.backcolor3.Font;
            GlobalVariables.TabCtrl.Font = colors1.backcolor3.Font;
            butrightborder.BackColor = colors1.backcolor3.BackColor;
            Class.Users.FontName = colors1.backcolor3.Font;
            this.panelheader.BackColor = colors1.backcolor3.BackColor;
            this.statusStrip.BackColor = colors1.backcolor3.BackColor;
            Class.Users.BackColors = colors1.backcolor3.BackColor;
            this.butfooter.BackColor = colors1.backcolor3.BackColor;
            Class.Users.ForeColors = colors1.backcolor3.ForeColor;
            
            toolStrip1.ForeColor = colors1.backcolor3.ForeColor;
            toolStrip1.Font = new System.Drawing.Font(colors1.backcolor3.Font.Name, colors1.backcolor3.Font.Size, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            lblMarquee1.ForeColor = colors1.backcolor3.ForeColor;
            lblheader.ForeColor = colors1.backcolor3.ForeColor;
            lbldashboard.ForeColor = Class.Users.Color1;
          
            butwaitapproval.ForeColor = Class.Users.Color1;
            butwaitapproval.BackColor = Class.Users.BackColors;
            butmdapproval.ForeColor = Class.Users.Color1;
            butmdapproval.BackColor = Class.Users.BackColors;
            Class.Users.ColorID = "Backcolor3_Click";
            
            News_Click(sender, e);
        }
        private void Backcolor4_Click(object sender, EventArgs e)
        {
          
            combosearch.Font = colors1.backcolor4.Font;
            GlobalVariables.MenuStrip1.Font = colors1.backcolor4.Font;
            GlobalVariables.Toolstrip1.Font = colors1.backcolor4.Font;
            GlobalVariables.TabCtrl.Font = colors1.backcolor4.Font;
            butrightborder.BackColor = colors1.backcolor4.BackColor;
            Class.Users.FontName = colors1.backcolor4.Font;
            this.panelheader.BackColor = colors1.backcolor4.BackColor;
             this.statusStrip.BackColor = colors1.backcolor4.BackColor;
            Class.Users.BackColors = colors1.backcolor4.BackColor;
            this.butfooter.BackColor = colors1.backcolor4.BackColor;
            Class.Users.ForeColors = colors1.backcolor4.ForeColor;
          
            toolStrip1.ForeColor = colors1.backcolor4.ForeColor;
            toolStrip1.Font = new System.Drawing.Font(colors1.backcolor4.Font.Name, colors1.backcolor4.Font.Size, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            lblMarquee1.ForeColor = colors1.backcolor5.ForeColor;
            lblheader.ForeColor = colors1.backcolor4.ForeColor;
            lbldashboard.ForeColor = Class.Users.Color1;
            butwaitapproval.ForeColor = Class.Users.Color1;
            butwaitapproval.BackColor = Class.Users.BackColors;
            butmdapproval.ForeColor = Class.Users.Color1;
            butmdapproval.BackColor = Class.Users.BackColors;
            Class.Users.ColorID = "Backcolor4_Click";
         
            News_Click(sender, e);
        }
        private void Backcolor5_Click(object sender, EventArgs e)
        {
          
            GlobalVariables.MenuStrip1.Font = colors1.backcolor5.Font;
            GlobalVariables.Toolstrip1.Font = colors1.backcolor5.Font;
            GlobalVariables.TabCtrl.Font = colors1.backcolor5.Font;
            combosearch.Font = colors1.backcolor5.Font;
            Class.Users.FontName = colors1.backcolor5.Font;
          
            butrightborder.BackColor = colors1.backcolor5.BackColor;
            this.panelheader.BackColor = colors1.backcolor5.BackColor;
            this.statusStrip.BackColor = colors1.backcolor5.BackColor;
            Class.Users.BackColors = colors1.backcolor5.BackColor;
            this.butfooter.BackColor = colors1.backcolor5.BackColor;
            Class.Users.ForeColors = colors1.backcolor5.ForeColor;
   
            toolStrip1.Font = new System.Drawing.Font(colors1.backcolor5.Font.Name, colors1.backcolor5.Font.Size, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));

            toolStrip1.ForeColor = colors1.backcolor5.ForeColor;
            lblMarquee1.ForeColor = colors1.backcolor5.ForeColor;
            lblheader.ForeColor = colors1.backcolor5.ForeColor;
            lbldashboard.ForeColor = Class.Users.Color1;
            butwaitapproval.ForeColor = Class.Users.Color1;
            butwaitapproval.BackColor = Class.Users.BackColors;
            butmdapproval.ForeColor = Class.Users.Color1;
            butmdapproval.BackColor = Class.Users.BackColors;
            Class.Users.ColorID = "Backcolor5_Click";
          
            News_Click(sender, e);
        }

        private void TabCtrl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TabCtrl.TabCount > 0)
            {
                GlobalVariables.CurrentForm = TabCtrl.SelectedTab.Controls.OfType<Form>().First();
                Class.Users.UserTime = 0; 
                string[] s;
                s = GlobalVariables.CurrentForm.ToString().Split(','); 
                Class.Users.ScreenName = s[1].Substring(7).TrimEnd();
                usercheck(Class.Users.HCompcode,Class.Users.HUserName,Class.Users.ScreenName);
                GlobalVariables.HeaderName.Text = TabCtrl.SelectedTab.Text;
                Class.Users.ScreenName = TabCtrl.SelectedTab.Text; 
                GlobalVariables.TabCtrl.Font= Class.Users.FontName;
               
            }
        }

        void ToolStripAccess.Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

 

        private void mastersToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

       
        private void nHDayEntryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Class.Users.UserTime = 0;
         
        }

        private void timer2_Tick(object sender, EventArgs e)
         {
            if (lblMarquee1.Left < 0 && (Math.Abs(lblMarquee1.Left) > lblMarquee1.Width))
                lblMarquee1.Left = panel1.Width;
            lblMarquee1.Left -= 5;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            Class.Users.Description = "Approval";
            //if (dataGridView1.Rows.Count <= 1)
            //{
                
            //    butwaitapproval.Text = "No Content";
            //}
            //else
            //{
            //    butwaitapproval.Text = "No Content";
                pop2(0, dataGridView1);
            //}
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();

        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow dv = dataGridView1.Rows[e.RowIndex];
                GridID = Convert.ToInt32(dv.Cells[2].Value.ToString());
                if (dataGridView1.Columns[e.ColumnIndex].Name == "Column1" && dataGridView1.CurrentCell is DataGridViewCheckBoxCell)
                {
                    bool isChecked = (bool)dataGridView1[e.ColumnIndex, e.RowIndex].EditedFormattedValue;

                
                    if (isChecked == false)
                    {
                        Class.Users.Description = "Approval";
                        //DataGridViewRow dv = dataGridView1.Rows[e.RowIndex];
                        //GridID = Convert.ToInt32("0" + dv.Cells[2].Value.ToString());
                        remarks = Convert.ToString(dv.Cells[14].Value.ToString());
                        string update = "update asptbladvpaymas set approval='T',mdapproval='F',ManagerRemarks='" + remarks + "'  where asptbladvpaymasid='" + GridID + "'";
                        Utility.ExecuteNonQuery(update); GridID = 0;
                        pop2(0, dataGridView1);
                        Models.Master mas = new Models.Master();
                        mas.pop("Approved", remarks.ToString(), ""); remarks = "";
                    }
                }
                if (e.ColumnIndex == 13)
                {
                    Class.Users.Paramid = Convert.ToInt64("0" + GridID);
                    
                    Master.Bank.ReportPopUp pop = new Master.Bank.ReportPopUp();
                   
                    pop.Show();

                }
            }
            catch(Exception EX) { }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow dv = dataGridView1.Rows[e.RowIndex];
                GridID = Convert.ToInt32(dv.Cells[2].Value.ToString());
                remarks = Convert.ToString(dv.Cells[14].Value.ToString());
                Class.Users.RowIndex = e.RowIndex;
            }
            catch (Exception ex) { }
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow dv = dataGridView2.Rows[e.RowIndex];
                GridID = Convert.ToInt32(dv.Cells[2].Value.ToString());
                remarks = Convert.ToString(dv.Cells[15].Value.ToString());
                mdstatus = Convert.ToString(dv.Cells[16].Value.ToString());
                Class.Users.RowIndex = e.RowIndex;
            }
            catch (Exception ex) { }
        }



        Int32 GridID = 0;string remarks = "",mdstatus="";
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow dv = dataGridView2.Rows[e.RowIndex];
                GridID = Convert.ToInt32("0" + dv.Cells[2].Value.ToString());
                mdstatus = Convert.ToString(dv.Cells[16].Value.ToString());
                if (dataGridView2.Columns[e.ColumnIndex].Name == "Approval0" && dataGridView2.CurrentCell is DataGridViewCheckBoxCell)
                {
                    bool isChecked = (bool)dataGridView2[e.ColumnIndex, e.RowIndex].EditedFormattedValue;
                    if (isChecked == false)
                    {
                        Class.Users.Description = "Approval2";
                        Class.Users.RowIndex = e.RowIndex;
                        remarks = Convert.ToString(dv.Cells[15].Value.ToString());
                        string update = "update asptbladvpaymas set approval='T',mdapproval='T',MDRemarks='" + remarks + "'  where asptbladvpaymasid='" + GridID + "'";
                        Utility.ExecuteNonQuery(update);
                        Models.Master mas = new Models.Master();
                        mas.pop("Approved", remarks.ToString(), "");
                        GridID = 0; remarks = "";
                        pop2(0, dataGridView2);
                    }
                }
                if (dataGridView2.Columns[e.ColumnIndex].Name == "Notes" && mdstatus != "")
                {                    
                    Class.Users.Description = "Approval2";
                    mdstatus = Convert.ToString(dv.Cells[16].Value.ToString());
                }
               
                if (e.ColumnIndex == 13)
                {

                    Class.Users.Paramid = Convert.ToInt64("0" + GridID);
                    Master.Bank.ReportPopUp pop = new Master.Bank.ReportPopUp();
                    pop.Show();

                }
            }
            catch(Exception ex) { }

        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView2.Rows[e.RowIndex].Cells["SNO1"].Value = (e.RowIndex + 1).ToString();
        }



        private void butmdapproval_Click(object sender, EventArgs e)
        {
            Class.Users.Description = "Approval2";
            pop2(0, dataGridView2);
        }


        string filepath = "D:\\temp.pdf"; FileStream FS = null;
        byte[] stdbytes; byte[] stdbytes2; byte[] stdbytes3; byte[] stdbytes4; Int64 std;
        public void checkcellvalue(int index, DataGridView dgrid)
        {
         
            if (dgrid.CurrentCell.ColumnIndex.Equals(2) || dgrid.CurrentCell.ColumnIndex.Equals(13) && index != -1)
            {
                if (dgrid.CurrentCell != null)
                {

                    try
                    {


                
                            filepath = "";
                            filepath = "D:\\temp1.pdf";
                            FS = new FileStream(filepath, System.IO.FileMode.Create);
                            FS.Write(stdbytes, 0, stdbytes.Length);
                            FS.Close();
                            Class.Users.StaticByts = stdbytes;
                            Class.Users.Paramid = Convert.ToInt64("0" + GridID);
                            Master.Bank.PopUp pop = new Master.Bank.PopUp();
                            pop.Show();
                      
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.ToString());
                    }
                }
            }

        }

        private void approvalCancelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (GridID >= 1)//manager Approval Cancel
            {
                Class.Users.Description = "Approval";
                string update = "update asptbladvpaymas set approval='R',ManagerRemarks='"+remarks+"' where asptbladvpaymasid='" + GridID + "'";
                Utility.ExecuteNonQuery(update);                
                GridID = 0; Models.Master mas = new Models.Master();
                mas.pop("Rejected ",remarks.ToString(), "");
                pop2(0, dataGridView1); remarks = "";
            }
            else
            {
                GridID = 0; remarks = "";
            }
        }

        private void approvalCancelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (GridID >= 1)//MD Approval
            {
                Class.Users.Description = "Approval2";
                string update = "update asptbladvpaymas set MDRemarks='" + remarks + "',mdapproval='R' where asptbladvpaymasid='" + GridID + "'";
                Utility.ExecuteNonQuery(update);               
                Models.Master mas = new Models.Master();
                mas.pop("Rejected ",remarks.ToString(), "");
                pop2(0, dataGridView2);
                GridID = 0; remarks = "";
            }
            else
            {
                GridID = 0; remarks = "";
            }
        }

        private void correctionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (GridID >= 1 && mdstatus != "")//MD Approval
            {
                Class.Users.Description = "Approval2";
                string update = "update asptbladvpaymas set approval='F',  mdstatus='" + mdstatus + "' where asptbladvpaymasid='" + GridID + "'";
                Utility.ExecuteNonQuery(update);              
                Models.Master mas = new Models.Master();               
                mas.pop("Correction ", mdstatus.ToString(), ""); mdstatus = "";
                pop2(0, dataGridView2);
                GridID = 0;

            }
            else
            {
                GridID = 0; remarks = "";
            }
        }

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }

    

        
        private void statusToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }
    }
}

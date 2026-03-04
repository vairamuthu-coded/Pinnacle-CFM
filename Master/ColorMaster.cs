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
using Pinnacle.UserControls;


namespace Pinnacle.Master
{
    public partial class ColorMaster : Form,ToolStripAccess
    {
        private static ColorMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        Pinnacle.UserControls.UCCListView ucclist=new UCCListView();
        public static ColorMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ColorMaster();

                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public ColorMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());

        }
        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    //for (int r = 0; r < dt1.Rows.Count; r++)
                    //{

                    //    if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                    //    if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                    //    if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                    //    if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                    //    if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                    //    if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                    //    if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                    //    if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                    //    if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                    //    if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                    //    if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                    //    if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                    //    if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }

                    //}
                }


            }
            else
            {

            }

        }
        private void ColorMaster_Load(object sender, EventArgs e)
        {
            GridLoad();empty(); txtcolor.Select();
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
                if (txtcolor.Text != "")
                {
                    //Models.Validate va = new Models.Validate();
                    //if (va.IsString(txtcolor.Text) == true)
                    //{

                        string chk = "";
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        string sel = "select ASPTBLCOLMASID    from  ASPTBLCOLMAS    WHERE COLORNAME='" + txtcolor.Text + "' and active='" + chk + "' and compcode='" + Class.Users.COMPCODE + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLCOLMAS");
                        DataTable dt = ds.Tables["ASPTBLCOLMAS"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtcolor.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtcolorid.Text) == 0 || Convert.ToInt32("0" + txtcolorid.Text) == 0)
                        {
                            string ins = "insert into ASPTBLCOLMAS(COLORNAME,active,compcode,createdby,modifiedby,ipaddress)  VALUES('" + txtcolor.Text.ToUpper() + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )";
                            Utility.ExecuteNonQuery(ins);
                            MessageBox.Show("Record Saved Successfully " + txtcolor.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            string up = "update  ASPTBLCOLMAS  set   COLORNAME='" + txtcolor.Text.ToUpper() + "' , active='" + chk + "' ,compcode='" + Class.Users.COMPCODE + "',modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where ASPTBLCOLMASID='" + txtcolorid.Text + "'";
                            Utility.ExecuteNonQuery(up);
                            MessageBox.Show("Record Updated Successfully " + txtcolor.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }

                    //}
                    //else
                    //{
                    //    MessageBox.Show("'COLORNAME  is Wrong'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);


                    //}
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("COLORNAME " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void ColorMaster_FormClosed(object sender, FormClosedEventArgs e)
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
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
            txtcolorid.Text = "";
            txtcolor.Text = "";
            txtcolor.Select();
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.ASPTBLCOLMASID, a.COLORNAME as  COLORNAME , a.active  from  ASPTBLCOLMAS a   order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCOLMAS");
                DataTable dt = ds.Tables["ASPTBLCOLMAS"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = i.ToString();
                        list.SubItems.Add(myRow["ASPTBLCOLMASID"].ToString());
                        list.SubItems.Add(myRow["COLORNAME"].ToString());
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

                    txtcolorid.Text = listView1.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " select a.ASPTBLCOLMASID, a.COLORNAME as COLORNAME , a.active    from  ASPTBLCOLMAS a    where a.ASPTBLCOLMASID=" + txtcolorid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCOLMAS");
                    DataTable dt = ds.Tables["ASPTBLCOLMAS"];

                    if (dt.Rows.Count > 0)
                    {
                        txtcolorid.Text = Convert.ToString(dt.Rows[0]["ASPTBLCOLMASID"].ToString());
                        txtcolor.Text = Convert.ToString(dt.Rows[0]["COLORNAME"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Search()
        {
            try
            {
                if (txtsearch.Text.ToUpper() != "")
                {
                    listView1.Items.Clear(); int iGLCount = 1;
                    string sel1 = "  SELECT  a.ASPTBLCOLMASID,a.COLORNAME as COLORNAME,a.active from ASPTBLCOLMAS a  where a.COLORNAME LIKE'%" + txtsearch.Text.ToUpper() + "%' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCOLMAS");
                    DataTable dt = ds.Tables["ASPTBLCOLMAS"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["ASPTBLCOLMASID"].ToString());
                            list.SubItems.Add(myRow["COLORNAME"].ToString());
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
            if (txtcolorid.Text != "")
            {
                string sel1 = "select a.ASPTBLCOLMASID from ASPTBLCOLMAS a join asptblpurdet b on a.ASPTBLCOLMASID=b.COLORNAME where a.ASPTBLCOLMASID='" + txtcolorid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCOLMAS");
                DataTable dt = ds.Tables["ASPTBLCOLMAS"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtcolor.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from ASPTBLCOLMAS where ASPTBLCOLMASID='" + Convert.ToInt64("0" + txtcolorid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtcolor.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;
                if (txtsearch.Text.Length > 1)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[2].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

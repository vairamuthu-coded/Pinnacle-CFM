using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Pinnacle.Master.Bank
{
    public partial class IFSCMaster : Form,ToolStripAccess
    {
        private static IFSCMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
       
        public static IFSCMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new IFSCMaster();
                GlobalVariables.CurrentForm = _instance;
                _instance.Font = Class.Users.FontName;
                return _instance;
            }
        }
        public IFSCMaster()
        {
            InitializeComponent();
           


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

        private void IFSCMaster_Load(object sender, EventArgs e)
        {
          
            
        }
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        public void bankname()
        {
            sb.Clear();
            sb.Append("SELECT  a.asptblbanmasid,a.bankname from asptblbanmas a where a.active='T'");
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblbanmas");
            DataTable dt = ds.Tables["asptblbanmas"];
            combobank.DataSource = dt;
            combobank.DisplayMember = "bankname";
            combobank.ValueMember = "asptblbanmasid";
           
        }
        public DataTable findbankname(string s)
        {
            sb.Clear();
            sb.Append("SELECT distinct  a.asptblbanmasid,a.bankname from asptblbanmas a where a.active='T' and a.bankname='"+s+"' ");
            DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblbanmas");
            DataTable dt = ds.Tables["asptblbanmas"];          
            return dt;
        }
        public void Saves()
        {
            try
            {
                sb.Clear();
                if (dataGridView1.Rows.Count > 0 && Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
                {
                    int cc = 0, i = 0;
                    cc = dataGridView1.Rows.Count;
                    if (cc >= 0)
                    {
                        progressBar3.Minimum = 0; progressBar3.Refresh();
                        progressBar3.Maximum = dataGridView1.Rows.Count;

                        for (i = 0; i < cc; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value.ToString() != "")
                            {
                                Models.Bank.Ifsc p = new Models.Bank.Ifsc();
                                sb.Clear();
                                p.active = "T";
                                if (dataGridView1.Rows[i].Cells[0].Value.ToString() != "")
                                {
                                    DataTable dt0 = findbankname(Convert.ToString(dataGridView1.Rows[i].Cells[0].Value.ToString().ToUpper().Trim()));
                                    p.bankname = dt0.Rows[0]["asptblbanmasid"].ToString();
                                    p.ifsc = dataGridView1.Rows[i].Cells[2].Value.ToString().ToUpper().Trim();
                                    p.branch = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value.ToString().ToUpper().Trim());
                                }
                                p.createdby = Convert.ToString(Class.Users.CREATED);
                                p.createdon = Convert.ToString(Class.Users.CREATED);
                                p.modifiedby = Convert.ToString(Class.Users.HUserName);
                                p.ipaddress = Class.Users.IPADDRESS;

                                sb.Clear();
                                sb.Append("select asptblifscmasid    from  asptblifscmas     WHERE ifsc='" + p.ifsc + "' ");
                                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblifscmas");
                                DataTable dt = ds.Tables["asptblifscmas"];
                                if (dt.Rows.Count != 0)
                                {
                                }
                                else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + p.asptblifscmasid) == 0 || Convert.ToInt32("0" + p.asptblifscmasid) == 0)
                                {
                                    sb.Clear();
                                    sb.Append("insert into asptblifscmas(ifsc, branch,bankname,active,compcode,createdby,modifiedby,ipaddress)  VALUES('" + p.ifsc + "','" + p.branch + "','" + p.bankname + "','" + p.active + "','" + Class.Users.COMPCODE + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )");
                                    Utility.ExecuteNonQuery(sb.ToString());

                                }
                                else
                                {
                                    sb.Clear();
                                    sb.Append("update  asptblifscmas  set ifsc='" + p.ifsc + "', branch='" + p.branch + "',bankname='" + p.bankname + "' , active='" + p.active + "' ,compcode='" + Class.Users.COMPCODE + "', modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblifscmasid='" + p.asptblifscmasid + "'");
                                    Utility.ExecuteNonQuery(sb.ToString());

                                }

                            }
                        }
                        if (i == cc)
                        {
                            MessageBox.Show("Record Saved Successfully ", "Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty(); progressBar3.Value = 0; lblprogress3.Text = "";
                        }
                    }
                 
                }
                else
                {
                    if (txtifsccode.Text == "")
                    {
                        MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtifsccode.Focus();
                        return;
                    }
                    if (combobank.SelectedValue == null)
                    {
                        MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combobank.Focus();
                        return;
                    }

                    else
                    {


                        string chk = ""; sb.Clear();
                        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
                        sb.Append("select a.asptblifscmasid    from  asptblifscmas a    WHERE a.ifsc='" + txtifsccode.Text + "' and a.branch='" + txtbranch.Text + "' and a.bankname='" + combobank.SelectedValue + "' and a.active='" + chk + "' and a.asptblifscmasid='" + txtifscid.Text + "'");
                        DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblifscmas");
                        DataTable dt = ds.Tables["asptblifscmas"];
                        if (dt.Rows.Count != 0)
                        {
                            MessageBox.Show("Child Record Found " + " Alert " + txtifsccode.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
                        }
                        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtifscid.Text) == 0 || Convert.ToInt32("0" + txtifscid.Text) == 0)
                        {
                            sb.Clear();
                            sb.Append("insert into asptblifscmas(ifsc, branch,bankname,active,compcode,createdby,modifiedby,ipaddress)  VALUES('" + txtifsccode.Text.ToUpper().Trim() + "','" + txtbranch.Text.ToUpper().Trim() + "','" + combobank.SelectedValue + "','" + chk + "','" + Class.Users.COMPCODE + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )");
                            Utility.ExecuteNonQuery(sb.ToString());
                            MessageBox.Show("Record Saved Successfully " + txtifsccode.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            sb.Clear();
                            sb.Append("update  asptblifscmas  set ifsc='" + txtifsccode.Text.ToUpper().Trim() + "', branch='" + txtbranch.Text.ToUpper().Trim() + "',bankname='" + combobank.SelectedValue + "' , active='" + chk + "' ,compcode='" + Class.Users.COMPCODE + "', modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblifscmasid='" + txtifscid.Text + "'");
                            Utility.ExecuteNonQuery(sb.ToString());
                            MessageBox.Show("Record Updated Successfully " + txtifsccode.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad();
                            empty();
                        }


                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("IFSC " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void IFSCMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        public void News()
        {
            GridLoad(); bankname(); empty();

        }
        private void empty()
        {
            txtifscid.Text = "";
            txtifsccode.Text = ""; Class.Users.UserTime = 0;
            combobank.Text = ""; combobank.SelectedIndex = -1;  
            txtbranch.Text = "";
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            do
            {
                int i = 0;
                for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }

            }
            while (dataGridView1.Rows.Count > 0);
            combobank.Select();
        }
        public void GridLoad()
        {
            try
            {
                

                Class.Users.SearchQuery = "select a.asptblifscmasid  AS ID,a.IFSC, a.BRANCH, b.BANKNAME , a.ACTIVE    from  asptblifscmas a  join asptblbanmas b on a.bankname=b.asptblbanmasid   order by 1";
                Class.Users.HideCols = new string[] { "ID" };
                listView1.Load_Details();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Searchs(int EditID)
        {
            try
            {
                if (EditID > 0)
                {

                    txtifscid.Text = EditID.ToString(); Class.Users.UserTime = 0;
                    DataTable dt = Utility.SQLQuery("select a.asptblifscmasid,a.ifsc, a.branch , b.bankname  , a.active    from  asptblifscmas a  join asptblbanmas b on a.bankname=b.asptblbanmasid where a.asptblifscmasID=" + txtifscid.Text + "");
                    if (dt.Rows.Count > 0)
                    {
                        txtifscid.Text = Convert.ToString(dt.Rows[0]["asptblifscmasid"].ToString());
                        txtifsccode.Text = Convert.ToString(dt.Rows[0]["ifsc"].ToString());                       
                        txtbranch.Text = Convert.ToString(dt.Rows[0]["branch"].ToString());
                        combobank.Text = Convert.ToString(dt.Rows[0]["bankname"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //private void ListView1_ItemActivate(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        if (listView1.Items.Count > 0)
        //        {

        //            txtcityid.Text = listView1.SelectedItems[0].SubItems[2].Text;
        //            string sel1 = "select a.asptblifscmasid,a.ifsc, b.branch , c.bankname  , a.active    from  asptblifscmas a  join GTSTATEMAST b on a.state=b.GTSTATEMASTID   join GTCOUNTRYMAST c on b.country=c.GTCOUNTRYMASTID  where a.asptblifscmasid=" + txtcityid.Text;
        //            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblifscmas");
        //            DataTable dt = ds.Tables["asptblifscmas"];

        //            if (dt.Rows.Count > 0)
        //            {
        //                txtcityid.Text = Convert.ToString(dt.Rows[0]["asptblifscmasid"].ToString());
        //                txtcity.Text = Convert.ToString(dt.Rows[0]["ifsc"].ToString());
        //                combostate.Text = Convert.ToString(dt.Rows[0]["branch"].ToString());
        //                combocountry.Text = Convert.ToString(dt.Rows[0]["bankname"].ToString());
        //                if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = true; checkactive.Checked = false; }
        //                Combostate_SelectedIndexChanged(sender,e);


        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //    }
        //}

        //public void Txtsearch_TextChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        int item0 = 0;int i = 1;
        //        if (txtsearch.Text.Length > 0)
        //        {
        //            listView1.Items.Clear();
        //            foreach (ListViewItem item in listfilter.Items)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                if (item.SubItems[3].ToString().Contains(txtsearch.Text))
        //                {
        //                    list.Text = i.ToString();
        //                    list.SubItems.Add(item.SubItems[1].Text);
        //                    list.SubItems.Add(item.SubItems[2].Text);
        //                    list.SubItems.Add(item.SubItems[3].Text);
        //                    list.SubItems.Add(item.SubItems[4].Text);
        //                    list.SubItems.Add(item.SubItems[5].Text);
        //                    list.SubItems.Add(item.SubItems[6].Text);
        //                    list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

        //                    listView1.Items.Add(list);
        //                }
        //                item0++;
        //            }
        //        }
        //        else
        //        {

        //                listView1.Items.Clear(); item0 = 0;
        //            foreach (ListViewItem item in listfilter.Items)
        //            {
        //                ListViewItem list = new ListViewItem();
        //                list.Text = i.ToString();
        //                list.SubItems.Add(item.SubItems[1].Text);
        //                list.SubItems.Add(item.SubItems[2].Text);
        //                list.SubItems.Add(item.SubItems[3].Text);
        //                list.SubItems.Add(item.SubItems[4].Text);
        //                list.SubItems.Add(item.SubItems[5].Text);
        //                list.SubItems.Add(item.SubItems[6].Text);
        //                list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

        //                listView1.Items.Add(list);

        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // MessageBox.Show("---" + ex.ToString());
        //    }
        //}




        public void Deletes()
        {
            if (txtifscid.Text != "")
            {
                string sel1 = "select a.asptblifscmasid from asptblifscmas a join asptblpartymas b on a.asptblifscmasid=b.bankname where a.asptblifscmasid='" + txtifscid.Text + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblifscmas");
                DataTable dt = ds.Tables["asptblifscmas"];
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Child Record Found.Can Not Delete." + txtifsccode.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else
                {

                    string del = "delete from asptblifscmas where asptblifscmasid='" + Convert.ToInt64("0" + txtifscid.Text) + "'";
                    Utility.ExecuteNonQuery(del);
                    MessageBox.Show("Record Deleted Successfully " + txtifsccode.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
            }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

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
            if (Class.Users.Log >= Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd")))
            {
                Class.Users.UserTime = 0;
                string filePath = string.Empty; dataGridView1.AllowUserToAddRows = false;
                string fileExt = string.Empty;
                OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
                if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
                {
                    filePath = file.FileName; //get the path of the file  
                    fileExt = Path.GetExtension(filePath); //get the file extension  
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        try
                        {
                            DataTable dtExcel = new DataTable();
                            dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                            dataGridView1.Visible = true;
                            dataGridView1.DataSource = dtExcel;
                            CommonFunctions.SetRowNumber(dataGridView1);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                    }
                }

                int cnt = dataGridView1.Rows.Count - 1;
                label48.Text = "Total Count  :" + cnt.ToString();
            }
            else
            {
                MessageBox.Show("pls Contact your Administrator." + Class.Users.Log.ToString(), "Register Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Dispose();
            }
        }
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            //if (fileExt.CompareTo(".xls") == 0)
            //    path = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            //else
            //    path = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  


            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (System.Data.OleDb.OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    System.Data.OleDb.OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
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

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bankname();
        }

        private void txtbranch_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || (char.IsWhiteSpace(e.KeyChar)) || e.KeyChar == (char)Keys.Back);

        }

        private void txtifsccode_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || (char.IsWhiteSpace(e.KeyChar)) || e.KeyChar == (char)Keys.Back);
        }

        private void combobank_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtbranch.Focus();
        }

        private void txtbranch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtifsccode.Focus();
        }

        private void combobank_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtifsccode_TextChanged(object sender, EventArgs e)
        {
            if (txtifscid.Text=="" && txtifsccode.Text.Length >= 6)
            {
                sb.Clear();
                sb.Append("select asptblifscmasid    from  asptblifscmas     WHERE   ifsc='" + txtifsccode.Text.ToUpper().Trim() + "' ");
                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblifscmas");
                DataTable dt = ds.Tables["asptblifscmas"];
                if (dt.Rows.Count > 0)
                {
                    txtifsccode.Text = "";txtbranch.Text = "";txtbranch.Focus();
                    MessageBox.Show("Child Recrod Found. IFSC  : " + txtifsccode.Text);
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Transactions.SKL
{
    public partial class BuyerShuffleEntry : Form,ToolStripAccess
    {
       
        public BuyerShuffleEntry()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            combofinyear.Text = Class.Users.Finyear;
        }
        private static BuyerShuffleEntry _instance;
        public static BuyerShuffleEntry Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new BuyerShuffleEntry();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        decimal listview2totalweight = 0;
        Models.Validate va = new Models.Validate();
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        string griddelrow = ""; ListView listfilter = new ListView();
        private void BuyerShuffleEntry_Load(object sender, EventArgs e)
        {         

            tabControl1.SelectTab(tabPagedel2);
            txtsearch.Select(); dateTimePicker1.Value = System.DateTime.Now.AddDays(0);
        }
        public void companyload()
        {
            try
            {
                string sel = "select a.gtcompmastid,a.compcode, a.compname from  gtcompmast a  where a.ptransaction ='COMPANY'  and a.active='T' order by 2 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];

                combocompcode.DisplayMember = "compcode";
                combocompcode.ValueMember = "gtcompmastid";
                combocompcode.DataSource = dt;


                combocodereport.DisplayMember = "compcode";
                combocodereport.ValueMember = "gtcompmastid";

                combocodereport.DataSource = dt;
                combocompname.DisplayMember = "compname";
                combocompname.ValueMember = "gtcompmastid";
                combocompname.DataSource = dt;
              
            }
            catch (Exception ex)
            {
                MessageBox.Show("companyload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void companyloadListview()
        {
            try
            {
                string sel = "select distinct a.gtcompmastid,a.compcode, a.compname from  gtcompmast a join HrSufTime b on b.compcode=a.gtcompmastid  order by 2 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];


                combocompcode1.DisplayMember = "compcode";
                combocompcode1.ValueMember = "gtcompmastid";
                combocompcode1.DataSource = dt;


            }
            catch (Exception ex)
            {
                MessageBox.Show("companyload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        public void autonumberload()
        {
            try
            {
                string sel = "select max(a.HrSufTimeid1)+1 as id,b.compname,b.compcode from HrSufTime a join gtcompmast b on a.compcode=b.gtcompmastid  where b.ptransaction='COMPANY'  and b.compcode='" + combocompcode.Text + "' group by   b.compname,b.compcode; ";

                DataSet ds = Utility.ExecuteSelectQuery(sel, "HrSufTime");
                DataTable dt = ds.Tables["HrSufTime"];
                int cnt = dt.Rows.Count;
                if (dt.Rows[0]["id"].ToString() == "")
                {

                    string sel1 = "select b.gtcompmastid,b.compcode, b.compname from  gtcompmast b  where b.ptransaction='COMPANY'  and b.compcode='" + combocompcode.Text + "'; ";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "gtcompmast");
                    DataTable dt1 = ds1.Tables["gtcompmast"];
                    combocompcode.DisplayMember = "compcode";
                    combocompcode.ValueMember = "gtcompmastid";
                    combocompcode.DataSource = dt1;
                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt1;

                    txtdocid.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + 1;
                    txtbuyershuffleid1.Text = "1";

                }
                else
                {

                    txtdocid.Text = combocompcode.Text + "/" + Class.Users.Finyear + "/" + dt.Rows[0]["id"].ToString();
                    txtbuyershuffleid1.Text = dt.Rows[0]["id"].ToString();
                    combocompcode.DisplayMember = "compcode";
                    combocompcode.ValueMember = "gtcompmastid";
                    combocompcode.DataSource = dt;

                    combocompname.DisplayMember = "compname";
                    combocompname.ValueMember = "gtcompmastid";
                    combocompname.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("autonumberload1: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void empty()
        {
            txtbuyershuffleid.Text = ""; txtbuyershuffleid1.Text = ""; combocompcode.SelectedIndex = -1; combocompcode.Text = "";
            combocompname.SelectedIndex = -1; combocompname.Text = ""; combomonth.Text = ""; combomonth.SelectedIndex = -1;
            listfilter.Items.Clear();
            combofinyear.Text = Class.Users.Finyear; dateTimePicker1.Value = System.DateTime.Now;
            txtdocid.Text = "";
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.MenuHighlight;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.EnableHeadersVisualStyles = false;
            do
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView1.Rows.Count > 1);
            do
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++) { try { dataGridView2.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView2.Rows.Count > 1);
            if (listView1.Items.Count > 0)
            {
                listView1.Items[0].Selected = true;
            }
            this.Font = Class.Users.FontName;
            this.BackColor = Class.Users.BackColors;
            panel4.BackColor = Class.Users.BackColors;
            Class.Users.UserTime = 0;
            panel1.BackColor = Class.Users.BackColors;
            panel8.BackColor = Class.Users.BackColors;
            lbltotal.ForeColor = Class.Users.BackColors;
            lblsearch.ForeColor = Class.Users.Color1;
            label6.ForeColor = Class.Users.Color1;
            label7.ForeColor = Class.Users.Color1;
            listView1.Font = Class.Users.FontName;
            Report.BackColor=Class.Users.BackColors;
          
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
        }
        public void News()
        {
            empty();
            GridLoad();
            companyload(); companyloadListview();
            tabControl1.SelectTab(tabPagedel2);
           
        }

        public void Saves()
        {
            throw new NotImplementedException();
        }

        public void Prints()
        {
            throw new NotImplementedException();
        }

        public void Searchs()
        {
            tabControl1.SelectTab(tabPagedel2);
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void Deletes()
        {
            try
            {
                if (txtbuyershuffleid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {


                        string del1 = "delete from HrSufTime where compcode='" + combocompcode.SelectedValue + "' and  HrSufTimeid='" + txtbuyershuffleid.Text + "';";
                        Utility.ExecuteNonQuery(del1);
                        string del = "delete from HrSufTimeDet where compcode='" + combocompcode.SelectedValue + "' and  HrSufTimeDetid='" + txtbuyershuffleid.Text + "';";
                        Utility.ExecuteNonQuery(del);
                        MessageBox.Show("Record Deleted Successfully " + txtbuyershuffleid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletes_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPagedel1"])//your specific tabname
            {
                combocompcode.Select();

            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPagedel2"])//your specific tabname
            {
                txtsearch.Select();

            }
        }


        private void buttsearch_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.HrSufTimeid,a.docid,a.finyear,a.compcode,a.date,a.month,a.active from HrSufTime a where a.date=date_format('" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d') order by a.HrSufTimeid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "HrSufTimeDet");
                DataTable dt = ds.Tables["HrSufTimeDet"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["HrSufTimeid"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["finyear"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["date"].ToString());
                        list.SubItems.Add(myRow["month"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("gridload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void butgetall_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "select a.HrSufTimeid,a.docid,a.finyear,a.compcode,a.date,a.month,a.active from HrSufTime a   order by a.HrSufTimeid desc;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "HrSufTimeDet");
                DataTable dt = ds.Tables["HrSufTimeDet"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["HrSufTimeid"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["finyear"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["date"].ToString());
                        list.SubItems.Add(myRow["month"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("gridload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            

            if (txtbuyershuffleid.Text == "")
            {
                autonumberload();
            }
        }



        // on text change of dtp, assign back to cell


        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }


        private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        if (txtbuyershuffleid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow != "")
                                {
                                    string del1 = "delete from  HrSufTimeDet     Where  HrSufTimeDetid='" + griddelrow + "';";
                                    Utility.ExecuteNonQuery(del1);

                                    griddelrow = "";
                                }
                                else
                                {
                                    dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                                }
                            }
                        }
                        else
                        {
                            dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                        }
                    }
                }

            }
            catch (Exception EX)
            {
                // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            }

        }
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
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
            string filePath = string.Empty;
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
                        dataGridView2.Visible = true;
                        dataGridView2.DataSource = dtExcel;
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
            tabControl1.SelectTab(DownLoadData);
            label48.Text = "Total Record   :" + dataGridView2.Rows.Count.ToString();
        }

        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        if (txtbuyershuffleid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow != "")
                                {
                                    string del1 = "delete from  HrSufTimeDet     Where  HrSufTimeDetid='" + griddelrow + "';";
                                    Utility.ExecuteNonQuery(del1);
                                    dataGridView1.Rows.RemoveAt(Convert.ToInt32(griddelrow));
                                    griddelrow = "";
                                }
                                else
                                {
                                    dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                                }
                            }
                        }
                        else
                        {
                            dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                        }
                    }
                }

            }
            catch (Exception EX)
            {
                // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (txtbuyershuffleid.Text != "")
            {
                griddelrow = "";
                griddelrow = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();

            }
        }

        private void refreshToolStripMenuItem_Click_2(object sender, EventArgs e)
        {
            companyload();
        }

        private void combocompcode1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear(); //date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')
                string sel1 = "select a.HrSufTimeid,a.docid,a.finyear,c.compcode,date_format(a.date,'%d-%m-%Y') as date,a.month, a.active from HrSufTime a  join gtcompmast c on a.compname=c.gtcompmastid  where c.compcode='" + combocompcode1.Text + "' order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "HrSufTime");
                DataTable dt = ds.Tables["HrSufTime"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["HrSufTimeid"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["finyear"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["date"].ToString());
                        list.SubItems.Add(myRow["month"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("gridload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void button2_Click(object sender, EventArgs e)
    {
        string sel2 = " select c.compcode, b.ddate,b.suftime,a.month,b.notes from HrSufTime a join HrSufTimeDet b on a.HrSufTimeid=b.HrSufTimeid join gtcompmast c on c.gtcompmastid=a.compcode where c.compcode='" + combocodereport.Text + "';";//where fromdate between '" + fromdatedateTimePicker1.Value.ToString("dd-MM-yyyy") + "' and '" + todatedateTimePicker2.Value.ToString("dd-MM-yyyy") + "'
        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "HrSufTimeDet");
        DataTable dt2 = ds2.Tables["HrSufTimeDet"];
        dataGridView3.DataSource = dt2;
    }

    private void button1_Click(object sender, EventArgs e)
    {
        if (dataGridView3.Rows.Count > 0)
        {
            Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
            xcelApp.Application.Workbooks.Add(Type.Missing);
            for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
            {
                xcelApp.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
            }

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {

                for (int j = 0; j < dataGridView3.Columns.Count; j++)
                {
                    xcelApp.Cells[i + 2, j + 1] = Convert.ToString(dataGridView3.Rows[i].Cells[j].Value);
                }
            }
            xcelApp.Columns.AutoFit();
            xcelApp.Visible = true;
        }
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
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                // empty();
                if (listView1.Items.Count > 0)
                {

                    txtbuyershuffleid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select a.HrSufTimeid, a.HrSufTimeid1,a.finyear,c.gtcompmastid compcode,c.compname,a.docid,a.date,a.month, a.active from HrSufTime a  join gtcompmast c on a.compname=c.gtcompmastid  where a.HrSufTimeid='" + txtbuyershuffleid.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HrSufTime");
                    DataTable dt = ds.Tables["HrSufTime"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtbuyershuffleid.Text = Convert.ToString(dt.Rows[0]["HrSufTimeid"].ToString());
                        txtbuyershuffleid1.Text = Convert.ToString(dt.Rows[0]["HrSufTimeid1"].ToString());
                        combofinyear.Text = Convert.ToString(dt.Rows[0]["finyear"].ToString());
                        combocompcode.SelectedValue = Convert.ToString(dt.Rows[0]["compcode"].ToString());
                        combocompname.Text = Convert.ToString(dt.Rows[0]["compname"].ToString());
                        txtdocid.Text = Convert.ToString(dt.Rows[0]["docid"].ToString());
                        dateTimePicker1.Text = Convert.ToString(dt.Rows[0]["date"].ToString());
                        combomonth.Text = Convert.ToString(dt.Rows[0]["month"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        string sel2 = "select a.HrSufTimeDetid, a.HrSufTimeid,a.HrSufTimeid1,a.compcode,a.ddate,a.suftime,a.notes from HrSufTimeDet a join HrSufTime b on a.HrSufTimeid = b.HrSufTimeid join gtcompmast d on a.compcode = d.gtcompmastid where a.HrSufTimeid='" + txtbuyershuffleid.Text + "';";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "HrSufTimeDet");
                        DataTable dt1 = ds2.Tables["HrSufTimeDet"];
                        if (dt1.Rows.Count > 0)
                        {
                            dataGridView1.DataSource = dt1;

                            for (i = 0; i < dt1.Rows.Count; i++)
                            {

                                if (Convert.ToInt64(dataGridView1.Rows[i].Cells[1].Value) > 0)
                                {
                                    dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["HrSufTimeDetid"].ToString());
                                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["HrSufTimeid"].ToString());
                                    dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["HrSufTimeid1"].ToString());
                                    dataGridView1.Rows[i].Cells[4].Value = Convert.ToInt64("0" + dt1.Rows[i]["compcode"].ToString());
                                    dataGridView1.Rows[i].Cells[5].Value = Convert.ToString(dt1.Rows[i]["ddate"].ToString());
                                    dataGridView1.Rows[i].Cells[6].Value = Convert.ToString(dt1.Rows[i]["suftime"].ToString());
                                    dataGridView1.Rows[i].Cells[7].Value = dt1.Rows[i]["notes"].ToString();
                                }
                            }
                            CommonFunctions.SetRowNumber(dataGridView1);

                        }
                        tabControl1.SelectTab(tabPagedel1);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ListView1_ItemActivate: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {

            try
            {
                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length >= 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[5].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[7].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

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
                //MessageBox.Show("---" + ex.ToString());
            }


        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear(); //date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')
                string sel1 = "select a.HrSufTimeid,a.docid,a.finyear,c.compcode,date_format(a.date,'%d-%m-%Y') as date,a.month, a.active from HrSufTime a  join gtcompmast c on a.compname=c.gtcompmastid  order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "HrSufTime");
                DataTable dt = ds.Tables["HrSufTime"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["HrSufTimeid"].ToString());
                        list.SubItems.Add(myRow["docid"].ToString());
                        list.SubItems.Add(myRow["finyear"].ToString());
                        list.SubItems.Add(myRow["compcode"].ToString());
                        list.SubItems.Add(myRow["date"].ToString());
                        list.SubItems.Add(myRow["month"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("gridload: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }
    }
}

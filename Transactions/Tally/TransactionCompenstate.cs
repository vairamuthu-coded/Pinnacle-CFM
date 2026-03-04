using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Pinnacle.Transactions.Tally
{
    public partial class TransactionCompenstate : Form, ToolStripAccess
    {
        private static TransactionCompenstate _instance;
        ListView listfilterred = new ListView();
        ListView listfilterscreen = new ListView();
        ListView listfilterred1 = new ListView();
        ListView listfilterscreen1 = new ListView();
        ListView listfilterdb = new ListView(); ListView listfilterslug = new ListView(); ListView listfilterexcel = new ListView();
        ListView tofilter = new ListView(); ListView allip3 = new ListView(); ListView allislug = new ListView(); ListView allip4 = new ListView();
        ListView COLUMNORDER = new ListView();

        int i = 0; int j = 1; string update1 = null;
        Models.Tally.TransactionCompenstate daimport = new Models.Tally.TransactionCompenstate();
        string Details = ""; string Details1 = "";
        string matchfield = "", matchfield1 = "";
        public TransactionCompenstate()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            string MODIFIED_BY = ", MODIFIED_BY ";
            string MODIFIED_ON = " MODIFIED_ON  ";
            string CREATED_BY = " CREATED_BY  ";
            string CREATED_ON = " CREATED_ON  ";
            string USERID = " USERID  ";
            string PROJECTID = " PROJECTID ";
            string IPADD = " IPADD ";

            string MODIFIED_BY1 = ",'" + Class.Users.HUserName + "'";
            string MODIFIED_ON1 = "TO_DATE('" + System.DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "','DD-MM-YYYY hh:mi:ss')";
            string CREATED_BY1 = "'" + Class.Users.HUserName + "'";
            string CREATED_ON1 = "TO_DATE('" + System.DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "','DD-MM-YYYY hh:mi:ss')";
            string USERID1 = "'" + Class.Users.USERID + "',";
            string PROJECTID1 = "'" + Class.Users.ProjectID + "',";
            string IPADD1 = "'" + Class.Users.IPADDRESS + "'";
            Details = MODIFIED_BY + "," + MODIFIED_ON + "," + CREATED_BY + "," + CREATED_ON + "," + USERID + "," + PROJECTID + "," + IPADD;
            Details1 = MODIFIED_BY1 + "," + MODIFIED_ON1 + "," + CREATED_BY1 + "," + CREATED_ON1 + "," + USERID1 + "" + PROJECTID1 + "" + IPADD1;
            GlobalVariables.DownLoads.Text = "UpLoad";
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
                case Keys.I | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    Imports();
                    return true; // signal that we've processed this key
                case Keys.D | Keys.Control:
                    // ... Process Shift+Ctrl+Alt+B ...
                    DownLoads();
                    return true; // signal that we've processed this key

            }
            // run base implementation
            return base.ProcessCmdKey(ref message, keys);
        }
        public static TransactionCompenstate Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new TransactionCompenstate();
                GlobalVariables.CurrentForm = _instance;
                return _instance;

            }
        }
        private void TranactionCompenstate_Load(object sender, EventArgs e)
        {


        }
        public void TableGridLoadd(string s)
        {
            try
            {
                if (s != "" || s == null)
                {

                    daimport.query = null;
                    dttbl2 = null;
                    daimport.query = "select * from " + Class.Users.ProjectID + "." + s + " ORDER BY 1";
                    daimport.ds = Utility.ExecuteSelectQuery(daimport.query, s);
                    dttbl2 = daimport.ds.Tables[s];
                    DBColumn.Items.Clear(); int j = 0; matchfield = ""; fromquery = "";
                    if (dttbl2.Columns.Count > 0)
                    {

                        i = 1;
                        for (j = 0; j < dttbl2.Columns.Count; j++)
                        {

                            if (DBColumn.Items.Contains(dttbl2.Columns[j].ToString()))
                            {
                                MessageBox.Show(dttbl2.Columns[j].ToString());
                            }
                            else
                            {
                                DBColumn.Items.Add(dttbl2.Columns[j].ToString());
                                if (dttbl2.Columns.Count == i)
                                {
                                    matchfield += "a." + dttbl2.Columns[j].ToString() + "";
                                    fromquery += "x." + dttbl2.Columns[j].ToString() + "";

                                }
                                if (dttbl2.Columns.Count != i)
                                {
                                    matchfield += "a." + dttbl2.Columns[j].ToString() + ",";
                                    fromquery += "x." + dttbl2.Columns[j].ToString() + ",";
                                }
                                i++;
                            }
                        }
                    }
                    else
                    {
                        DBColumn.Items.Clear(); matchfield = "";
                        i = 1;
                        for (j = 0; j < dttbl2.Columns.Count; j++)
                        {

                            if (DBColumn.Items.Contains(dttbl2.Columns[j].ToString()))
                            {
                                MessageBox.Show(dttbl2.Columns[j].ToString());
                            }
                            else
                            {
                                DBColumn.Items.Add(dttbl2.Columns[j].ToString());

                                if (dttbl2.Columns.Count == i)
                                {
                                    matchfield += "a." + dttbl2.Columns[j].ToString() + "";

                                }
                                if (dttbl2.Columns.Count != i)
                                {
                                    matchfield += "a." + dttbl2.Columns[j].ToString() + ",";
                                }
                                i++;
                            }
                            matchfield += "a.refno,";
                        }
                    }
                }
                ///dbcol = false;
            }
            catch (Exception EX) { }
        }
        public void TableGridLoad(string s)
        {
            try
            {
                if (s != "" || s == null)
                {
                    dttbl2.Columns.Clear();
                    dttbl2.Rows.Clear();
                    daimport.query = null;
                    daimport.query = "select * from " + Class.Users.ProjectID + "." + s + " ORDER BY 1";
                    daimport.ds = Utility.ExecuteSelectQuery(daimport.query, s);
                    dttbl2 = daimport.ds.Tables[s];
                    if (dttbl2 != null)
                    {
                        TableColumn.Items.Clear(); int j = 0; matchfield1 = "";

                        i = 1; matchfield1 = ""; toquery = "";
                        for (j = 0; j < dttbl2.Columns.Count; j++)
                        {

                            if (TableColumn.Items.Contains(dttbl2.Columns[j].ToString()))
                            {
                                MessageBox.Show(dttbl2.Columns[j].ToString());
                            }
                            else
                            {
                                TableColumn.Items.Add(dttbl2.Columns[j].ToString());
                                if (dttbl2.Columns.Count == i)
                                {
                                    matchfield1 += "a." + dttbl2.Columns[j].ToString() + "";
                                    toquery += "x." + dttbl2.Columns[j].ToString() + "";

                                }
                                if (dttbl2.Columns.Count != i)
                                {
                                    matchfield1 += "a." + dttbl2.Columns[j].ToString() + ",";
                                    toquery += "x." + dttbl2.Columns[j].ToString() + ",";
                                }
                                i++;
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("No Data Found in Customer / Supplier ");
                    }

                }

            }
            catch (Exception EX) { }
        }

        public void GridLoad()
        {
            try
            {
                listfilterdb.Items.Clear(); int r = 1, i = 0, j = 0; ;

                string sel2 = "show  tables  where Tables_in_" + Class.Users.ProjectID + " like'%my_%' ";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "my_");
                DataTable dt2 = ds2.Tables["my_"];
                if (dt2 != null)
                {

                    string sel4 = "select * from  " + dt2.Rows[0]["Tables_in_bank"].ToString() + " ";
                    DataSet ds4 = Utility.ExecuteSelectQuery(sel4, dt2.Rows[0]["Tables_in_bank"].ToString());
                    DataTable dt4 = ds4.Tables[dt2.Rows[0]["Tables_in_bank"].ToString()];
                    FromTable1.Items.Clear();
                    for (j = 0; j < dt2.Rows.Count; j++)
                    {
                        if (dt2.Rows[j]["Tables_in_" + Class.Users.ProjectID + ""].ToString().Substring(0, 3) == "my_")
                        {
                            FromTable1.Items.AddRange(dt2.Rows[j]["Tables_in_" + Class.Users.ProjectID + ""].ToString().Remove(0, 3));
                        }
                    }
                    int k = 1;
                    if (dataGridView2.Rows.Count <= 1 && dt4.Rows.Count > 0)
                    {
                        for (j = 0; j < dt4.Columns.Count; j++)
                        {

                            dataGridView2.Rows.Add();
                            dataGridView2.Rows[j].Cells[0].Value = k.ToString();
                            dataGridView2.Rows[j].Cells[1].Value = dt4.Columns[j].ToString();
                            k++;
                        }
                        CommonFunctions.SetRowNumber(dataGridView2);
                    }
                }
                else { MessageBox.Show("pls UPLoad MyCompany Data "); }

                string sel1 = "show  tables  where Tables_in_" + Class.Users.ProjectID + " like'%to_%' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "to_" + buttosupplier.Text.ToLower().Trim());
                DataTable dt = ds.Tables[buttosupplier.Text.ToLower().Trim()];
                if (dt != null)
                {
                    string sel3 = "select * from " + dt2.Rows[dt2.Rows.Count]["Tables_in_" + Class.Users.ProjectID + ""].ToString() + " order by 1";
                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, buttosupplier.Text.ToLower().Trim());
                    DataTable dt3 = ds3.Tables[buttosupplier.Text.ToLower().Trim()];

                    if (dt3.Rows.Count > 0)
                    {
                        TableName.Items.Clear();
                        for (j = 0; j < dt.Rows.Count; j++)
                        {
                            if (dt.Rows[j]["Tables_in_" + Class.Users.ProjectID + ""].ToString().Substring(0, 3) == "to_")
                            {
                                TableName.Items.Add(dt.Rows[j]["Tables_in_" + Class.Users.ProjectID + ""].ToString().Remove(0, 3));
                            }
                        }



                    }

                    lblemptot.Text = " Total Table Count   :" + dataGridView2.Rows.Count;
                }
            }
            catch (Exception ex)
            {

            }
        }
        DataTable dttbl2 = new DataTable();


        public void News()
        {

            GlobalVariables.DownLoads.Text = "UpLoad";
            Class.Users.Bisconnectclear = false;
            allip3.Items.Clear();
            allislug.Items.Clear();
            allip3.Items.Clear(); progressBar1.Maximum = 0;
            Class.Users.TableName = null; Class.Users.TableNameGrid = null; Class.Users.TableName = ""; Class.Users.TableNameGrid = "";
            Class.Users.TableNameSubGrid = null; Class.Users.Prefix = null; Class.Users.Prefix = null;
            Class.Users.Description = null; Class.Users.Description = null;
            Class.Users.FieldName = null; Class.Users.FieldName = null;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName; listView2.Font = Class.Users.FontName;
            Class.Users.TableName = null; Class.Users.TableNameGrid = null; Class.Users.TableName = ""; Class.Users.TableNameGrid = "";
            Class.Users.TableNameSubGrid = null; Class.Users.Prefix = null; Class.Users.Prefix = null;
            Class.Users.Description = null; Class.Users.Description = null;
            Class.Users.FieldName = null; Class.Users.FieldName = null;

            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;

            if (dataGridView2.Rows.Count > 1)
            {
                daimport.GridRowRemove(dataGridView1);
                daimport.GridRowRemove(dataGridView2);

            }
            GridLoad(); TableLoad();

            listView1.Items.Clear();
            listView2.Items.Clear();

            Class.Users.CompCode1 = "";

        }

        public void Saves()
        {

        }

        public void Prints()
        {

            ////string sel = "SELECT  DISTINCT  A.AUTOGENERATEID,A.TX_VIEW_ID  FROM " + Class.Users.ProjectID + ".AUTOGENERATE A  WHERE A.TX_VIEW_ID='" + Class.Users.TableNameGrid + "' AND A.PROJECTID='" + Class.Users.ProjectID + "' ";
            ////DataSet ds = Utility.ExecuteSelectQuery(sel, Class.Users.TableNameGrid);
            ////DataTable dt3 = ds.Tables[Class.Users.TableNameGrid];
            ////if (dt3.Rows.Count > 0)
            ////{
            ////    string sel1 = "select to_char(count(*)) as total  from PSSDEMO." + Class.Users.TableName;
            ////    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, Class.Users.TableName);
            ////    DataTable dt1 = ds1.Tables[Class.Users.TableName];

            ////    string ins = "update " + Class.Users.ProjectID + ".autogenerate set TX_VIEW_ID='" + dt3.Rows[0]["TX_VIEW_ID"].ToString() + "',LASTNO='" + dt1.Rows[0]["total"].ToString() + "' where AUTOGENERATEID=" + dt3.Rows[0]["AUTOGENERATEID"].ToString();
            ////    Utility.ExecuteNonQuery(ins);


            ////}


        }
        public void Searchs()
        {

        }

        public void Deletes()
        {

        }

        public void Imports()
        {
            //int l = 0; int k = 0; string sel = ""; int cnt = 0; Class.Users.Bisconnectclear = false;
            //try
            //{

            //    //progressBar1.Visible = true;
            //    fromgridtotable = "";
            //    Class.Users.Sequenceno = 0;
            //    //lblprocessbar.Visible = true;
            //    string nongridid = "";
            //    string nongridid1 = ""; string gridid = ""; string gridid1 = "";
            //    cnt = 0; int p = 0; int q = 0; update1 = "";
            //    int colcount = dataGridView1.Rows.Count; bool savefalse = false; Class.Users.UserTime = 0;
            //    if (Class.Users.TableName != "")
            //    {
            //        if (dataGridView2.Rows.Count > 0)
            //        {
            //            Cursor.Current = Cursors.WaitCursor;
            //            double sequence = 0; string sequence1 = "";
            //            //slugTable(dataGridView3.Rows[0].Cells[2].FormattedValue.ToString(), dataGridView3.Rows[0].Cells[3].FormattedValue.ToString(), dataGridView3.Rows[0].Cells[4].FormattedValue.ToString());
            //            if (colcount > 0)
            //            {
            //                if (fromtable == "" || fromtable == null)
            //                {

            //                    foreach (DataGridViewRow item in dataGridView2.Rows)
            //                    {
            //                        if (fromtable == "")
            //                        {
            //                            fromtable = item.Cells[3].FormattedValue.ToString();
            //                        }
            //                        else
            //                        {
            //                            fromtable += "," + item.Cells[3].FormattedValue.ToString();
            //                        }
            //                        i++;

            //                    }
            //                }

            //                if (Class.Users.TableName != null)
            //                {

            //                    DataTable dtseq = daimport.FindSlug(Class.Users.TableName);
            //                    if (Convert.ToInt32("0" + dtseq.Rows[0]["CNT"].ToString()) == 0 || Class.Users.Bisconnectclear == true)
            //                    {
            //                        if (Class.Users.TableNameGrid != "")
            //                        {
            //                            DataTable dttbl2 = daimport.SlugData(Class.Users.ProjectID, Class.Users.TableNameGrid, Class.Users.Description, Class.Users.FieldName, Class.Users.Prefix);
            //                            if (dttbl2.Rows[0]["LASTNO"].ToString() == "") { Class.Users.Sequenceno = 14; }
            //                            else
            //                            {
            //                                Class.Users.Sequenceno = Convert.ToDouble(dttbl2.Rows[0]["LASTNO"].ToString());
            //                            }
            //                            sequence = Convert.ToInt64(dttbl2.Rows[0]["ZEROPADDING"].ToString());
            //                            //textMessagebox.Refresh();
            //                            //textMessagebox.Text = "Sequence No" + sequence;

            //                        }
            //                        else
            //                        {
            //                            sequence = Class.Users.Sequenceno;
            //                            //textMessagebox.Refresh();
            //                            //textMessagebox.Text = "Sequence No" + sequence;
            //                        }
            //                        if (Convert.ToDouble(sequence) >= 1 && Class.Users.Sequenceno == 14)
            //                        {
            //                            switch (sequence)
            //                            {
            //                                case 1:
            //                                    sequence1 = "1";
            //                                    break;
            //                                case 2:

            //                                    sequence1 = "1";
            //                                    break;
            //                                case 3:

            //                                    sequence1 = "101";
            //                                    break;
            //                                case 4:

            //                                    sequence1 = "1001";
            //                                    break;
            //                                case 5:
            //                                    sequence1 = "10001";
            //                                    break;
            //                                case 6:
            //                                    sequence1 = "100001";
            //                                    break;
            //                                case 7:
            //                                    sequence1 = "1000001";
            //                                    break;
            //                                case 8:
            //                                    sequence1 = "10000001";
            //                                    break;
            //                                case 9:

            //                                    sequence1 = "100000001";
            //                                    break;
            //                                case 10:

            //                                    sequence1 = "1000000001";
            //                                    break;
            //                                case 11:
            //                                    sequence1 = "10000000001";
            //                                    break;
            //                                case 12:
            //                                    sequence1 = "100000000001";
            //                                    break;
            //                                case 13:
            //                                    sequence1 = "1000000000001";
            //                                    break;
            //                                case 14:
            //                                    sequence1 = "2000000000001";
            //                                    break;
            //                                default:
            //                                    sequence1 = "0";
            //                                    break;
            //                            }
            //                            Class.Users.Sequenceno = Convert.ToDouble(sequence1);
            //                        }
            //                    }
            //                    if (Convert.ToInt32("0" + dtseq.Rows[0]["CNT"].ToString()) > 0)
            //                    {
            //                        daimport.dropsequence(Class.Users.ProjectID, Class.Users.TableName);
            //                        //textMessagebox.Refresh(); lblmaxid.Refresh();
            //                        //textMessagebox.Text = "Droped Sequence";

            //                        daimport.CreateSequence(Class.Users.ProjectID, Class.Users.TableName, Class.Users.Sequenceno);

            //                        //textMessagebox.Text = "Created Sequence" + Class.Users.ProjectID + Class.Users.TableName + "SEQ" + " Sequence No :" + Class.Users.Sequenceno;

            //                    }
            //                    if (Convert.ToInt32("0" + dtseq.Rows[0]["CNT"].ToString()) == 0)
            //                    {
            //                       // textMessagebox.Refresh(); lblmaxid.Refresh();

            //                        daimport.CreateSequence(Class.Users.ProjectID, Class.Users.TableName, Class.Users.Sequenceno);

            //                       // textMessagebox.Text = "Created Sequence" + Class.Users.ProjectID + Class.Users.TableName + "SEQ" + " Sequence No :" + Class.Users.Sequenceno;

            //                    }
            //                    DataTable dtseq0 = daimport.FindTrigger(Class.Users.TableName, Class.Users.ProjectID);
            //                    if (dtseq0 == null || dtseq0.Rows[0]["CNT"].ToString() == "0")
            //                    {
            //                       // textMessagebox.Refresh();
            //                        daimport.CreateTrigger(Class.Users.TableName, Class.Users.ProjectID);
            //                       // textMessagebox.Refresh();
            //                       // textMessagebox.Text = "Created Trigger" + Class.Users.ProjectID + Class.Users.TableName + "TRI";
            //                    }
            //                }

            //                l = 0; int n = 0;
            //                //progressBar1.Minimum = 0;
            //                i = 0; DataTable dt1 = new DataTable();
            //                //progressBar1.Maximum = dataGridView1.Rows.Count;
            //                dt1.Rows.Clear(); k = 0; l = 0; int totcnt = dataGridView1.Rows.Count;
            //                for (k = 0; k < dataGridView1.Rows.Count; k++)
            //                {
            //                    Class.Users.UserTime = 0;
            //                    if (fromtable != "")
            //                    {
            //                        for (l = 0; l < dataGridView1.Columns.Count; l++)
            //                        {
            //                            Cursor.Current = Cursors.WaitCursor;
            //                            Class.Users.UserTime = 0;
            //                            if (Class.Users.TableName != null)
            //                            {
            //                                if (totable == "" && Class.Users.TableName != null)
            //                                {
            //                                    totable = "'" + dataGridView1.Rows[k].Cells[l].Value + "'";
            //                                    if (dataGridView2.Rows[0].Cells[5].FormattedValue.ToString().Trim() == "YES")
            //                                    {
            //                                        DataTable dt2 = daimport.GridSelect(dataGridView2.Rows[l].Cells[2].FormattedValue.ToString().Trim(), dataGridView2.Rows[l].Cells[4].FormattedValue.ToString().Trim(), dataGridView1.Rows[k].Cells[l].FormattedValue.ToString().Trim());
            //                                        if (dt2 == null || Convert.ToInt32(dt2.Rows.Count) <= 0)
            //                                        {
            //                                            //textMessagebox.Refresh();
            //                                           // textMessagebox.Text = "No Data Found this Table   :'" + dataGridView2.Rows[l].Cells[2].FormattedValue + "' 's      " + " Field Name is " + dataGridView2.Rows[l].Cells[4].FormattedValue + "='" + dataGridView1.Rows[k].Cells[l].Value + "  " + "ROW  :  " + k.ToString() + "  COLUMN  :" + l.ToString() + "   Insert Query    : " + sel;
            //                                            // MessageBox.Show("No Data Found this Table   :'" + dataGridView2.Rows[l].Cells[2].FormattedValue + "' 's      " + " Field Name is " + dataGridView2.Rows[l].Cells[4].FormattedValue + "='" + dataGridView1.Rows[k].Cells[l].Value + "  " + "ROW  :  " + k.ToString() + "  COLUMN  :" + l.ToString() + "   Insert Query    : " + sel + "    Excel  Data Invalid", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                                            cnt = 0; colcount = 0; totable = ""; totableupdate = "";
            //                                            cnt = k + 1; l = dataGridView1.Columns.Count;
            //                                        }
            //                                        if (dt2.Rows.Count >= 1)
            //                                        {
            //                                            string c = dt2.Rows[0][dataGridView2.Rows[l].Cells[2].FormattedValue + "ID"].ToString();
            //                                            totable = "'" + c + "'";
            //                                            nongridid = "  where  ";
            //                                            nongridid1 = dataGridView2.Rows[l].Cells[3].FormattedValue.ToString();
            //                                            totableupdate += " where " + dataGridView2.Rows[l].Cells[3].FormattedValue.ToString() + "='" + c + "'";

            //                                        }
            //                                    }
            //                                    if (dataGridView2.Rows[0].Cells[5].FormattedValue.ToString().Trim() == "")
            //                                    {
            //                                        nongridid = "  where  ";
            //                                        nongridid1 = dataGridView2.Rows[l].Cells[3].FormattedValue.ToString();
            //                                        totableupdate += " where " + dataGridView2.Rows[l].Cells[4].FormattedValue.ToString() + "='" + dataGridView1.Rows[k].Cells[l].Value + "'";
            //                                    }
            //                                }
            //                                else
            //                                {
            //                                    if (Class.Users.TableName != null)
            //                                    {
            //                                        if (dataGridView1.Rows[k].Cells[l].Value.ToString().Trim().Length >= 10)
            //                                        {
            //                                            if (dataGridView1.Rows[k].Cells[l].Value.ToString().Substring(2, 1) == "&")//|| dataGridView1.Rows[k].Cells[l].Value.ToString().Substring(2, 1) == "/"
            //                                            {
            //                                                string Dates = dataGridView1.Rows[k].Cells[l].Value.ToString().Replace("&", "-");
            //                                                totable += ",to_date('" + Dates.Substring(0, 10) + "','dd-MM-yyyy')";
            //                                            }
            //                                            else
            //                                            {
            //                                                if (dataGridView2.Rows[l].Cells[2].FormattedValue.ToString().Trim() != "")// dataGridView1.Rows[l].Cells[2].FormattedValue.ToString() NOW CHANGES FOR ID IS NOT NULL
            //                                                {
            //                                                    DataTable dt2 = daimport.GridSelect(dataGridView2.Rows[l].Cells[2].FormattedValue.ToString().Trim(), dataGridView2.Rows[l].Cells[4].FormattedValue.ToString().Trim(), dataGridView1.Rows[k].Cells[l].FormattedValue.ToString().Trim());
            //                                                    if (dt2 == null || Convert.ToInt32(dt2.Rows.Count) <= 0)
            //                                                    {
            //                                                        Cursor = Cursors.Default;// textMessagebox.Refresh();
            //                                                        //MessageBox.Show("No Data Found this Table   :'" + dataGridView2.Rows[l].Cells[2].FormattedValue + "' 's      " + " Field Name is " + dataGridView2.Rows[l].Cells[4].FormattedValue + "='" + dataGridView1.Rows[k].Cells[l].Value + "  " + "ROW  :  " + k.ToString() + "  COLUMN  :" + l.ToString() + "   Insert Query    : " + daimport.query + "    Excel  Data Invalid", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                                                        //textMessagebox.Text = "No Data Found this Table   :'" + dataGridView2.Rows[l].Cells[2].FormattedValue + "' 's      " + " Field Name is " + dataGridView2.Rows[l].Cells[4].FormattedValue + "='" + dataGridView1.Rows[k].Cells[l].Value + "  " + "ROW  :  " + k.ToString() + "  COLUMN  :" + l.ToString() + "   Insert Query    : " + daimport.query + "    Excel  Data Invalid";
            //                                                        cnt = 0; colcount = 0; totable = ""; totableupdate = "";
            //                                                        cnt = k + 1; l = dataGridView1.Columns.Count;
            //                                                    }
            //                                                    if (dt2.Rows.Count >= 1)
            //                                                    {
            //                                                        string c = dt2.Rows[0][dataGridView2.Rows[l].Cells[2].FormattedValue + "ID"].ToString();
            //                                                        totable += ",'" + dt2.Rows[0][dataGridView2.Rows[l].Cells[2].FormattedValue + "ID"].ToString() + "'";
            //                                                        daimport.query = "";
            //                                                    }
            //                                                }
            //                                                else
            //                                                {
            //                                                    totable += ",'" + dataGridView1.Rows[k].Cells[l].Value.ToString() + "'";
            //                                                }
            //                                            }
            //                                        }
            //                                        else
            //                                        {
            //                                            if (dataGridView2.Rows[l].Cells[2].FormattedValue.ToString().Trim() != "")
            //                                            {
            //                                                daimport.query = "";
            //                                                DataTable dt2 = daimport.GridSelect(dataGridView2.Rows[l].Cells[2].FormattedValue.ToString().Trim(), dataGridView2.Rows[l].Cells[4].FormattedValue.ToString().Trim(), dataGridView1.Rows[k].Cells[l].FormattedValue.ToString().Trim());
            //                                                if (dt2 == null || Convert.ToInt32(dt2.Rows.Count) <= 0)
            //                                                {
            //                                                    Cursor = Cursors.Default;
            //                                                    //textMessagebox.Refresh();
            //                                                   // textMessagebox.Text = "No Data Found this Table: '" + dataGridView2.Rows[l].Cells[2].FormattedValue + "' 's      " + " Field Name is " + dataGridView2.Rows[l].Cells[4].FormattedValue + "='" + dataGridView1.Rows[k].Cells[l].Value + "  " + "ROW: " + k.ToString() + "  COLUMN: " + l.ToString() + "   Insert Query    : " + daimport.query + "    Excel Data Invalid";
            //                                                    colcount = 0; totable = ""; totableupdate = "";
            //                                                    cnt = k + 1; l = dataGridView1.Columns.Count;
            //                                                }
            //                                                if (dt2.Rows.Count >= 1)
            //                                                {
            //                                                    string c = dt2.Rows[0][dataGridView2.Rows[l].Cells[2].FormattedValue + "ID"].ToString();
            //                                                    totable += ",'" + dt2.Rows[0][dataGridView2.Rows[l].Cells[2].FormattedValue + "ID"].ToString() + "'";

            //                                                }
            //                                            }
            //                                            else
            //                                            {
            //                                                totable += ",'" + dataGridView1.Rows[k].Cells[l].Value + "'";
            //                                            }
            //                                        }
            //                                    }
            //                                    else
            //                                    {
            //                                        Cursor = Cursors.Default;
            //                                        MessageBox.Show("Table Name is Empty");
            //                                    }
            //                                }
            //                            }
            //                        }
            //                    }
            //                    if (totable != "" && fromtable != "" && Class.Users.TableName != null)
            //                    {
            //                        update1 = totableupdate;
            //                        daimport.query = "";
            //                        DataTable dt3 = daimport.FindDuplicate(Class.Users.TableName, Class.Users.ProjectID, update1);
            //                        if (dt3 == null)
            //                        {
            //                            Cursor = Cursors.Default;
            //                           // textMessagebox.Refresh();
            //                            ///textMessagebox.Text = "No Data Found this Table   :'" + Class.Users.TableName + "' 's      " + " Field Name is  Empty. pls select  Correct Table name .    ROW  :  " + k.ToString() + "  COLUMN  :" + l.ToString() + "   Insert Query    : " + daimport.query + "    Excel  Data Invalid";
            //                            totable = ""; totableselect = ""; totableupdate = ""; totableid = ""; nongridid = ""; nongridid1 = ""; totableupdategrid = "";
            //                            cnt = 0; colcount = 0; cnt = k + 1; l = dataGridView1.Columns.Count;
            //                            savefalse = true;
            //                        }
            //                        else if (dt3.Rows.Count != 0)
            //                        {

            //                            totable = ""; totableselect = ""; totableupdate = ""; totableid = ""; nongridid = ""; nongridid1 = ""; totableupdategrid = "";
            //                            cnt = 0; colcount = 0;
            //                        }
            //                        else if (dt3.Rows.Count == 0)
            //                        {
            //                            if (dataGridView2.Rows[0].Cells[5].FormattedValue.ToString() == "YES")
            //                            {
            //                                Cursor.Current = Cursors.Default;
            //                                string ins = "insert into " + Class.Users.ProjectID + "." + Class.Users.TableName + "(" + fromtable + ")values(" + totable + ")";
            //                                //textMessagebox.Refresh();
            //                                //textMessagebox.Text = ins;
            //                                Utility.ExecuteNonQuery(ins);
            //                                totable = ""; totableselect = ""; totableupdate = ""; totableid = ""; nongridid = ""; nongridid1 = ""; totableupdategrid = "";
            //                                savefalse = true;
            //                            }
            //                            if (dataGridView2.Rows[0].Cells[5].FormattedValue.ToString() == "")
            //                            {
            //                                Cursor.Current = Cursors.Default;
            //                                string ins = "insert into " + Class.Users.ProjectID + "." + Class.Users.TableName + "(" + fromtable + Details + ")values(" + totable + Details1 + ")";
            //                                //textMessagebox.Refresh();
            //                                //textMessagebox.Text = ins;
            //                                Utility.ExecuteNonQuery(ins);
            //                                totable = ""; totableselect = ""; totableupdate = ""; totableid = ""; nongridid = ""; nongridid1 = ""; totableupdategrid = "";
            //                                savefalse = true;
            //                            }
            //                        }
            //                        else
            //                        {
            //                            totable = ""; totableselect = ""; totableupdate = ""; totableid = ""; nongridid = ""; nongridid1 = ""; totableupdategrid = "";
            //                        }

            //                        //lblprocessbar.Refresh();
            //                        //lblprocessbar.Text = "Total Rows :" + totcnt + " of " + k.ToString();
            //                        //textMessagebox.Text = lblprocessbar.Text;
            //                        //lblprocessbar.Refresh(); textMessagebox.Refresh();

            //                    }
            //                    cnt++;
            //                }


            //                if (savefalse == false)
            //                {
            //                    MessageBox.Show("Child Record Found."); lblchild.Refresh();
            //                    lblchild.Text = "Child Record Found.";
            //                    //textMessagebox.Refresh();
            //                    //textMessagebox.Text = lblchild.Text;
            //                }
            //                if (savefalse == true)
            //                {
            //                    MessageBox.Show("Record Saved Successfully." + dataGridView1.Rows.Count.ToString());
            //                    //textMessagebox.Refresh();
            //                    //textMessagebox.Text = "Record Saved Successfully." + dataGridView1.Rows.Count.ToString();
            //                }
            //            }

            //        }
            //        else
            //        {
            //            Cursor = Cursors.Default;
            //            MessageBox.Show("No Data Found Excel Field. ", "Excel Column count: '" + dataGridView2.Rows.Count + "'", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //        }
            //    }
            //    //if (Convert.ToString(Class.Users.TableName) != null && Class.Users.TableNameGrid != null && Class.Users.Prefix != null)
            //    //{
            //    //    //if (allislug.Items.Count >= 1 && allislug.Items.Count > 0)
            //    //    //{

            //    //    //    if (dataGridView3.Rows[0].Cells[2].FormattedValue.ToString() != "")
            //    //    //    {
            //    //    //        totable = "";
            //    //    //        DataTable dt2 = daimport.TableMax(Class.Users.TableName, Class.Users.ProjectID);
            //    //    //        dataGridView3.Rows[0].Cells[5].Value = dt2.Rows[0]["LASTNO"].ToString();
            //    //    //        daimport.query = null;
            //    //    //        DataTable dt3 = daimport.SlugData(Class.Users.ProjectID, Class.Users.TableNameGrid, dataGridView3.Rows[0].Cells[2].FormattedValue.ToString(), dataGridView3.Rows[0].Cells[3].FormattedValue.ToString(), dataGridView3.Rows[0].Cells[4].FormattedValue.ToString());
            //    //    //        if (dt3 == null || Convert.ToInt32(dt3.Rows.Count) <= 0)
            //    //    //        {
            //    //    //            MessageBox.Show("No data found in Parent Table.Pls Check Excel Data  :'" + dataGridView3.Rows[0].Cells[2].FormattedValue + "' 's      " + " Field Name is " + dataGridView3.Rows[0].Cells[3].FormattedValue + "='" + dataGridView3.Rows[0].Cells[1].Value + "        " + "ROW  :  " + k.ToString() + "  COLUMN  :" + l.ToString() + "   Insert Query    : " + sel + "    Excel  Data Invalid", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    //    //            fromtable = ""; totable = ""; totableupdate = ""; daimport.query = "";
            //    //    //        }
            //    //    //        if (dt3.Rows.Count > 0)
            //    //    //        {
            //    //    //            Cursor.Current = Cursors.WaitCursor;
            //    //    //            for (k = 0; k < dataGridView3.Rows.Count - 1; k++)
            //    //    //            {

            //    //    //                if (allislug.Items.Count > 0)
            //    //    //                {
            //    //    //                    for (l = 0; l < dataGridView3.Columns.Count - 1; l++)
            //    //    //                    {
            //    //    //                        if (Class.Users.TableNameGrid != null && dataGridView3.Rows[0].Cells[1].FormattedValue.ToString() != "")
            //    //    //                        {
            //    //    //                            if (l == 0)
            //    //    //                            {
            //    //    //                                totablegrid = allislug.Items[l].SubItems[0].Text;
            //    //    //                                totable = allislug.Items[l].SubItems[0].Text + " = '" + dt3.Rows[0]["AUTOGENERATEID"].ToString() + "'";
            //    //    //                            }
            //    //    //                            if (l >= 1)
            //    //    //                            {

            //    //    //                                totableupdate += "  " + allislug.Items[l].SubItems[0].Text + "='" + dataGridView3.Rows[k].Cells[l].Value + "' ,";
            //    //    //                            }
            //    //    //                        }
            //    //    //                    }
            //    //    //                }
            //    //    //            }
            //    //    //            if (totable != "" && allislug.Items.Count > 0 && Class.Users.TableNameGrid != null)
            //    //    //            {
            //    //    //                update1 = ""; int ss, sss;
            //    //    //                ss = Convert.ToInt32(totableupdate.Length.ToString());
            //    //    //                sss = Convert.ToInt32(ss - 1);
            //    //    //                update1 = totableupdate.Substring(0, sss).ToString();
            //    //    //                daimport.query = "update " + Class.Users.ProjectID + ".autogenerate set " + update1 + " where " + totable + " AND PREFIX='" + Class.Users.Prefix + "'";
            //    //    //                Utility.ExecuteNonQuery(daimport.query);
            //    //    //                totable = ""; fromgridtotable = ""; update1 = ""; //progressBar1.Minimum = 0;
            //    //    //            }


            //    //    //        }
            //    //    //        else
            //    //    //        {
            //    //    //            Cursor = Cursors.Default;
            //    //    //            MessageBox.Show(" Columns Field's  : " + dataGridView2.Columns.Count.ToString() + "   Table Columns Field's :   " + allislug.Items.Count.ToString(), "  Invalid  ", MessageBoxButtons.OK, MessageBoxIcon.Information); ;

            //    //    //        }

            //    //    //    }

            //    //    //}
            //    //}

            //    Cursor = Cursors.Default; lblchild.Text = cnt.ToString();
            //    cnt = 0; colcount = 0; 
            //    fromtable = ""; totable = ""; totableupdate = ""; totableupdategrid = "";// News();
            //    DataTable DT = daimport.FindTrigger(Class.Users.TableName, Class.Users.ProjectID);
            //    if (DT.Rows.Count > 0)
            //    {
            //        daimport.dropsequence(Class.Users.ProjectID, Class.Users.TableName);
            //        daimport.DropTrigger(Class.Users.ProjectID, Class.Users.TableName);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    // Cursor = Cursors.Default;
            //    MessageBox.Show(ex.Message + "-----" + k.ToString() + "  COLUMN  :" + l.ToString() + "   Insert Query    : " + sel + "    pls Select Correct Table or Column", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    fromtable = ""; fromslugtale = "";
            //    totable = ""; totableupdate = ""; totableupdategrid = ""; cnt = k + 1; l = dataGridView1.Columns.Count;
            //}

        }

        public void Pdfs()
        {


        }

        public void ChangePasswords()
        {

        }

        public void DownLoads()
        {
            Master.Bank.PopUp1 pop = new Master.Bank.PopUp1();
            pop.Show();
            //try
            //{


            //    DataTable dtgridview = new DataTable();

            //    if (dataGridView1.Rows.Count > 0)
            //    {
            //        daimport.GridRowRemove(dataGridView1);
            //    }

            //    if (dataGridView1.Rows.Count <= 0)
            //    {
            //        dtgridview.Rows.Clear(); dtgridview.Columns.Clear();

            //        int i = 0;

            //        System.Data.OleDb.OleDbConnection OledbConn;
            //        System.Data.OleDb.OleDbCommand OledbCmd;
            //        System.Data.OleDb.OleDbDataAdapter OledbAdapter;
            //        string filePath = string.Empty; string fileExt = string.Empty;
            //        OpenFileDialog file = new OpenFileDialog(); string path = "";
            //        if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            //        {
            //            filePath = file.FileName; //get the path of the file  
            //            fileExt = Path.GetExtension(filePath); //get the file extension  
            //            if (fileExt.CompareTo(".xls") == 0)
            //                path = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            //            else
            //                path = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  

            //            OledbConn = new System.Data.OleDb.OleDbConnection(path);
            //            string qry1 = "Select * from [Sheet1$]";
            //            OledbAdapter = new OleDbDataAdapter(qry1, OledbConn);
            //            OledbAdapter.Fill(dtgridview);
            //            if (dtgridview.Rows.Count > 0)
            //            {

            //                dataGridView1.DataSource = dtgridview;
            //                int k = 1;
            //                if (dataGridView2.Rows.Count <= 1)
            //                {
            //                    for (int j = 0; j < dtgridview.Columns.Count; j++)
            //                    {

            //                        dataGridView2.Rows.Add();
            //                        dataGridView2.Rows[j].Cells[0].Value = k.ToString();
            //                        dataGridView2.Rows[j].Cells[1].Value = dtgridview.Columns[j].ToString();

            //                        k++;


            //                    }
            //                }
            //            }

            //        }

            //        int cnt = dataGridView1.Rows.Count;
            //        lbldowncount.Text = "Total Excel Rows : " + cnt.ToString();
            //        if (dataGridView1.Rows.Count > 0)
            //        {
            //            tabControl1.SelectTab(tabPage3);
            //        }
            //        else
            //        {
            //            tabControl1.SelectTab(tabPage4);
            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //    //MessageBox.Show(ex.ToString());
            //}

        }
        public void GridView(DataGridView dataGridView1, Button myTablename)
        {


            try
            {

                DataTable dtgridview = new DataTable();

                if (dataGridView1.Rows.Count > 0)
                {
                    daimport.GridRowRemove(dataGridView1);
                }

                if (dataGridView1.Rows.Count <= 0)
                {
                    dtgridview.Rows.Clear(); dtgridview.Columns.Clear();

                    int i = 0;

                    System.Data.OleDb.OleDbConnection OledbConn;
                    System.Data.OleDb.OleDbCommand OledbCmd;
                    System.Data.OleDb.OleDbDataAdapter OledbAdapter;
                    string filePath = string.Empty; string fileExt = string.Empty;
                    OpenFileDialog file = new OpenFileDialog(); string path = "";
                    if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
                    {
                        filePath = file.FileName; //get the path of the file  
                        fileExt = Path.GetExtension(filePath); //get the file extension  
                        if (fileExt.CompareTo(".xls") == 0)
                            path = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                        else
                            path = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  

                        OledbConn = new System.Data.OleDb.OleDbConnection(path);
                        string qry1 = "Select * from [Sheet1$]";
                        OledbAdapter = new OleDbDataAdapter(qry1, OledbConn);
                        OledbAdapter.Fill(dtgridview);
                        if (dtgridview.Rows.Count > 0)
                        {

                            dataGridView1.DataSource = dtgridview;
                            int k = 1;
                            if (dataGridView2.Rows.Count <= 1)
                            {
                                for (int j = 0; j < dtgridview.Columns.Count; j++)
                                {

                                    dataGridView2.Rows.Add();
                                    dataGridView2.Rows[j].Cells[0].Value = k.ToString();
                                    dataGridView2.Rows[j].Cells[1].Value = dtgridview.Columns[j].ToString();

                                    k++;


                                }
                            }
                        }

                    }

                    int cnt = dataGridView1.Rows.Count;
                    lbldowncount.Text = "Total Excel Rows : " + cnt.ToString();

                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }

            Models.Validate val = new Models.Validate();
            if (dataGridView1.Rows.Count > 1 && Class.Users.CompCode1 != "")
            {
                if (dataGridView2.Columns.Count >= 7)
                {
                    listView2.Items.Clear();
                    string frmtable = "", frmrow = "", frmdata = ""; int i = 0;
                    foreach (DataGridViewColumn row in dataGridView1.Columns)
                    {
                        //if (val.IsStringdotspace(row.HeaderText) == true)
                        //{
                        if (i == 0)
                        {
                            if (row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") == "date")
                            {
                                frmtable += row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") + "id  bigint NOT NULL AUTO_INCREMENT primary key";
                                frmtable += "," + row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") + "1 date";
                                frmrow += row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") + "1";
                            }
                            else
                            {
                                MessageBox.Show("First Column Should be Date Column in XL. Pls Change in XL ");
                                return;
                            }
                        }
                        if (i >= 1)
                        {
                            if (row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") == "credit" || row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") == "debit")
                            {
                                
                                frmtable += "," + row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") + "  decimal(12,2) not null default 0.00";
                                frmrow += "," + row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "");
                                if (row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") == "credit")
                                {
                                    frmtable += "," + "unmatched varchar(250) default null";
                                }
                            }

                            else
                            {

                                frmtable += "," + row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") + "  varchar(500) default null";
                                frmrow += "," + row.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "");
                            }
                        }
                        i++;

                    }
                    if (frmtable != "")
                    {
                        if (myTablename.Text == "MY COMPANY")
                        {

                            string cre = "drop table if exists my_" + Class.Users.CompCode1 + "; ";
                            Utility.ExecuteNonQuery(cre);
                            cre = "";
                            cre = "create table if not exists my_" + Class.Users.CompCode1 + " (" + frmtable + ");";
                            Utility.ExecuteNonQuery(cre);
                        }
                        else
                        {

                            string cre = "drop table if exists  to_" + Class.Users.CompCode1 + ";";
                            Utility.ExecuteNonQuery(cre);
                            cre = "";
                            cre = "create table if not exists to_" + Class.Users.CompCode1 + "(" + frmtable + ");";
                            Utility.ExecuteNonQuery(cre);
                        }

                        i = 0; int j = 0;

                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value.ToString() != "")
                            {
                                foreach (DataGridViewColumn col in dataGridView1.Columns)
                                {


                                    if (i + j == 0)
                                    {
                                        frmdata += "'" + Convert.ToDateTime(row.Cells[i + j].Value.ToString().Substring(0, 10)).ToString("yyyy-MM-dd") + "'";
                                    }
                                    if (i + j >= 1)
                                    {
                                        if (col.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") == "debit")
                                        {
                                            frmdata += ",'" + Convert.ToDecimal("0" + row.Cells[i + j].Value) + "'";
                                        }
                                        if (col.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") == "credit")
                                        {
                                            frmdata += ",'" + Convert.ToDecimal("0" + row.Cells[i + j].Value) + "'";
                                        }

                                        if (col.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") != "debit" && col.HeaderText.ToLower().Replace("#", "").Trim().Replace(" ", "") != "credit")
                                        {
                                            frmdata += ",'" + row.Cells[i + j].Value + "'";
                                        }
                                    }

                                    i++;

                                }
                                string ins = "";
                                if (myTablename.Text == "MY COMPANY")
                                {
                                    ins = "";
                                    ins = "insert into my_" + Class.Users.CompCode1 + " (" + frmrow + ") values(" + frmdata + ");";
                                    Utility.ExecuteNonQuery(ins);
                                }
                                else
                                {
                                    ins = "";
                                    ins = "insert into to_" + Class.Users.CompCode1 + " (" + frmrow + ") values(" + frmdata + ");";
                                    Utility.ExecuteNonQuery(ins);
                                }
                                frmdata = "";
                                i++; ins = ""; i = 0;
                            }
                        }
                        frmtable = "";

                        do
                        {
                            for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
                        }
                        while (dataGridView1.Rows.Count > 0);
                        Class.Users.CompCode1 = "";
                        GridLoad();
                        tabControl1.SelectTab(tabPage4);
                    }
                    else
                    {
                        MessageBox.Show("No Data Found  in frmtable");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Minimum Excel Column should be 7. ","7 Columns In Excel ");
                    return;
                }
            }
            else
            {
                MessageBox.Show("No Data Found");
                return;
            }
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Class.Users.UserTime = 0;
            int IDX = e.RowIndex;
            if (dataGridView2.Columns[e.ColumnIndex].Name == "Show")
            {

                try
                {
                    tabControl1.SelectTab(tabPage1);
                    listfilterred.Items.Clear(); listfilterscreen.Items.Clear();
                    listfilterred1.Items.Clear(); listfilterscreen1.Items.Clear();
                    Class.Users.UserTime = 0; DataTable dt = new DataTable();
                    DataTable dt1 = new DataTable();
                    listView1.Items.Clear(); listView2.Items.Clear();
                    int i = 0; int gridcount = 1; Class.Users.TableNameGrid = ""; Class.Users.DocID = "";
                    Class.Users.UniqueID = ""; textBox1.Text = ""; Class.Users.TableNameSubGrid = "";
                    if (dataGridView2.Rows[0].Cells[4].EditedFormattedValue.ToString() != "" && dataGridView2.Rows[0].Cells[5].EditedFormattedValue.ToString() != "")
                    {
                        Class.Users.CompCode1 = "";
                        Class.Users.CompCode1 = dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower();
                        Class.Users.DocID = dataGridView2.Rows[0].Cells[4].EditedFormattedValue.ToString();
                        Class.Users.UniqueID = dataGridView2.Rows[0].Cells[5].EditedFormattedValue.ToString();
                        foreach (DataGridViewRow gridrow in dataGridView2.Rows)
                        {
                            if (gridrow.Cells[4].EditedFormattedValue.ToString() != "" && gridrow.Cells[5].EditedFormattedValue.ToString() != "")
                            {
                                Class.Users.TableNameGrid += " a." + gridrow.Cells[4].EditedFormattedValue.ToString() + " = b." + gridrow.Cells[5].EditedFormattedValue.ToString() + " and ";

                            }
                            if (gridrow.Cells[4].EditedFormattedValue.ToString() != "" && gridrow.Cells[5].EditedFormattedValue.ToString() != "")
                            {
                                Class.Users.TableNameSubGrid += " b." + gridrow.Cells[4].EditedFormattedValue.ToString() + " = a." + gridrow.Cells[5].EditedFormattedValue.ToString() + " and ";

                            }

                        }
                        if (Class.Users.TableNameGrid != "")
                        {
                            update1 = ""; int ss, sss; textBox1.Text = "";
                            ss = Convert.ToInt32(Class.Users.TableNameGrid.Length.ToString());
                            sss = Convert.ToInt32(ss - 4);
                            update1 = Class.Users.TableNameGrid.Substring(0, sss).ToString();
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct " + matchfield + ",case when b." + Class.Users.UniqueID + " is null then 'No' else 'Yes' end sts from my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower().Trim() + " a left join to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower().Trim() + " b on " + update1 + " ;";
                            DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower());
                            dt = ds.Tables["to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower()];
                            textBox1.Text += Class.Users.Query.ToString();
                            update1 = ""; ss = 0; sss = 0; Class.Users.Query = "";
                            ss = Convert.ToInt32(Class.Users.TableNameSubGrid.Length.ToString());
                            sss = Convert.ToInt32(ss - 4);
                            update1 = Class.Users.TableNameSubGrid.Substring(0, sss).ToString();
                            Class.Users.Query = "select distinct " + matchfield1 + ", case when b." + Class.Users.DocID + " is null then 'No' else 'Yes' end sts from to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower().Trim() + " a left join my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower().Trim() + "  b on " + update1 + " ";

                            DataSet ds1 = Utility.ExecuteSelectQuery(Class.Users.Query, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower().Trim());
                            dt1 = ds1.Tables["to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString().ToLower().Trim()];
                            textBox1.Text += Class.Users.Query.ToString();
                            update1 = ""; Class.Users.Query = "";
                            i = 1;
                            if (dt != null)
                            {
                                listView1.Items.Clear(); listfilterred.Items.Clear(); listfilterscreen.Items.Clear();
                                lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                                lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                                i = 1;
                                string[] split = textBox1.Text.Split(';');
                                string sel = "select " + fromquery + ",x.sts from(" + split[0] + ") x";
                                listviewxl3(sel, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString(), "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString(), listView1);

                                
                            }
                            else
                            {
                                MessageBox.Show("No Data Found in MyCompany");
                            }
                            if (dt1 != null)
                            {
                                listView2.Items.Clear(); listfilterred1.Items.Clear(); listfilterscreen1.Items.Clear();
                                lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                                i = 1;
                                
                                    string[] split = textBox1.Text.Split(';');
                                    string sel = "select " + toquery + ",x.sts from(" + split[1] + ") x";
                                    listviewxl4(sel, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString(), "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString(),listView2);

                                    
                            }
                            else
                            {
                                MessageBox.Show("No Data Found in Customer or Supplier");

                            }

                        }
                    }
                    else
                    {
                        MessageBox.Show("Pls Select Matching Fields");
                    }
                }
                catch (Exception ex) { ex.ToString(); }
            }
            if (IDX == 0 && dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue != null)
            {


                Cursor.Current = Cursors.WaitCursor;
                GridColumnRefresh(dataGridView2, IDX);
                Cursor.Current = Cursors.Default;



            }
            //if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue != null)
            //{



            //    if (e.RowIndex % 2 == 0)
            //    {
            //        dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.AntiqueWhite;

            //    }
            //    else
            //    {
            //        dataGridView2.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightPink;

            //    }



            //}
            if (IDX >= 1)
            {
                dataGridView2.Rows[e.RowIndex].Cells[2].ReadOnly = true;
            }
            if (celcrear == true)
            {

                Class.Users.Indexer = IDX;
                GridCellClear(dataGridView2, IDX);
                celcrear = false;

            }


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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int item0 = 0;
            //    if (txtsearch.Text.Length > 0)
            //    {
            //        listView1.Items.Clear();

            //        foreach (ListViewItem item in listfilterdb.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (listfilterdb.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text))
            //            {

            //                list.SubItems.Add(listfilterdb.Items[item0].SubItems[1].Text);
            //                list.SubItems.Add(listfilterdb.Items[item0].SubItems[2].Text);


            //                if (item0 % 2 == 0)
            //                {
            //                    list.BackColor = Color.White;

            //                }
            //                else
            //                {
            //                    list.BackColor = Color.WhiteSmoke;

            //                }



            //                listView1.Items.Add(list);
            //            }

            //            item0++;
            //        }
            //        lblemptot.Text = "Total Rows    :" + listView1.Items.Count;
            //    }
            //    else
            //    {
            //        //try
            //        //{
            //        //    listView1.Items.Clear();
            //        //    foreach (ListViewItem item in listfilterdb.Items)
            //        //    {
            //        //        this.listView1.Items.Add((ListViewItem)item.Clone());
            //        //        item0++;
            //        //    }
            //        //    lblemptot.Text = "Total Rows    :" + listView1.Items.Count;
            //        //}
            //        //catch (Exception ex)
            //        //{

            //        //}
            //    }
            //}
            //catch (Exception ex)
            //{
            //    // MessageBox.Show("---" + ex.ToString());
            //}
        }



        private void lvlogs_ItemChecked(object sender, ItemCheckedEventArgs e)
        {


        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int item0 = 0;
            //    if (txtsearch.Text.Length > 0)
            //    {
            //        listView1.Items.Clear();
            //        foreach (ListViewItem item in listfilterexcel.Items)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            if (listfilterexcel.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text))
            //            {
            //                list.Text = "";
            //                list.SubItems.Add(listfilterexcel.Items[item0].SubItems[1].Text);
            //                list.SubItems.Add(listfilterexcel.Items[item0].SubItems[2].Text);
            //                list.SubItems.Add(listfilterexcel.Items[item0].SubItems[3].Text);

            //                if (item0 % 2 == 0)
            //                {
            //                    list.BackColor = Color.White;

            //                }
            //                else
            //                {
            //                    list.BackColor = Color.WhiteSmoke;

            //                }



            //                listView1.Items.Add(list);
            //            }

            //            item0++;
            //        }
            //        lblemptot.Text = "Total Rows    :" + listView1.Items.Count;
            //    }
            //    else
            //    {
            //        try
            //        {
            //            listView1.Items.Clear();
            //            foreach (ListViewItem item in listfilterexcel.Items)
            //            {
            //                this.listView1.Items.Add((ListViewItem)item.Clone());
            //                item0++;
            //            }
            //            lblemptot.Text = "Total Rows    :" + listView1.Items.Count;
            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("---" + ex.ToString());
            //}
        }

        private void butfindtablename_Click(object sender, EventArgs e)
        {

        }

        private void listView2_ItemChecked(object sender, ItemCheckedEventArgs e)
        {


        }

      
        void list1Green(string s, string tbl)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";
                hed = null; hed = s.Split(',');
                string[] hed1 = fromquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s; DataTable dt = null;
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 1;

                    if (dt.Rows.Count > 0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed1.Length; k++)
                                {
                                    ws.Cells[1, k] = hed1[k].Replace("x.", "").ToUpper();
                                }
                                i = 2; j = 0;
                                ws.Range["A1", "H1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "H1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                k = 1; int l = 1;
                                foreach (ListViewItem item in listfilterscreen.Items)
                                {
                                    for (k = 1; k < item.SubItems.Count; k++)
                                    {
                                        ws.Cells[i, l] = item.SubItems[l].Text;
                                        l++;
                                    }
                                    l = 1;
                                    ws.Range["A" + i, "H" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();
                                    progressBar1.Value = j;
                                    i++;j++;

                                }
                                ws.Range["A" + i, "H" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }


        }
        void List1Red(string s,string tbl)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";              
                hed = null; hed = s.Split(',');
                string[] hed1 = fromquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s; DataTable dt = null;
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 1;

                    if (dt.Rows.Count > 0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed1.Length; k++)
                                {
                                    ws.Cells[1, k] = hed1[k].Replace("x.", "").ToUpper();

                                }
                                i = 2; j = 0;
                                ws.Range["A1", "H1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "H1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                k = 1; int l = 1;
                                foreach (ListViewItem item in listfilterred.Items)
                                {
                                    for (k = 1; k < item.SubItems.Count; k++)
                                    {
                                        ws.Cells[i, l] = item.SubItems[l].Text;
                                        l++;
                                    }
                                    l = 1;
                                    ws.Range["A" + i, "H" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                                    ws.Range["F" + i, "F" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                    ws.Range["F" + i, "F" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();
                                    progressBar1.Value = j;
                                    i++; j++;

                                }
                                ws.Range["A" + i, "H" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }

        }
        void list1All(string s, string tbl, string tbl2)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";
               
                hed = null; hed = s.Split(',');
                string[] hed1 = toquery.Split(','); string[] hed2 = fromquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s;
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                DataTable dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {                   
                   
                    i = 1;k = 0;
                    foreach (DataRow row in dt.Rows)
                    {                       

                        if (row["sts"].ToString() == "No")
                        {
                            if (row.ItemArray[5].ToString() == "")
                            {
                                unmatched = "No Data Found ,";
                            }
                            if (Convert.ToDateTime(row.ItemArray[1].ToString()).ToString("yyyy-MM-dd").Substring(0, 10) != "")
                            {
                                DataTable dt2 = null;
                                dt2 = checkdata(hed1[1].ToString(), tbl2, Convert.ToDateTime(dt.Rows[k]["date1"].ToString()).ToString("yyyy-MM-dd"));
                                if (dt2 != null) { unmatched += hed1[1].ToString().Replace("x.", "") + ","; }
                                else { }
                            }
                            if (row.ItemArray[5].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2=null;
                                if (hed1.Length == 8)
                                {                                   
                                    dt2 = checkdata(hed1[4].ToString(), tbl2,dt.Rows[k]["vchno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[4].ToString().Replace("x.", "") + ","; }
                                    else  { }
                                }
                                else
                                {                                    
                                    dt2 = checkdata(hed1[3].ToString(), tbl2, dt.Rows[k]["vchno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[3].ToString().Replace("x.", "") + ","; }
                                    else { }
                                }
                            }
                            if (row.ItemArray[6].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2;
                                if (hed1.Length == 8)
                                {
                                     dt2 = checkdata(hed1[6].ToString(), dt.Rows[k]["debit"].ToString(), tbl2,hed1[4].ToString(),dt.Rows[k]["vchno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[6].ToString().Replace("x.", "") + ","; }
                                    else { }
                                }
                                else
                                {
                                    dt2 = checkdata(hed1[7].ToString(), dt.Rows[k]["debit"].ToString(), tbl2, hed1[3].ToString(), dt.Rows[k]["vchno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[7].ToString().Replace("x.", "") + ","; }
                                    else { }
                                }
                            }
                            if (row.ItemArray[7].ToString() != "")
                            {
                                Class.Users.Query = "";
                                DataTable dt2;
                                if (hed1.Length == 8)
                                {
                                    dt2 = checkdata(hed1[7].ToString(), dt.Rows[k]["credit"].ToString(), tbl2, hed1[4].ToString(), dt.Rows[k]["vchno"].ToString());

                                    if (dt2 != null) { unmatched += hed1[7].ToString().Replace("x.", "") + ","; }
                                    else { }
                                }
                                else
                                {
                                    dt2 = checkdata(hed1[8].ToString(), dt.Rows[k]["credit"].ToString(), tbl2, hed1[3].ToString(), dt.Rows[k]["vchno"].ToString());

                                    if (dt2 != null) { unmatched += hed1[8].ToString().Replace("x.", "") + ""; }
                                    else { }
                                }
                            }
                            dt.Rows[k]["unmatched"] = unmatched;
                           
                            unmatched = "";
                           

                        }
                       
                        Class.Users.Query = ""; unmatched = "";
                        decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                        lblprogress1.Text = "" + (per).ToString("N0") + " %";
                        lblprogress1.Refresh();
                        progressBar1.Value = j;
                        i++; k++;

                  
                    }
                    if (dt.Rows.Count > 0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed2.Length; k++)
                                {
                                    ws.Cells[1, k] = hed2[k].Replace("x.", "").ToUpper();
                                }
                                i = 2; j = 0;
                                ws.Range["A1", "I1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "I1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                k = 1; int l = 1;
                                foreach (DataRow item in dt.Rows)
                                {                                    
                                    for (k = 1; k < item.ItemArray.Length; k++)
                                    {
                                        if (item["sts"].ToString() == "No")
                                        {

                                            ws.Range["A" + i, "I" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                                            ws.Range["H" + i, "H" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbHotPink;
                                            ws.Range["H" + i, "H" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                            ws.Cells[i, l] = item[l].ToString();
                                            l++;
                                        }
                                        else
                                        {
                                            ws.Cells[i, l] = item[l].ToString();
                                            l++;

                                            ws.Range["A" + i, "I" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                           
                                        }

                                    }
                                    //ws.Range["E" + i, "E" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbHotPink;
                                    //ws.Range["E" + i, "E" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                    l = 1;



                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();

                                    progressBar1.Value = j;
                                    i++; j++;

                                }
                                ws.Range["A" + i, "J" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }

        }
        void list2Green(string s, string tbl)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";
                hed = null; hed = s.Split(',');
                string[] hed1 = toquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s; DataTable dt = null;
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 1;
                    //if (dt.Rows.Count > 0)
                    //{
                    //    progressBar1.Minimum = 0;
                    //    progressBar1.Maximum = dt.Rows.Count;
                    //    i = 0;
                    //    using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                    //        if (sfd.ShowDialog() == DialogResult.OK)
                    //        {
                    //            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    //            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                    //            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                    //            app.Visible = false;
                    //            for (k = 1; k < hed1.Length; k++)
                    //            {
                    //                ws.Cells[1, k] = hed1[k].Replace("x.", "").ToUpper();
                    //            }
                    //            i = 2; j = 0;
                    //            ws.Range["A1", "J1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                    //            ws.Range["A1", "J1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                    //            k = 1; int l = 1;
                    //            foreach (DataRow item in dt.Rows)
                    //            {
                    //                for (k = 1; k < item.ItemArray.Length; k++)
                    //                {
                    //                    if (item["sts"].ToString() == "No")
                    //                    {
                    //                        ws.Range["A" + i, "J" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;

                    //                        ws.Cells[i, l] = item[l].ToString();
                    //                        l++;
                    //                    }
                    //                    else
                    //                    {
                    //                        ws.Range["A" + i, "J" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;

                    //                        ws.Cells[i, l] = item[l].ToString();
                    //                        l++;
                    //                    }
                    //                }
                    //                l = 1;
                    //                decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                    //                lblprogress1.Text = "" + (per).ToString("N0") + " %";
                    //                lblprogress1.Refresh();
                    //                progressBar1.Value = j;
                    //                i++; j++;

                    //            }
                    //            ws.Range["A" + i, "J" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                    //            ws.Columns.AutoFit();
                    //            wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                    //            app.Quit();
                    //            MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                    //        }
                    //}
                    if (dt.Rows.Count > 0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed1.Length; k++)
                                {
                                    ws.Cells[1, k] = hed1[k].Replace("x.", "").ToUpper();
                                }
                                i = 2; j = 0;
                                ws.Range["A1", "I1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "I1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                k = 1; int l = 1;
                                foreach (ListViewItem item in listfilterscreen1.Items)
                                {
                                    for (k = 1; k < item.SubItems.Count; k++)
                                    {
                                        ws.Cells[i, l] = item.SubItems[l].Text;
                                        l++;
                                    }
                                    l = 1;
                                    ws.Range["A" + i, "I" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();
                                    progressBar1.Value = j;
                                    i++;j++;

                                }
                                ws.Range["A" + i, "I" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }


        }
        void List2Red(string s, string tbl)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";
                hed = null; hed = s.Split(',');
                string[] hed1 = toquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s; DataTable dt = null;
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 1;

                    if (dt.Rows.Count > 0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed1.Length; k++)
                                {
                                    ws.Cells[1, k] = hed1[k].Replace("x.", "").ToUpper();
                                }
                                i = 2; j = 0;
                                ws.Range["A1", "I1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "I1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                k = 1; int l = 1;
                                foreach (ListViewItem item in listfilterred1.Items)
                                {
                                    for (k = 1; k < item.SubItems.Count; k++)
                                    {
                                        ws.Cells[i, l] = item.SubItems[l].Text;
                                        l++;
                                    }
                                    l = 1;
                                    ws.Range["A" + i, "I" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                                    ws.Range["G" + i, "G" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                    ws.Range["G" + i, "G" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();
                                    progressBar1.Value = j;
                                    i++;j++;

                                }
                                ws.Range["A" + i, "I" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }

        }
        void list2All(string s, string tbl, string tbl2)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";
                hed = null; hed = s.Split(',');
                string[] hed1 = fromquery.Split(','); string[] hed2 = toquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s; DataTable dt = null;
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {
                    i = 1;
                    foreach (DataRow row in dt.Rows)
                    {
                        unmatched = "";
                        if (row["sts"].ToString() == "No")
                        {
                            if (row.ItemArray[5].ToString() == "")
                            {
                                unmatched = "No Data Found ,";
                            }
                            if (Convert.ToDateTime(row.ItemArray[1].ToString()).ToString("yyyy-MM-dd").Substring(0, 10) != "")
                            {
                                DataTable dt2 = null;
                                dt2 = checkdata(hed1[1].ToString(), tbl, Convert.ToDateTime(dt.Rows[k]["date1"].ToString()).ToString("yyyy-MM-dd"));
                                if (dt2 != null) { unmatched += hed1[1].ToString().Replace("x.", "") + ","; }
                                else { }

                            }
                            if (row.ItemArray[5].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2 = null;
                                if (hed1.Length == 8)
                                {
                                    dt2 = checkdata(hed1[4].ToString(), tbl, dt.Rows[k]["refno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[4].ToString().Replace("x.", "") + ","; }
                                    else { }


                                }
                                else
                                {
                                    dt2 = checkdata(hed1[5].ToString(), tbl, dt.Rows[k]["refno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[5].ToString().Replace("x.", "") + ","; }
                                    else { }
                                }
                            }
                            if (dt.Rows[k]["debit"].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2 = null;
                                if (hed1.Length == 8)
                                {
                                    dt2 = checkdata(hed1[6].ToString(), dt.Rows[j]["debit"].ToString(), tbl, hed1[4].ToString(), dt.Rows[k]["refno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[6].ToString().Replace("x.", "") + ","; }
                                    else { }
                                    //Class.Users.Query = "select distinct " + hed1[6].ToString() + " from " + tbl + "  x where  " + hed1[4].ToString() + "='" + dt.Rows[j]["refno"].ToString() + "' and " + hed1[6].ToString() + "='" + dt.Rows[j]["debit"].ToString() + "';";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                    //dt2 = ds2.Tables[tbl];
                                    //if (dt2.Rows.Count <= 0)
                                    //{
                                    //    unmatched += hed1[6].ToString().Replace("x.", "") + ",";
                                    //}
                                    //else
                                    //{


                                    //}
                                }
                                else
                                {
                                    dt2 = checkdata(hed1[6].ToString(), dt.Rows[k]["debit"].ToString(), tbl, hed1[5].ToString(), dt.Rows[k]["refno"].ToString());
                                    if (dt2 != null) { unmatched += hed1[6].ToString().Replace("x.", "") + ","; }
                                    else { }
                                }
                            }
                            if (dt.Rows[k]["credit"].ToString() != "")
                            {
                                Class.Users.Query = "";
                                DataTable dt2=null;
                                    if (hed1.Length == 8)
                                    {
                                        dt2 = checkdata(hed1[6].ToString(), dt.Rows[k]["credit"].ToString(), tbl, hed1[4].ToString(), dt.Rows[k]["refno"].ToString());
                                        if (dt2 != null) { unmatched += hed1[6].ToString().Replace("x.", "") + ""; }
                                        else { }
                                       
                                    }
                                    else
                                    {
                                        dt2 = checkdata(hed1[7].ToString(), dt.Rows[k]["credit"].ToString(), tbl, hed1[5].ToString(), dt.Rows[k]["refno"].ToString());
                                        if (dt2 != null) { unmatched += hed1[7].ToString().Replace("x.", "") + ""; }
                                        else { }
                                        
                                    }
                            }

                           
                        }                   
                        if (hed1.Length == 8)
                        {
                            string un = hed1[5].Replace("x.", "");
                            dt.Rows[k][un] = unmatched;
                            
                        }
                        else
                        {
                            dt.Rows[k]["unmatched"] = unmatched;
                           
                        }
                                    
                        Class.Users.Query = ""; unmatched = "";
                        decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                        lblprogress1.Text = "" + (per).ToString("N0") + " %";
                        lblprogress1.Refresh();
                        progressBar1.Value = j;
                        i++; k++;                      
                    }
                    if (dt.Rows.Count > 0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed2.Length; k++)
                                {
                                    ws.Cells[1, k] = hed2[k].Replace("x.", "").ToUpper();
                                }
                                i = 2; j = 0;
                                ws.Range["A1", "J1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "J1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                k = 1; int l = 1;
                                foreach (DataRow item in dt.Rows)
                                {
                                    for (k = 1; k < item.ItemArray.Length; k++)
                                    {
                                        if (item["sts"].ToString() == "No")
                                        {
                                            ws.Range["A" + i, "J" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                                            ws.Range["I" + i, "I" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbHotPink;
                                            ws.Range["I" + i, "I" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                            ws.Cells[i, l] = item[l].ToString();
                                            l++;
                                        }
                                        else
                                        {
                                            ws.Range["A" + i, "J" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                            ws.Cells[i, l] = item[l].ToString();
                                            l++;
                                        }
                                    }
                                  
                                    l = 1;
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();
                                    progressBar1.Value = j;
                                    i++; j++;

                                }
                                ws.Range["A" + i, "J" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }

        }
        private DataTable checkdata(string param1,string tbl,string val)
        {

           string qry = "select distinct " + param1 + " from " + tbl + " x where " + param1 + "='" + val + "';";
            DataSet ds0 = Utility.ExecuteSelectQuery(qry, tbl);
            DataTable dt0 = ds0.Tables[tbl];
            return dt0;
        }
        private DataTable checkdata(string param1,string val1, string tbl, string param2,string val2)
        {

            string qry = "select distinct " + param1 + " from " + tbl + " x where " + param2 + "='" + val2 + "' and " + param1 + "='" + val1 + "';";
            DataSet ds0 = Utility.ExecuteSelectQuery(qry, tbl);
            DataTable dt0 = ds0.Tables[tbl];
            return dt0;
        }
        void listviewxl3(string s, string tbl, string tbl2,ListView listview)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";
                listview.Items.Clear(); listview.Columns.Clear();
                hed = null; hed = s.Split(',');
                string[] hed1 = toquery.Split(','); string[] hed2 = fromquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s;
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                DataTable dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {
                    listfilterred.Items.Clear(); listfilterscreen.Items.Clear();
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 0;
                    foreach (string heders in hed2)
                    {
                        listview.Columns.Add(heders.ToString().Replace("x.","").ToUpper());
                        if (i == 0)
                        {
                            listview.Columns[0].Width = 30;
                        }
                        if (i == 2)
                        {
                            listview.Columns[2].Width = 50;
                        }
                        if (i == 1 || i>2)
                        {
                            listview.Columns[i].Width = 100;
                        }
                        i++;
                    }
                    i = 1;
                        foreach (DataRow row in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();                       
                        list.Text = i.ToString();
                        list.SubItems.Add(row.ItemArray[1].ToString().Substring(0, 10));
                        list.SubItems.Add(row.ItemArray[2].ToString());
                        list.SubItems.Add(row.ItemArray[3].ToString());
                        list.SubItems.Add(row.ItemArray[4].ToString());
                        list.SubItems.Add(row.ItemArray[5].ToString());
                      
                        if (row["sts"].ToString() == "No")
                        {
                            if (row.ItemArray[5].ToString() == "")
                            {
                                unmatched = "No Data Found ,";
                            }
                            if (Convert.ToDateTime(row.ItemArray[1].ToString()).ToString("yyyy-MM-dd").Substring(0, 10) != "")
                            {
                                DataTable dt2 = null;
                                dt2 = checkdata(hed1[1].ToString(),tbl2, Convert.ToDateTime(dt.Rows[k]["date1"].ToString()).ToString("yyyy-MM-dd"));
                                //Class.Users.Query = "";
                                //Class.Users.Query = "select distinct " + hed1[1].ToString() + " from " + tbl2 + " x where " + hed1[1].ToString() + "='" + Convert.ToDateTime(dt.Rows[j]["date1"].ToString()).ToString("yyyy-MM-dd") + "';";
                                //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl2);
                                //DataTable dt2 = ds2.Tables[tbl2];
                                if (dt2 != null) { unmatched += hed1[1].ToString().Replace("x.", "") + ","; }
                                else
                                {


                                }
                            }
                            if (row.ItemArray[5].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2=null;
                                if (hed1.Length == 8)
                                {
                                    
                                    dt2 = checkdata(hed1[4].ToString(), tbl2, dt.Rows[k]["vchno"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[4].ToString() + " from " + tbl2 + "  x where  " + hed1[4].ToString() + "='" + dt.Rows[j]["vchno"].ToString() + "' ;";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl2);
                                    //dt2 = ds2.Tables[tbl2];
                                    if (dt2 != null) { unmatched += hed1[4].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }

                                }
                                else
                                {
                                    dt2 = checkdata(hed1[3].ToString(), tbl2, dt.Rows[k]["vchno"].ToString());
                                    //Class.Users.Query = "select distinct " + hed1[3].ToString() + " from " + tbl2 + "  x where  " + hed1[3].ToString() + "='" + dt.Rows[j]["vchno"].ToString() + "' ;";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl2);
                                    //dt2 = ds2.Tables[tbl2];

                                    if (dt2 != null) { unmatched += hed1[3].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                            }
                            if (row.ItemArray[7].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2;
                                if (hed1.Length == 8)
                                {
                                    
                                    dt2 = checkdata(hed1[4].ToString(), tbl2, dt.Rows[k]["vchno"].ToString(), hed1[6].ToString(), dt.Rows[k]["debit"].ToString());
                                    //Class.Users.Query = "select distinct " + hed1[4].ToString() + " from " + tbl2 + "  x where  " + hed1[4].ToString() + "='" + dt.Rows[j]["vchno"].ToString() + "' and " + hed1[6].ToString() + "='" + dt.Rows[j]["debit"].ToString() + "';";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl2);
                                    //dt2 = ds2.Tables[tbl2];
                                    if (dt2 != null) { unmatched += hed1[4].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                                else
                                {
                                    dt2 = checkdata(hed1[8].ToString(), dt.Rows[k]["debit"].ToString(),tbl2,hed1[3].ToString(), dt.Rows[k]["vchno"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[8].ToString() + " from " + tbl2 + " x where " + hed1[3].ToString() + "='" + dt.Rows[j]["vchno"].ToString() + "' and " + hed1[8].ToString() + "='" + dt.Rows[j]["debit"].ToString() + "'";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl2);
                                    //dt2 = ds2.Tables[tbl2];

                                    if (dt2 != null) { unmatched += hed1[8].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                            }
                            if (row.ItemArray[8].ToString() != "")
                            {
                                Class.Users.Query = "";
                                DataTable dt2;
                                if (hed1.Length == 8)
                                {
                                    dt2 = checkdata(hed1[7].ToString(), tbl2, dt.Rows[k]["vchno"].ToString(), hed1[4].ToString(), dt.Rows[k]["debit"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[4].ToString() + " from " + tbl2 + "  x where  " + hed1[4].ToString() + "='" + dt.Rows[j]["vchno"].ToString() + "' and " + hed1[7].ToString() + "='" + dt.Rows[j]["debit"].ToString() + "';";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl2);
                                    //dt2 = ds2.Tables[tbl2];
                                    if (dt2 != null) { unmatched += hed1[4].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                                else
                                {
                                    dt2 = checkdata(hed1[9].ToString(), dt.Rows[k]["credit"].ToString(), tbl2, hed1[3].ToString(), dt.Rows[k]["vchno"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[9].ToString() + " from " + tbl2 + " x where " + hed1[3].ToString() + "='" + dt.Rows[j]["vchno"].ToString() + "'  and " + hed1[9].ToString() + "='" + dt.Rows[j]["credit"].ToString() + "'";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl2);
                                    //dt2 = ds2.Tables[tbl2];

                                    if (dt2 != null) { unmatched += hed1[9].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                            }
                            dt.Rows[k]["unmatched"] = unmatched;
                            
                            list.ForeColor = Color.Red;
                            list.SubItems.Add(row["debit"].ToString());
                            list.SubItems.Add(row["credit"].ToString());
                            list.SubItems.Add(unmatched);
                            
                            listfilterred.Items.Add((ListViewItem)list.Clone());
                           
                        }
                        else
                        {
                            list.ForeColor = Color.DarkGreen;
                           
                            list.SubItems.Add(row["debit"].ToString());
                            list.SubItems.Add(row["credit"].ToString());
                            list.SubItems.Add("");
                            listfilterscreen.Items.Add((ListViewItem)list.Clone());
                           
                        }
                      
                        fdebitamt += Convert.ToDecimal("0" + row["Debit"].ToString());
                        fcreditamt += Convert.ToDecimal("0" + row["Credit"].ToString());
                       
                        listview.Items.Add(list);
                        unmatched = "";
                        lblFromcreditamount.Refresh(); lblFromdebitamount.Refresh();
                        Class.Users.Query = ""; unmatched = "";
                        decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                        lblprogress1.Text = "" + (per).ToString("N0") + " %";
                        lblprogress1.Refresh();
                        progressBar1.Value = j;
                        i++; k++;

                        //j++;
                    }
                    
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }

        }
        void listviewxl4(string s, string tbl, string tbl2,ListView listView)
        {
            try
            {
                Class.Users.UserTime = 0;
                int i = 0; int gridcount = 1, k = 0, m = 0; string unmatched = "";
                listView.Items.Clear(); listView.Columns.Clear();
                hed = null; hed = s.Split(',');
                string[] hed1 = fromquery.Split(','); string[] hed2 = toquery.Split(',');
                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = s; DataTable dt = null;
                 DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                dt = ds.Tables[tbl];
                i = 1; int j = 0;
                if (dt != null && tbl != "")
                {
                    listfilterred1.Items.Clear(); listfilterscreen1.Items.Clear();
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 0; 
                    foreach (string heders in hed2)
                    {
                        listView.Columns.Add(heders.ToString().Replace("x.", "").ToUpper());
                        if (i == 0)
                        {
                            listView.Columns[0].Width = 30;
                        }                        
                        else
                        {
                            listView.Columns[i].Width = 100;
                        }
                        i++;
                    }
                    i = 1;
                    foreach (DataRow row in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();

                        list.Text = i.ToString();
                        list.SubItems.Add(row.ItemArray[1].ToString().Substring(0, 10));
                        list.SubItems.Add(row.ItemArray[2].ToString());
                            list.SubItems.Add(row.ItemArray[3].ToString());
                            list.SubItems.Add(row.ItemArray[4].ToString());
                            list.SubItems.Add(row.ItemArray[5].ToString());
                            list.SubItems.Add(row.ItemArray[6].ToString());
                      

                        if (row["sts"].ToString() == "No")
                        {
                            if (row.ItemArray[5].ToString() == "")
                            {
                                unmatched = "No Data Found ,";
                            }
                            if (Convert.ToDateTime(row.ItemArray[1].ToString()).ToString("yyyy-MM-dd").Substring(0, 10) != "")
                            {
                                DataTable dt2 = null;
                                dt2 = checkdata(hed1[1].ToString(), tbl, Convert.ToDateTime(dt.Rows[k]["date1"].ToString()).ToString("yyyy-MM-dd"));

                                //Class.Users.Query = "";
                                //Class.Users.Query = "select distinct " + hed1[1].ToString() + " from " + tbl + " x where " + hed1[1].ToString() + "='" + Convert.ToDateTime(dt.Rows[k]["date1"].ToString()).ToString("yyyy-MM-dd") + "';";
                                //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                //DataTable dt2 = ds2.Tables[tbl];
                                if (dt2.Rows.Count <= 0)
                                {
                                    unmatched += hed1[1].ToString().Replace("x.", "") + ","; }
                                else
                                {
                                    
                                }
                            }
                            if (row.ItemArray[5].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2 = null;
                                if (hed1.Length == 8)
                                {
                                   
                                    dt2 = checkdata(hed1[4].ToString(), tbl,dt.Rows[k]["refno"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[4].ToString() + " from " + tbl + "  x where  " + hed1[4].ToString() + "='" + dt.Rows[k]["refno"].ToString() + "' ;";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                    //dt2 = ds2.Tables[tbl];
                                    if (dt2.Rows.Count <= 0)
                                    {
                                        unmatched += hed1[4].ToString().Replace("x.", "") + ","; }
                                    else
                                    {

                                       
                                    }

                                }
                                else
                                {
                                    dt2 = checkdata(hed1[5].ToString(), tbl, dt.Rows[k]["refno"].ToString());
                                    //Class.Users.Query = "select distinct " + hed1[5].ToString() + " from " + tbl + "  x where  " + hed1[5].ToString() + "='" + dt.Rows[k]["refno"].ToString() + "' ;";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                    //dt2 = ds2.Tables[tbl];

                                    if (dt2.Rows.Count<=0) {
                                        unmatched += hed1[5].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                            }
                            if (dt.Rows[k]["debit"].ToString() != "")
                            {
                                Class.Users.Query = ""; DataTable dt2;
                                if (hed1.Length == 8)
                                {
                                    dt2 = checkdata(hed1[6].ToString(), dt.Rows[k]["debit"].ToString(), tbl, hed1[4].ToString(), dt.Rows[k]["refno"].ToString());
                                    //Class.Users.Query = "select distinct " + hed1[6].ToString() + " from " + tbl + "  x where  " + hed1[4].ToString() + "='" + dt.Rows[j]["refno"].ToString() + "' and " + hed1[6].ToString() + "='" + dt.Rows[j]["debit"].ToString() + "';";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                    //dt2 = ds2.Tables[tbl];
                                    if (dt2.Rows.Count <= 0)
                                    {
                                        unmatched += hed1[6].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                                else
                                {
                                    dt2 = checkdata(hed1[6].ToString(), dt.Rows[k]["debit"].ToString(), tbl, hed1[5].ToString(), dt.Rows[k]["refno"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[7].ToString() + " from " + tbl + " x where " + hed1[5].ToString() + "='" + dt.Rows[k]["refno"].ToString() + "' and " + hed1[7].ToString() + "='" + dt.Rows[k]["debit"].ToString() + "'";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                    //dt2 = ds2.Tables[tbl];

                                    if (dt2.Rows.Count <= 0)
                                    {
                                         unmatched += hed1[7].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                            }
                            if (dt.Rows[k]["credit"].ToString() != "")
                            {
                                Class.Users.Query = "";
                                DataTable dt2;
                                if (hed1.Length == 8)
                                {
                                    dt2 = checkdata(hed1[7].ToString(), dt.Rows[k]["debit"].ToString(), tbl, hed1[4].ToString(), dt.Rows[k]["refno"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[7].ToString() + " from " + tbl + "  x where  " + hed1[4].ToString() + "='" + dt.Rows[k]["refno"].ToString() + "' and " + hed1[7].ToString() + "='" + dt.Rows[k]["debit"].ToString() + "';";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                    //dt2 = ds2.Tables[tbl];
                                    if (dt2.Rows.Count <= 0)
                                    {
                                         unmatched += hed1[7].ToString().Replace("x.", ""); }
                                    else
                                    {

                                    }
                                }
                                else
                                {
                                    dt2 = checkdata(hed1[7].ToString(), dt.Rows[k]["credit"].ToString(), tbl, hed1[5].ToString(), dt.Rows[k]["refno"].ToString());

                                    //Class.Users.Query = "select distinct " + hed1[8].ToString() + " from " + tbl + " x where " + hed1[5].ToString() + "='" + dt.Rows[k]["refno"].ToString() + "'  and " + hed1[8].ToString() + "='" + dt.Rows[k]["credit"].ToString() + "'";
                                    //DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                                    //dt2 = ds2.Tables[tbl];

                                    if (dt2.Rows.Count <= 0) { unmatched += hed1[8].ToString().Replace("x.", "") + ","; }
                                    else
                                    {


                                    }
                                }
                            }

                            
                            list.ForeColor = Color.Red;
                            list.SubItems.Add(row["Debit"].ToString());
                            list.SubItems.Add(row["Credit"].ToString());
                            list.SubItems.Add(unmatched);
                            listfilterred1.Items.Add((ListViewItem)list.Clone());

                        }
                        else
                        {
                            list.ForeColor = Color.DarkGreen;
                          
                            list.SubItems.Add(row["Debit"].ToString());
                            list.SubItems.Add(row["Credit"].ToString());
                            list.SubItems.Add("");
                            listfilterscreen1.Items.Add((ListViewItem)list.Clone());

                        }
                        if (hed1.Length == 8)
                        {
                            string un = hed1[5].Replace("x.", "");
                            dt.Rows[k][un] = unmatched;
                        }
                        else
                        {
                            dt.Rows[k]["unmatched"] = unmatched;
                        }
                        unmatched = "";

                        fdebitamt += Convert.ToDecimal("0" + row["Debit"].ToString());
                        fcreditamt += Convert.ToDecimal("0" + row["Credit"].ToString());
                      
                        listView.Items.Add(list);


                        lblFromcreditamount.Refresh(); lblFromdebitamount.Refresh();
                        Class.Users.Query = ""; unmatched = "";

                        decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                        lblprogress1.Text = "" + (per).ToString("N0") + " %";
                        lblprogress1.Refresh();
                        progressBar1.Value = j;
                        i++; k++;

                        //j++;
                    }
                   
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

        }
        private void butimport_Click(object sender, EventArgs e)
        {

        }


        private void listView1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //Class.Users.UserTime = 0;
            //if (e.Item.Checked == true)
            //{
            //    ListViewItem it = new ListViewItem();
            //    e.Item.SubItems[3].Text = "✔";
            //    e.Item.SubItems[3].ForeColor = Color.Green;
            //    it.SubItems.Add(e.Item.SubItems[2].Text);
            //    it.SubItems.Add(e.Item.Checked.ToString());


            //    allip3.Items.Add(it);
            //    lblfieldcount.Text = "Selected Table Columns Count  :  " + allip3.Items.Count;
            //}
            //if (e.Item.Checked == false && e.Item.SubItems[3].Text == "✔")
            //{
            //    e.Item.SubItems[3].Text = "✖";
            //    e.Item.SubItems[3].ForeColor = Color.DarkRed;
            //    e.Item.Checked = false;
            //    for (int c = 0; c < allip3.Items.Count; c++)
            //    {
            //        if (allip3.Items[c].SubItems[1].Text == e.Item.SubItems[2].Text)
            //        {
            //            allip3.Items[c].Remove();
            //            c--;
            //            lblfieldcount.Text = "Selected Table Columns Count  :  " + allip3.Items.Count;
            //        }
            //    }
        }



        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {

        }
        void TableLoad()
        {
            if (buttosupplier.Text != "")
            {
                string sel1 = "show  tables  where Tables_in_" + Class.Users.ProjectID + " like'%my_%' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, buttosupplier.Text.ToLower().Trim());
                DataTable dt = ds.Tables[buttosupplier.Text.ToLower().Trim()];
                combodrop.DataSource = null;
                if (dt.Rows.Count > 0 || dt != null)
                {

                    combodrop.DataSource = dt;
                    combodrop.DisplayMember = "Tables_in_" + Class.Users.ProjectID + "";
                    combodrop.ValueMember = "Tables_in_" + Class.Users.ProjectID + "";
                }
            }
        }
        decimal fdebitamt, fcreditamt, tdebitamt, tcreditamt = 0;
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Class.Users.CompCode1 = "";
            
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                if (Class.Users.HUserName == "VAIRAM")
                {
                    textBox1.Visible = true;
                }
                else
                {
                    textBox1.Visible = false;
                }
            }
            
        }

        private void chkall_CheckedChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int i = 0;
            //    if (chkall.Checked == true)
            //    {
            //        for (i = 0; i < listView1.Items.Count; i++)
            //        {
            //            listView1.Items[i].Checked = true;
            //        }
            //    }
            //    if (chkall.Checked == false)
            //    {
            //        for (i = 0; i < listView1.Items.Count; i++)
            //        {
            //            listView1.Items[i].Checked = false;
            //        }

            //        fromtable = ""; allip3.Items.Clear();

            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        private void checkallgrid_CheckedChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    int i = 0;
            //    if (checkallgrid.Checked == true)
            //    {
            //        for (i = 0; i < listView4.Items.Count; i++)
            //        {
            //            listView4.Items[i].Checked = true;
            //        }
            //    }
            //    if (checkallgrid.Checked == false)
            //    {
            //        for (i = 0; i < listView4.Items.Count; i++)
            //        {
            //            listView4.Items[i].Checked = false;
            //        }

            //        fromgridtable = ""; allip3grid.Items.Clear();

            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        private void listView4_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //if (e.Item.Checked == true)
            //{
            //    ListViewItem it = new ListViewItem();
            //    e.Item.SubItems[3].Text = "✔";
            //    e.Item.SubItems[3].ForeColor = Color.Green;
            //    it.SubItems.Add(e.Item.SubItems[2].Text);
            //    it.SubItems.Add(e.Item.Checked.ToString());


            //    allip3grid.Items.Add(it);
            //}
            //if (e.Item.Checked == false && e.Item.SubItems[3].Text == "✔")
            //{
            //    e.Item.SubItems[3].Text = "✖";
            //    e.Item.SubItems[3].ForeColor = Color.DarkRed;
            //    e.Item.Checked = false;
            //    for (int c = 0; c < allip3grid.Items.Count; c++)
            //    {
            //        if (allip3grid.Items[c].SubItems[1].Text == e.Item.SubItems[2].Text)
            //        {
            //            allip3grid.Items[c].Remove();
            //            c--;
            //        }
            //    }
            //}
        }

        private void lvlogsgrid_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            //if (e.Item.Checked == true)
            //{
            //    ListViewItem it = new ListViewItem();

            //    e.Item.SubItems[3].Text = "✔";
            //    e.Item.SubItems[3].ForeColor = Color.Green;
            //   Class.Users.TableNameGrid = "";
            //    it.SubItems.Add(e.Item.SubItems[2].Text);
            //    it.SubItems.Add(e.Item.Checked.ToString());
            //    lbltablenamegrid.Text = "";
            //    lbltablenamegrid.Text = "Table Name :-  " + e.Item.SubItems[2].Text;              
            //    Class.Users.TableNameGrid = e.Item.SubItems[2].Text;
            //    string sel = "select * from " + e.Item.SubItems[2].Text + ";";
            //    DataSet ds = Utility.ExecuteSelectQuery(sel, Class.Users.TableNameGrid);
            //    DataTable dttbl = ds.Tables[Class.Users.TableNameGrid];
            //    int i = 1;
            //    for (int j = 0; j < dttbl.Columns.Count; j++)
            //    {
            //        ListViewItem ittbl = new ListViewItem();
            //        ittbl.Text = "";
            //        ittbl.SubItems.Add(i.ToString());
            //        ittbl.SubItems.Add(dttbl.Columns[j].ToString());
            //        ittbl.SubItems.Add("");

            //        if (i % 2 == 0)
            //        {
            //            ittbl.BackColor = Color.White;
            //        }
            //        else
            //        {
            //            ittbl.BackColor = Color.WhiteSmoke;
            //        }

            //        i++;


            //        listView4.Items.Add(ittbl);
            //    }
            //    // allip4.Items.Add(it);
            //}
            //if (e.Item.Checked == false)
            //{
            //    listView4.Items.Clear();
            //    e.Item.SubItems[3].Text = "";
            //    e.Item.SubItems[3].ForeColor = Color.DarkRed;
            //    e.Item.Checked = false;
            //    //for (int c = 0; c < allip4.Items.Count; c++)
            //    //{

            //    //    if (lvlogs.Items[c].SubItems[2].Text == e.Item.SubItems[2].Text)
            //    //    {
            //    //        allip4.Items[c].Remove();
            //    //        c--;
            //    //    }
            //    //}
            //}
        }

        private void butNonGridTable_Click(object sender, EventArgs e)
        {

        }
        // DataTable dtgridview1 = new DataTable();
        private void butGridTable_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {




        }

        private void txtsearchtable_TextChanged(object sender, EventArgs e)
        {

        }

        private void tableNameToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        // bool dbcol = false;
        void GridColumnRefresh(DataGridView grid, int idx)
        {

            if (grid.Rows[idx].Cells[2].FormattedValue.ToString() != "" && grid.Columns[2].Name == "FromTable1")
            {
                int lines = idx;
                //lblfieldcount.Refresh();
                //lblfieldcount.Text = lines.ToString();
                if (lines % 2 == 0)
                {
                    grid.Rows[idx].DefaultCellStyle.BackColor = Color.AntiqueWhite;
                    // Class.Users.TableName = grid.Rows[idx].Cells[2].FormattedValue.ToString();

                }
                else
                {
                    grid.Rows[idx].DefaultCellStyle.BackColor = Color.LightPink;
                    //Class.Users.TableNameGrid = grid.Rows[idx].Cells[2].FormattedValue.ToString();

                }
                string tablenames = dataGridView2.Rows[idx].Cells[2].EditedFormattedValue.ToString();
                TableGridLoadd("my_" + tablenames);
                TableGridLoad("to_" + tablenames);
            }

        }

        void GridCellClear(DataGridView grid, int idx)
        {
            //if (grid.Rows[idx].Cells[4].FormattedValue.ToString() != "")
            //{

            grid.Rows[idx].Cells[5].Value = string.Empty;
            grid.Rows[idx].Cells[4].Value = string.Empty;
            celcrear = false;
            //}
            //else
            //{

            //    grid.Rows[idx].Cells[3].Value = string.Empty;
            //    celcrear = false;
            //}

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //this.dataGridView2.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        private void txtslugsearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabControl3_SelectedIndexChanged(object sender, EventArgs e)
        {


        }


        private void dataGridView3_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //this.dataGridView3.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }



        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == 2)
            //{
            //    dataGridView3.Rows[0].Cells[4].Value = DBNull.Value;
            //    Prefix.DataSource = null;
            //    string sel = "SELECT  DISTINCT A.PREFIX FROM " + Class.Users.ProjectID + ".AUTOGENERATE A  WHERE   A.TX_VIEW_ID='" + Class.Users.TableNameGrid + "' AND  A.PROJECTID='" + Class.Users.ProjectID + "'  AND  A.DESCRIPTION='" + dataGridView3.Rows[0].Cells[2].FormattedValue.ToString() + "'  ORDER BY 1";
            //    DataSet ds2 = Utility.ExecuteSelectQuery(sel, "AUTOGENERATE");
            //    DataTable dt2 = ds2.Tables["AUTOGENERATE"];
            //    if (dt2.Rows.Count > 0)
            //    {
            //        Prefix.DataSource = dt2;
            //        Prefix.DisplayMember = "PREFIX";
            //        Prefix.ValueMember = "PREFIX";
            //    }
            //    if (Class.Users.TableNameGrid == "")
            //    {
            //        Class.Users.TableNameGrid = dataGridView3.Rows[0].Cells[1].FormattedValue.ToString();
            //    }
            //}
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {



        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Class.Users.UserTime = 0;





            }
            catch (Exception EX) { }
            Cursor = Cursors.Default;
        }

        public void Searchs(int id)
        {
            throw new NotImplementedException();
        }

        public void Deletes(int id)
        {
            throw new NotImplementedException();
        }

        private void listViewRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //listView1.Items.Clear();

        }

        private void dBColumnRefresToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //dbcol = true;
        }

        private void lvlogs_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }
        bool celcrear = false;
        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            celcrear = true;


        }

        private void dataGridView2_DataError_1(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void butsubmit_Click(object sender, EventArgs e)
        {

        }

        private void FilterRedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listfilterred.Items.Count > 0)
            {

                MessageBox.Show("MyCompany- UnMatched Rows are : " + listfilterred.Items.Count.ToString());

            }
        }

        private void FilterGreenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listfilterscreen.Items.Count > 0)
            {

                MessageBox.Show("MyCompany- Matched Rows are : " + listfilterscreen.Items.Count.ToString());
            }
        }

        private void filterRedToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (listfilterred1.Items.Count > 0)
            {

                MessageBox.Show("Customer or Supplier -UnMatched Rows are : " + listfilterred1.Items.Count.ToString());

            }
        }

        private void filterGreenToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (listfilterscreen1.Items.Count > 0)
            {

                MessageBox.Show("Customer or Supplier - Matched Rows are : " + listfilterscreen1.Items.Count.ToString());

            }
        }
      
       
        string toquery = "", fromquery = "";

       
        private void exportToXLRedToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
            {
                string[] split = textBox1.Text.Split(';');
                string sel = "select " + toquery + ",x.sts from(" + split[1] + ") x  where x.sts='No'";
                List2Red(sel,"to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
            }
            else { MessageBox.Show("No data Found", "Null"); }
        }

        private void exportToXLGreenToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
            {
                string[] split = textBox1.Text.Split(';');
                string sel = "select " + toquery + ",x.sts from(" + split[1] + ") x  where x.sts='Yes'";
                list2Green(sel, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
            }
            else { MessageBox.Show("No data Found", "Null"); }
        }
        private void allToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
            {
                string[] split = textBox1.Text.Split(';');
                string sel = "select " + toquery + ",x.sts from(" + split[1] + ") x";
                list2All(sel, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString(), "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());

            }
            else { MessageBox.Show("No data Found", "Null"); }
        }
        private void exportToXLRedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
            {
                string[] split = textBox1.Text.Split(';');
                string sel = "select " + fromquery + ",x.sts from(" + split[0] + ") x  where x.sts='No'";
                List1Red(sel, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());

            }
            else { MessageBox.Show("No data Found", "Null"); }
        }

        private void exportToXLGreenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
            {
                string[] split = textBox1.Text.Split(';');
                string sel = "select " + fromquery + ",x.sts from(" + split[0] + ") x  where x.sts='Yes'";
                list1Green(sel, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
            }           
        }
        private void allToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
            {
                string[] split = textBox1.Text.Split(';');
                string sel = "select " + fromquery + ",x.sts from(" + split[0] + ") x";
                list1All(sel, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString(), "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());

            }
            else { MessageBox.Show("No data Found","Null"); }
        }
        string[] hed;
        private void unMachedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

            try
            {
                //listfilterred.Items.Clear(); listfilterscreen.Items.Clear();
                //listfilterred1.Items.Clear(); listfilterscreen1.Items.Clear();
                Class.Users.UserTime = 0;

                //listView1.Items.Clear(); listView2.Items.Clear();
                int i = 0; int gridcount = 1, k = 0; string unmatched = "";
                
                hed = null; hed = fromquery.Split(',');
                string[] split = textBox1.Text.Split(';');               
                Class.Users.Query = "";
                Class.Users.Query = "select " + fromquery + ",x.sts from(" + split[0] + ") x where x.sts='No'";
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                DataTable dt = ds.Tables["my_"+ dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                i = 1; int j = 0; 
                if (dt != null && dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
                {
                     listfilterred.Items.Clear(); listfilterscreen.Items.Clear();
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 1;
                    foreach (DataRow row in dt.Rows)
                    {
                       
                        lblFromcreditamount.Refresh(); lblFromdebitamount.Refresh();                     
                        Class.Users.Query = ""; unmatched = "";              
                        if (row.ItemArray[5].ToString() == "")
                        {
                            unmatched = "No Data Found ,";
                        }
                        if (Convert.ToDateTime(row.ItemArray[1].ToString()).ToString("yyyy-MM-dd").Substring(0,10) != "")
                        {
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.date1 from to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() + " a where a.date1='" + Convert.ToDateTime(dt.Rows[j]["date1"].ToString()).ToString("yyyy-MM-dd") + "';";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "date,";
                            }
                        }
                        if (row.ItemArray[5].ToString() != "")
                        {
                          Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.refno from to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() + " a where  a.refno='" + dt.Rows[j]["vchno"].ToString() + "';";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "vchno,";
                            }
                        }
                        if (row.ItemArray[7].ToString() != "")
                        {
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.Debit from to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()+" a where a.refno='" + dt.Rows[j]["vchno"].ToString() + "' and a.debit='" + dt.Rows[j]["debit"].ToString() + "'";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "debit,";
                            }
                        }
                        if (row.ItemArray[8].ToString() != "")
                        {
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.credit from to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() + " a where a.refno='" + dt.Rows[j]["vchno"].ToString() + "' and a.debit='" + dt.Rows[j]["debit"].ToString() + "' and a.credit='" + dt.Rows[j]["credit"].ToString() + "'";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "credit";
                            }
                        }
                        dt.Rows[j]["unmatched"] = unmatched;
                       
                        unmatched = "";
                       
                        fdebitamt += Convert.ToDecimal("0" + row["Debit"].ToString());
                        fcreditamt += Convert.ToDecimal("0" + row["Credit"].ToString());                                     
                       
                        i++; j++;
                    }
                    if (dt.Rows.Count>0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed.Length; k++)
                                {
                                    ws.Cells[1, k] = hed[k].Replace("x.","").ToUpper();
                                }                               
                                i = 2;  j = 0;
                                ws.Range["A1", "J1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "J1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;

                                k = 1; int l = 1;
                                foreach (DataRow item in dt.Rows)
                                {
                                    for (k = 1; k < item.ItemArray.Length; k++)
                                    {
                                        ws.Cells[i, l] = item[l].ToString();
                                        l++;
                                        //ws.Cells[i, 1] = item.ItemArray[1].ToString().Substring(0, 10);
                                    }
                                    l = 1;
                                    
                                    ws.Range["A" + i, "J" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                                   
                                  
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();

                                    progressBar1.Value = j;
                                    i++; j++;

                                }
                                ws.Range["A" + i, "J" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }
        }

        string unmatched = "";
        private void unMatchedToolStripMenuItem_Click(object sender, EventArgs e)
        {            

            try
            {
              
                Class.Users.UserTime = 0;

                
                int i = 0; int gridcount = 1;



                string[] split = textBox1.Text.Split(';');
                Class.Users.Query = "";
                Class.Users.Query = "select " + toquery + ",x.sts from(" + split[1] + ") x where x.sts='No'";
                DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                DataTable dt = ds.Tables["to_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                i = 1; int j = 0,k=0;
                string[] hed;
                hed = null; hed = toquery.Split(',');
               
                 if (dt != null && dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
                {
                    listfilterred.Items.Clear(); listfilterscreen.Items.Clear();
                    lblFromcreditamount.Text = ""; lblFromdebitamount.Text = ""; fdebitamt = 0; fcreditamt = 0;
                    lblTocreditamount.Text = ""; lblTodebitamount.Text = ""; tdebitamt = 0; tcreditamt = 0;
                    i = 1;
                    foreach (DataRow row in dt.Rows)
                    {

                        lblFromcreditamount.Refresh(); lblFromdebitamount.Refresh();
                        Class.Users.Query = ""; unmatched = "";
                        if (dt.Rows[j].ItemArray[2].ToString()  == "" || dt.Rows[j].ItemArray[3].ToString() == "")
                        {
                            unmatched = "No Data Found ,";

                        }
                        if (dt.Rows[j].ItemArray[1].ToString().Substring(0, 10) != "")
                        {
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.date1 from my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() + " a where a.date1='" + Convert.ToDateTime(dt.Rows[j]["date1"].ToString()).ToString("yyyy-MM-dd") + "';";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "date,";
                            }
                        }
                        if (dt.Rows[j].ItemArray[3].ToString().ToString() != "")
                        {
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.vchno from my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() + " a where  a.vchno='" + dt.Rows[j]["refno"].ToString() + "';";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "refno,";
                            }
                        }
                        if (dt.Rows[j]["debit"].ToString() != "")
                        {
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.Debit from my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() + " a where a.vchno='" + dt.Rows[j]["refno"].ToString() + "' and a.debit='" + dt.Rows[j]["debit"].ToString() + "'";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "debit,";
                            }
                        }
                        if (dt.Rows[j]["credit"].ToString() != "")
                        {
                            Class.Users.Query = "";
                            Class.Users.Query = "select distinct a.credit from my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() + " a where a.vchno='" + dt.Rows[j]["refno"].ToString() + "' and a.debit='" + dt.Rows[j]["debit"].ToString() + "' and a.credit='" + dt.Rows[j]["credit"].ToString() + "'";
                            DataSet ds2 = Utility.ExecuteSelectQuery(Class.Users.Query, "my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString());
                            DataTable dt2 = ds2.Tables["my_" + dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString()];
                            if (dt2.Rows.Count > 0) { }
                            else
                            {
                                unmatched += "credit";
                            }
                        }
                        dt.Rows[j]["unmatched"] = unmatched;

                        unmatched = "";

                        fdebitamt += Convert.ToDecimal("0" + row["Debit"].ToString());
                        fcreditamt += Convert.ToDecimal("0" + row["Credit"].ToString());


                        i++; j++;
                    }
                    if (dt.Rows.Count > 0)
                    {
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = dt.Rows.Count;
                        i = 0;
                        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
                                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)app.ActiveSheet;
                                app.Visible = false;
                                for (k = 1; k < hed.Length; k++)
                                {
                                    ws.Cells[1, k] = hed[k].Replace("x.", "").ToUpper() ;
                                    //ws.Cells[1, i] = "Date";
                                }
                                i = 2; j = 0;
                                ws.Range["A1", "J1"].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Range["A1", "J1"].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite;
                                k = 1;int l = 1;
                                foreach (DataRow item in dt.Rows)
                                {
                                    for (k = 1; k < item.ItemArray.Length; k++)
                                    {                                       
                                        ws.Cells[i, l] = item[l].ToString();
                                        l++;
                                        //ws.Cells[i, 1] = item.ItemArray[1].ToString().Substring(0, 10);
                                    }
                                    l =1;
                                   
                                    ws.Range["A" + i, "J" + i].Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;

                                   // k++;
                                    decimal per = Convert.ToDecimal(100 / GenFun.ToInt(dt.Rows.Count)) * (j + 1);
                                    lblprogress1.Text = "" + (per).ToString("N0") + " %";
                                    lblprogress1.Refresh();

                                    progressBar1.Value = j;
                                    i++; j++;

                                }
                                ws.Range["A" + i, "J" + i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbNavy;
                                ws.Columns.AutoFit();
                                wb.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, (Microsoft.Office.Interop.Excel.XlSaveAsAccessMode)Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                                app.Quit();
                                MessageBox.Show("Completed"); progressBar1.Maximum = 0;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Data Found in MyCompany");
                }

            }

            catch (Exception ex) { ex.ToString(); }
        }

        private void all2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //if (dataGridView2.Rows[0].Cells[2].EditedFormattedValue.ToString() != "")
            //{
              
               
            //    string[] split = textBox1.Text.Split(';');
            //    string sel = "select " + fromquery + ",x.sts from(" + split[0] + ") x";
               
            //}
            //else { MessageBox.Show("No data Found", "Null"); }
        }
        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void butdroptable_Click(object sender, EventArgs e)
        {
            try
            {
                string tabl1 = "", sel3 = "", sel4 = "", tabl2 = "";
                tabl2 = combodrop.Text;
                tabl1 = combodrop.Text.Replace("my", "to");
                sel4 = "drop table " + tabl2 + "";
                Utility.ExecuteNonQuery(sel4);
                sel3 = "drop table " + tabl1 + "";
                Utility.ExecuteNonQuery(sel3);
                
                MessageBox.Show("Table Droped " + sel3.ToString() + "--" + sel4.ToString());
                combodrop.Text = "";
                GridLoad(); TableLoad(); 
            }catch(Exception ex)
            {

            }
        }

        private void butmycompany_Click(object sender, EventArgs e)
        {
            if (Class.Users.CompCode1 != "")
            {
                
                GridView(dataGridView1, butmycompany);
                Class.Users.CompCode1 = "";
            }
            else
            {
                MessageBox.Show("Pls Select Customer or Supplier Name ","Customer");
            }
           
        }


        private void buttosupplier_Click(object sender, EventArgs e)
        {
            if (Class.Users.CompCode1 != "")
            {
                GridView(dataGridView3, buttosupplier);
                Class.Users.CompCode1 = "";
            }
            else
            {
                MessageBox.Show("Pls Select Customer or Supplier Name ", "Customer");
            }

        }
        private void button2_Click_2(object sender, EventArgs e)
        {

        }
    }
}
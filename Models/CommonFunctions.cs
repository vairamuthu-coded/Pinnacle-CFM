using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle
{

    public static class CommonFunctions
    {
        public static string query { get; set; }
        public static DataSet ds { get; set; }
        public static DataTable dt { get; set; }
        public static DataTable dt1 { get; set; }
        public static DataTable select(string qry, string tbl)
        {
            query = qry;
            ds = Utility.ExecuteSelectQuery(query, tbl);
            dt = ds.Tables[tbl];
            return dt;
        }
       
        public static DataSet select1(string qry, string tbl)
        {
            query = qry;
            ds = Utility.ExecuteSelectQuery(query, tbl);
            return ds;
        }
        public static void SetRowNumber(DataGridView Grid)
        {
            int rowNumber = 1;
            Grid.Font = Class.Users.FontName;
            Grid.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            Grid.ColumnHeadersDefaultCellStyle.Font = Class.Users.FontName;
            Grid.ColumnHeadersDefaultCellStyle.Font= new Font(Class.Users.FontName.FontFamily, Class.Users.FontName.Size, FontStyle.Bold);
            Grid.RowTemplate.DefaultCellStyle.Font = Class.Users.FontName;
            Grid.RowTemplate.DefaultCellStyle.ForeColor = Color.Black;
            Grid.ColumnHeadersDefaultCellStyle.ForeColor = Class.Users.Color1;
            Grid.DefaultCellStyle.ForeColor = Class.Users.BackColors;
            Grid.RowHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            Grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;//AutoSizeToFirstHeader

            foreach (DataGridViewRow Row in Grid.Rows)
            {
                if (Row.IsNewRow) continue;
                Row.HeaderCell.Value = rowNumber.ToString();
                Row.DefaultCellStyle.BackColor = rowNumber % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                rowNumber = rowNumber + 1;
            }
            Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
           
        }
        public static void RemoveColumn(DataGridView Grid, string[] s)
        {
            foreach (string str in s)
            {
                Grid.Columns[str.ToString()].Visible = false;
            }
        }
        //public static void AddColumn(DataGridView Grid, string[] g, string[] s, string[] ss, int cnt)
        //{
        //    Grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders;//AutoSizeToFirstHeader

        //    int index2 = Grid.Rows.Count - 1;
        //    int index1 = Grid.Columns.Count - 1;
        //    int i = 0;
        //    if (index1 <= 0)
        //    {
        //        Grid.Columns.Add("", "ColorName");
        //        foreach (string str in s)
        //        {
        //            if (index1 <= 0)
        //            {
        //                Grid.Rows.Add(s[i].ToString(), s[i].ToString());
        //            }
        //            i++;
        //        }
        //        i = 0;
        //        foreach (string str1 in ss)
        //        {
        //            if (index1 <= 0)
        //            {
        //                Grid.Columns.Add(ss[i].ToString(), ss[i].ToString());
        //            }
        //            i++;

        //        }
        //    }
        //    if (index2 > 0)
        //    {
        //        for (i = 0; i < ss.Length; i++)
        //        {
        //            for (int j = index2; j <= i; j++)
        //            {
        //                if (i == j)
        //                {
        //                    Grid.Columns.Add("", ss[i].ToString());
        //                }
        //            }
        //        }
        //    }
        //}

        public static void AddColumn(DataGridView Grid, string[] g, string[] s, string[] ss, Int32[] width, int cnt)
        {
            
            Grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;//AutoSizeToFirstHeader
            int grow = Grid.Rows.Count;
            int gcol = Grid.Columns.Count;
            int wid = width.Length;
            int i = 0, j = 0;
            if (gcol <= 0)
            {
                i = 0;
                foreach (string str1 in ss)
                {
                    if (gcol <= 0)
                    {
                        Grid.Columns.Add(ss[i].ToString(), ss[i].ToString());
                        if (wid > i)
                        {
                            Grid.Columns[i].Width = width[i];
                            Grid.Columns[0].ReadOnly = true;
                        }
                        //else if (Grid.Columns[i].HeaderText == "REMARKS")
                        //{
                        //    Grid.Columns[i].Width = 120;
                        //}
                        else
                        {
                            Grid.Columns[i].Width = 50;
                        }
                    }
                    i++;
                }
                i = 0;
                foreach (string str in s)
                {
                    if (gcol <= 0)
                    {
                        Grid.Rows.Add(s[i].ToString(), "");
                    }
                    i++;
                }
               
               
            }
            if (gcol > 0)
            {
                //Color
                i = 0;j = 0;
                for (i = 0; i < s.Length; i++)
                {
                    for (j = grow; j <= i; j++)
                    {
                        if (i == j)
                        {
                            Grid.Rows.Add(s[i].ToString(), "");
                        }
                    }
                }

                //Size
                i = 0;j = 0;
                for (i = 0; i < ss.Length; i++)
                {
                    for (j = gcol; j <= i; j++)
                    {
                        if (i == j)
                        {
                            Grid.Columns.Add("", ss[i].ToString());
                        }
                        //if (Grid.Columns[j].HeaderText == "REMARKS")
                        //{
                        //    Grid.Columns[i].Width = 120;
                        //}
                    }
                }
            }
          
            Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);//AutoSizeToAllHeaders
            SetRowNumber(Grid);
        }
        public static void AddColumnDynamic(DataGridView Grid, string[] s, string[] ss, string query, string[] hide, Int32[] width)
        {

           // Grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;//
            int grow = Grid.Rows.Count;
            int gcol = Grid.Columns.Count;
            int wid = width.Length;
            int i = 1, j = 0, cnt = 0; bool check = false;
            DataSet ds4 = Utility.ExecuteSelectQuery(query, Class.Users.TableNameGrid);
            DataTable dt4 = ds4.Tables[Class.Users.TableNameGrid];
            i = 0;
            if (gcol == 1 || gcol <= 0)
            {
                foreach (DataColumn str1 in dt4.Columns)
                {
                    Grid.Columns.Add(i.ToString(), str1.ToString().ToUpper());
                   
                        if (wid > i)
                        {
                            Grid.Columns[i].Width = width[i];
                        }
                    
                    i++;
                }
            }
            if (ss.Length > 0 || Grid.Columns.Count > 1)
            {
                i = 0; j = 0; cnt = 1;
                if (dt4.Columns.Count < Grid.Columns.Count)
                {
                    check = true;
                }
                foreach (DataRow row in dt4.Rows)
                {

                    Grid.Rows.Add();
                    if (check == false)
                    {

                        for (j = 0; j < Grid.Columns.Count; j++)
                        {
                            Grid.Rows[i].Cells[j].Value = row[j].ToString();
                            Grid.Rows[i].DefaultCellStyle.Font = new Font(Class.Users.FontName.FontFamily, Class.Users.FontName.Size, FontStyle.Bold);

                        }
                    }
                    else
                    {
                        for (j = 0; j < Grid.Columns.Count - 1; j++)
                        {
                            Grid.Rows[i].Cells[j + 1].Value = row[j].ToString();
                        }
                    }
                    i++;
                }
                i = 0; j = 0;
                if (hide != null)
                {
                    foreach (string str2 in hide)
                    {
                        j = 0;
                        foreach (DataGridViewColumn str3 in Grid.Columns)
                        {
                            if (str3.HeaderText == str2.ToString())
                            {
                                Grid.Columns[j].Visible = false;
                            }
                            j++;
                        }
                        i++;
                    }
                }

            }
            SetRowNumber(Grid);
           // Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
           // Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);//AutoSizeToAllHeaders
        }
        //public static void AddColumnDynamic(DataGridView Grid, string[] s, string[] ss, string query, string[] hidem int [] width)
        //{
           
        //    Grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;//
        //    int grow = Grid.Rows.Count;
        //    int gcol = Grid.Columns.Count;
        //    int i = 1, j = 0, cnt = 0; bool check = false;
        //    DataSet ds4 = Utility.ExecuteSelectQuery(query, Class.Users.TableNameGrid);
        //    DataTable dt4 = ds4.Tables[Class.Users.TableNameGrid];
        //    i = 0;
        //    if (gcol == 1 || gcol <= 0)
        //    {
        //        foreach (DataColumn str1 in dt4.Columns)
        //        {
        //            Grid.Columns.Add(i.ToString(), str1.ToString().ToUpper());                    
        //            i++;
        //        }
        //    }
        //    if (ss.Length > 0 || Grid.Columns.Count > 1)
        //    {
        //        i = 0; j = 0; cnt = 1;
        //        if (dt4.Columns.Count < Grid.Columns.Count)
        //        {
        //            check = true;
        //        }
        //        foreach (DataRow row in dt4.Rows)
        //        {

        //            Grid.Rows.Add();
        //            if (check == false)
        //            {

        //                for (j = 0; j < Grid.Columns.Count; j++)
        //                {
        //                    Grid.Rows[i].Cells[j].Value = row[j].ToString();
        //                    Grid.Rows[i].DefaultCellStyle.Font = new Font(Class.Users.FontName.FontFamily, Class.Users.FontName.Size, FontStyle.Bold);

        //                }
        //            }
        //            else
        //            {
        //                for (j = 0; j < Grid.Columns.Count - 1; j++)
        //                {
        //                    Grid.Rows[i].Cells[j + 1].Value = row[j].ToString();
        //                }
        //            }
        //            i++;
        //        }
        //        i = 0; j = 0;
        //        if (hide != null)
        //        {
        //            foreach (string str2 in hide)
        //            {
        //                j = 0;
        //                foreach (DataGridViewColumn str3 in Grid.Columns)
        //                {
        //                    if (str3.HeaderText == str2.ToString())
        //                    {
        //                        Grid.Columns[j].Visible = false;
        //                    }
        //                    j++;
        //                }
        //                i++;
        //            }
        //        }
              
        //    }
        //    Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        //    Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);//AutoSizeToAllHeaders
        //}
        public static void AddGridColumn(DataGridView Grid, string query, string[] hide,Int32[] width,string tablename)
        {
            Grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;//
            int grow = Grid.Rows.Count;
            int gcol = 1;
            int wid = width.Length;
            int i = 1, j = 0, cnt = 0; bool check = false;
            DataSet ds4 = Utility.ExecuteSelectQuery(query, tablename);
            DataTable dt4 = ds4.Tables[tablename];
            if (dt4 != null)
            {
                i = 0;
                if (gcol == 1 || gcol <= 0)
                {
                    foreach (DataColumn str1 in dt4.Columns)
                    {
                        Grid.Columns.Add(i.ToString(), str1.ToString().Trim());

                        if (wid > i)
                        {
                            Grid.Columns[i].Width = width[i];
                        }

                        i++;
                    }
                }
                i = 0; j = 0;
                if (hide != null)
                {
                    foreach (string str2 in hide)
                    {
                        j = 0;
                        foreach (DataGridViewColumn str3 in Grid.Columns)
                        {
                            if (str3.HeaderText == str2.ToString().Trim())
                            {
                                Grid.Columns[j].Visible = false;
                            }

                            j++;
                        }
                        i++;
                    }
                }
                SetRowNumber(Grid);
                Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);//AutoSizeToAllHeaders
            }
        }
        public static void AddColumnDynamic1(DataGridView Grid, string query)
        {
            Grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;//
            int grow = Grid.Rows.Count;
            int gcol = Grid.Columns.Count;
           
            int i = 1, j = 0, cnt = 0; bool check = false;
            DataSet ds4 = Utility.ExecuteSelectQuery(query, Class.Users.TableNameGrid);
            DataTable dt4 = ds4.Tables[Class.Users.TableNameGrid];
            i = 0;         
            if (Grid.Columns.Count > 1)
            {
                i = 0; j = 0; cnt = 1;
                if (dt4.Columns.Count < Grid.Columns.Count)
                {
                    check = true;
                }
                if (Grid.Rows.Count <= 0)
                {
                    foreach (DataRow row in dt4.Rows)
                    {
                        Grid.Rows.Add();
                      
                        if (check == false)
                        {
                            for (j = 0; j < Grid.Columns.Count; j++)
                            {
                               
                                if (j == 0)
                                {
                                    Grid.Rows[i].Cells[j].ReadOnly = false;
                                }
                                else
                                {
                                    Grid.Rows[i].Cells[j].ReadOnly = true;
                                    Grid.Rows[i].Cells[j].Value = row[j].ToString();
                                    Grid.Rows[i].DefaultCellStyle.Font = new Font(Class.Users.FontName.FontFamily, Class.Users.FontName.Size, FontStyle.Bold);
                                }
                            }
                        }
                        else
                        {
                            for (j = 0; j < Grid.Columns.Count - 1; j++)
                            {
                                if (j == 0)
                                {
                                    Grid.Rows[i].Cells[j].ReadOnly = false;
                                }
                                else
                                {
                                    Grid.Rows[i].Cells[j].ReadOnly = true;
                                    Grid.Rows[i].Cells[j + 1].Value = row[j].ToString();
                                }
                            }
                        }
                        i++;
                    }
                }
                i = 0; j = 0;
            

              
            }
           // Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Grid.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);//AutoSizeToAllHeaders
            SetRowNumber(Grid);
        }
        public static DataTable ReadExcelIntoDataTable(string FileName, string SheetName)
        {
            DataTable RetVal = new DataTable();
            string strConnString;
            strConnString = ("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + (FileName + ";"));
            string strSQL;
            strSQL = ("SELECT * FROM [" + (SheetName + "$]"));
            OdbcDataAdapter y = new OdbcDataAdapter(strSQL, strConnString);
            y.Fill(RetVal);
            return RetVal;
        }
        public static DataTable GetDistinctDataTable(DataTable Dt, string[] Columns)
        {
            DataView Dv = new DataView(Dt);
            DataTable RetVal = new DataTable();
            RetVal = Dv.ToTable(true, Columns);
            return RetVal;
        }
        public static void ShowPopUpForm(object sender, object ParentFrm)
        {           
            bool Exsist = false;
            for (var i = 0; i < GlobalVariables.TabCtrl.TabPages.Count; i++)
            {
                if (GlobalVariables.TabCtrl.TabPages[i].Name == ((Form)sender).Name.Trim())
                {
                    Exsist = true; 
                    GlobalVariables.CurrentForm = sender;
                    GlobalVariables.TabCtrl.SelectedIndex = i;
                    GlobalVariables.TabCtrl.Focus();
                }
            }
            if (Exsist == false)
            {
                TabPage tbp = new TabPage();
                ((Form)sender).TopLevel = false; 
                ((Form)sender).Visible = true;
               // ((Form)sender).Font = Class.Users.FontName;
                ((Form)sender).Font=new Font(Class.Users.FontName.FontFamily.Name, Class.Users.FontName.Size, FontStyle.Bold);
                ((Form)sender).FormBorderStyle = FormBorderStyle.None;
                ((Form)sender).Dock = DockStyle.Fill;
                GlobalVariables.CurrentForm = sender;
                tbp.Controls.Add(((Form)sender));
                GlobalVariables.TabCtrl.TabPages.Add(tbp);
                tbp.Name = ((Form)sender).Name.Trim();
                tbp.Font = Class.Users.FontName;
                tbp.Font = new Font(Class.Users.FontName.FontFamily.Name, Class.Users.FontName.Size, FontStyle.Bold);
                tbp.Text = ((Form)sender).Name.Trim();
                GlobalVariables.TabCtrl.SelectedTab = tbp;
                ((ToolStripAccess)sender).News();
            }
           
        }
        public static void ShowForm(object sender, object ParentFrm)
        {
            //Form Frm = ((Form)sender);
            //Frm.MdiParent = ((Form)ParentFrm);
            //Frm.FormBorderStyle = FormBorderStyle.None;
            //Frm.Dock = DockStyle.Fill;
            //Frm.StartPosition = FormStartPosition.CenterScreen;
            //Frm.Show();
            bool Exsist = false;
            for (var i = 0; i < GlobalVariables.TabCtrl .TabPages.Count; i++)
            {
                if (GlobalVariables.TabCtrl .TabPages[i].Name == ((Form)sender).Name.Trim())
                {
                    Exsist = true;
                    GlobalVariables.CurrentForm = sender;
                    GlobalVariables.TabCtrl .SelectedIndex = i;
                    GlobalVariables.TabCtrl .Focus();
                }
            }
            if (Exsist == false)
            {
                TabPage tbp = new TabPage();
                ((Form)sender).TopLevel = false;
                ((Form)sender).Visible = true;
               //  ((Form)sender).Font = Class.Users.FontName;
               ((Form)sender).Font = new Font(Class.Users.FontName.FontFamily.Name, Class.Users.FontName.Size, FontStyle.Bold);
                ((Form)sender).FormBorderStyle = FormBorderStyle.None;
                ((Form)sender).Dock = DockStyle.Fill;
                GlobalVariables.CurrentForm = sender;
                tbp.Controls.Add(((Form)sender));
                GlobalVariables.TabCtrl .TabPages.Add(tbp);
               // tbp.Font = Class.Users.FontName;
                tbp.Font = new Font(Class.Users.FontName.FontFamily.Name, Class.Users.FontName.Size, FontStyle.Bold);
                tbp.Name = ((Form)sender).Name.Trim();
                tbp.Text = ((Form)sender).Name.Trim();
                GlobalVariables.TabCtrl .SelectedTab = tbp;
                ((ToolStripAccess)sender).News();
            }
        }
        public static void Controls_Enter(object sender, EventArgs e)
        {
            try
            {
                if (sender is ComboBox)
                {
                    ((ComboBox)sender).BackColor = Color.Moccasin;
                }
                else if (sender is TextBox)
                {
                    ((TextBox)sender).BackColor = Color.Moccasin;
                    ((TextBox)sender).SelectionStart = ((TextBox)sender).Text.Length;
                }
                else if (sender is CheckedListBox)
                {
                    ((CheckedListBox)sender).BackColor = Color.Moccasin;
                }
            }
            catch { }
        }
        public static bool Chk_GridDuplicate(DataGridView Grid, string CurrCtrlValue, params string[] ColumnRange)
        {
            string ChkColumn = "";
            string ChkColumn1 = "";
            foreach (DataGridViewRow Row in Grid.Rows)
            {
                if (Row.Index != Grid.CurrentCell.RowIndex)
                {
                    ChkColumn = "";
                    ChkColumn1 = "";
                    foreach (string Str in ColumnRange)
                    {
                        ChkColumn += GenFun.IsNull(Row.Cells[Str].Value).Trim();
                        if (Grid.Columns[Grid.CurrentCell.ColumnIndex].Name == Str)
                        {
                            ChkColumn1 += CurrCtrlValue.Trim();
                        }
                        else
                        {
                            ChkColumn1 += GenFun.IsNull(Grid.Rows[Grid.CurrentCell.RowIndex].Cells[Str].Value).Trim();
                        }
                    }
                    if (ChkColumn == ChkColumn1)
                    {
                        MessageBox.Show("Duplication Not Allowed..!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }
            return true;
        }
        public static bool Chk_GridDuplicate1(DataGridView Grid, DataGridViewRow Row, string[] ColumnRange, string[] ColumnRange1)
        {
            foreach (DataGridViewRow Row1 in Grid.Rows)
            {
                string ChkColumn = "";
                string ChkColumn1 = "";
                foreach (string Col in ColumnRange)
                {
                    ChkColumn += GenFun.IsNull(Row1.Cells[Col].Value).Trim();
                }
                foreach (string Col in ColumnRange1)
                {
                    ChkColumn1 += GenFun.IsNull(Row.Cells[Col].Value).Trim();
                }
                if (ChkColumn == ChkColumn1)
                {
                    return false;
                }
            }
            return true;
        }
        public static void Controls_Leave(object sender, EventArgs e)
        {
            if (sender is ComboBox)
            {
                ((ComboBox)sender).BackColor = Color.White;
            }
            else if (sender is TextBox)
            {
                ((TextBox)sender).BackColor = Color.White;
            }
            else if (sender is CheckedListBox)
            {
                ((CheckedListBox)sender).ClearSelected();
                ((CheckedListBox)sender).BackColor = Color.White;
            }
            else if (sender is DataGridView)
            {
                ((DataGridView)sender).ClearSelection();
            }
        }
        public static void Controls_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (sender is TextBox || sender is ComboBox || sender is DateTimePicker || sender is CheckedListBox || sender is CheckBox || sender is NumericUpDown)
                {
                    SendKeys.Send("{TAB}");
                }
            }
        }
        public static void Move_GridCtrl(object CurCtrl, DataGridView Grid, string NextColName, bool NextRow = false, string ValueMemberCol = "", bool ApplyText = true)
        {
            if (Grid.CurrentCell != null)
            {
                if (CurCtrl.GetType() == typeof(UCComboBox))
                {
                    if (ApplyText == true)
                    {
                        Grid.CurrentCell.Value = GenFun.IsNull(((UCComboBox)CurCtrl).Text).Trim();
                    }
                    if (ValueMemberCol != "")
                    {
                        if ((((UCComboBox)CurCtrl).SelectedIndex != -1))
                        {
                            Grid.Rows[Grid.CurrentCell.RowIndex].Cells[ValueMemberCol].Value = ((UCComboBox)CurCtrl).SelectedValue;
                        }
                        else
                        {
                            Grid.Rows[Grid.CurrentCell.RowIndex].Cells[ValueMemberCol].Value = 0;
                        }
                    }
                }
                else if (CurCtrl.GetType() == typeof(UCTextBox) && ApplyText == true)
                {
                    Grid.CurrentCell.Value = GenFun.IsNull(((UCTextBox)CurCtrl).Text).Trim();
                }
                else if (CurCtrl.GetType() == typeof(UCDateTimePicker) && ApplyText == true)
                {
                    Grid.CurrentCell.Value = GenFun.IsNull(((UCDateTimePicker)CurCtrl).Text).Trim();
                }
                if (NextRow == true)
                {
                    if (Grid.CurrentCell.RowIndex == Grid.Rows.Count - 1)
                    {
                        if (Grid.DataSource != null)
                        {
                            ((DataTable)Grid.DataSource).Rows.Add();
                        }
                        else
                        {
                            Grid.Rows.Add();
                        }
                    }
                    Grid.CurrentCell = Grid[NextColName, Grid.CurrentCell.RowIndex + 1];
                }
                else
                {
                    Grid.CurrentCell = Grid[NextColName, Grid.CurrentCell.RowIndex];
                }
            }
        }
        public static void Enter_Cells(DataGridView Grid, Control Ctrl)
        {
            if (Grid.CurrentCell.Value != null)
            {
                Ctrl.Text = Grid.CurrentCell.Value.ToString().Trim();
            }
            else
            {
                Ctrl.Text = "";
            }
            Ctrl.BackColor = Color.Moccasin;
            Rectangle Rct = Grid.GetCellDisplayRectangle(Grid.CurrentCell.ColumnIndex, Grid.CurrentCell.RowIndex, true);
            Ctrl.Left = Grid.Left + Rct.Left;
            Ctrl.Top = Grid.Top + Rct.Top;
            Ctrl.Width = Rct.Width;
            Ctrl.Height = Rct.Height;
            Ctrl.Visible = true;
            Ctrl.BringToFront();
            Ctrl.Focus();
        }
        public static void CellCtrl_Resize(DataGridView Grid, Control Ctrl)
        {
            Rectangle Rct = Grid.GetCellDisplayRectangle(Grid.CurrentCell.ColumnIndex, Grid.CurrentCell.RowIndex, true);
            Ctrl.Left = Grid.Left + Rct.Left;
            Ctrl.Top = Grid.Top + Rct.Top;
            Ctrl.Width = Rct.Width;
            Ctrl.Height = Rct.Height;
        }
        public static void Clear_Controls(this Control RootCtrl)
        {
            foreach (Control Ctrl in RootCtrl.Controls)
            {
                if (Ctrl.GetType() == typeof(CheckBox))
                {
                    ((CheckBox)Ctrl).Checked = false;
                }
                else if (Ctrl.GetType() == typeof(TextBox))
                {
                    ((TextBox)Ctrl).Text = "";
                    ((TextBox)Ctrl).Tag = "";
                }
                else if (Ctrl.GetType() == typeof(ComboBox))
                {
                    if (((ComboBox)Ctrl).DropDownStyle == ComboBoxStyle.DropDownList && ((ComboBox)Ctrl).Items.Count > 0)
                    {
                        ((ComboBox)Ctrl).SelectedIndex = 0;
                    }
                    else
                    {
                        ((ComboBox)Ctrl).SelectedIndex = -1;
                    }
                }
                else if (Ctrl.GetType() == typeof(DateTimePicker))
                {
                    ((DateTimePicker)Ctrl).Value = DateTime.Now;
                }
                else if (Ctrl.GetType() == typeof(TextBox))
                {
                    ((PictureBox)Ctrl).Image = null;
                    ((PictureBox)Ctrl).BackgroundImage = null;
                }
                else if (Ctrl.GetType() == typeof(DataGridView))
                {
                    if (((DataGridView)Ctrl).ColumnCount == 0)
                    {
                        ((DataGridView)Ctrl).DataSource = null;
                        ((DataGridView)Ctrl).Rows.Clear();
                    }
                    ((DataGridView)Ctrl).ClearSelection();
                }
                else if (Ctrl.GetType() == typeof(CheckedListBox))
                {
                    ((CheckedListBox)Ctrl).DataSource = null;
                    ((CheckedListBox)Ctrl).Items.Clear();
                    ((CheckedListBox)Ctrl).ClearSelected();
                }
                if (Ctrl.Controls != null)
                {
                    Clear_Controls(Ctrl);
                }

            }
        }
        public static bool Data_Validation(Control RootCtrl, ErrorProvider ErrProvider)
        {
            foreach (Control Ctrl in RootCtrl.Controls)
            {
                if (ErrProvider.GetError(Ctrl) != "")
                {
                    return false;
                }
            }
            return true;
        }
        public static Hashtable GetEachCtrlValuesIntoHashTable(this Control RootCtrl)
        {
            Hashtable Rtn = new Hashtable();
            Rtn = GetCtrlValuesIntoHashTable(RootCtrl, Rtn);
            return Rtn;
        }
        public static Hashtable GetCtrlValuesIntoHashTable(this Control RootCtrl, Hashtable Rtn)
        {
            foreach (Control Ctrl in RootCtrl.Controls)
            {
                if (Ctrl.GetType() == typeof(TextBox))
                {
                    Rtn.Add("@" + ((TextBox)Ctrl).Name.ToString(), ((TextBox)Ctrl).Text.ToString().Trim());
                }
                else if (Ctrl.GetType() == typeof(ComboBox))
                {
                    if (Convert.ToInt32(((ComboBox)Ctrl).SelectedValue) > 0)
                    {
                        Rtn.Add("@" + ((ComboBox)Ctrl).Name.ToString() + "ID", ((ComboBox)Ctrl).SelectedValue);
                        Rtn.Add("@" + ((ComboBox)Ctrl).Name.ToString(), ((ComboBox)Ctrl).Text.ToString().Trim());
                    }
                    else
                    {
                        Rtn.Add("@" + ((ComboBox)Ctrl).Name.ToString(), ((ComboBox)Ctrl).Text);
                        Rtn.Add("@" + ((ComboBox)Ctrl).Name.ToString() + "FL", ((ComboBox)Ctrl).Text.Substring(0, 1));
                    }
                    if (Convert.ToInt32(((ComboBox)Ctrl).SelectedIndex) != -1)
                    {
                        Rtn.Add("@" + ((ComboBox)Ctrl).Name.ToString() + "Idx", ((ComboBox)Ctrl).SelectedIndex + 1);
                    }
                }
                else if (Ctrl.GetType() == typeof(DateTimePicker))
                {
                    Rtn.Add("@" + ((DateTimePicker)Ctrl).Name.ToString(), ((DateTimePicker)Ctrl).Value);
                }
                else if (Ctrl.GetType() == typeof(CheckBox))
                {
                    if (((CheckBox)Ctrl).CheckState == CheckState.Checked)
                    {
                        Rtn.Add("@" + ((CheckBox)Ctrl).Name.ToString(), "Yes");
                        Rtn.Add("@" + ((CheckBox)Ctrl).Name.ToString() + "FL", "Y");
                    }
                    else
                    {
                        Rtn.Add("@" + ((CheckBox)Ctrl).Name.ToString(), "No");
                        Rtn.Add("@" + ((CheckBox)Ctrl).Name.ToString() + "FL", "N");
                    }
                }
                else if (Ctrl.GetType() == typeof(NumericUpDown))
                {
                    Rtn.Add("@" + ((NumericUpDown)Ctrl).Name.ToString(), ((NumericUpDown)Ctrl).Value);
                }
                else if (Ctrl.GetType() == typeof(PictureBox))
                {
                    Rtn.Add("@" + ((PictureBox)Ctrl).Name.ToString(), ConvertImageToByte(((PictureBox)Ctrl).Image));
                }
                if (Ctrl.Controls != null)
                {
                    GetCtrlValuesIntoHashTable(Ctrl, Rtn);
                }
            }
            return Rtn;
        }
        public static byte[] ConvertImageToByte(Image Img)
        {
            if (Img != null)
            {
                MemoryStream ms = new MemoryStream();
                Img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                return ms.ToArray();
            }
            else
            {
                return null;
            }
        }
        public static DataTable GetEachCtrlValuesIntoDataTable(this Control RootCtrl)
        {
            DataTable Rtn = new DataTable();
            DataRow Row = null;
            Row = Rtn.NewRow();
            foreach (Control Ctrl in RootCtrl.Controls)
            {
                if (Ctrl.GetType() == typeof(TextBox))
                {
                    Rtn.Columns.Add(((TextBox)Ctrl).Name.ToString());
                    Row[((TextBox)Ctrl).Name.ToString()] = ((TextBox)Ctrl).Text;
                }
                else if (Ctrl.GetType() == typeof(ComboBox))
                {
                    if (Convert.ToInt32(((ComboBox)Ctrl).SelectedValue) > 0)
                    {
                        Rtn.Columns.Add(((ComboBox)Ctrl).Name.ToString());
                        Row[((ComboBox)Ctrl).Name.ToString()] = ((ComboBox)Ctrl).SelectedValue;
                    }
                    else
                    {
                        Rtn.Columns.Add(((ComboBox)Ctrl).Name.ToString());
                        Row[((ComboBox)Ctrl).Name.ToString()] = ((ComboBox)Ctrl).Text;
                    }
                }
                else if (Ctrl.GetType() == typeof(DateTimePicker))
                {
                    Rtn.Columns.Add(((DateTimePicker)Ctrl).Name.ToString());
                    Row[((DateTimePicker)Ctrl).Name.ToString()] = ((DateTimePicker)Ctrl).Value;
                }
                if (Ctrl.Controls != null)
                {
                    GetEachCtrlValuesIntoDataTable(Ctrl);
                }
            }
            Rtn.Rows.Add(Row);
            return Rtn;
        }
        public static DataTable GetValuesIntoDataTable(object sender, params String[] columnNames)
        {
            DataTable Rtn = new DataTable();
            DataRow Row;
            foreach (var Col in columnNames)
            {
                Rtn.Columns.Add(Col);
            }
            if (sender.GetType() == typeof(CheckedListBox))
            {
                foreach (var Item in ((CheckedListBox)sender).CheckedItems)
                {
                    Row = Rtn.NewRow();
                    foreach (var Col in columnNames)
                    {
                        Row[Col] = ((DataRowView)Item)[Col];
                    }
                    Rtn.Rows.Add(Row);
                }
            }
            return Rtn;
        }

        public static bool Validate_StringEmpty(object sender, ErrorProvider ErrProvider)
        {
            if (sender.GetType() == typeof(ComboBox))
            {
                if (((ComboBox)sender).SelectedIndex == -1)
                {
                    ErrProvider.SetError(((ComboBox)sender), "Please Select Vaild Data..!");
                    return false;
                }
                else
                {
                    ErrProvider.SetError(((ComboBox)sender), null);
                }

            }
            else if (sender.GetType() == typeof(TextBox))
            {
                if (((TextBox)sender).Text == string.Empty)
                {
                    ErrProvider.SetError(((TextBox)sender), "Please Enter Vaild Data..!");
                    return false;
                }
                else
                {
                    ErrProvider.SetError(((TextBox)sender), null);
                }
            }
            return true;
        }

        public static bool Validate_NumericEmptyReplace(object sender)
        {

            if (sender.GetType() == typeof(TextBox))
            {
                if (((TextBox)sender).Text == string.Empty)
                {
                    ((TextBox)sender).Text = (0).ToString();
                    return false;
                }
            }
            return true;
        }
        public static void SetColumnsOrder(this DataTable table, params String[] columnNames)
        {
            int columnIndex = 0;
            foreach (var columnName in columnNames)
            {
                table.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }
        }

        public static Image ConvertToImage(byte[] Field)
        {
            Image Img = null;
            try
            {

                if (Field.GetUpperBound(0) != -1)
                {
                    MemoryStream ms = new MemoryStream(Field);
                    Image returnImage = Image.FromStream(ms);
                    return returnImage;
                }
            }
            catch
            {
                Img = null;
            }
            return Img;
        }
        public static bool Chk_DeleltePassword()
        {
            //if (GlobalVariables.Reqd_DelPassword == true)
            //{
            //    var PassFrm = new Frm_Password();
            //    PassFrm.Password = GlobalVariables.DelPassword;
            //    PassFrm.ShowDialog();
            //    if (GlobalVariables.Rtn_Password == GlobalVariables.DelPassword)
            //    {
            //        return true;
            //    }
            //    else
            //    {
            //        return false;
            //    }
            //}
            //else
            //{
                return true;
            //}
        }
        public static bool Chk_EditPassword()
        {
            //if (GlobalVariables.Reqd_EditPassword == true)
            //{
            //    var PassFrm = new Frm_Password();
            //    PassFrm.Password = GlobalVariables.SavePassword;
            //    PassFrm.ShowDialog();
            //    if (GlobalVariables.Rtn_Password == GlobalVariables.DelPassword)
            //    {
            //        return true;
            //    }
            //    else
            //    {
            //        return false;
            //    }
            //}
            //else
            //{
                return true;
            //}
        }
        public static void FilterGrid(DataGridView Sender, string SearchColumn, string SearchStr)
        {
            ((DataTable)Sender.DataSource).DefaultView.RowFilter = String.Format("[{0}] LIKE '%{1}%'", SearchColumn, SearchStr);
        }
        public static void BindValuesIntoEachCtrl(Form RootCtrl, DataTable Dt)
        {
            if (Dt.Rows.Count > 0)
            {
                foreach (DataColumn Column in Dt.Columns)
                {
                    Control[] Ctrl = RootCtrl.Controls.Find(Column.ColumnName.ToString(), true);
                    for (int i = 0; i <= Ctrl.GetUpperBound(0); i++)
                    {
                        if (Dt.Rows[0][Column.ColumnName] != DBNull.Value)
                        {
                            if (Ctrl[i].GetType() == typeof(TextBox))
                            {
                                ((TextBox)Ctrl[i]).Text = Dt.Rows[0][Column.ColumnName].ToString();

                            }
                            else if (Ctrl[i].GetType() == typeof(ComboBox))
                            {
                                if (Dt.Columns[Column.ColumnName].DataType == typeof(Int32))
                                {
                                    ((ComboBox)Ctrl[i]).SelectedItem = Convert.ToInt32(Dt.Rows[0][Column.ColumnName]) - 1;
                                }
                                else if (((ComboBox)Ctrl[i]).FindString(Dt.Rows[0][Column.ColumnName].ToString().Trim()) != -1)
                                {
                                    ((ComboBox)Ctrl[i]).SelectedIndex = ((ComboBox)Ctrl[i]).FindString(Dt.Rows[0][Column.ColumnName].ToString().Trim());
                                }
                                else if (((ComboBox)Ctrl[i]).FindString(Dt.Rows[0][Column.ColumnName].ToString().Substring(1, 1)) != -1)
                                {
                                    ((ComboBox)Ctrl[i]).SelectedIndex = ((ComboBox)Ctrl[i]).FindString(Dt.Rows[0][Column.ColumnName].ToString().Substring(1, 1));
                                }
                            }
                            else if (Ctrl[i].GetType() == typeof(DateTimePicker))
                            {

                                ((DateTimePicker)Ctrl[i]).Value = Convert.ToDateTime(Dt.Rows[0][Column.ColumnName]);

                            }
                            else if (Ctrl[i].GetType() == typeof(CheckBox))
                            {
                                if (Dt.Rows[0][Column.ColumnName].ToString() == "Yes")
                                {
                                    ((CheckBox)Ctrl[i]).Checked = true;
                                }
                                else
                                {
                                    ((CheckBox)Ctrl[i]).Checked = false;
                                }
                            }
                            else if (Ctrl[i].GetType() == typeof(NumericUpDown))
                            {
                                ((NumericUpDown)Ctrl[i]).Value = Convert.ToDecimal(Dt.Rows[0][Column.ColumnName]);
                            }
                            else if (Ctrl[i].GetType() == typeof(PictureBox) && Dt.Columns[Column.ColumnName].DataType == typeof(byte[]))
                            {
                                ((PictureBox)Ctrl[i]).Image = ConvertToImage((byte[])Dt.Rows[0][Column.ColumnName]);
                                ((PictureBox)Ctrl[i]).SizeMode = PictureBoxSizeMode.StretchImage;
                            }
                        }
                    }
                }
            }

        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.UserControls
{
    public partial class UCCListView : UserControl
    {
        public UCCListView()
        {
            InitializeComponent();
        }
        public void Load_Details()
        {
            try
            {
                TxtSearch.Text = ""; panel1.BackColor = Class.Users.BackColors;
                panel2.BackColor = Class.Users.BackColors;Class.Users.UserTime = 0;
                DGLoadDetails.Font = Class.Users.FontName;               
                DGLoadDetails.Columns.Clear();
                if (Class.Users.SearchQuery != "")
                {
                    Utility.Load_DataGrid(DGLoadDetails, Class.Users.SearchQuery);
                    HideColumns();
                    lbltotal.Refresh(); lbltotal.Text = "  Total  :" + DGLoadDetails.Rows.Count;
                    SetRowNumber(DGLoadDetails);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error..!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
   
        private void HideColumns()
        {
            try
            {
                if (Class.Users.HideCols != null)
                {
                    foreach (String Str in Class.Users.HideCols)
                    {
                        DGLoadDetails.Columns[Str].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error..!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
       public Label TotalCount
        {
            get { return lbltotal; }
            set { lbltotal = value; }
        }
        private void DGLoadDetails_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= -1)
            {
                ((ToolStripAccess)GlobalVariables.CurrentForm).Searchs(Convert.ToInt32(DGLoadDetails.Rows[e.RowIndex].Cells["ID"].Value));
                //((ToolStripAccess)GlobalVariables.CurrentForm).Form_Edit(Convert.ToInt32(DGLoadDetails.Rows[e.RowIndex].Cells["ID"].Value));

            }
        }
        private void DGLoadDetails_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex >= -1)
            //{
            //    try
            //    {
            //        ((ToolStripAccess)GlobalVariables.CurrentForm).Searchs(Convert.ToInt32(DGLoadDetails.Rows[e.RowIndex].Cells["ID"].Value));

            //    }
            //    catch (Exception ex) { }

            //}
        }
        private void DGLoadDetails_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (DGLoadDetails.CurrentCell != null)
            {
                if (e.KeyCode == Keys.Enter && DGLoadDetails.CurrentCell.RowIndex != -1 && DGLoadDetails.Columns[DGLoadDetails.CurrentCell.ColumnIndex].Name != "ChkSelect")
                {
                      ((ToolStripAccess)GlobalVariables.CurrentForm).Searchs(Convert.ToInt32(DGLoadDetails.Rows[DGLoadDetails.CurrentCell.RowIndex].Cells["ID"].Value));

                }
            }
        }
  
        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (DGLoadDetails.CurrentCell != null)
                {

                    int ColIdx = DGLoadDetails.CurrentCell.ColumnIndex;
                    ((DataTable)DGLoadDetails.DataSource).DefaultView.RowFilter = String.Format("Convert([{0}], System.String) LIKE '{1}%'", DGLoadDetails.Columns[DGLoadDetails.CurrentCell.ColumnIndex].Name, TxtSearch.Text.ToString());
                    if (DGLoadDetails.RowCount > 0)
                    {
                        DGLoadDetails.CurrentCell = DGLoadDetails[ColIdx, 0];
                        lbltotal.Refresh(); lbltotal.Text = "  Total  :" + DGLoadDetails.Rows.Count;
                        SetRowNumber(DGLoadDetails);
                    }
                   
                }
                else
                {
                    ((DataTable)DGLoadDetails.DataSource).DefaultView.RowFilter = null;
                    SetRowNumber(DGLoadDetails);
                }
            }
            catch { }
        }

        private void DGLoadDetails_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            DGLoadDetails.ClearSelection();
        }

        private void DGLoadDetails_KeyPress(object sender, KeyPressEventArgs e)
        {
            char key = (char)e.KeyChar;
            if (e.KeyChar == 8 && TxtSearch.Text.ToString().Length >= 1)
            {
                TxtSearch.Text = TxtSearch.Text.ToString().Substring(0, TxtSearch.Text.ToString().Length - 1);
            }
            else if (char.IsLetterOrDigit(key) || char.IsWhiteSpace(key))
            {
                TxtSearch.Text += e.KeyChar;
            }
        }
        public static void SetRowNumber(DataGridView Grid)
        {
            int rowNumber = 1;
            Grid.Font = Class.Users.FontName;
            Grid.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            Grid.ColumnHeadersDefaultCellStyle.Font = Class.Users.FontName;
            Grid.ColumnHeadersDefaultCellStyle.Font = new Font(Class.Users.FontName.FontFamily, Class.Users.FontName.Size, FontStyle.Bold);
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
        private void DGLoadDetails_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //this.DGLoadDetails.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
            //DGLoadDetails.Rows[0].DefaultCellStyle.BackColor = Class.Users.Color1;
            //DGLoadDetails.Rows[0].DefaultCellStyle.ForeColor = Class.Users.Color2;

        }

      

        //private void CMnuEdit_Click(object sender, EventArgs e)
        //{
        //    if (DGLoadDetails.CurrentCell != null)
        //    {
        //        ((ToolStripAccess)GlobalVariables.CurrentForm).Form_Edit(Convert.ToInt32(DGLoadDetails.Rows[DGLoadDetails.CurrentCell.RowIndex].Cells["ID"].Value));
        //    }
        //}

        //private void CMnuDelete_Click(object sender, EventArgs e)
        //{
        //    if (DGLoadDetails.CurrentCell != null)
        //    {
        //        ((ToolStripAccess)GlobalVariables.CurrentForm).Form_Delete(Convert.ToInt32(DGLoadDetails.Rows[DGLoadDetails.CurrentCell.RowIndex].Cells["ID"].Value));
        //    }
        //}

        //private void CMnuPreview_Click(object sender, EventArgs e)
        //{
        //    if (DGLoadDetails.CurrentCell != null)
        //    {
        //        ((ToolStripAccess)GlobalVariables.CurrentForm).Form_Preview(Convert.ToInt32(DGLoadDetails.Rows[DGLoadDetails.CurrentCell.RowIndex].Cells["ID"].Value));
        //    }
        //}

    }
}

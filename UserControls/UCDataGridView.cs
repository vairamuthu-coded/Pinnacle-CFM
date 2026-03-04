using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle
{
    public partial class UCDataGridView : DataGridView
    {
        public bool ShowRowNumber { get; set; }
        public bool MergeCells { get; set; }
        public bool RowIndicater { get; set; }
        public bool UserDeleteRow { get; set; } = true;
        public UCDataGridView()
        {
            InitializeComponent();
            ShowRowNumber = true;
            RowIndicater = true;
            MergeCells = false;
        }

        private void UCDataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (RowIndicater == true)
            {
                e.PaintCells(e.ClipBounds, DataGridViewPaintParts.All);
                e.PaintHeader(DataGridViewPaintParts.Background

                    | DataGridViewPaintParts.Border

                    | DataGridViewPaintParts.Focus

                    | DataGridViewPaintParts.SelectionBackground

                    | DataGridViewPaintParts.ContentForeground);

                e.Handled = true;
            }
        }

        private void UCDataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (ShowRowNumber == true)
            {
                this.Rows[e.RowIndex].HeaderCell.Value = (e.RowIndex + 1).ToString();
            }
            if (MergeCells == true)
            {
                if (e.RowIndex == 0)
                    return;
                if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }
            }
        }

        private void UCDataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (MergeCells == true)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                if (e.RowIndex < 1 || e.ColumnIndex < 0)
                    return;
                if (IsTheSameCellValue(e.ColumnIndex, e.RowIndex))
                {
                    e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                }
                else
                {
                    e.AdvancedBorderStyle.Top = this.AdvancedCellBorderStyle.Top;
                }
            }
        }
        bool IsTheSameCellValue(int column, int row)
        {
            DataGridViewCell cell1 = this[column, row];
            DataGridViewCell cell2 = this[column, row - 1];
            if (cell1.Value == null || cell2.Value == null)
            {
                return false;
            }
            return cell1.Value.ToString() == cell2.Value.ToString();
        }

        private void UCDataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Delete)
            {
                if (CurrentCell != null && UserDeleteRow == true)
                {
                    this.Rows.Remove(this.CurrentRow);
                }
            }
        }
    }
}


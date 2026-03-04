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
    public partial class UCPopUpBox : Form
    {
        public DataTable DataSourse { get; set; }
        public string DisplayMember { get; set; }
        public string ValueMember { get; set; }
        public DataGridViewRow ReturnRow { get; set; }
        public string[] HideCols { get; set; }
      

        public UCPopUpBox()
        {
            InitializeComponent();
        }

        private void UCPopUpBox_Load(object sender, EventArgs e)
        {
            //DataSourse.DefaultView.RowFilter = null;
            //DGPopUp.DataSource = null;
            //DGPopUp.DataSource = DataSourse;
            //TxtSearch.Text = "";
            //TxtSearch.Enabled = false;
            //if (HideCols != null)
            //{
            //    foreach (string str in HideCols)
            //    {
            //        DGPopUp.Columns[str].Visible = false;
            //    }
            //}
            //if (DGPopUp.Rows.Count > 0)
            //{
            //    DGPopUp.CurrentCell = DGPopUp[DisplayMember, 0];
            //}
        }

        private void DGPopUp_KeyPress(object sender, KeyPressEventArgs e)
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
            else if (e.KeyChar == 27)
            {
                this.Close();
            }
        }

        private void DGPopUp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                ReturnRow = DGPopUp.Rows[e.RowIndex];
                this.Close();
            }
        }

        private void DGPopUp_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (DGPopUp.CurrentCell != null)
            {
                if (e.KeyCode == Keys.Enter && DGPopUp.CurrentCell.RowIndex != -1)
                {
                    ReturnRow = DGPopUp.Rows[DGPopUp.CurrentCell.RowIndex];
                    this.Close();
                }
            }
        }
    }
}

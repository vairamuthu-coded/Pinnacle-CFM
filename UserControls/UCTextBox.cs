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
    public partial class UCTextBox : TextBox
    {
        public DataTable DataSource { get; set; }
        public string DisplayMember { get; set; }
        public string ValueMember { get; set; }
        public ErrorProvider ErrProvider { get; set; }
        public string[] HidePopColumns { get; set; }
        public string ValidateType { get; set; }
        public string DataType { get; set; } = "String";
        [DefaultValue(typeof(Color), "Moccasin")]
        public Color FocusColor { get; set; }
        public bool EnterKeyMoveNextTab { get; set; } = true;
        public UCTextBox()
        {
            InitializeComponent();
            FocusColor = Color.Moccasin;
        }


        private void UCTextBox_Enter(object sender, EventArgs e)
        {
            this.BackColor = FocusColor;
            //this.SelectionStart = this.Text.Length;
            if (ErrProvider != null && (this.Text != null || this.Text != string.Empty))
            {
                ErrProvider.SetError(this, "");
            }
            if (DataSource != null && DisplayMember != null && ValueMember != null)
            {
                ((TextBox)sender).AutoCompleteCustomSource = CustomSource();
            }
        }

        private void UCTextBox_Leave(object sender, EventArgs e)
        {
            this.BackColor = Color.White;
            if (DataSource != null && DisplayMember != null && ValueMember != null)
            {
                DataRow[] Row = DataSource.Select(string.Format("{0} = '{1}'", DisplayMember, this.Text.ToString().Trim()), ValueMember);
                if (Row.GetUpperBound(0) != -1)
                {
                    this.Tag = Row[0][DisplayMember].ToString().Trim();
                }
            }
        }

        private AutoCompleteStringCollection CustomSource()
        {
            var CustomSource = new AutoCompleteStringCollection();
            foreach (DataRow Dr in DataSource.Rows)
            {
                CustomSource.Add(Dr[DisplayMember].ToString());
            }
            return CustomSource;
        }

        private void UCTextBox_Validating(object sender, CancelEventArgs e)
        {
            if (ValidateType == "Empty")
            {
                if (this.Text == string.Empty)
                {
                    if (ErrProvider != null)
                    {
                        ErrProvider.SetError(this, "Please Enter Valid Data..!");
                    }
                }
                else
                {
                    if (ErrProvider != null)
                    {
                        ErrProvider.SetError(this, "");
                    }
                }
            }
        }

        private void UCTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (DataType == "Numeric")
            {
                e.Handled = !char.IsNumber(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != '.' && e.KeyChar != '-' && e.KeyChar != '+';
            }
        }
        private void UCTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && EnterKeyMoveNextTab == true)
            {
                if (this.Multiline == false)
                {
                    SendKeys.Send("{TAB}");
                    e.Handled = true;
                }
            }
            else if (e.KeyCode == Keys.Down)
            {
                if (DataSource != null && DisplayMember != null && ValueMember != null)
                {
                    var Frm = new UCPopUpBox();
                    Frm.DataSourse = DataSource;
                    Frm.DisplayMember = DisplayMember;
                    Frm.ValueMember = ValueMember;
                    Frm.ShowDialog();
                    DataSource.DefaultView.RowFilter = null;
                    e.Handled = true;
                }
            }
        }
    }
}

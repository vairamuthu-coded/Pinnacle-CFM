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
    public partial class UCComboBox : ComboBox
    {
        [DefaultValue(typeof(Color), "Moccasin")]
        public Color FocusColor { get; set; }
        public ErrorProvider ErrProvider { get; set; }
        public string ValidateType { get; set; }
        public bool EnterKeyMoveNextTab { get; set; } = true;
        public string MasterForm { get; set; }
        public string Tittle { get; set; }
        public int PopUpWidth { get; set; } = 420;
        public int PopUpHeight { get; set; } = 320;
        public string[] HideCols { get; set; }
        public DataGridViewRow ReturnRow { get; set; }
        public object CurrForm { get; set; }
        public string SQLQuery { get; set; }
        public UCComboBox()
        {
            InitializeComponent();
            FocusColor = Color.Moccasin;
        }

        protected override void OnPaintBackground(PaintEventArgs pevent)
        {
            base.OnPaintBackground(pevent);
            using (var brush = new SolidBrush(BackColor))
            {
                pevent.Graphics.FillRectangle(brush, ClientRectangle);
                pevent.Graphics.DrawRectangle(Pens.DarkGray, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
            }
        }
        private void UCComboBox_Enter(object sender, EventArgs e)
        {
            this.BackColor = FocusColor;
            this.SelectionStart = this.Text.Length;
            if (ErrProvider != null && (this.Text != null || this.Text != string.Empty))
            {
                ErrProvider.SetError(this, "");
            }
        }

        private void UCComboBox_Leave(object sender, EventArgs e)
        {
            this.BackColor = Color.White;
        }

        private void UCComboBox_Validating(object sender, CancelEventArgs e)
        {
            if (ValidateType == "Empty" && this.Text == string.Empty)
            {
                if (ErrProvider != null)
                {
                    ErrProvider.SetError(this, "Please Select Valid Data");
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

        private void UCComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && EnterKeyMoveNextTab == true)
            {
                SendKeys.Send("{TAB}");
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.N && MasterForm != null)
            {
                try
                {
                    Type Call = Type.GetType(MasterForm);
                    var Frm = Activator.CreateInstance(Call);
                    if (Frm != null)
                    {
                        CurrForm = GlobalVariables.CurrentForm;
                        ((Form)Frm).ShowDialog();
                        GlobalVariables.CurrentForm = CurrForm;
                        this.DataSource = Utility.SQLQuery(this.SQLQuery);
                    }
                }
                catch { }
            }
            else if (e.KeyCode == Keys.Down && this.DropDownStyle == ComboBoxStyle.DropDown && this.Text.ToString().Trim() == "")
            {
                if (this.DataSource != null && DisplayMember != null && ValueMember != null)
                {
                    var Frm = new UCPopUpBox();
                    Frm.DataSourse = (DataTable)this.DataSource;
                    Frm.DisplayMember = this.DisplayMember;
                    Frm.ValueMember = this.ValueMember;
                    Frm.Width = this.PopUpWidth;
                    Frm.Height = this.PopUpHeight;
                    Frm.HideCols = this.HideCols;
                    Frm.ShowDialog();
                    if (Frm.ReturnRow != null)
                    {
                        this.SelectedIndex = this.FindStringExact(Frm.ReturnRow.Cells[DisplayMember].Value.ToString().Trim());
                    }
                    ((DataTable)this.DataSource).DefaultView.RowFilter = null;
                    if (EnterKeyMoveNextTab == true)
                    {
                        SendKeys.Send("{TAB}");
                        e.Handled = true;
                    }
                    else
                    {
                        SendKeys.Send("{ENTER}");
                        e.Handled = true;
                    }
                }
            }
        }
    }
}

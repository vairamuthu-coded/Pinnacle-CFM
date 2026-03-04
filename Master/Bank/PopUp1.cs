using Microsoft.Office.Interop.Word;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Pinnacle.Master.Bank
{
    public partial class PopUp1 : Form
    {
        public PopUp1()
        {
            InitializeComponent();
            this.Font = Class.Users.FontName;
        
        }
       

            private void PopUp1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
            this.Hide();

        }

        private void PopUp1_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
        }

        private void PopUp1_Load(object sender, System.EventArgs e)
        {
            
                string sel = "select a.partycode,a.asptblpartymasid from asptblpartymas a order by a.partycode asc";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
                DataTable dt = ds.Tables["asptblpartymas"];

                if (dt != null)
                {

                    if (dt.Rows[0]["partycode"].ToString() != "")
                    {
                        combo_compcode.DataSource = dt;
                        combo_compcode.DisplayMember = "partycode";
                        combo_compcode.ValueMember = "asptblpartymasid";
                    }
                    else
                    {
                        combo_compcode.DataSource = null;
                    }
                }
           
        }

        private void btn_sumbit_Click(object sender, System.EventArgs e)
        {
            Class.Users.CompCode1 = "";
            Class.Users.CompCode1 = combo_compcode.Text.Replace(".", "").Replace(" ", "").ToLower();
            this.Close();
        }

        private void combo_compcode_SelectedIndexChanged(object sender, System.EventArgs e)
        {

        } 

        private void refreshToolStripMenuItem_Click(object sender, System.EventArgs e)
        {
            PopUp1_Load(sender,e);
        }
    }
}

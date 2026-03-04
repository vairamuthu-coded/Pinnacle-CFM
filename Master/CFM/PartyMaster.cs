using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master.CFM
{
    public partial class PartyMaster : Form,ToolStripAccess
    {
        private static PartyMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        public static PartyMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new PartyMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public PartyMaster()
        {
            InitializeComponent();          
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }
        void empty()
        {
            this.BackColor = Class.Users.BackColors;

            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
        }
      
        private void PartyMaster_Load(object sender, EventArgs e)
        {
            
        }

        public void compid(string s)
        {
            string sel = "select max(a.asptblpartymasid1) as id  from  asptblpartymas a  where a.PartyCode='" + s.ToUpper() + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblpartymas");
            DataTable dt = ds.Tables["asptblpartymas"];
            Int64 partycount = 1;
            if (dt.Rows.Count >= 1)
            {
                txtpartyid1.Text = "";
                Int64 cc = Convert.ToInt64("0" + dt.Rows[0]["id"].ToString()) + partycount;
                txtpartyid1.Text = cc.ToString();
            }
            else
            {
                txtpartyid1.Text = "1";
            }
        }
        public DataTable state(Int64 s)
        {
            string sel = "select b.gtstatemastid,b.statename from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where b.active='T' and a.gtcitymastid='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];
            return dt;
        }
        public DataTable city()
        {
            string sel = "select a.gtcitymastid,a.cityname from  gtcitymast a  join gtstatemast b on a.state=b.gtstatemastid   join gtcountrymast c on b.country=c.gtcountrymastid   where a.active='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcitymast");
            DataTable dt = ds.Tables["gtcitymast"];

            combocity.DisplayMember = "cityname";
            combocity.ValueMember = "gtcitymastid";
            combocity.DataSource = dt;
            return dt;
        }
        public DataTable country(Int64 s)
        {
            string sel1 = "SELECT  b.gtcountrymastid,b.countryname from gtstatemast a join gtcountrymast b on a.country=b.gtcountrymastid where a.gtstatemastid='" + s + "' ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtstatemast");
            DataTable dt = ds.Tables["gtstatemast"];

            return dt;
        }
     

        private void PartyMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

       

        private void News_Click(object sender, EventArgs e)
        {

          
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {

        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {

        }

        private void Deletes_Click(object sender, EventArgs e)
        {

        }

        public void News()
        {
            throw new NotImplementedException();
        }

        public void Saves()
        {
            //try
            //{
            //    sb.Clear();
            //    if (txtifsccode.Text == "")
            //    {
            //        MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.txtifsccode.Focus();
            //        return;
            //    }
            //    if (combobank.SelectedValue == null)
            //    {
            //        MessageBox.Show("'State Name  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); this.combobank.Focus();
            //        return;
            //    }

            //    else
            //    {


            //        string chk = ""; sb.Clear();
            //        if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; checkactive.Checked = false; }
            //        sb.Append("select a.asptblifscmasid    from  asptblifscmas a    WHERE a.ifsc='" + comboifsc.sele + "' and a.branch='" + txtbranch.Text + "' and a.bankname='" + combobank.SelectedValue + "' and a.active='" + chk + "' and a.asptblifscmasid='" + txtifscid.Text + "'");
            //        DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "asptblifscmas");
            //        DataTable dt = ds.Tables["asptblifscmas"];
            //        if (dt.Rows.Count != 0)
            //        {
            //            MessageBox.Show("Child Record Found " + " Alert " + txtifsccode.Text, "Alert", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); empty();
            //        }
            //        else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtifscid.Text) == 0 || Convert.ToInt32("0" + txtifscid.Text) == 0)
            //        {
            //            sb.Clear();
            //            sb.Append("insert into asptblifscmas(ifsc, branch,bankname,active,createdby,modifiedby,ipaddress)  VALUES('" + txtifsccode.Text.ToUpper() + "','" + txtbranch.Text.ToUpper() + "','" + combobank.SelectedValue + "','" + chk + "','" + Class.Users.HUserName + "','" + Class.Users.HUserName + "','" + Class.Users.IPADDRESS + "' )");
            //            Utility.ExecuteNonQuery(sb.ToString());
            //            MessageBox.Show("Record Saved Successfully " + txtifsccode.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //            GridLoad(); empty();
            //        }
            //        else
            //        {
            //            sb.Clear();
            //            sb.Append("update  asptblifscmas  set ifsc='" + txtifsccode.Text.ToUpper() + "', branch='" + txtbranch.Text.ToUpper() + "',bankname='" + combobank.SelectedValue + "' , active='" + chk + "' , modifiedby='" + Class.Users.HUserName + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblifscmasid='" + txtifscid.Text + "'");
            //            Utility.ExecuteNonQuery(sb.ToString());
            //            MessageBox.Show("Record Updated Successfully " + txtifsccode.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //            GridLoad();
            //            empty();
            //        }


            //    }

            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show("IFSC " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //}
        }

        public void Prints()
        {
            throw new NotImplementedException();
        }

        public void Searchs()
        {
            throw new NotImplementedException();
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

        public void Deletes()
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
            throw new NotImplementedException();
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
            throw new NotImplementedException();
        }

        public void GridLoad()
        {
            throw new NotImplementedException();
        }
    }
}

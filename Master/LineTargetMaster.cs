using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class LineTargetMaster : Form,ToolStripAccess
    {
        public LineTargetMaster()
        {
            InitializeComponent();

         
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
          
        }
        private static LineTargetMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Sewing sew = new Models.Sewing();
        Models.UserRights sm = new Models.UserRights();
        Int64 griddelrow;
     
    
        public static LineTargetMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new LineTargetMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        ListView listfilter = new ListView();

       
        private void LineTargetMaster_Load(object sender, EventArgs e)
        {
            News();
            compcodeload();
            GridLoad(); shiftload(); lineload();
             BuyerLoad(); GroupLoad(); SectionLoad();
            txtsearch.Select();
        }
        public void News()
        {
            empty(); 



        }

       

        public void StyleGroupLoad(string s)
        {
            try
            {

                DataTable dt = sew.StyleGroup(s);
                combostylename.DisplayMember = "stylename";
                combostylename.ValueMember = "asptblstymasid";
                combostylename.DataSource = dt;
                combostylename.Text = ""; combostylename.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void BuyerLoad()
        {
            try
            {

                DataTable dt = sew.Buyer();
                combobuyer.DisplayMember = "buyername";
                combobuyer.ValueMember = "asptblbuymasid";
                combobuyer.DataSource = dt;
                combobuyer.Text = ""; combobuyer.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void GroupLoad()
        {
            try
            {

                DataTable dt = sew.Group();
                combogroup.DisplayMember = "stylegroup";
                combogroup.ValueMember = "asptblstygrpmasid";
                combogroup.DataSource = dt;
                combogroup.Text = ""; combogroup.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void SectionLoad()
        {
            try
            {

                DataTable dt = sew.section();
                combosection.DisplayMember = "section";
                combosection.ValueMember = "asptblsecmasid";
                combosection.DataSource = dt;


                combosection.Text = ""; combosection.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void compcodeload()
        {
            try
            {

                DataTable dt = mas.comcode();
                compcode.DisplayMember = "compcode";
                compcode.ValueMember = "gtcompmastid";
                compcode.DataSource = dt;
                compcode.Text = ""; compcode.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void compcodeload(string s)
        {
            try
            {

                DataTable dt = mas.findcomcode(s);
                combocompname.DisplayMember = "compname";
                combocompname.ValueMember = "gtcompmastid";
                combocompname.DataSource = dt;
               
            }
            catch (Exception EX)
            { }
        }
        public void lineload()
        {
            try
            {

                DataTable dt = sew.Line();
                combolineno.DisplayMember = "lineno";
                combolineno.ValueMember = "asptbllinemasid";
                combolineno.DataSource = dt;
                combolineno.Text = ""; combolineno.SelectedIndex = -1;
            }
            catch (Exception EX)
            { }
        }
        public void shiftload()
        {
            try
            {

                DataTable dt = sew.shift();
                comboshiftno.DisplayMember = "shiftno";
                comboshiftno.ValueMember = "asptblshitypeid";
                comboshiftno.DataSource = dt;

            }
            catch (Exception EX)
            { }
        }








        public void Saves()
        {
            Int64 maxid = 0; Int64 j = 0;

            try
            {

                if (txtsearch.Text != "")
                {
                    Models.stylegroupModel c1 = new Models.stylegroupModel();
                    c1.active = "";
                    Models.Validate va = new Models.Validate();
                    c1.asptblstygrpmasid = Convert.ToInt64("0" + txtlinetargetid.Text);
                    c1.stylegroup = Convert.ToString(txtsearch.Text);
                    if (checkactive.Checked == true)
                        c1.active = "T";
                    else
                        c1.active = "F";

                    c1.compcode = Convert.ToInt64(Class.Users.COMPCODE);
                    c1.username = Convert.ToInt64(Class.Users.USERID);
                    c1.createdby = Convert.ToString(Class.Users.HUserName);
                    c1.createdon = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString());
                    c1.modified = Convert.ToDateTime(System.DateTime.Now.ToLongTimeString());
                    c1.modifiedby = Class.Users.HUserName;
                    c1.ipAddress = Class.Users.IPADDRESS;






                    string sel = "select asptblstygrpmasid    from  asptblstygrpmas   WHERE   stylegroup='" + c1.stylegroup + "'and active='" + c1.active + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblstygrpmas");
                    DataTable dt = ds.Tables["asptblstygrpmas"];
                    if (dt.Rows.Count != 0)
                    {

                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtlinetargetid.Text) == 0 || Convert.ToInt32("0" + txtlinetargetid.Text) == 0)
                    {
                        string ins = "insert into asptblstygrpmas(stylegroup,active,compcode,username,createdby,createdon,modified,ipaddress) values('" + c1.stylegroup + "','" + c1.active + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + Convert.ToDateTime(c1.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Class.Users.IPADDRESS + "');";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptblstygrpmasid) as asptblstygrpmasid   from  asptblstygrpmas ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblstygrpmas");
                        DataTable dt2 = ds2.Tables["asptblstygrpmas"]; maxid = 0;
                        maxid = Convert.ToInt64(dt2.Rows[0]["asptblstygrpmasid"].ToString());
                    }
                    else
                    {
                        string up = "update  asptblstygrpmas  set   stylegroup='" + c1.stylegroup + "',active='" + c1.active + "', compcode='" + Class.Users.COMPCODE + "',  username='" + Class.Users.USERID + "',createdby='" + Class.Users.HUserName + "', modified='" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblstygrpmasid='" + c1.asptblstygrpmasid + "';";
                        Utility.ExecuteNonQuery(up);
                        maxid = 0;
                        maxid = Convert.ToInt64(txtlinetargetid.Text);

                    }
                    int i = 0;
                    Models.stylegroupModeldet c = new Models.stylegroupModeldet();

                    int cc = dataGridView1.Rows.Count;
                    if (cc >= 1)
                    {
                        for (i = 0; i < cc; i++)
                        {
                            if (Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value) >= 1)
                            {
                                if (txtlinetargetid.Text == "") { c.asptblstygrpmasid = Convert.ToInt64("0" + maxid); }
                                else { c.asptblstygrpmasid = Convert.ToInt64("0" + txtlinetargetid.Text); }
                                c.stylename = Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value);
                                c.stylegroup = Convert.ToString(txtsearch.Text);
                                c.notes = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);

                                string sel1 = "select asptblstygrpdetid    from  asptblstygrpdet   where  asptblstygrpmasid='" + c.asptblstygrpmasid + "' and  stylename='" + c.stylename + "'  and stylegroup='" + c.stylegroup + "'and notes='" + c.notes + "';";
                                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblstygrpdet");
                                DataTable dt1 = ds1.Tables["asptblstygrpdet"];
                                if (dt1.Rows.Count != 0)
                                {

                                }
                                else if (dt1.Rows.Count != 0 && Convert.ToInt64("0" + c.asptblstygrpdetid) == 0 || Convert.ToInt64("0" + c.asptblstygrpdetid) == 0)
                                {

                                    string ins1 = "insert into asptblstygrpdet(asptblstygrpmasid,stylename,stylegroup,notes) values('" + c.asptblstygrpmasid + "' ,'" + c.stylename + "' ,'" + c.stylegroup + "','" + c.notes + "' );";
                                    Utility.ExecuteNonQuery(ins1);
                                }
                                else
                                {
                                    string up1 = "update  asptblstygrpdet  set asptblstygrpmasid='" + c.asptblstygrpmasid + "', stylename='" + c.stylename + "',stylegroup='" + c.stylegroup + "',notes='" + c.notes + "'  where asptblstygrpdetid='" + c.asptblstygrpdetid + "';";
                                    Utility.ExecuteNonQuery(up1);
                                }
                            }
                        }

                    }


                    if (txtlinetargetid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtlinetargetid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        News();

                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtlinetargetid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        News();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Saves_Click " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }



        private void empty()
        {
            txtlinetargetid.Text = "";
            txtsearch.Text = "";
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.EnableHeadersVisualStyles = false;
            listView1.Font = Class.Users.FontName;
            this.Font = Class.Users.FontName;
            do
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView1.Rows.Count > 1);
            if (listView1.Items.Count > 0)
            {
                listView1.Items[0].Selected = true;
            }
        }
        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear(); //date_format('" + dateTimePicker3.Value.ToString("yyyy-MM-dd") + "', '%Y-%m-%d')
                string sel1 = "select a.asptblstygrpmasid,a.stylegroup, a.active from asptblstygrpmas a   order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblstygrpmas");
                DataTable dt = ds.Tables["asptblstygrpmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblstygrpmasid"].ToString());
                        list.SubItems.Add(myRow["stylegroup"].ToString());
                        list.SubItems.Add(myRow["active"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        listView1.Items.Add(list);
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        i++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GridLoad: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ListView1_ItemActivate(object sender, EventArgs e)
        {

        }




        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();


        }

        public void Deletes()
        {
            try
            {
                if (txtlinetargetid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {


                        string del1 = "delete from asptblstygrpmas where   asptblstygrpmasid='" + txtlinetargetid.Text + "';";
                        Utility.ExecuteNonQuery(del1);
                        string del = "delete from asptblstygrpdet where  asptblstygrpdetid='" + txtlinetargetid.Text + "';";
                        Utility.ExecuteNonQuery(del);
                        MessageBox.Show("Record Deleted Successfully " + txtlinetargetid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletes_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


       

        private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        if (txtlinetargetid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow > 0)
                                {
                                    string del1 = "delete from  asptblstygrpdet     Where  asptblstygrpdetid='" + griddelrow + "';";
                                    Utility.ExecuteNonQuery(del1);

                                    griddelrow = 0;
                                }
                                else
                                {
                                    dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception EX)
            {
                // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            }

        }








        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        if (txtlinetargetid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow > 0)
                                {
                                    string del1 = "delete from  asptblstygrpdet     Where  asptblstygrpdetid='" + griddelrow + "';";
                                    Utility.ExecuteNonQuery(del1);
                                    dataGridView1.Rows.RemoveAt(Convert.ToInt32(griddelrow));
                                    griddelrow = 0;
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
            //try
            //{
            //    if (txtstylegroupdetid.Text != "")
            //    {
            //        griddelrow = 0;
            //        griddelrow = Convert.ToInt64("0" + dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
            //    }
            //   // dataGridView1.BeginEdit(true);
            //}
            //catch(Exception ex)
            //{

            //}
        }



        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length >= 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    listView1.Items.Clear();
                    item0 = listfilter.Items.Count;
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        


                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                        list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listView1.Items.Add(list);
                       
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }

        }


       
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }

      
        
        public void Prints()
        {
        }

        public void Searchs()
        {
        }

        public void Searchs(int EditID)
        {
        }

    
        public void ReadOnlys()
        {
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
            GlobalVariables.MdiPanel.Show();
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        void ToolStripAccess.GridLoad()
        {
            throw new NotImplementedException();
        }

        private void compcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            compcodeload(compcode.Text);
        }

        private void compcode_Click(object sender, EventArgs e)
        {
           
        }

        private void LineTargetMaster_FormClosing(object sender, FormClosingEventArgs e)
        {
            _instance = null;
        }

        private void combogroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            StyleGroupLoad(combogroup.Text);
        }
    }
}

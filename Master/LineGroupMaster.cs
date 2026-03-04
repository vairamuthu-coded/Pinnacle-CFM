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
    public partial class LineGroupMaster : Form,ToolStripAccess
    {
        public LineGroupMaster()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }
        private static LineGroupMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Int64 griddelrow;
        public static LineGroupMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new LineGroupMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        ListView listfilter = new ListView();

     
        private void linegroupMaster_Load(object sender, EventArgs e)
        {
            News();
            GridLoad();
            txtsearch.Select();
        }
        public void News()
        {
            empty();
            GridLoad(); linenameload();



        }

        public void linenameload()
        {
            try
            {
                string sel = "select  asptbllinemasid,lineno from  asptbllinemas  order by 1 ;";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllinemas");
                DataTable dt = ds.Tables["asptbllinemas"];

                linename.DisplayMember = "lineno";
                linename.ValueMember = "asptbllinemasid";
                linename.DataSource = dt;

            }
            catch (Exception EX)
            { }
        }

        public void Saves()
        {
            Int64 maxid = 0; Int64 j = 0;

            try
            {

                if (txtlinegroup.Text != "")
                {
                    Models.LineGroupModel c1 = new Models.LineGroupModel();
                    c1.active = "";
                    Models.Validate va = new Models.Validate();
                    c1.asptbllingrpmasid = Convert.ToInt64("0" + txtlinegroupid.Text);
                    c1.linegroup = Convert.ToString(txtlinegroup.Text);
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






                    string sel = "select asptbllingrpmasid    from  asptbllingrpmas   WHERE   linegroup='" + c1.linegroup + "'and active='" + c1.active + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllingrpmas");
                    DataTable dt = ds.Tables["asptbllingrpmas"];
                    if (dt.Rows.Count != 0)
                    {
                       
                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtlinegroupid.Text) == 0 || Convert.ToInt32("0" + txtlinegroupid.Text) == 0)
                    {
                        string ins = "insert into asptbllingrpmas(linegroup,active,compcode,username,createdby,createdon,modified,ipaddress) values('" + c1.linegroup + "','" + c1.active + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + Convert.ToDateTime(c1.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Class.Users.IPADDRESS + "');";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptbllingrpmasid) as asptbllingrpmasid   from  asptbllingrpmas ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbllingrpmas");
                        DataTable dt2 = ds2.Tables["asptbllingrpmas"]; maxid = 0;
                        maxid = Convert.ToInt64(dt2.Rows[0]["asptbllingrpmasid"].ToString());
                    }
                    else
                    {
                        string up = "update  asptbllingrpmas  set   linegroup='" + c1.linegroup + "',active='" + c1.active + "', compcode='" + Class.Users.COMPCODE + "',  username='" + Class.Users.USERID + "',createdby='" + Class.Users.HUserName + "', modified='" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptbllingrpmasid='" + c1.asptbllingrpmasid + "';";
                        Utility.ExecuteNonQuery(up);
                        maxid = 0;
                        maxid = Convert.ToInt64(txtlinegroupid.Text);

                    }
                    int i = 0;
                    Models.LineGroupdetModel c = new Models.LineGroupdetModel();

                    int cc = dataGridView1.Rows.Count - 1;
                    if (cc >= 0)
                    {
                        for (i = 0; i < cc; i++)
                        {
                            if (Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value) >= 1)
                            {
                                if (txtlinegroupid.Text == "") { c.asptbllingrpmasid = Convert.ToInt64("0" + maxid); }
                                else { c.asptbllingrpmasid = Convert.ToInt64("0" + txtlinegroupid.Text); }
                                c.asptbllingrpdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[1].Value);
                                c.linename = Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value);
                                c.linegroup = Convert.ToString(txtlinegroup.Text);
                                c.notes = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);

                                string sel1 = "select asptbllingrpdetid    from  asptbllingrpdet   where  asptbllingrpmasid='" + c.asptbllingrpmasid + "' and  linename='" + c.linename + "'  and linegroup='" + c.linegroup + "' and notes='" + c.notes + "' ;";//
                                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbllingrpdet");
                                DataTable dt1 = ds1.Tables["asptbllingrpdet"];
                                if (dt1.Rows.Count != 0)
                                {

                                }
                                else if (dt1.Rows.Count != 0 && Convert.ToInt64("0" + c.asptbllingrpdetid) == 0 || Convert.ToInt64("0" + c.asptbllingrpdetid) == 0)
                                {
                                    string ins1 = "insert into asptbllingrpdet(asptbllingrpmasid,linename,linegroup,notes) values('" + c.asptbllingrpmasid + "' ,'" + c.linename + "' ,'" + c.linegroup + "','" + c.notes + "' );";
                                    Utility.ExecuteNonQuery(ins1);
                                }
                                else
                                {
                                    string up1 = "update  asptbllingrpdet  set asptbllingrpmasid='" + c.asptbllingrpmasid + "', linename='" + c.linename + "',linegroup='" + c.linegroup + "',notes='" + c.notes + "'  where asptbllingrpdetid='" + c.asptbllingrpdetid + "';";
                                    Utility.ExecuteNonQuery(up1);
                                }
                            }
                        }

                    }


                    if (txtlinegroupid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtlinegroupid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        News();

                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtlinegroupid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            txtlinegroupid.Text = "";
            txtlinegroup.Text = "";
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            listView1.Font = Class.Users.FontName;
            this.Font = Class.Users.FontName;
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.EnableHeadersVisualStyles = false;

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
                string sel1 = "select a.asptbllingrpmasid,a.linegroup, a.active from asptbllingrpmas a   order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllingrpmas");
                DataTable dt = ds.Tables["asptbllingrpmas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbllingrpmasid"].ToString());
                        list.SubItems.Add(myRow["linegroup"].ToString());
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
            try
            {
                // empty();
                if (listView1.Items.Count > 0)
                {

                    txtlinegroupid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select asptbllingrpmasid,linegroup, active from asptbllingrpmas    where asptbllingrpmasid='" + txtlinegroupid.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbllingrpmas");
                    DataTable dt = ds.Tables["asptbllingrpmas"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtlinegroupid.Text = Convert.ToString(dt.Rows[0]["asptbllingrpmasid"].ToString());
                        txtlinegroup.Text = Convert.ToString(dt.Rows[0]["linegroup"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        string sel2 = "select a.asptbllingrpdetid, a.asptbllingrpmasid,b.asptbllinemasid,a.linegroup,a.notes from asptbllingrpdet a join asptbllinemas b on b.asptbllinemasid = a.linename    where a.asptbllingrpmasid='" + txtlinegroupid.Text + "';";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbllingrpdet");
                        DataTable dt1 = ds2.Tables["asptbllingrpdet"];
                        if (dt1.Rows.Count > 0)
                        {

                            dataGridView1.Rows.Clear();
                            for (i = 0; i < dt1.Rows.Count; i++)
                            {
                                dataGridView1.Rows.Add();
                                if (Convert.ToInt64(dt1.Rows[i]["asptbllingrpdetid"].ToString()) > 0)
                                {

                                    dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptbllingrpdetid"].ToString());
                                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptbllingrpmasid"].ToString());
                                    dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptbllinemasid"].ToString());
                                    dataGridView1.Rows[i].Cells[4].Value = Convert.ToString(dt1.Rows[i]["linegroup"].ToString());
                                    dataGridView1.Rows[i].Cells[5].Value = dt1.Rows[i]["notes"].ToString();

                                }
                            }


                        }

                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("ListView1_ItemActivate: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            GridLoad();


        }

        public void Deletes()
        {
            try
            {
                if (txtlinegroupid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {


                        string del1 = "delete from asptbllingrpmas where   asptbllingrpmasid='" + txtlinegroupid.Text + "';";
                        Utility.ExecuteNonQuery(del1);
                        string del = "delete from asptbllingrpdet where  asptbllingrpdetid='" + txtlinegroupid.Text + "';";
                        Utility.ExecuteNonQuery(del);
                        MessageBox.Show("Record Deleted Successfully " + txtlinegroupid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        GridLoad(); empty();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletes_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void buttnew_Click(object sender, EventArgs e)
        {
            News();
        }
        //private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        //{
        //    //this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        //}


        private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        if (txtlinegroupid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow > 0)
                                {
                                    string del1 = "delete from  asptbllingrpdet     Where  asptbllingrpdetid='" + griddelrow + "';";
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

                        if (txtlinegroupid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow > 0)
                                {
                                    string del1 = "delete from  asptbllingrpdet     Where  asptbllingrpdetid='" + griddelrow + "';";
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
            try
            {
                if (Convert.ToInt64(txtlinegroupid.Text) > 0)
                {
                    griddelrow = 0;
                   // griddelrow = Convert.ToInt64("0" + dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
                }
                // dataGridView1.BeginEdit(true);
            }
            catch (Exception ex)
            {

            }
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
                        if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
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

                    listView1.Items.Clear(); item0 = listfilter.Items.Count;
                    foreach (ListViewItem item in listfilter.Items)
                    {

                        this.listView1.Items.Add((ListViewItem)item.Clone());
                        if (item0 % 2 == 0)
                        {
                            item.BackColor = System.Drawing.Color.White;

                        }
                        else
                        {
                            item.BackColor = System.Drawing.Color.WhiteSmoke;
                        }

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


        private void linenameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void refreshToolStripMenuItem_Click_3(object sender, EventArgs e)
        {

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
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
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            empty();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

    }
}

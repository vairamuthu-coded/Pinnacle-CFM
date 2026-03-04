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
    public partial class StyleGroupMaster : Form,ToolStripAccess
    {
        public StyleGroupMaster()
        {
            InitializeComponent();
                       Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }
        private static StyleGroupMaster _instance; Int64 griddelrow;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        
        public static StyleGroupMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new StyleGroupMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        ListView listfilter = new ListView();

        public void usercheck(string s, string ss, string sss)
        {
           
        }
        private void stylegroupMaster_Load(object sender, EventArgs e)
        {
            News();
          
            txtsearch.Select();
        }
        public void News()
        {
            empty();
            GridLoad(); stylenameload();



        }

        public void stylenameload()
        {
            try
            {
                string sel = "select ASPTBLSTYMASid,stylename from  ASPTBLSTYMAS where active='T' order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLSTYMAS");
                DataTable dt = ds.Tables["ASPTBLSTYMAS"]; 

                stylename.DisplayMember = "stylename";
                stylename.ValueMember = "ASPTBLSTYMASid";
                stylename.DataSource = dt;

            }
            catch (Exception EX)
            { }
        }

        public void Saves()
        {
            Int64 maxid = 0; Int64 j = 0;

            try
            {

                if (txtstylegroup.Text != "")
                {
                    Models.stylegroupModel c1 = new Models.stylegroupModel();
                    c1.active = "";
                    Models.Validate va = new Models.Validate();
                    c1.asptblstygrpmasid = Convert.ToInt64("0" + txtstylegroupid.Text);
                    c1.stylegroup = Convert.ToString(txtstylegroup.Text);
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
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtstylegroupid.Text) == 0 || Convert.ToInt32("0" + txtstylegroupid.Text) == 0)
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
                        maxid = Convert.ToInt64(txtstylegroupid.Text);

                    }
                    int i = 0;
                    Models.stylegroupModeldet c = new Models.stylegroupModeldet();

                    int cc = dataGridView1.Rows.Count-1;
                    if (cc >= 0)
                    {
                        for (i = 0; i < cc; i++)
                        {
                            if (Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value) >= 1)
                            {
                                if (dataGridView1.Rows[i].Cells[1].Value == null && txtstylegroupid.Text == "")
                                {

                                    c.asptblstygrpdetid = 0;
                                    c.asptblstygrpmasid = Convert.ToInt64("0" + maxid);
                                }
                                if (dataGridView1.Rows[i].Cells[0].Value == null && txtstylegroupid.Text != "")
                                {

                                    c.asptblstygrpdetid = 0;
                                    c.asptblstygrpmasid = Convert.ToInt64("0" + txtstylegroupid.Text);
                                }
                                else
                                { 
                                    c.asptblstygrpmasid = Convert.ToInt64("0" + txtstylegroupid.Text);
                              
                                    c.asptblstygrpdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[1].Value.ToString());
                                }
                              
                                c.stylename = Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value);
                                c.stylegroup = Convert.ToString(txtstylegroup.Text);
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


                    if (txtstylegroupid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtstylegroupid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        News();

                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtstylegroupid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            txtstylegroupid.Text = "";
            txtstylegroup.Text = "";
            this.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            GlobalVariables.HideCols = new string[] { "SNo", "asptblstygrpDetid", "asptblstygrpmasid" };
            CommonFunctions.RemoveColumn(dataGridView1, GlobalVariables.HideCols);
            CommonFunctions.SetRowNumber(dataGridView1);

            do
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView1.Rows.Count > 1);
            if (listView1.Items.Count > 0)
            {
                listView1.Items[0].Selected = true;
            }
            this.Font = Class.Users.FontName;
            listView1.Font = Class.Users.FontName;
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
                        list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                        listView1.Items.Add(list);
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
               
                if (listView1.Items.Count > 0)
                {

                    txtstylegroupid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select asptblstygrpmasid,stylegroup, active from asptblstygrpmas    where asptblstygrpmasid='" + txtstylegroupid.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblstygrpmas");
                    DataTable dt = ds.Tables["asptblstygrpmas"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtstylegroupid.Text = Convert.ToString(dt.Rows[0]["asptblstygrpmasid"].ToString());
                        txtstylegroup.Text = Convert.ToString(dt.Rows[0]["stylegroup"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        string sel2 = "select B.asptblstygrpdetid, B.asptblstygrpmasid,C.asptblstymasid,A.stylegroup,B.notes from asptblstygrpmas a join asptblstygrpdet b on a.asptblstygrpmasid=b.asptblstygrpmasid join asptblstymas c on c.asptblstymasid = b.stylename   where a.asptblstygrpmasid='" + txtstylegroupid.Text + "';";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblstygrpdet");
                        DataTable dt1 = ds2.Tables["asptblstygrpdet"];
                        if (dt1.Rows.Count > 0)
                        {
                            dataGridView1.Rows.Clear();
                            for (i = 0; i < dt1.Rows.Count; i++)
                            {
                                if (Convert.ToInt64(dt1.Rows[i]["asptblstygrpdetid"].ToString()) > 0)
                                {
                                    dataGridView1.Rows.Add();
                                    dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblstygrpdetid"].ToString());
                                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblstygrpmasid"].ToString());
                                    dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblstymasid"].ToString());
                                    dataGridView1.Rows[i].Cells[4].Value = Convert.ToString(dt1.Rows[i]["stylegroup"].ToString());
                                    dataGridView1.Rows[i].Cells[5].Value = dt1.Rows[i]["notes"].ToString();
                                }
                            }

                            CommonFunctions.SetRowNumber(dataGridView1);
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


            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }

        public void Deletes()
        {
            try
            {
                if (txtstylegroupid.Text != "")
                {
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        string sel1 = "select distinct a.asptblstygrpmasid from asptblstygrpmas a join asptblpur b on a.asptblstygrpmasid=b.stylename where a.asptblstygrpmasid='" + txtstylegroupid.Text + "'";
                        DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblstygrpmas");
                        DataTable dt = ds.Tables["asptblstygrpmas"];
                        if (dt.Rows.Count > 0)
                        {
                            MessageBox.Show("Child Record Found.Can Not Delete." + txtstylegroupid.Text, " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                        }
                        else
                        {


                            string del1 = "delete from asptblstygrpmas where   asptblstygrpmasid='" + txtstylegroupid.Text + "';";
                            Utility.ExecuteNonQuery(del1);
                            string del = "delete from asptblstygrpdet where  asptblstygrpdetid='" + txtstylegroupid.Text + "';";
                            Utility.ExecuteNonQuery(del);
                            MessageBox.Show("Record Deleted Successfully " + txtstylegroupid.Text, " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            GridLoad(); empty();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletes_Click: " + ex.Message, " Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

                        if (txtstylegroupid.Text != "")
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

                        if (txtstylegroupid.Text != "")
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
            try
            {
                //if (txtstylegroupid.Text != "")
                //{
                //    griddelrow = 0;
                //    griddelrow = Convert.ToInt64(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
                //}
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
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            if (item0 % 2 == 0) { item.BackColor = Color.White; } else { item.BackColor = Color.WhiteSmoke; }
                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }else
                {
                    ListView ll = new ListView();

                    listView1.Items.Clear(); 
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        

                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            if (item0 % 2 == 0) { item.BackColor = Color.White; } else { item.BackColor = Color.WhiteSmoke; }
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

       
        private void styleNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();
        }

        private void refreshToolStripMenuItem_Click_3(object sender, EventArgs e)
        {
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
        }

        private void rowDeleteToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        if (txtstylegroupid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            dataGridView1.BeginEdit(true);
                            if (confirmation == DialogResult.Yes)
                            {
                                if (griddelrow >= 1)
                                {
                                    string del1 = "delete from  asptblsizegroupDet     Where  asptblsizegroupdetid='" + griddelrow + "';";
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
            empty();
            this.Hide();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);

        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }


        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3 && dataGridView1.CurrentCell.Value != null)
            {
                
               mas.checkduplicate(e.ColumnIndex,dataGridView1);
                //foreach(DataGridViewRow row in dataGridView1.Rows)
                //{
                //    if (row.Index == this.dataGridView1.CurrentCell.RowIndex)
                //    {
                //        continue;
                //    }
                //    if(this.dataGridView1.CurrentCell.Value == null)
                //    {
                //        continue;
                //    }
                //    if(row.Cells[3].Value != null && row.Cells[3].Value.ToString()==dataGridView1.CurrentCell.Value.ToString())
                //    {
                //        MessageBox.Show("Duplicate not allowed ");
                //        dataGridView1.CurrentCell.Value = null;
                //    }
                //}
            }
        }
    }
}

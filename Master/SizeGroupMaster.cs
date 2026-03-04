using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class SizeGroupMaster : Form,ToolStripAccess
    {
        public SizeGroupMaster()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
        }
        private static SizeGroupMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Int64 griddelrow = 0;
       
        public static SizeGroupMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new SizeGroupMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }



        ListView listfilter = new ListView();

        private void SizeGroupMaster_Load(object sender, EventArgs e)
        {
            News();
           
            txtsearch.Select(); 
        }
        public void News()
        {
            empty();
            GridLoad(); sizenameload();



        }

        public void sizenameload()
        {
            try
            {
                string sel = "select asptblsizmasid,sizename from  asptblsizmas  order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsizmas");
                DataTable dt = ds.Tables["asptblsizmas"];
                sizename.DisplayMember = "sizename";
                sizename.ValueMember = "asptblsizmasid";
                sizename.DataSource = dt;
               
            }
            catch (Exception EX)
            { }
        }








        public void Saves()
        {
            Int64 maxid = 0; Int64 j = 0;

            try
            {
                
                if (txtsizegroup.Text != "")
                {
                    Models.SizeGroupModel c1 = new Models.SizeGroupModel();
                    c1.active = "";
                    Models.Validate va = new Models.Validate();
                    c1.asptblsizgrpid = Convert.ToInt64("0" + txtsizegroupdetid.Text);
                    c1.sizegroup = Convert.ToString(txtsizegroup.Text.ToUpper());
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
                    string sel = "select asptblsizgrpid    from  asptblsizgrp   WHERE   sizegroup='" + c1.sizegroup + "'and active='" + c1.active + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsizgrp");
                    DataTable dt = ds.Tables["asptblsizgrp"];
                    if (dt.Rows.Count != 0)
                    {

                    }
                    else if (dt.Rows.Count != 0 && Convert.ToInt32("0" + txtsizegroupdetid.Text) == 0 || Convert.ToInt32("0" + txtsizegroupdetid.Text) == 0)
                    {
                        string ins = "insert into asptblsizgrp(sizegroup,active,compcode,username,createdby,createdon,modified,ipaddress) values('" + c1.sizegroup + "','" + c1.active + "','" + Class.Users.COMPCODE + "','" + Class.Users.USERID + "','" + Class.Users.HUserName + "','" + Convert.ToDateTime(c1.createdon).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "','" + Class.Users.IPADDRESS + "');";
                        Utility.ExecuteNonQuery(ins);
                        string sel2 = "select max(asptblsizgrpid) as asptblsizgrpid   from  asptblsizgrp ;";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblsizgrp");
                        DataTable dt2 = ds2.Tables["asptblsizgrp"]; maxid = 0;
                        maxid = Convert.ToInt64(dt2.Rows[0]["asptblsizgrpid"].ToString());
                    }
                    else
                    {
                        string up = "update  asptblsizgrp  set   sizegroup='" + c1.sizegroup + "',active='" + c1.active + "', compcode='" + Class.Users.COMPCODE + "',  username='" + Class.Users.USERID + "',createdby='" + Class.Users.HUserName + "', modified='" + Convert.ToDateTime(c1.modified).ToString("yyyy-MM-dd hh:mm:ss") + "',ipaddress='" + Class.Users.IPADDRESS + "' where asptblsizgrpid='" + c1.asptblsizgrpid + "';";
                        Utility.ExecuteNonQuery(up);
                        maxid = 0;
                        maxid = Convert.ToInt64(txtsizegroupdetid.Text);

                    }
                    int i = 0;
                    Models.SizeGroupDetModel c = new Models.SizeGroupDetModel();
                    
                        int cc = dataGridView1.Rows.Count-1;
                        if (cc >= 1)
                        {
                        for (i = 0; i < cc; i++)
                        {
                            if (Convert.ToInt64(dataGridView1.Rows[i].Cells[3].Value) >= 1)
                            {
                                if (txtsizegroupdetid.Text == "") { c.asptblsizgrpid = Convert.ToInt64("0" + txtsizegroupdetid.Text); }
                                else { c.asptblsizgrpid = Convert.ToInt64("0" + txtsizegroupdetid.Text); c.asptblsizgrpdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[1].Value); }
                                c.sizename = Convert.ToInt64("0"+dataGridView1.Rows[i].Cells[3].Value);
                                c.sizegroup = Convert.ToString(txtsizegroup.Text);
                                c.notes = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);

                                string sel1 = "select asptblsizgrpdetid    from  asptblsizgrpDet   where  asptblsizgrpid='" + c.asptblsizgrpid + "' and   sizename='" + c.sizename + "'  and  sizegroup='" + c.sizegroup + "' and notes='" + c.notes + "';";
                                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptblsizgrpDet");
                                DataTable dt1 = ds1.Tables["asptblsizgrpDet"];
                                if (dt1.Rows.Count != 0)
                                {

                                }
                                else if (dt1.Rows.Count != 0 && Convert.ToInt64("0" + c.asptblsizgrpdetid) == 0 || Convert.ToInt64("0" + c.asptblsizgrpdetid) == 0)
                                {

                                    string ins1 = "insert into asptblsizgrpDet(asptblsizgrpid,sizename,sizegroup,notes) values('" + c.asptblsizgrpid + "' ,'" + c.sizename + "' ,'" + c.sizegroup + "','" + c.notes + "' );";
                                    Utility.ExecuteNonQuery(ins1);
                                }
                                else
                                {
                                    string up1 = "update  asptblsizgrpDet  set asptblsizgrpid='" + c.asptblsizgrpid + "', sizename='" + c.sizename + "',sizegroup='" + c.sizegroup + "',notes='" + c.notes + "'  where asptblsizgrpdetid='" + c.asptblsizgrpdetid + "';";
                                    Utility.ExecuteNonQuery(up1);
                                }
                            }
                        }

                        }
                   
                  
                    if (txtsizegroupdetid.Text == "")
                    {
                        MessageBox.Show("Record Saved Successfully " + txtsizegroupdetid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        News();

                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtsizegroupdetid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            txtsizegroupdetid.Text = ""; 
            txtsizegroup.Text = "";
            butheader.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
                GlobalVariables.HideCols = new string[] { "SNo", "asptblsizgrpDetid", "asptblsizgrpid", "sizegroup" };
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
                string sel1 = "select a.asptblsizgrpid,a.sizegroup, a.active from asptblsizgrp a   order by 1;";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizgrp");
                DataTable dt = ds.Tables["asptblsizgrp"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptblsizgrpid"].ToString());
                        list.SubItems.Add(myRow["sizegroup"].ToString());
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
                // empty();
                if (listView1.Items.Count > 0)
                {

                    txtsizegroupdetid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = "select asptblsizgrpid,sizegroup, active from asptblsizgrp    where asptblsizgrpid='" + txtsizegroupdetid.Text + "';";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizgrp");
                    DataTable dt = ds.Tables["asptblsizgrp"];
                    int i = 1;
                    if (dt.Rows.Count > 0)
                    {
                        txtsizegroupdetid.Text = Convert.ToString(dt.Rows[0]["asptblsizgrpid"].ToString());                              
                        txtsizegroup.Text = Convert.ToString(dt.Rows[0]["sizegroup"].ToString());
                        if (dt.Rows[0]["active"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        string sel2 = "select a.asptblsizgrpdetid, a.asptblsizgrpid,b.asptblsizmasid AS sizename,a.sizegroup,a.notes from asptblsizgrpDet a join asptblsizmas b on b.asptblsizmasid = a.sizename  where a.asptblsizgrpid='" + txtsizegroupdetid.Text + "';";
                        DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptblsizgrpDet");
                        DataTable dt1 = ds2.Tables["asptblsizgrpDet"];
                        if (dt1.Rows.Count > 0)
                        {
                            
                                dataGridView1.Rows.Clear();
                           
                         
                            for (i = 0; i < dt1.Rows.Count; i++)

                            {
                               
                                if (Convert.ToInt64(dt1.Rows[i]["asptblsizgrpdetid"].ToString()) > 0)
                                {
                                    dataGridView1.Rows.Add();
                                    dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblsizgrpdetid"].ToString());
                                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["asptblsizgrpid"].ToString());
                                    dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["sizename"].ToString());


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
         

        }

        public void Deletes()
        {
            try
            {
                if (txtsizegroupdetid.Text != "")
                {

                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        
                            string sel1 = "select a.asptblsizgrpid from asptblsizgrp a join ASPTBLPUR b on a.asptblsizgrpid=b.sizegroup where a.asptblsizgrpid='" + txtsizegroupdetid.Text + "'";
                            DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptblsizgrp");
                            DataTable dt = ds.Tables["asptblsizgrp"];
                        if (dt.Rows.Count > 0)
                        {
                            MessageBox.Show("Child Record Found.Can Not Delete.", " Alert Message ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                        }
                        else
                        {

                            string del = "delete from asptblsizgrp where asptblsizgrpid='" + Convert.ToInt64("0" + txtsizegroupdetid.Text) + "'";
                            Utility.ExecuteNonQuery(del);
                            string del1 = "delete from asptblsizgrpdet where asptblsizgrpid='" + Convert.ToInt64("0" + txtsizegroupdetid.Text) + "'";
                            Utility.ExecuteNonQuery(del1);
                            MessageBox.Show("Record Deleted Successfully ", " Delete Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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


        private void buttnew_Click(object sender, EventArgs e)
        {
            News();
        }
        //private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        //{
        //    //this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        //}


        //private void rowDeleteToolStripMenuItem1_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
        //        {
        //            if (oneCell.Selected)
        //            {

        //                if (txtsizegroupdetid.Text != "")
        //                {
        //                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        //                    if (confirmation == DialogResult.Yes)
        //                    {
        //                        if (griddelrow >0)
        //                        {
        //                            string del1 = "delete from  asptblsizgrpDet     Where  asptblsizgrpdetid='" + griddelrow + "';";
        //                            Utility.ExecuteNonQuery(del1);

        //                            griddelrow = 0;
        //                        }
        //                        else
        //                        {
        //                            dataGridView1.Rows.RemoveAt(Convert.ToInt32(oneCell.RowIndex));
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //    }
        //    catch (Exception EX)
        //    {
        //        // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
        //    }

        //}
    

       
      

      
      

        private void rowDeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected && griddelrow>0)
                    {



                        if (txtsizegroupdetid.Text != null)
                        {
                            string sel = "select a.sizename from  ASPTBLPURDET1 a join asptblsizmas b on a.sizename=b.asptblsizmasid   Where  b.asptblsizmasid='" + oneCell.Value + "'; ";
                            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLPURDET1");
                            DataTable dt = ds.Tables["ASPTBLPURDET1"];
                            if (dt.Rows.Count > 0)
                            {
                                MessageBox.Show("Child Record Found.Can Can not Delete.");
                            }
                            else
                            {
                                if (Convert.ToInt64("0"+ griddelrow) > 0)
                                {
                                    string del1 = "delete from  asptblsizgrpDet     Where  asptblsizgrpdetid='" + griddelrow + "';";
                                    Utility.ExecuteNonQuery(del1);
                                    dataGridView1.Rows.RemoveAt(Convert.ToInt32(oneCell.RowIndex));
                                    griddelrow = 0;
                                    GridLoad();

                                }
                            }
                        }
                        else
                        {
                            dataGridView1.Rows.RemoveAt(Convert.ToInt32(griddelrow));
                            dataGridView1.Update(); return;
                        }
                        
                    }
                    //}
                    //else
                    //{
                    //    dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                    //}
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
                if (e.ColumnIndex == 3)
                {
                    if (txtsizegroupdetid.Text != "")
                    {
                        griddelrow = 0;
                        if (dataGridView1.Rows[e.RowIndex].Cells[1].EditedFormattedValue != "")
                        {
                            griddelrow = Convert.ToInt64("0" + dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
                        }
                    }
                }
                dataGridView1.BeginEdit(true);
            }
            catch (Exception ex)
            {

            }
        }

        private void refreshToolStripMenuItem_Click_2(object sender, EventArgs e)
        {
            
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; 
                if (txtsearch.Text.Length >= 1)
                {
                    listView1.Items.Clear();
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
                            listView1.Items.Add(list);
                          
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                //else
                //{
                //    ListView ll = new ListView();
                //    listView1.Items.Clear();

                //    foreach (ListViewItem item in listfilter.Items)
                //    {

                //        this.listView1.Items.Add((ListViewItem)item.Clone());

                //        item0++;
                //    }
                //    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                //}


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }

        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GridLoad();

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }
        List<string> sizearray = new List<string>();
        string[] siz;
        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == 3)
            //{
            //    int cc = 0;
            //    foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
            //    {
            //        if (oneCell.Selected)
            //        {
                       
            //            string s = dataGridView1.Rows[oneCell.RowIndex].Cells[3].EditedFormattedValue.ToString();
            //            if (!(sizearray.Contains(s)) && s != null)
            //            {
            //                sizearray.Add(s);
            //                siz = sizearray.ToArray<string>();
            //            }
            //            else
            //            {
            //                 MessageBox.Show("Duplicate");
                          
            //            }
            //        }
            //    }
                
            //}
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
           
            if (e.ColumnIndex == 3)
            {
                //foreach (DataGridViewRow row in dataGridView1.Rows)
                //{
                //    IEnumerable<DataGridViewCell> cellsWithValusInRows = from DataGridViewCell cell in row.Cells where string.IsNullOrEmpty((string)cell.Value)
                //                                                         select cell;
                //}

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {

                        string s = dataGridView1.Rows[oneCell.RowIndex].Cells[3].EditedFormattedValue.ToString();
                        if (!(sizearray.Contains(s)) && s != null)
                        {
                            sizearray.Add(s);
                            siz = sizearray.ToArray<string>();
                        }
                        else
                        {
                            dataGridView1.Rows[oneCell.RowIndex].Cells[3].Value = null;
                            MessageBox.Show("Duplicate" + s);
                           
                        }
                    }
                }

            }
        }

        private void dataGridView1_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void SizeGroupMaster_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
        }

        private void lblsearch_Click(object sender, EventArgs e)
        {

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

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }

        private void dataGridView1_DataError_1(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = false;
        }
    }
}

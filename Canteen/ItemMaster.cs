using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Canteen
{
    public partial class ItemMaster : Form,ToolStripAccess
    {
        public ItemMaster()
        {
            InitializeComponent();
          //  usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
           
        }
        public void Searchs(int EditID)
        {

        }
        private static ItemMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
        byte[]bytes; OpenFileDialog open = new OpenFileDialog();
        public static ItemMaster Instance
        {
            get { if (_instance == null) _instance = new ItemMaster(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }



        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                        //if (dt1.Rows[r]["NEWS"].ToString() == "T") { this.News.Visible = true; } else { this.News.Visible = false; }
                        //if (dt1.Rows[r]["SAVES"].ToString() == "T") { this.Saves.Visible = true; } else { this.Saves.Visible = false; }
                        //if (dt1.Rows[r]["PRINTS"].ToString() == "T") { this.Prints.Visible = true; } else { this.Prints.Visible = false; }
                        //if (dt1.Rows[r]["READONLY"].ToString() == "T") { this.Enabled = true; } else { this.Enabled = false; }
                        //if (dt1.Rows[r]["SEARCH"].ToString() == "T") { this.Searchs.Visible = true; } else { this.Searchs.Visible = false; }
                        //if (dt1.Rows[r]["DELETES"].ToString() == "T") { this.Deletes.Visible = true; } else { this.Deletes.Visible = false; }
                        //if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = true; } else { this.TreeButtons.Visible = false; }
                        //if (dt1.Rows[r]["GLOBALSEARCH"].ToString() == "T") { this.GlobalSearchs.Visible = true; } else { this.GlobalSearchs.Visible = false; }
                        //if (dt1.Rows[r]["LOGIN"].ToString() == "T") { this.Logins.Visible = true; } else { this.Logins.Visible = false; }
                        //if (dt1.Rows[r]["CHANGEPASSWORD"].ToString() == "T") { this.ChangePasswords.Visible = true; } else { this.ChangePasswords.Visible = false; }
                        //if (dt1.Rows[r]["CHANGESKIN"].ToString() == "T") { ChangeSkins.Visible = true; } else { ChangeSkins.Visible = false; }
                        //if (dt1.Rows[r]["DOWNLOAD"].ToString() == "T") { this.DownLoads.Visible = true; } else { this.DownLoads.Visible = false; }
                        //if (dt1.Rows[r]["Pdf"].ToString() == "T") { this.Pdfs.Visible = true; } else { this.Pdfs.Visible = false; }
                        //if (dt1.Rows[r]["Imports"].ToString() == "T") { this.Imports.Visible = true; } else { this.Imports.Visible = false; }
                    }
                }


            }
            else
            {

                GlobalVariables.Toolstrip1.Enabled = false;
            }

        }
        private void News_Click(object sender, EventArgs e)
        {
            News();
        }

        private void Btnsave_Click(object sender, EventArgs e)
        {

          
        }
       public void GridLoad()
        {
            try
            {
                listcanitem.Items.Clear();
                string sel1 = "  SELECT A.ASPTBLCANITEMMASID,  A.ITEMCODE,A.ITEMNAME1, B.CATEGORY,A.EMPLOYEECOST, A.CONTRACTORCOST,A.ACTIVE  FROM  ASPTBLCANITEMMAS A  JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY=B.ASPTBLCANCATEGORYMASID ORDER BY 1";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANITEMMAS");
                DataTable dt1 = ds1.Tables["ASPTBLCANITEMMAS"];
                if (dt1.Rows.Count > 0)
                {
                    int mycount = 1;
                    foreach (DataRow myRow in dt1.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = mycount.ToString();
                        list.SubItems.Add(myRow["ASPTBLCANITEMMASID"].ToString());
                        list.SubItems.Add(myRow["ITEMCODE"].ToString());
                        list.SubItems.Add(myRow["ITEMNAME1"].ToString());
                        list.SubItems.Add(myRow["CATEGORY"].ToString());
                        
                            list.SubItems.Add(myRow["EMPLOYEECOST"].ToString());
                        list.SubItems.Add(myRow["CONTRACTORCOST"].ToString());                       
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        if (mycount % 2 == 0)
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        else
                        {
                            list.BackColor = Color.White;
                        }
                        listcanitem.Items.Add(list);
                        mycount++;
                    }
                    lblcanitemtotal.Text = "Total Rows    :" + listcanitem.Items.Count;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void Txtitemsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
               
                if (txtitemsearch.Text != "")
                {
                    int iGLCount = 1;
                    string sel1 = "  SELECT A.ASPTBLCANITEMMASID,  A.ITEMCODE,A.ITEMNAME1, B.CATEGORY,A.EMPLOYEECOST,A.CONTRACTORCOST,A.ACTIVE  FROM  ASPTBLCANITEMMAS A  JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY=B.ASPTBLCANCATEGORYMASID  WHERE A.ITEMNAME1 LIKE'%" + txtitemsearch.Text + "%'  or B.CATEGORY LIKE'%" + txtitemsearch.Text + "%'  OR A.ACTIVE LIKE'%" + txtitemsearch.Text + "%'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANITEMMAS");
                    DataTable dt = ds.Tables["ASPTBLCANITEMMAS"];
                    if (dt.Rows.Count > 0)
                    {
                        listcanitem.Items.Clear();
                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["ASPTBLCANITEMMASID"].ToString());
                            list.SubItems.Add(myRow["ITEMCODE"].ToString());
                            list.SubItems.Add(myRow["ITEMNAME1"].ToString());
                            list.SubItems.Add(myRow["CATEGORY"].ToString());
                            
                          
                            list.SubItems.Add(myRow["EMPLOYEECOST"].ToString());
                            list.SubItems.Add(myRow["CONTRACTORCOST"].ToString());
                            list.SubItems.Add(myRow["ACTIVE"].ToString());
                            listcanitem.Items.Add(list);
                            iGLCount++;
                        }
                        lblcanitemtotal.Text = "Total Count    :" + listcanitem.Items.Count;
                    }
                    else
                    {
                        listcanitem.Items.Clear();
                    }
                }
                else
                {
                    listcanitem.Items.Clear();
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void listcanitem_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listcanitem.Items.Count > 0)
                {
                    pictureitem.Image = null;
                    txtitemid.Text = listcanitem.SelectedItems[0].SubItems[1].Text;
                    string sel1 = "SELECT A.ASPTBLCANITEMMASID,  A.ITEMCODE,A.ITEMNAME1,B.CATEGORY,A.EMPLOYEECOST,A.CONTRACTORCOST,A.ITEMIMAGE,A.ACTIVE  FROM  ASPTBLCANITEMMAS A  JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY=B.ASPTBLCANCATEGORYMASID   WHERE A.ASPTBLCANITEMMASID=" + txtitemid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANITEMMAS");
                    DataTable dt = ds.Tables["ASPTBLCANITEMMAS"];
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow myRow in dt.Rows)
                        {
                            txtitemid.Text = Convert.ToString(myRow["ASPTBLCANITEMMASID"].ToString());

                            txtitemcode.Text = Convert.ToString(myRow["CATEGORY"].ToString()) + txtitemid.Text;
                            txtitemname.Text = Convert.ToString(myRow["ITEMNAME1"].ToString());
                            combocategory.Text = Convert.ToString(myRow["CATEGORY"].ToString());
                            txtempcost.Text = Convert.ToString(myRow["EMPLOYEECOST"].ToString());
                            txtcontcost.Text = Convert.ToString(myRow["CONTRACTORCOST"].ToString());
                           
                            if (myRow["ITEMIMAGE"].ToString() != "")
                            {

                                bytes = (byte[])myRow["ITEMIMAGE"];
                                Image img = Models.Device.ByteArrayToImage(bytes);
                                pictureitem.Image = img;


                            }
                            if (dt.Rows[0]["ACTIVE"].ToString() == "T") { checkactive.Checked = true; } else { checkactive.Checked = false; }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ItemMaster_Load(object sender, EventArgs e)
        {
            string sel1 = "  SELECT A.ASPTBLCANCATEGORYMASID,  A.Category  FROM  ASPTBLCANCATEGORYMAS A   ORDER BY 1";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANCATEGORYMAS");
            DataTable dt1 = ds1.Tables["ASPTBLCANCATEGORYMAS"];
            combocategory.DataSource = dt1;
            combocategory.ValueMember = "ASPTBLCANCATEGORYMASID";
            combocategory.DisplayMember = "Category";
         
            combocategory.SelectedIndex = -1;
           
            GridLoad(); this.txtitemname.Select(); 
        }

        private void Pictureitem_Click(object sender, EventArgs e)
        {
            try
            {
                bytes = null;
                PictureBox p = sender as PictureBox;
                if (p != null)
                {
                    open.Filter = "(*.jpg;*.jpeg;*.bmp;)| *.jpg; *.jpeg; *.bmp";
                    if (open.ShowDialog() == DialogResult.OK)
                    {

                        p.Image = new Bitmap(open.FileName);
                        bytes = Models.Device.ImageToByteArray(p);
                      

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Label7_Click(object sender, EventArgs e)
        {

        }

        private void Saves_Click(object sender, EventArgs e)
        {
            Btnsave_Click(sender,e);
        }

        private void Combocategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtitemid.Text == "")
                {
                    string sel = "select MAX(A.ASPTBLCANITEMMASID)+1 AS  ID FROM ASPTBLCANITEMMAS A ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLCANITEMMAS");
                    DataTable dt = ds.Tables["ASPTBLCANITEMMAS"];
                    txtitemcode.Text = combocategory.Text + dt.Rows[0]["ID"].ToString();
                }
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message.ToString());
            }
        }

        private void Txtitemcost_TextChanged(object sender, EventArgs e)
        {

        }

        private void Label6_Click(object sender, EventArgs e)
        {

        }

        private void Checkactive_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Label4_Click(object sender, EventArgs e)
        {

        }

        private void Txtitemname_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void Txtitemid_TextChanged(object sender, EventArgs e)
        {

        }

        public void News()
        {
            txtitemid.Text = "";
            txtitemcode.Text = "";
            txtitemname.Text = "";
            combocategory.Text = "";
            txtcontcost.Text = "";
            txtempcost.Text = "";
            checkactive.Checked = true;
            pictureitem.Image = null;
        }

        public void Saves()
        {
            try
            {
                string chk = "";

                if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }
                string sel = "select A.ASPTBLCANITEMMASID FROM ASPTBLCANITEMMAS A  WHERE  A.ITEMCODE='" + txtitemcode + "' AND  A.ITEMNAME1='" + txtitemname.Text + "' AND A.CATEGORY='" + txtitemcode.Text + "' AND A.EMPLOYEECOST=" + txtempcost.Text + " AND A.CONTRACTORCOST=" + txtcontcost.Text + " AND  A.ACTIVE='" + chk + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLCANITEMMAS");
                DataTable dt = ds.Tables["ASPTBLCANITEMMAS"];
                if (dt.Rows.Count != 0)
                {
                    MessageBox.Show("Child Record Found    :" + txtitemname.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else if (dt.Rows.Count == 0 && Convert.ToInt32("0" + txtitemid.Text) == 0 || Convert.ToInt32("0" + txtitemid.Text) == 0)
                {
                    string ins = "INSERT INTO ASPTBLCANITEMMAS(ITEMCODE,ITEMNAME1,CATEGORY,EMPLOYEECOST,CONTRACTORCOST,ITEMCOST,EMPLOYEETYPE,ACTIVE, USERNAME,  MODIFIED,CREATEDON,IPADDRESS)VALUES('" + txtitemcode.Text.ToUpper() + "','" + txtitemname.Text.ToUpper() + "'," + combocategory.SelectedValue + "," + txtempcost.Text + "," + txtcontcost.Text + ",'" + chk + "'," + Class.Users.USERID + ",'" + Convert.ToString(Class.Users.CREATED) + "','" + Convert.ToString(Class.Users.CREATED) + "','" + Class.Users.IPADDRESS + "' )";
                    Utility.ExecuteNonQuery(ins);
                    MessageBox.Show("Record Saved Successfully     :" + txtitemname.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    string up = "UPDATE ASPTBLCANITEMMAS SET   ITEMCODE='" + txtitemcode.Text.ToUpper() + "', ITEMNAME1='" + txtitemname.Text.ToUpper() + "', CATEGORY='" + combocategory.SelectedValue + "' ,EMPLOYEECOST='" + txtempcost.Text + "',CONTRACTORCOST=" + txtcontcost.Text + ",ITEMIMAGE=:EMPIMAGE,ACTIVE='" + chk + "', USERNAME=" + Class.Users.USERID + ",  MODIFIED='" + Convert.ToString(Class.Users.CREATED).ToString() + "',IPADDRESS='" + Class.Users.IPADDRESS + "' WHERE  ASPTBLCANITEMMASID=" + txtitemid.Text;
                    Utility.ExecuteNonQuery(up, bytes);
                    MessageBox.Show("Record Updated      :" + txtitemname.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                GridLoad();
                News();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        public void Prints()
        {
            
        }

        public void Searchs()
        {
            
        }

        public void Deletes()
        {
            
        }

        public void ReadOnlys()
        {
            
        }

        public void Imports()
        {
            
        }

        public void Pdfs()
        {
            
        }

        public void ChangePasswords()
        {
            
        }

        public void DownLoads()
        {
            
        }

        public void ChangeSkins()
        {
            
        }

        public void Logins()
        {
            
        }

        public void GlobalSearchs()
        {
            
        }

        public void TreeButtons()
        {
            
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
            this.Hide();
        }

      
    }
}

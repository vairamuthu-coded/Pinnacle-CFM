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
    public partial class CanteenCategoryMaster : Form,ToolStripAccess
    {
        public CanteenCategoryMaster()
        {
            InitializeComponent();
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
         
            //Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
        }
        private static CanteenCategoryMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
        OpenFileDialog open = new OpenFileDialog();
        byte[] bytes;
        public static CanteenCategoryMaster Instance
        {
            get { if (_instance == null) _instance = new CanteenCategoryMaster(); GlobalVariables.CurrentForm = _instance; return _instance; }

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

        private void Saves_Click(object sender, EventArgs e)
        {
            Saves();
        }
        public void GridLoad()
        {
            try
            {
                listcategory.Items.Clear();
                string sel1 = "  SELECT A.ASPTBLCANCATEGORYMASID,  A.Category,A.ACTIVE  FROM  ASPTBLCANCATEGORYMAS A   ORDER BY 1";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANCATEGORYMAS");
                DataTable dt1 = ds1.Tables["ASPTBLCANCATEGORYMAS"];
                if (dt1.Rows.Count > 0)
                {
                    int mycount = 1;
                    foreach (DataRow myRow in dt1.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = mycount.ToString();
                        list.SubItems.Add(myRow["ASPTBLCANCATEGORYMASID"].ToString());
                        list.SubItems.Add(myRow["Category"].ToString());                       
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        if (mycount % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                      
                        listcategory.Items.Add(list);
                        mycount++;
                    }
                    lblcatetotal.Text = "Total Rows    :" + listcategory.Items.Count;
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
   

        private void Txtcancategorysearch_Click(object sender, EventArgs e)
        {
            try
            {
                listcategory.Items.Clear();
                if (txtcancategorysearch.Text != "")
                {
                    int iGLCount = 1;
                    string sel1 = "  SELECT  A.ASPTBLCANCATEGORYMASID,A.Category,A.ACTIVE FROM   ASPTBLCANCATEGORYMAS A WHERE A.Category LIKE'%" + txtcancategorysearch.Text + "%'  or A.ACTIVE LIKE'%" + txtcancategorysearch.Text + "%'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANCATEGORYMAS");
                    DataTable dt = ds.Tables["ASPTBLCANCATEGORYMAS"];
                    if (dt.Rows.Count > 0)
                    {

                        foreach (DataRow myRow in dt.Rows)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = iGLCount.ToString();
                            list.SubItems.Add(myRow["ASPTBLCANCATEGORYMASID"].ToString());                         
                            list.SubItems.Add(myRow["Category"].ToString());
                            list.SubItems.Add(myRow["ACTIVE"].ToString());
                            listcategory.Items.Add(list);
                            iGLCount++;
                        }
                        lblcatetotal.Text = "Total Count    :" + listcategory.Items.Count;
                    }
                }
                else
                {
                    GridLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Listcategory_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listcategory.Items.Count > 0)
                {
                    pictureitem.Image = null;
                    txtcategoryid.Text = listcategory.SelectedItems[0].SubItems[1].Text;
                    string sel1 = " SELECT A.ASPTBLCANCATEGORYMASID,A.Category,A.CATEGORYIMAGE,A.ACTIVE  FROM  ASPTBLCANCATEGORYMAS A WHERE A.ASPTBLCANCATEGORYMASID=" + txtcategoryid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANCATEGORYMAS");
                    DataTable dt = ds.Tables["ASPTBLCANCATEGORYMAS"];
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow myRow in dt.Rows)
                        {
                            txtcategoryid.Text = Convert.ToString(myRow["ASPTBLCANCATEGORYMASID"].ToString());

                            txtcategory.Text = Convert.ToString(myRow["Category"].ToString());
                            if (myRow["CATEGORYIMAGE"].ToString() != "")
                            {

                                bytes = (byte[])myRow["CATEGORYIMAGE"];
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

       

        private void CanteenCategoryMaster_Load(object sender, EventArgs e)
        {
            GridLoad();txtcategory.Select();
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

        public void News()
        {
            txtcategoryid.Text = "";
            txtcategory.Text = "";
            checkactive.Checked = true;
            pictureitem.Image = null;
        }

        public void Saves()
        {
            try
            {
                string chk = "";

                if (checkactive.Checked == true) { chk = "T"; } else { chk = "F"; }
                string sel = "select A.ASPTBLCANCATEGORYMASID FROM ASPTBLCANCATEGORYMAS A  WHERE  A.CATEGORY='" + txtcategory.Text + "'  AND A.ACTIVE='" + chk + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLCANCATEGORYMAS");
                DataTable dt = ds.Tables["ASPTBLCANCATEGORYMAS"];
                if (dt.Rows.Count != 0)
                {
                    MessageBox.Show("Child Record Found    :" + txtcategory.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else if (dt.Rows.Count == 0 && Convert.ToInt32("0" + txtcategoryid.Text) == 0 || Convert.ToInt32("0" + txtcategoryid.Text) == 0)
                {
                    string ins = "INSERT INTO ASPTBLCANCATEGORYMAS(CATEGORY,ACTIVE, USERNAME,  MODIFIED,CREATEDON,IPADDRESS)VALUES('" + txtcategory.Text + "','" + chk + "'," + Class.Users.USERID + ",'" + Convert.ToString(Class.Users.CREATED) + "','" + Convert.ToString(Class.Users.CREATED) + "','" + Class.Users.IPADDRESS + "' )";
                    Utility.ExecuteNonQuery(ins);
                    MessageBox.Show("Record Saved Successfully     :" + txtcategory.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    string up = "UPDATE ASPTBLCANCATEGORYMAS SET   CATEGORY='" + txtcategory.Text.ToUpper() + "',CATEGORYIMAGE=:EMPIMAGE,ACTIVE='" + chk + "', USERNAME=" + Class.Users.USERID + ",  MODIFIED='" + Convert.ToString(Class.Users.CREATED).ToString() + "',IPADDRESS='" + Class.Users.IPADDRESS + "' WHERE  ASPTBLCANCATEGORYMASID=" + txtcategoryid.Text;
                    Utility.ExecuteNonQuery(up, bytes);
                    MessageBox.Show("Record Updated      :" + txtcategory.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

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
    public partial class TokenCancel : Form,ToolStripAccess
    {
        public TokenCancel()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            
        }
        public void Searchs(int EditID)
        {

        }
        private static TokenCancel _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        Models.MenuName c = new Models.MenuName();
         OpenFileDialog open = new OpenFileDialog();
        public static TokenCancel Instance
        {
            get { if (_instance == null) _instance = new TokenCancel(); GlobalVariables.CurrentForm = _instance; return _instance; }

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

        private void TokenCancel_Load(object sender, EventArgs e)
        {
            GridLoad();txtTokenNumber.Focus();
        }

        private void TxtTokenNumber_TextChanged(object sender, EventArgs e)
        {
            
        }
       public void GridLoad()
        {
            try
            {
                listcanitem.Items.Clear();
               


                    string sel1 = "  SELECT DISTINCT  C.ASPTBLCANTOKENID,C.TOKENNO,D.IDCARDNO,C.EMPID, D.EMPNAME,C.TOKENNOCANCEL,A.EMPLOYEECOST  FROM ASPTBLCANITEMMAS A JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY = B.ASPTBLCANCATEGORYMASID  JOIN ASPTBLCANTOKEN C ON C.ITEMNAME1 = A.ASPTBLCANITEMMASID  JOIN ASPTBLEMP D ON D.ASPTBLEMPID = C.EMPID where C.TOKENNOCANCEL = 'F' ORDER BY 1";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANTOKEN");
                    DataTable dt1 = ds1.Tables["ASPTBLCANTOKEN"];
                if (dt1.Rows.Count > 0)
                {
                    int mycount = 1;
                    foreach (DataRow myRow in dt1.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.Text = mycount.ToString();
                        list.SubItems.Add(myRow["ASPTBLCANTOKENID"].ToString());
                        list.SubItems.Add(myRow["TOKENNO"].ToString());
                        list.SubItems.Add(myRow["IDCARDNO"].ToString());
                        list.SubItems.Add(myRow["EMPID"].ToString());
                        list.SubItems.Add(myRow["EMPNAME"].ToString());
                        list.SubItems.Add(myRow["TOKENNOCANCEL"].ToString());
                        list.SubItems.Add(myRow["EMPLOYEECOST"].ToString());
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
        private void Saves_Click(object sender, EventArgs e)
        {

        }

        private void Listcanitem_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listcanitem.Items.Count > 0)
                {
                    txttokennum.Text = listcanitem.SelectedItems[0].SubItems[1].Text;
                    txtTokenNumber.Text = listcanitem.SelectedItems[0].SubItems[2].Text;
                    if (listcanitem.SelectedItems[0].SubItems[6].Text == "T")
                    
                        checkactive.Checked = true;
                    else
                        checkactive.Checked = false;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Txtitemsearch_TextChanged(object sender, EventArgs e)
        {
            if (txtitemsearch.Text.Length >= 1)
            {
                string sel1 = "  SELECT DISTINCT  C.ASPTBLCANTOKENID,C.TOKENNO,D.IDCARDNO,C.EMPID, D.EMPNAME,C.TOKENNOCANCEL,A.EMPLOYEETYPE  FROM ASPTBLCANITEMMAS A JOIN ASPTBLCANCATEGORYMAS B ON A.CATEGORY = B.ASPTBLCANCATEGORYMASID  JOIN ASPTBLCANTOKEN C ON C.ITEMNAME1 = A.ASPTBLCANITEMMASID  JOIN ASPTBLEMP D ON D.ASPTBLEMPID = C.EMPID where C.TOKENNOCANCEL='F' AND D.IDCARDNO LIKE'%" + txtitemsearch.Text + "%'  OR D.EMPNAME LIKE'%" + txtitemsearch.Text + "%'  ORDER BY 1";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLCANTOKEN");
                DataTable dt1 = ds1.Tables["ASPTBLCANTOKEN"];
                if (dt1.Rows.Count > 0)
                {
                    int mycount = 1;
                    listcanitem.Items.Clear();
                    foreach (DataRow myRow in dt1.Rows)
                    {

                        ListViewItem list = new ListViewItem();
                        list.Text = mycount.ToString();
                        list.SubItems.Add(myRow["ASPTBLCANTOKENID"].ToString());
                        list.SubItems.Add(myRow["TOKENNO"].ToString());
                        list.SubItems.Add(myRow["IDCARDNO"].ToString());
                        list.SubItems.Add(myRow["EMPID"].ToString());
                        list.SubItems.Add(myRow["EMPNAME"].ToString());
                        list.SubItems.Add(myRow["TOKENNOCANCEL"].ToString());
                        list.SubItems.Add(myRow["EMPLOYEETYPE"].ToString());
                        listcanitem.Items.Add(list);
                        mycount++;
                    }
                    lblcanitemtotal.Text = "Total Rows    :" + listcanitem.Items.Count;

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

        public void News()
        {
            
        }

        public void Saves()
        {
            try
            {
                if (txtTokenNumber.Text.Length == 19)
                {
                    string chk = "";
                    if (checkactive.Checked == true)
                        chk = "T";
                    else
                        chk = "F";
                    Int64 tokenid = Convert.ToInt64("0" + txttokennum.Text);
                    string sel4 = "   SELECT A.ASPTBLCANTOKENID FROM ASPTBLCANTOKEN A  WHERE A.TOKENNO='" + txtTokenNumber.Text + "' and A.TOKENNOCANCEL='" + chk + "'";
                    DataSet ds4 = Utility.ExecuteSelectQuery(sel4, "ASPTBLCANTOKEN");
                    DataTable dt4 = ds4.Tables["ASPTBLCANTOKEN"];

                    if (dt4.Rows.Count != 0)
                    {
                        MessageBox.Show("This Token already Cancelled. " + txtTokenNumber.Text + " \n Date : " + dt4.Rows[0]["MODIFIED"].ToString(), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                        txtTokenNumber.Text = ""; txtTokenNumber.Focus();txttokennum.Text = ""; return;
                    }
                    else 
                    {

                      //  DialogResult result1 = MessageBox.Show("Do You want to Cancel??\n' THIS TOKEN NO  :" + txtTokenNumber.Text, "Canteen", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                       // if (result1.Equals(DialogResult.OK))
                       // {

                            string up = "UPDATE  ASPTBLCANTOKEN A SET A.TOKENNOCANCEL='" + chk + "'   WHERE  A.ASPTBLCANTOKENID='" + txttokennum.Text + "' ";
                            Utility.ExecuteNonQuery(up); this.lblcancelalert.Show();
                            MessageBox.Show("This Token are Cancelled Successfully " + txtTokenNumber.Text, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtTokenNumber.Text = ""; txtTokenNumber.Focus(); txttokennum.Text = "";

                        GridLoad();
                        //}
                        //else
                        //{
                        //    txtTokenNumber.Text = ""; txtTokenNumber.Focus();
                        //}


                    }
                }
            }
            catch (Exception EX)
            {
                MessageBox.Show("Invalid" + EX.Message.ToString());
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

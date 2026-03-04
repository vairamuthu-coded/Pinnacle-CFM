using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Pinnacle.Models;

namespace Pinnacle.Fuel
{
    public partial class FuelRateMaster : Form,ToolStripAccess
    {
        public FuelRateMaster()
        {
            InitializeComponent();

            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd/MM/yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            //usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
        }

        public void Searchs(int EditID)
        {

        }
        public void ReadOnlys()
        {

        }
        private static FuelRateMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();Int64 MAXID = 0;
        byte[] bytes; private int rowIndex = 0; ListView listfilter = new ListView(); string griddelrow = "";
        public static FuelRateMaster Instance
        {
            get { if (_instance == null) _instance = new FuelRateMaster(); GlobalVariables.CurrentForm = _instance; return _instance; }

        }
        Pinnacle.Models.FuelRateMasterModel c = new FuelRateMasterModel();

        public void usercheck(string s, string ss, string sss)
        {

            DataTable dt1 = sm.headerdropdowns(s, ss, sss);
            if (dt1.Rows.Count > 0)
            {
                if (dt1.Rows[0]["Menuname"].ToString() == Class.Users.ScreenName)
                {
                    for (int r = 0; r < dt1.Rows.Count; r++)
                    {

                        if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { combocompcode.Enabled = false; combobunk.Enabled = true; } else { combocompcode.Enabled = false; combobunk.Enabled = false; }
                      

                    }
                }

            }
            else
            {

                GlobalVariables.Toolstrip1.Enabled = false;
            }

        }
        private void FuelRateMaster_Load(object sender, EventArgs e)
        {
            empty();
            finyear();
           
            COMPCODE(); bunkfind();
            comboitem();
            GridLoad();
            combocompcode.Select();
        }
       
        void comboitem()
        {
            try
            {
                string sel1 = "SELECT DISTINCT  A.GTGENITEMMASTID,A.ITEMNAME FROM GTGENITEMMAST A WHERE A.FT='T' AND A.ACTIVE='T'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTGENITEMMAST");
                DataTable dt = ds.Tables["GTGENITEMMAST"];

                ITEMNAME.ValueMember = "GTGENITEMMASTID";
                ITEMNAME.DisplayMember = "ITEMNAME";
                ITEMNAME.DataSource = dt;
            }
            catch (Exception EX) { }
        }
        void finyear()
        {

            DataTable dt = mas.finyear();
            combofinyear.ValueMember = "gtfinancialyearid";
            combofinyear.DisplayMember = "finyear";
            combofinyear.DataSource = dt;
            Class.Users.Finyear = dt.Rows[0]["finyear"].ToString();
        }

        void COMPCODE()
        {

            DataTable dt = mas.compcodefind(Class.Users.HCompcode);

            combocompcode.DisplayMember = "COMPCODE";
            combocompcode.ValueMember = "GTCOMPMASTID";
            combocompcode.DataSource = dt;
       
        }
        void bunkfind()
        {

            DataTable dt = mas.bunkfind(Class.Users.HCompcode);
         
            combobunk.DisplayMember = "BUNKNAME";
            combobunk.ValueMember = "ASPTBLPETMASID";
            combobunk.DataSource = dt;
        }
        public void GridLoad()
        {

            listView1.Items.Clear();
            try
            {
                string sel = "SELECT A.ASPTBLFUELRATEMASID,A.ASPTBLFUELRATEID,B.COMPCODE,A.FUELDATE,C.BUNKNAME FROM ASPTBLFUELRATEMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID LEFT JOIN ASPTBLPETMAS C ON C.ASPTBLPETMASID=A.BUNKNAME WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND A.FUELDATE=to_date('" + dateTimePicker3.Value.ToShortDateString() + "','dd-MM-yyyy')order by ASPTBLFUELRATEMASID DESC";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLFUELRATEMAS");
                DataTable dt = ds.Tables["ASPTBLFUELRATEMAS"];

                if (dt.Rows.Count > 0)
                {
                    int iGLCount = 1;
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        ListViewItem list1 = new ListViewItem();

                        list1.SubItems.Add(iGLCount.ToString());
                        list1.SubItems.Add(dt.Rows[j]["ASPTBLFUELRATEMASID"].ToString());
                        list1.SubItems.Add(dt.Rows[j]["ASPTBLFUELRATEID"].ToString());
                        list1.SubItems.Add(dt.Rows[j]["COMPCODE"].ToString());                       
                        list1.SubItems.Add(dt.Rows[j]["FUELDATE"].ToString());
                        list1.SubItems.Add(dt.Rows[j]["BUNKNAME"].ToString());
                        iGLCount++;
                        listView1.Items.Add(list1);
                    }
                }
                lblattcount.Text = "Total Count    :" + listView1.Items.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

        }
        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; listView1.Items.Clear(); int item1 = 1;
                if (txtsearch.Text.Length >= 1 && listfilter.Items.Count > 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text.ToUpper()))
                        {


                            list.Text = item1.ToString();
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);                     
               
                            listView1.Items.Add(list);


                        }
                        item0++; item1++;
                    }
                    lblattcount.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {
                    ListView ll = new ListView();
                    listView1.Items.Clear(); listView1.BackColor = System.Drawing.Color.LightSteelBlue;
                    foreach (ListViewItem item in listfilter.Items)
                    {


                        this.listView1.Items.Add((ListViewItem)item.Clone());



                        item0++;
                    }
                    lblattcount.Text = "Total Count: " + listView1.Items.Count;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }

           
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                txtfuelrateid.Text = ""; //combobunk.Enabled = false; 

                txtfuelrateid.Text = Convert.ToString(listView1.SelectedItems[0].SubItems[2].Text);

                string sel = "SELECT A.ASPTBLFUELRATEMASID,A.ASPTBLFUELRATEID,C.FINYR AS FINYEAR, B.COMPCODE,A.FUELDATE,C.BUNKNAME,A.MONTHNAME FROM ASPTBLFUELRATEMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN  gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  LEFT JOIN ASPTBLPETMAS C ON C.ASPTBLPETMASID=A.BUNKNAME  WHERE A.ASPTBLFUELRATEMASID='" + txtfuelrateid.Text + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLFUELRATEMAS");
                DataTable dt = ds.Tables["ASPTBLFUELRATEMAS"];

              
                if (dt.Rows.Count > 0)
                {


                    txtfuelrateid.Text = Convert.ToString(dt.Rows[0]["ASPTBLFUELRATEMASID"].ToString());
                    txtfuelrateid1.Text = Convert.ToString(dt.Rows[0]["ASPTBLFUELRATEID"].ToString());
                    combofinyear.Text = Convert.ToString(dt.Rows[0]["FINYEAR"].ToString());
                    combocompcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                    combobunk.Text = Convert.ToString(dt.Rows[0]["BUNKNAME"].ToString());
                    txtmonthname.Text = Convert.ToString(dt.Rows[0]["MONTHNAME"].ToString());
                    dateTimePicker1.Value = Convert.ToDateTime(dt.Rows[0]["FUELDATE"].ToString());
                    
                    string sel1 = " SELECT B.ASPTBLFUELRATEMASDETID,B.ASPTBLFUELRATEMASID,B.ASPTBLFUELRATEID, B.COMPCODE,C.GTGENITEMMASTID as ITEMNAME,B.FUELRATE2,B.FUELDATE,B.BUNKNAME  FROM  ASPTBLFUELRATEMAS A JOIN ASPTBLFUELRATEMASDET B ON A.ASPTBLFUELRATEMASID=B.ASPTBLFUELRATEMASID   JOIN GTGENITEMMAST C ON C.GTGENITEMMASTID = B.ITEMNAME WHERE  A.ASPTBLFUELRATEMASID ='" + txtfuelrateid.Text + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLFUELRATEMASDET");
                    DataTable dt1 = ds1.Tables["ASPTBLFUELRATEMASDET"];
                  
                    if (dt1.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dt1;
                      
                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            if (Convert.ToInt64(dataGridView1.Rows[i].Cells[1].Value) > 0)
                            {
                                dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["ASPTBLFUELRATEMASDETID"].ToString());
                                dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["ASPTBLFUELRATEMASID"].ToString());
                                dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["ASPTBLFUELRATEID"].ToString());
                                dataGridView1.Rows[i].Cells[4].Value = Convert.ToInt64("0" + dt1.Rows[i]["COMPCODE"].ToString());
                                dataGridView1.Rows[i].Cells[5].Value = Convert.ToInt64("0" + dt1.Rows[i]["ITEMNAME"].ToString());
                                dataGridView1.Rows[i].Cells[6].Value = Convert.ToDecimal("0" + dt1.Rows[i]["FUELRATE2"].ToString());
                                dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(dt1.Rows[i]["FUELDATE"].ToString());
                                dataGridView1.Rows[i].Cells[8].Value = Convert.ToString(dt1.Rows[i]["BUNKNAME"].ToString());
                                dataGridView1.Rows[i].Cells[9].Value = dt1.Rows[i]["NOTES"].ToString();
                                //totlitres += Convert.ToInt32(dt1.Rows[i]["LITRES"].ToString());

                            }
                        }
                        //  txttotlitre.Text = totlitres.ToString();
                    }
                    else
                    {
                        do
                        {
                            for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
                        }
                        while (dataGridView1.Rows.Count > 1);


                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        private void empty()
        {

            txtfuelrateid.Text = "";
            txtfuelrateid1.Text = ""; 
            listfilter.Items.Clear();
            txtmonthname.Text = "";
            dateTimePicker1.Value= DateTime.Today.AddDays(0); dateTimePicker3.Value = DateTime.Today.AddDays(0);

            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.MenuHighlight;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.EnableHeadersVisualStyles = false;
            do
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView1.Rows.Count > 1);
        }

        bool ch = false; int rowcount = 1; int rowcount1 = 0;
        private bool Checks()
        {
            Models.Validate va = new Models.Validate();
            rowcount = 1; rowcount1 = 0;
            if (combocompcode.SelectedValue == null)
            {
                MessageBox.Show("CompCode is empty");

                this.Focus();
                this.combocompcode.Focus(); return false;
            }
            if (dateTimePicker1.Value.ToString() == null)
            {
                MessageBox.Show("EmpNmae is empty");
                this.dateTimePicker1.Focus();
                return false;
            }
            if (combobunk.SelectedValue.ToString() == null)
            {
                MessageBox.Show("BunkName is empty");
                this.combobunk.Focus();
                return false;
            }
            Models.FuelRateMasterModelDet c1 = new FuelRateMasterModelDet();
            rowcount1 = dataGridView1.Rows.Count-1;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
               
                c1.ITEMNAME = Convert.ToInt64("0" + row.Cells["ITEMNAME"].Value);
                c1.COMPCODE = Convert.ToInt64("0" + combocompcode.SelectedValue.ToString());
                c1.FUELRATE2 = Convert.ToString(row.Cells["FUELRATE2"].Value);
                c1.FUELDATE = Convert.ToDateTime(dateTimePicker1.Value.ToString());
                if (c1.ITEMNAME > 1)
                {
                    if (va.IsInteger(Convert.ToInt64("0" + c1.ITEMNAME).ToString()) == false)
                    {

                        MessageBox.Show("ItemName Field is empty");

                        return false;
                    }
                    if (va.IsDecimal(Convert.ToDecimal(c1.FUELRATE2).ToString()) == false)
                    {
                        MessageBox.Show("FuelRate Field is empty");
                        return false;
                    }
                    if (c1.FUELRATE2.Length > 6)
                    {
                        MessageBox.Show("FuelRate Field is empty");
                        row.Cells["FUELRATE2"].Value = "";
                        return false;
                    }
                }
                if (rowcount == rowcount1)
                {
                    if (c1.ITEMNAME > 1)
                    {
                        return true;
                    }
                }

                rowcount++;

            }
            return ch;
        }
       public void Saves() { 
            if (Checks() == true)
            {
//combocompcode_SelectedIndexChanged(sender,e);        
                DataTable dt = new DataTable();
                c.ASPTBLFUELRATEMASID = Convert.ToInt64("0" + txtfuelrateid.Text);
                c.ASPTBLFUELRATEID = Convert.ToInt64("0"+txtfuelrateid1.Text);
                c.COMPCODE = Convert.ToInt64("0" + combocompcode.SelectedValue);
                c.BUNKNAME = Convert.ToInt64("0" + combobunk.SelectedValue);
                c.FINYEAR= Convert.ToInt64("0" + combofinyear.SelectedValue);
                c.FUELDATE = Convert.ToDateTime(dateTimePicker1.Value.ToString("dd-MM-yyyy"));
             
                c.MONTHNAME = txtmonthname.Text;
                c.USERNAME = Class.Users.USERID;
                c.MODIFIED = Class.Users.CREATED;
                c.MODIFIED = Class.Users.CREATED;
                c.IPADDRESS = Class.Users.IPADDRESS;
              
                c.FUELTOKEN =  Class.Users.Finyear + "/" + Class.Users.HCompcode + "/" + txtfuelrateid1.Text;
                DataTable dt1 = c.select(c.FINYEAR, c.COMPCODE, c.FUELDATE,c.BUNKNAME);
                if (dt1.Rows.Count != 0)
                {
                   
                    MAXID = Convert.ToInt64("0" + txtfuelrateid.Text); 
                }
                else if (dt1.Rows.Count != 0 && c.ASPTBLFUELRATEMASID == 0 || c.ASPTBLFUELRATEMASID == 0)
                {
                    c = new FuelRateMasterModel(c.ASPTBLFUELRATEID, c.FINYEAR, c.COMPCODE, c.FUELDATE, c.USERNAME, c.MODIFIED, c.CREATEDON, c.IPADDRESS, c.FUELTOKEN, c.BUNKNAME, c.MONTHNAME);
                    string sel = "SELECT MAX(A.ASPTBLFUELRATEMASID) ID FROM  ASPTBLFUELRATEMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE A.COMPCODE=" + Class.Users.COMPCODE + " and A.bunkname=" + combobunk.SelectedValue;
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLFUELRATEMAS");
                    dt = ds.Tables["ASPTBLFUELRATEMAS"];
                    MAXID = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                }
                else
                {
                    c = new FuelRateMasterModel(c.ASPTBLFUELRATEID, c.FINYEAR, c.COMPCODE, c.FUELDATE, c.USERNAME, c.MODIFIED, c.IPADDRESS, c.FUELTOKEN, c.BUNKNAME, c.ASPTBLFUELRATEMASID,c.MONTHNAME);
                  
                }

                Models.FuelRateMasterModelDet c1 = new FuelRateMasterModelDet();
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {

                    if (txtfuelrateid.Text == "") { c1.ASPTBLFUELRATEMASID = Convert.ToInt64("0" + MAXID);  c1.ASPTBLFUELRATEID = Convert.ToInt64("0"+txtfuelrateid1.Text); }
                    else { c1.ASPTBLFUELRATEMASID = Convert.ToInt64("0" + txtfuelrateid.Text); c1.ASPTBLFUELRATEID = Convert.ToInt64(txtfuelrateid1.Text); }
                    c1.ASPTBLFUELRATEMASDETID = Convert.ToInt64("0" + row.Cells["ASPTBLFUELRATEMASDETID"].Value);
                    c1.COMPCODE = Convert.ToInt64("0" + combocompcode.SelectedValue);
                    c1.BUNKNAME = Convert.ToInt64("0" + combobunk.SelectedValue);
                    c1.ITEMNAME = Convert.ToInt64("0" + row.Cells["ITEMNAME"].Value);
                    c1.FUELRATE2 = Convert.ToString(row.Cells["FUELRATE2"].FormattedValue); 
                    c1.FUELDATE = Convert.ToDateTime(dateTimePicker1.Value.ToString("dd-MM-yyyy"));
                    if (c1.ITEMNAME > 1)
                    {
                        ch = true;
                        DataTable dt2 = c1.select(c1.COMPCODE, c1.ITEMNAME, c1.FUELRATE2, c1.FUELDATE, c1.BUNKNAME);
                        if (dt2.Rows.Count != 0) {  }
                        else if (dt2.Rows.Count != 0 && c1.ASPTBLFUELRATEMASDETID == 0 || c1.ASPTBLFUELRATEMASDETID == 0 && c1.ASPTBLFUELRATEMASID >= 1)
                        {
                            c1 = new Models.FuelRateMasterModelDet(c1.ASPTBLFUELRATEMASID, c1.ASPTBLFUELRATEID, c1.COMPCODE, c1.ITEMNAME, c1.FUELRATE2, c1.FUELDATE, c1.BUNKNAME);

                        }
                        else
                        {
                            c1 = new Models.FuelRateMasterModelDet(c1.ASPTBLFUELRATEMASID, c1.ASPTBLFUELRATEID, c1.COMPCODE, c1.ITEMNAME, c1.FUELRATE2, c1.FUELDATE, c1.BUNKNAME, c1.ASPTBLFUELRATEMASDETID);

                        }
                    }

                }
                if (txtfuelrateid.Text == "" && MAXID >=1)
                {
                    MessageBox.Show("Record Saved Successfully." + txtfuelrateid1.Text, "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
                if (txtfuelrateid.Text != "" && MAXID >= 1)
                {
                    
                    MessageBox.Show("Record Updated Successfully." + txtfuelrateid1.Text, "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GridLoad(); empty();
                }
                if (txtfuelrateid.Text == "" && MAXID == 0)
                {
                    MessageBox.Show("Child Record Found.." + txtfuelrateid1.Text, "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     empty();
                }

            }
            
        }
        //public void autono()
        //{

        //    if (Class.Users.HCompcode == "FLFD")
        //    {
        //        string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
        //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
        //        DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
        //        Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
        //        if (cnt == 0)
        //        {
        //            string sel2 = "  SELECT  MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO    FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  AND  A.ACTIVE='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
        //            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
        //            DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
        //            txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString();
        //            txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //        else
        //        {

        //            txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //    }

        //    if (Class.Users.HCompcode == "FLF")
        //    {
        //        string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";

        //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
        //        DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
        //        Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
        //        if (cnt == 0)
        //        {
        //            string sel2 = "  SELECT  MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO   FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  AND  A.ACTIVE='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
        //            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
        //            DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
        //            txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //        else
        //        {

        //            txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //    }
        //    if (Class.Users.HCompcode == "AGF")
        //    {
        //        string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
        //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
        //        DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
        //        Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
        //        if (cnt == 0)
        //        {
        //            string sel2 = " SELECT  MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO   FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T' AND  A.ACTIVE='T' WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "' ";
        //            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
        //            DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
        //            txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //        else
        //        {

        //            txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //    }

        //    if (Class.Users.HCompcode == "AGFM")
        //    {
        //        string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";

        //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
        //        DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
        //        Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
        //        if (cnt == 0)
        //        {
        //            string sel2 = "   MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO   FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T' AND  A.ACTIVE='T' WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "' ";
        //            DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
        //            DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
        //            txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //        else
        //        {

        //            txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString();
        //            txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
        //            return;
        //        }
        //    }
        //}
        private void combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {
            string id1 = "";
            string sel = "SELECT MAX(TO_NUMBER(A.ASPTBLFUELRATEMASID)+1) AS  ID FROM  ASPTBLFUELRATEMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE C.FINYR='" + Class.Users.Finyear + "' AND B.COMPCODE='" + Class.Users.HCompcode + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLFUELRATEMAS");
            DataTable dt = ds.Tables["ASPTBLFUELRATEMAS"];

           
            Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
            if (cnt == 0)
            {
                string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLFUELRATEMASID)+1) AS  ID FROM  ASPTBLFUELRATEMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE C.FINYR='" + Class.Users.Finyear + "'";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLFUELRATEMAS");
                DataTable dt1 = ds1.Tables["ASPTBLFUELRATEMAS"];
                id1 = Convert.ToString(dt1.Rows[0]["ID"].ToString());
                txtfueltoken.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + id1;
                txtfuelrateid1.Text = id1;
                return;
            }
            else
            {

                id1 = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString();
                txtfueltoken.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + id1;
                txtfuelrateid1.Text = id1;
                return;
            }

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNO"].Value = (e.RowIndex + 1).ToString();
        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            GridLoad();empty();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboitem();
        }

        private void buttsearch_Click(object sender, EventArgs e)
        {

            listView1.Items.Clear(); listfilter.Items.Clear();
            try
            {
                string sel = "SELECT A.ASPTBLFUELRATEMASID,A.ASPTBLFUELRATEID,B.COMPCODE,A.FUELDATE,C.BUNKNAME FROM ASPTBLFUELRATEMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID LEFT JOIN ASPTBLPETMAS C ON C.ASPTBLPETMASID=A.BUNKNAME WHERE B.COMPCODE='" + Class.Users.HCompcode+"' AND A.FUELDATE=to_date('" + dateTimePicker3.Value.ToShortDateString() + "','dd-MM-yyyy') order by ASPTBLFUELRATEMASID DESC";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLFUELRATEMAS");
                DataTable dt = ds.Tables["ASPTBLFUELRATEMAS"];
                int iGLCount = 1;
                if (dt.Rows.Count > 0)
                {
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(iGLCount.ToString());
                        list.SubItems.Add(dt.Rows[j]["ASPTBLFUELRATEMASID"].ToString());
                        list.SubItems.Add(dt.Rows[j]["ASPTBLFUELRATEID"].ToString());
                        list.SubItems.Add(dt.Rows[j]["COMPCODE"].ToString());
                        list.SubItems.Add(Convert.ToDateTime(dt.Rows[j]["FUELDATE"].ToString()).ToString("dd-MM-yyyy"));
                        list.SubItems.Add(dt.Rows[j]["BUNKNAME"].ToString());
                        listView1.Items.Add(list);
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        iGLCount++;
                    }
                }
                lblattcount.Text = "Total Count    :" + listView1.Items.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }


           
        }

        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {
                        dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                        if (txtfuelrateid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                string del1 = "DELETE   FROM  ASPTBLFUELRATEMASDET     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  ASPTBLFUELRATEMASDETID='" + griddelrow + "'";
                                Utility.ExecuteNonQuery(del1);

                                griddelrow = "";
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows[e.RowIndex].Cells[2].Value != null)
            {
                if (txtfuelrateid.Text != "")
                {
                    griddelrow = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                   

                }


            }
        }

        private void DownLoads_Click(object sender, EventArgs e)
        {

        }

        //private void Exit_Click(object sender, EventArgs e)
        //{
           

        //    this.Hide(); GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
        //}

        public void Deletes()
        {
            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirmation == DialogResult.Yes)
            {
                string del1 = "DELETE   FROM  ASPTBLFUELRATEMAS     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  ASPTBLFUELRATEMASID='" + txtfuelrateid.Text + "'";
                Utility.ExecuteNonQuery(del1);
                string del2 = "DELETE   FROM  ASPTBLFUELRATEMASDET     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  ASPTBLFUELRATEMASID='" + txtfuelrateid.Text + "'";
                Utility.ExecuteNonQuery(del2);

                MessageBox.Show("Record Deleted");
                GridLoad(); empty();
            }
        }

        private void Imports_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            string header = "";//rbHeaderYes.Checked ? "YES" : "NO";
            string conStr, sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }
        }

        

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            txtmonthname.Text = "";
            txtmonthname.Text = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dateTimePicker1.Value.Month) + "-" + System.DateTime.Now.Year;

        }

        public void News()
        {
            empty(); GridLoad();
        }

       

        public void Prints()
        {
           
        }

        public void Searchs()
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

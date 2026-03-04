using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
namespace Pinnacle.Fuel
{
    public partial class FuelToken : Form,ToolStripAccess
    {
        public FuelToken()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            
        }

        public void ReadOnlys()
        {

        }
        private static FuelToken _instance; bool litreenable = false; 
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        byte[] bytes; private int rowIndex = 0;
        string griddelrow = ""; ListView listfilter = new ListView();
        public static FuelToken Instance
        {
            get { if (_instance == null) _instance = new FuelToken(); 
                GlobalVariables.CurrentForm = _instance;              
                return _instance; }

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


                   
                        if (dt1.Rows[r]["SEARCH"].ToString() == "T") { GlobalVariables.Searchs.Visible = false; litreenable = true; } else { GlobalVariables.Searchs.Visible = false; litreenable = false; }
                      

                    }
                }


            }
            else
            {


                //this.Saves.Visible = false;
                //this.Prints.Visible = false;
                //this.Prints.Visible = false;
                //this.Searchs.Visible = false;
                //this.Deletes.Visible = false;
                //this.TreeButtons.Visible = false;
                //this.GlobalSearchs.Visible = false;
                //this.Logins.Visible = false;
                //this.ChangePasswords.Visible = false;
                //this.ChangeSkins.Visible = false;
                //this.DownLoads.Visible = false;
                //this.Pdfs.Visible = false;
                //this.Imports.Visible = false;
                // btnsavesIN.Visible = false;
            }


        }

        private void Label8_Click(object sender, EventArgs e)
        {

        }

        private void TextBox5_TextChanged(object sender, EventArgs e)
        {

        }
        public void autono()
        {
           
            if (Class.Users.HCompcode == "FLFD")
            {
                string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1+1)) AS  ID FROM ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID  JOIN gtfinancialyear C ON C.GTFINANCIALYEARID = A.FINYEAR  and C.CURRENTFINYR = 'T' JOIN ASPTBLINOUTMAS D ON D.COMPCODE=A.COMPCODE AND  D.COMPCODE=B.GTCOMPMASTID  AND  A.FINYEAR=C.GTFINANCIALYEARID AND D.FINYEAR=A.FINYEAR AND   D.ACTIVE='T' WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "' AND  a.TOKENDATE >=  substr(D.MODIFIED,0,10)";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
                DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];             
                Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                if (cnt == 0)
                {
                    string sel2 = "  SELECT  MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO    FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  AND  A.ACTIVE='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
                    DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
                    txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
                else
                {

                    txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
            }

            if (Class.Users.HCompcode == "FLF")
            {
                 string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";

                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
                DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                if (cnt == 0)
                {
                    string sel2 = "  SELECT  MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO   FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  AND  A.ACTIVE='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
                    DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
                    txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
                else
                {

                    txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
            }
            if (Class.Users.HCompcode == "AGF")
            {
                  string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
                DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                if (cnt == 0)
                {
                    string sel2 = " SELECT  MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO   FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T' AND  A.ACTIVE='T' WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "' ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
                    DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
                    txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
                else
                {

                    txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
            }

            if (Class.Users.HCompcode == "AGFM")
            {
                 string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";

                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
                DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                if (cnt == 0)
                {
                    string sel2 = "SELECT MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO   FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T' AND  A.ACTIVE='T' WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "' ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
                    DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
                    txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
                else
                {
                  
                    txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString();
                    txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text; 
                    return;
                }
            }
            if (Class.Users.HCompcode == "AGFMGII")
            {
                string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHTOKENID1)+1) AS  ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T'  WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";

                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
                DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                if (cnt == 0)
                {
                    string sel2 = "SELECT MAX(TO_NUMBER(A.FUELTOKEN)) AS  INWARDNO   FROM  ASPTBLINOUTMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T' AND  A.ACTIVE='T' WHERE A.FINYEAR='" + combofinyear.SelectedValue + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "' ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLINOUTMAS");
                    DataTable dt2 = ds2.Tables["ASPTBLINOUTMAS"];
                    txtfuelid1.Text = dt2.Rows[0]["INWARDNO"].ToString(); txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
                else
                {
                    txtfuelid1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString();
                    txttokenno.Text = combofinyear.Text + "/" + combocompcode.Text + "/" + txtfuelid1.Text;
                    return;
                }
            }
        }

       

        bool ch = false;int rowcount = 0;
        private bool Checks()
        {
            Models.Validate va = new Models.Validate();
            rowcount = 0;
            rowcount = dataGridView1.Rows.Count;
            if (combobunk.Text == "")
            {
                MessageBox.Show("Bunk Name is empty");
                this.Focus();
                this.combovechineno.Focus(); return false;
            }

          

            if (txtempname.Text == "")
            {
                MessageBox.Show("Name is empty");
                this.Focus();
                this.txtempname.Select(); return false;
            }
            if (Convert.ToInt64(combovechineno.SelectedValue) <= 0)
            {
                MessageBox.Show("Vechile Number is empty");
                this.Focus();
                this.combovechineno.Focus(); return false;
            }
            if(Convert.ToInt64("0"+combovechineno.SelectedValue) == 0)
            {
                MessageBox.Show("Vechile Number is empty");
                this.Focus();
                this.combovechineno.Focus(); return false;
            }
          
            if (combovechiletype.SelectedValue == null)
            {
                MessageBox.Show("Vehicle Type is empty");
                this.combovechiletype.Focus();
                return false;
            }
            //if (Convert.ToInt64("0" + txtkm.Text) < 1)
            //{
            //    MessageBox.Show("pls Check Kilo meter.Previous KM : " + txtprekm.Text);
            //    return false;
            //}
            Models.FuelTokenModelDet c1 = new Models.FuelTokenModelDet();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {

                // MessageBox.Show(row.Cells["ITEMNAME"].Value.ToString());

                c1.ITEMNAME = Convert.ToInt64(row.Cells["ITEMNAME"].Value.ToString());
                // c1.KM = Convert.ToInt64("0" + row.Cells["KM"].Value.ToString());
                c1.LITRES = Convert.ToString(row.Cells["LITRES"].Value.ToString());

                if (va.IsInteger(Convert.ToInt64("0" + c1.ITEMNAME).ToString()) == false)
                {

                    MessageBox.Show("ItemName Field is empty");

                    return false;
                }
                if (va.IsInteger(Convert.ToInt64("0" + c1.KM).ToString()) == false)
                {
                    MessageBox.Show("KiloMeter Field is empty");
                    return false;
                }
                if (c1.LITRES == "0")
                {
                    MessageBox.Show("Litres Field is empty");
                    return false;
                }
                if (va.IsStringNumberic1(Convert.ToString(c1.LITRES).ToString()) == false)
                {
                    MessageBox.Show("Litres Field is empty");
                    return false;
                }
                if (c1.LITRES == "")
                {
                    MessageBox.Show("Litres Field is empty");
                    return false;
                }
                
                return true;
            }
                return ch;
        }
       public void Saves()
        {
            try
            {
                Models.FuelTokenModel c = new Models.FuelTokenModel();

                if (Checks() == true)
                {
                    DataTable dt = new DataTable();
                   
                    if (txtfuelid.Text == "") { autono(); c.ASPTBLVEHTOKENID1 = Convert.ToInt64("0" + txtfuelid1.Text); txtfuelid.Text = ""; }
                    else { c.ASPTBLVEHTOKENID = Convert.ToInt64("0" + txtfuelid.Text); c.ASPTBLVEHTOKENID1 = Convert.ToInt64("0" + txtfuelid1.Text); }
                    c.FINYEAR = Convert.ToInt64(combofinyear.SelectedValue);
                    c.VCategory = Convert.ToString(txtvcategory.Text);
                    c.BUNKNAME = Convert.ToInt64("0" + combobunk.SelectedValue); 
                    c.COMPCODE = Convert.ToInt64(combocompcode.SelectedValue);
                    c.TOKENNO = txttokenno.Text;
                    c.TOKENDATE = txttokendate.Value;
                    c.VEHICLENO = Convert.ToInt64("0" + combovechineno.SelectedValue);
                    c.VEHICLETYPE = Convert.ToInt64("0" + combovechiletype.SelectedValue);
                    c.EMPNAME = Convert.ToInt64("0" + txtempname.SelectedValue.ToString());
                    c.EMPNAME1 = txtempname.Text;
                    if (checkactive.Checked) c.ACTIVE = "T"; else c.ACTIVE = "F";
                    if (checkcancel.Checked==true) 
                    { c.TOKENCANCEL = "T"; txtlastkm.Text = txtprekm.Text; }
                    else { c.TOKENCANCEL = "F";  }
                    c.TOTALLITRES = Convert.ToInt64("0" + txttotlitre.Text);
                    c.REMARKS = Convert.ToString(txtremarks.Text);
                    c.USERNAME = Class.Users.USERID;
                    c.MODIFIED = Class.Users.CREATED;
                    c.MODIFIED = Class.Users.CREATED;
                    c.CREATEDON = Class.Users.CREATED;
                    c.IPADDRESS = Class.Users.IPADDRESS;
                    c.LOILKM = Convert.ToInt64("0" + txtluboil.Text);                  
                    c.FUELTOKEN = Class.Users.Finyear + "/" + Class.Users.HCompcode + "/" + txtfuelid1.Text;
                    c.PREKM = Convert.ToInt64("0" + txtprekm.Text);
                    c.LASTKM = Convert.ToInt64("0" + txtlastkm.Text);
                    Int64 KM = Convert.ToInt64("0" + txtlastkm.Text) - Convert.ToInt64("0" + txtprekm.Text);
                    txtkm.Text = KM.ToString();
                    Int64 RKM = Convert.ToInt64("0" + txtloilrtotal.Text) + Convert.ToInt64("0" + txtkm.Text);
                    txtluboilremaining.Text = RKM.ToString();
                    c.LOILRKM = RKM;
                    c.TOTALKM = Convert.ToInt64("0" + KM.ToString());
                    if (Convert.ToInt64("0" + txtluboilremaining.Text) > Convert.ToInt64("0" + txtluboil.Text))
                    {
                        c.LOILRKM = 0;
                    }
                    DataTable dt1 = c.select(c.ASPTBLVEHTOKENID, c.ASPTBLVEHTOKENID1, c.FINYEAR,c.BUNKNAME, c.COMPCODE, c.TOKENNO, c.TOKENDATE, c.VEHICLENO, c.VEHICLETYPE, c.EMPNAME, c.EMPNAME1, c.ACTIVE, c.TOKENCANCEL, c.TOTALLITRES,c.REMARKS,c.PREKM,c.LASTKM,c.TOTALKM,c.LOILKM,c.LOILRKM);
                    if (dt1.Rows.Count != 0)
                    {

                    }
                    else if (dt1.Rows.Count != 0 && c.ASPTBLVEHTOKENID == 0 || c.ASPTBLVEHTOKENID == 0)
                    {
                        c = new Models.FuelTokenModel(c.ASPTBLVEHTOKENID1, c.FINYEAR, c.VCategory, c.BUNKNAME, c.COMPCODE, c.TOKENNO, c.TOKENDATE, c.VEHICLENO, c.VEHICLETYPE, c.EMPNAME, c.EMPNAME1, c.ACTIVE, c.TOKENCANCEL, c.TOTALLITRES, c.REMARKS, c.USERNAME, c.MODIFIED, c.CREATEDON, c.IPADDRESS, c.FUELTOKEN, c.PREKM, c.LASTKM, c.TOTALKM, c.LOILKM, c.LOILRKM);
                        string sel1 = "SELECT MAX(A.ASPTBLVEHTOKENID) ID FROM  ASPTBLVEHTOKEN A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE A.COMPCODE=" + Class.Users.COMPCODE;
                        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
                        dt = ds.Tables["ASPTBLVEHTOKEN"];
                    }
                    else
                    {

                        c = new Models.FuelTokenModel(c.ASPTBLVEHTOKENID1, c.FINYEAR, c.VCategory, c.BUNKNAME, c.COMPCODE, c.TOKENNO, c.TOKENDATE, c.VEHICLENO, c.VEHICLETYPE, c.EMPNAME, c.EMPNAME1, c.ACTIVE, c.TOKENCANCEL, c.TOTALLITRES, c.REMARKS, c.USERNAME, c.MODIFIED, c.IPADDRESS, c.FUELTOKEN, c.PREKM, c.LASTKM, c.TOTALKM,c.LOILKM,c.LOILRKM, c.ASPTBLVEHTOKENID);
                    }

                        Models.FuelTokenModelDet c1 = new Models.FuelTokenModelDet();
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                       
                        if (txtfuelid.Text == "") { c1.ASPTBLVEHTOKENID = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()); c1.ASPTBLVEHTOKENID1 = Convert.ToInt64("0" + txtfuelid1.Text); }
                        else { c1.ASPTBLVEHTOKENID = Convert.ToInt64("0" + txtfuelid.Text); c1.ASPTBLVEHTOKENID1 = Convert.ToInt64("0" + txtfuelid1.Text); }
                        c1.ASPTBLVEHTOKENDETID = Convert.ToInt64("0" + row.Cells["ASPTBLVEHTOKENDETID"].Value);
                        c1.ITEMNAME = Convert.ToInt64("0" + row.Cells["ITEMNAME"].Value);
                        c1.ITEMDESC = Convert.ToString(row.Cells["ITEMNAME"].FormattedValue);
                        c1.KM = Convert.ToInt64("0" + row.Cells["KM"].Value);
                        c1.LITRES = Convert.ToString(row.Cells["LITRES"].Value).ToUpper();
                        c1.NOTES = Convert.ToString(row.Cells["NOTES"].Value);
                        c1.COMPCODE = Class.Users.COMPCODE;
                        if (c1.ITEMNAME > 1)
                        {

                            DataTable dt2 = c1.select(c1.ASPTBLVEHTOKENDETID, c1.ASPTBLVEHTOKENID, c1.ITEMNAME, c1.KM, c1.LITRES, c1.NOTES, c1.COMPCODE, c.ASPTBLVEHTOKENID1);
                            if (dt2.Rows.Count != 0) { }
                            else if (dt2.Rows.Count != 0 && c1.ASPTBLVEHTOKENDETID == 0 || c1.ASPTBLVEHTOKENDETID == 0)
                            {
                                c1 = new Models.FuelTokenModelDet(c1.ASPTBLVEHTOKENID, c1.ITEMNAME, c1.ITEMDESC, c1.KM, c1.LITRES, c1.NOTES, c1.COMPCODE, c1.ASPTBLVEHTOKENID1);
                            }
                            else
                            {
                                c1 = new Models.FuelTokenModelDet(c1.ASPTBLVEHTOKENID, c1.ITEMNAME, c1.ITEMDESC, c1.KM, c1.LITRES, c1.NOTES, c1.COMPCODE, c1.ASPTBLVEHTOKENID1, c1.ASPTBLVEHTOKENDETID);
                            }
                        }
                    }

                    if (txtfuelid.Text == "")
                        {
                            MessageBox.Show("Record Saved Successfully." + txtfuelid1.Text, "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                        else
                        {
                            MessageBox.Show("Record Updated Successfully." + txtfuelid1.Text, "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            GridLoad(); empty();
                        }
                   
                }
                else
                {
                    this.dataGridView1.AllowUserToAddRows = true;
                   // MessageBox.Show("pls enter GridView Row.   ", "Informmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch(Exception ex)
            {
                string sel1 = "ROLLBACK";
                Utility.ExecuteNonQuery(sel1);
               // MessageBox.Show(ex.Message, "Informmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void GridLoad()
        {

            listView1.Items.Clear(); listfilter.Items.Clear(); int r = 1; dateTimePicker3.Text = System.DateTime.Now.ToString("dd-MM-yyyy");
            try
            {
                string sel = "SELECT A.ASPTBLVEHTOKENID,A.ASPTBLVEHTOKENID1,A.FINYEAR,C.COMPCODE,A.TOKENNO,A.TOKENDATE,(SELECT  AA.VEHICLENO    FROM HRVEHMAST AA  JOIN HRVEHTYPEMAST Ab on AA.VEHICLETYPE=AB.HRVEHTYPEMASTID     JOIN GTCOMPMAST AC ON AC.GTCOMPMASTID=AA.COMPCODE  AND AA.ACTIVE='T'    WHERE AA.HRVEHMASTID=A.VEHICLENO  UNION ALL    SELECT BA.VEHICLENO  FROM ASPTBLVEHMAS BA  JOIN HRVEHTYPEMAST BB ON BA.VEHICLETYPE=BB.HRVEHTYPEMASTID   JOIN GTCOMPMAST BC ON BC.GTCOMPMASTID=BA.COMPCODE    AND BA.ACTIVE='T'      WHERE BA.ASPTBLVEHMASID=A.VEHICLENO) AS VEHICLENO,  A.VEHICLETYPE,CONCAT(E.FNAME, concat('-', F.MIDCARD)) AS EMPNAME, A.EMPNAME1,A.ACTIVE,A.TOKENCANCEL  FROM ASPTBLVEHTOKEN A     JOIN     GTCOMPMAST C ON C.GTCOMPMASTID=A.COMPCODE JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME   JOIN  HREMPLOYDETAILS F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID AND  E.IDCARDNO = F.IDCARD WHERE C.COMPCODE='" + Class.Users.HCompcode + "' AND A.TOKENDATE = to_date('" + dateTimePicker3.Text + "', 'dd-MM-yyyy') order by ASPTBLVEHTOKENID DESC";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHTOKEN");
                DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                int iGLCount = 1;
                if (dt.Rows.Count > 0)
                {
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(iGLCount.ToString());
                        list.SubItems.Add(dt.Rows[j]["ASPTBLVEHTOKENID"].ToString());
                        list.SubItems.Add(dt.Rows[j]["ASPTBLVEHTOKENID1"].ToString());
                        list.SubItems.Add(dt.Rows[j]["TOKENNO"].ToString());
                        list.SubItems.Add(dt.Rows[j]["EMPNAME"].ToString());
                        list.SubItems.Add(dt.Rows[j]["VEHICLENO"].ToString());
                        if (dt.Rows[j]["ACTIVE"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");
                        if (dt.Rows[j]["TOKENCANCEL"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");
                        list.SubItems.Add(Convert.ToDateTime(dt.Rows[j]["TOKENDATE"].ToString()).ToString("dd-MM-yyyy"));
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        if (r % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        r++;
                        listView1.Items.Add(list);
                        iGLCount++;
                    }
                }
                else
                {
                   // MessageBox.Show("No Data Found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

        }

        public void Prints()
        {


            if (Convert.ToInt64("0" + txtfuelid.Text) >= 1)
            {
                Class.Users.Paramid = Convert.ToInt64(txtfuelid.Text);
                Report.FuelPrint p = new Report.FuelPrint();
                p.Show();
               
            }
            else
            {
                MessageBox.Show("Pls select one record from Listiview.   ","Informmation",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }

        private void empty()
        {
            txtfuelid.Text = ""; txtmidcard.Text = ""; txtempname.Text = "";
            txtfuelid1.Text = ""; txttokendate.Value= DateTime.Today.AddDays(0); 
            dateTimePicker3.Value = DateTime.Today.AddDays(0);
            txttokenno.Text = ""; listfilter.Items.Clear();
            combovechineno.Enabled = true; txtempname.Enabled = true; txtmidcard.Enabled = true;
            combovechineno.SelectedIndex = -1; combovechineno.Text = "";
            comboempname.Text = ""; txtloilrtotal.Text = "";
            combovechiletype.Text = "";
            comboempname.SelectedIndex = -1; txtluboilremaining.Text = "";
            checkactive.Checked = true;
            checkcancel.Checked = false;
            txttotlitre.Text = "";
            txtremarks.Text = "";
            pictureBox1.Image = null;txtkm.Text = "";txtlastkm.Text = "";txtprekm.Text = "";
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.MenuHighlight;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.EnableHeadersVisualStyles = false;
            do
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++) {try {   dataGridView1.Rows.RemoveAt(i);  }catch (Exception) { }}
            }
            while (dataGridView1.Rows.Count > 1);
            if (listView1.Items.Count > 1)
            {
                listView1.Items[0].Selected = true;
            }
            combovechineno.Select();
            if (Class.Users.HUserName == "VAIRAM")
            {
                txtprekm.Visible = true; lblprekm.Visible = true; txtkm.Visible = true; lbltotkm.Visible = true;
                lblluboil.Visible = true;txtluboil.Visible = true;txtloilrtotal.Visible = true;
            }
            else
            {
                txtprekm.Enabled = false; lblprekm.Enabled = false; txtkm.Visible = false; lbltotkm.Visible = false;
                lblluboil.Visible = false; txtluboil.Visible = false; txtloilrtotal.Visible = false;
            }


        }

        private void Butprint1_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {

                //printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                //printDocument1.Print();
             
              
            
            }
            else
            {
               
            }
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
        void bunkname()
        {

            DataTable dt = mas.bunkname(Class.Users.HCompcode);
            if (dt.Rows.Count == 1) { combobunk.Enabled = false; } else { combobunk.Enabled = true; }
            combobunk.ValueMember = "ASPTBLPETMASID";
            combobunk.DisplayMember = "BUNKNAME";
            combobunk.DataSource = dt;
        }
        void COMPCODE()
        {

            DataTable dt = mas.compcodefind(Class.Users.HCompcode);
          
            combocompcode.DisplayMember = "COMPCODE";
            combocompcode.ValueMember = "GTCOMPMASTID";
            combocompcode.DataSource = dt;
            //txtbunkname.Text = dt.Rows[0]["ASPTBLPETMASID"].ToString();
        }
        void EMPNAME()
        {


            string selemp = "select DISTINCT a.hremploymastid as asptblempid,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as empname from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid  join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname";
            DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
            DataTable dt = dsemp.Tables["hremploymast"];
            txtempname.DisplayMember = "EMPNAME";
            txtempname.ValueMember = "asptblempid";
            txtempname.DataSource = dt;
            //comboempname.DisplayMember = "EMPNAME";
            //comboempname.ValueMember = "ASPTBLEMPID";
            //comboempname.DataSource = dt;
            //txtempname.DisplayMember = "EMPNAME";
            //txtempname.ValueMember = "asptblempid";
            //txtempname.DataSource = dt;
        }
        void EMPNAME(Int64 eid)
        {

            string selemp = "select DISTINCT a.hremploymastid as asptblempid,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as empname from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid  join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname where a.hremploymastid='"+eid+"' ";
            DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
            DataTable dt = dsemp.Tables["hremploymast"];
            comboempname.DisplayMember = "EMPNAME";
            comboempname.ValueMember = "ASPTBLEMPID";
            comboempname.DataSource = dt;
            txtempname.DisplayMember = "EMPNAME";
            txtempname.ValueMember = "ASPTBLEMPID";
            txtempname.DataSource = dt;
        }
        void vechineno()
        {
              string sel1 = "SELECT DISTINCT X.HRVEHMASTID,X.VEHICLENO FROM(SELECT A.HRVEHMASTID,A.VEHICLENO  FROM HRVEHMAST A  JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID where A.ACTIVE='T' UNION ALL SELECT AA.ASPTBLVEHMASID AS HRVEHMASTID,AA.VEHICLENO  FROM ASPTBLVEHMAS AA  JOIN GTCOMPMAST AB ON AA.COMPCODE=AB.GTCOMPMASTID AND AA.ACTIVE='T' AND AB.COMPCODE='" + Class.Users.HCompcode + "' UNION ALL SELECT AC.ASPTBLVEHMASID AS HRVEHMASTID,AC.VEHICLENO  FROM ASPTBLVEHMAS AC  JOIN GTCOMPMAST AD ON AC.COMPCODE=AD.GTCOMPMASTID AND AC.ACTIVE1='T' )X  ORDER BY 2";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRVEHMAST");
            DataTable dt = ds.Tables["HRVEHMAST"];
            combovechineno.ValueMember = "HRVEHMASTID";
            combovechineno.DisplayMember = "VEHICLENO";
            combovechineno.DataSource = dt;

        }
        void vechineno(string veh)
        {
            try
            {
                // string sel = "SELECT DISTINCT X.HRVEHTYPEMASTID,X.VEHTYPE,X.VCATEGORY,X.VEHICLENO,X.PARTYNAME FROM(SELECT AB.HRVEHTYPEMASTID,AB.VEHTYPE,AA.VCATEGORY,AA.VEHICLENO,'' AS PARTYNAME FROM HRVEHMAST AA  JOIN HRVEHTYPEMAST AB on AA.VEHICLETYPE=AB.HRVEHTYPEMASTID   where AA.ACTIVE='T'   AND AA.VCATEGORY='COMPANY'    UNION ALL     SELECT  BB.HRVEHTYPEMASTID,BB.VEHTYPE,BA.VCATEGORY,BA.VEHICLENO,BA.PARTYNAME FROM ASPTBLVEHMAS BA    JOIN HRVEHTYPEMAST BB ON BA.VEHICLETYPE=BB.HRVEHTYPEMASTID     AND BA.ACTIVE='T' )X WHERE X.VEHICLENO='" + veh + "'";
                string sel = "SELECT DISTINCT X.HRVEHTYPEMASTID,X.VEHTYPE,X.VCATEGORY,X.VEHICLENO,X.HREMPLOYMASTID ,X.FNAME,X.MIDCARD FROM(SELECT  DISTINCT AB.HRVEHTYPEMASTID,AB.VEHTYPE,AA.VCATEGORY,AA.VEHICLENO,0 AS HREMPLOYMASTID ,'' AS FNAME,'' AS MIDCARD FROM HRVEHMAST AA  JOIN HRVEHTYPEMAST AB on AA.VEHICLETYPE=AB.HRVEHTYPEMASTID   where AA.ACTIVE='T'   AND AA.VCATEGORY='COMPANY'     UNION ALL     SELECT DISTINCT  BB.HRVEHTYPEMASTID,BB.VEHTYPE,BA.VCATEGORY,BA.VEHICLENO,BC.HREMPLOYMASTID ,BC.FNAME,BD.MIDCARD FROM ASPTBLVEHMAS BA      JOIN HRVEHTYPEMAST BB ON BA.VEHICLETYPE=BB.HRVEHTYPEMASTID     AND BA.ACTIVE='T'   JOIN  hremploymast BC ON  BA.PARTYNAME=BC.HREMPLOYMASTID  JOIN hremploydetails BD ON BD.HREMPLOYMASTID=BC.HREMPLOYMASTID   AND BD.IDACTIVE='YES' )X    WHERE X.VEHICLENO='" + veh + "'";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "HRVEHTYPEMAST");
                DataTable dt = ds.Tables["HRVEHTYPEMAST"];
                combovechiletype.ValueMember = "HRVEHTYPEMASTID";
                combovechiletype.DisplayMember = "VEHTYPE";
                combovechiletype.DataSource = dt;
                txtvcategory.Text = dt.Rows[0]["VCATEGORY"].ToString();
                if (dt.Rows[0]["FNAME"].ToString() != "")
                {

                    //comboempname.DisplayMember = "FNAME";
                    //comboempname.ValueMember = "HREMPLOYMASTID";
                    //comboempname.DataSource = dt;
                    comboempname.Text = dt.Rows[0]["FNAME"].ToString();
                    txtmidcard.Text = dt.Rows[0]["MIDCARD"].ToString();
                    
                }
                else
                {
                    comboempname.DataSource = null;
                    txtmidcard.Text = "";
                    comboempname.Text = "";
                    comboempname.SelectedIndex = -1;
                }
                if (veh.Length >= 6 && txtfuelid.Text == "")
                {
                    string selemp1 = "SELECT X.VEHICLENO, X.LASTKM,X.PREKM,X.LOILRKM FROM(  SELECT   MAX( AC.ASPTBLVEHTOKENID) ASASPTBLVEHTOKENID ,AA.VEHICLENO,  AC.LASTKM,AC.PREKM,AC.LOILRKM  FROM HRVEHMAST AA    JOIN ASPTBLVEHTOKEN AC ON   AC.VEHICLENO=AA.HRVEHMASTID  where AA.ACTIVE='T'   AND AA.VCATEGORY='COMPANY'  AND AA.VEHICLENO='" + veh + "'     AND  AC.ASPTBLVEHTOKENID=(SELECT   MAX(A2.ASPTBLVEHTOKENID) AS MAXID  FROM HRVEHMAST A1    JOIN ASPTBLVEHTOKEN A2 ON  A2.VEHICLENO=A1.HRVEHMASTID          where A1.ACTIVE='T'   AND A1.VCATEGORY='COMPANY'  AND A1.VEHICLENO='" + veh + "') GROUP BY AA.VEHICLENO, AC.LASTKM,AC.PREKM ,AC.LOILRKM     UNION ALL    SELECT MAX(BD.ASPTBLVEHTOKENID) AS ASPTBLVEHTOKENID,    BA.VEHICLENO,BD.LASTKM,BD.PREKM,BD.LOILRKM   FROM ASPTBLVEHMAS BA    JOIN ASPTBLVEHTOKEN BD ON  BD.VEHICLENO=BA.ASPTBLVEHMASID   WHERE  BA.VEHICLENO='" + veh + "'  AND BD.ASPTBLVEHTOKENID=(SELECT   MAX(A2.ASPTBLVEHTOKENID) AS MAXID  FROM ASPTBLVEHMAS A1    JOIN ASPTBLVEHTOKEN A2 ON  A2.VEHICLENO=A1.ASPTBLVEHMASID    where A1.ACTIVE='T'    AND A1.VEHICLENO='" + veh + "') GROUP BY BA.VEHICLENO,BD.LASTKM  ,BD.PREKM,BD.LOILRKM        )X    WHERE X.VEHICLENO='" + veh + "'";
                    DataSet dsemp1 = Utility.ExecuteSelectQuery(selemp1, "ASASPTBLVEHTOKEN");
                    DataTable dtemp1 = dsemp1.Tables["ASASPTBLVEHTOKEN"];
                    if (dtemp1 != null || dtemp1.Rows.Count >= 0)
                    {

                        txtprekm.Text = dtemp1.Rows[0]["LASTKM"].ToString();
                        txtloilrtotal.Text = dtemp1.Rows[0]["LOILRKM"].ToString();
                    }
                    
                }
            }
            catch(Exception ex)
            {
               // MessageBox.Show(ex.Message) ;
            }
        }
      
        private void FuelToken_Load(object sender, EventArgs e)
        {
            
            finyear(); 
            vechineno();
            COMPCODE(); EMPNAME(); bunkname();
            comboitem(); 
            GridLoad();
         
            combovechineno.Select(); empty();
           // usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);

        }





        public void News()
        {
            finyear();
            vechineno();
            COMPCODE(); EMPNAME(); bunkname();
            comboitem(); empty();
            GridLoad();
            combovechineno.Select();

        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void Combocompcode_SelectedIndexChanged(object sender, EventArgs e)
        {

           
        }

        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                txtfuelid.Text = ""; txtmidcard.Text = "";
                txttokendate.Text = "";

                txtfuelid.Text = Convert.ToString(listView1.SelectedItems[0].SubItems[2].Text);
                string sel0 = "SELECT A.ASPTBLVEHTOKENID,A.ASPTBLVEHTOKENID1,G.FINYR AS FINYEAR,C.COMPCODE,A.TOKENNO,A.TOKENDATE, H.BUNKNAME,A.LASTKM,A.PREKM,A.TOTALKM,LOILRKM,( SELECT  AA.VEHICLENO    FROM HRVEHMAST AA  JOIN HRVEHTYPEMAST Ab on AA.VEHICLETYPE=AB.HRVEHTYPEMASTID       JOIN GTCOMPMAST AC ON AC.GTCOMPMASTID=AA.COMPCODE   JOIN ASPTBLVEHTOKEN AD  ON AD.VEHICLENO=AA.HRVEHMASTID  AND AA.ACTIVE='T'  WHERE  AD.ASPTBLVEHTOKENID='" + txtfuelid.Text + "'   UNION ALL    SELECT BA.VEHICLENO  FROM ASPTBLVEHMAS BA JOIN HRVEHTYPEMAST BB ON BA.VEHICLETYPE=BB.HRVEHTYPEMASTID   JOIN GTCOMPMAST BC ON BC.GTCOMPMASTID=BA.COMPCODE  JOIN ASPTBLVEHTOKEN BD  ON BD.VEHICLENO=BA.ASPTBLVEHMASID   AND BA.ACTIVE='T'    WHERE BD.ASPTBLVEHTOKENID='" + txtfuelid.Text + "'    )AS VEHICLENO, B.VEHTYPE AS VEHICLETYPE,A.EMPNAME,A.EMPNAME1,    A.ACTIVE,A.TOKENCANCEL,A.REMARKS    FROM ASPTBLVEHTOKEN A JOIN  HRVEHTYPEMAST B ON B.HRVEHTYPEMASTID=A.VEHICLETYPE  JOIN     GTCOMPMAST C ON C.GTCOMPMASTID=A.COMPCODE JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME  JOIN  HREMPLOYDETAILS F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID AND  E.IDCARDNO = F.IDCARD JOIN GTFINANCIALYEAR  G ON G.GTFINANCIALYEARID=A.FINYEAR JOIN ASPTBLPETMAS  H ON H.COMPCODE=C.GTCOMPMASTID AND H.COMPCODE=A.COMPCODE AND H.ASPTBLPETMASID=A.BUNKNAME WHERE A.ASPTBLVEHTOKENID='" + txtfuelid.Text + "' ";
                DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLVEHTOKEN");
                DataTable dt = ds0.Tables["ASPTBLVEHTOKEN"];
                if (dt.Rows.Count > 0)
                {
                    txtfuelid.Text = Convert.ToString(dt.Rows[0]["ASPTBLVEHTOKENID"].ToString());
                    txtfuelid1.Text = Convert.ToString(dt.Rows[0]["ASPTBLVEHTOKENID1"].ToString());
                    combofinyear.Text = Convert.ToString(dt.Rows[0]["FINYEAR"].ToString());
                    combocompcode.Text = Convert.ToString(dt.Rows[0]["COMPCODE"].ToString());
                    txttokenno.Text = Convert.ToString(dt.Rows[0]["TOKENNO"].ToString());
                    txttokendate.Text = dt.Rows[0]["TOKENDATE"].ToString();
                    combobunk.Text = Convert.ToString(dt.Rows[0]["BUNKNAME"].ToString());
                    combovechineno.Text = Convert.ToString(dt.Rows[0]["VEHICLENO"].ToString());
                    combovechiletype.Text = Convert.ToString(dt.Rows[0]["VEHICLETYPE"].ToString());
                    txtprekm.Text = dt.Rows[0]["LASTKM"].ToString();
                   // vechineno(Convert.ToString(dt.Rows[0]["VEHICLENO"].ToString()));
                    EMPNAME(Convert.ToInt64(dt.Rows[0]["EMPNAME"].ToString()));

                    // combovechineno.Enabled = false; comboempname.Enabled = false; txtempname.Enabled = false; txtmidcard.Enabled = false;
                    if (dt.Rows[0]["ACTIVE"].ToString() == "T") checkactive.Checked = true; else checkactive.Checked = false; ;
                    if (dt.Rows[0]["TOKENCANCEL"].ToString() == "T")
                        checkcancel.Checked = true;

                    else checkcancel.Checked = false;
                    txtremarks.Text = Convert.ToString(dt.Rows[0]["REMARKS"].ToString());
                    //txtlastkm.Text = Convert.ToString(dt.Rows[0]["LASTKM"].ToString());
                    txtprekm.Text = Convert.ToString(dt.Rows[0]["LASTKM"].ToString());
                    txtloilrtotal.Text = Convert.ToString(dt.Rows[0]["LOILRKM"].ToString());
                    QRCoder.QRCodeGenerator qc = new QRCoder.QRCodeGenerator();
                    var mydata1 = qc.CreateQrCode(txttokenno.Text, QRCoder.QRCodeGenerator.ECCLevel.L);
                    var code1 = new QRCoder.QRCode(mydata1);
                    pictureBox1.Image = code1.GetGraphic(50, Color.Black, Color.White, true);

                    string sel1 = "SELECT B.ASPTBLVEHTOKENDETID,A.ASPTBLVEHTOKENID,A.ASPTBLVEHTOKENID1,B.COMPCODE, C.GTGENITEMMASTID AS ITEMNAME,B.KM,B.LITRES,B.NOTES FROM ASPTBLVEHTOKEN A JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID   JOIN GTGENITEMMAST C ON C.GTGENITEMMASTID = B.ITEMNAME WHERE A.ASPTBLVEHTOKENID='" + txtfuelid.Text + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKENDET");
                    DataTable dt1 = ds1.Tables["ASPTBLVEHTOKENDET"];
                    if (dt1.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dt1;
                        int totlitres = 0;
                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            if (Convert.ToInt64(dataGridView1.Rows[i].Cells[1].Value) > 0)
                            {
                                dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt1.Rows[i]["ASPTBLVEHTOKENDETID"].ToString());
                                dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt1.Rows[i]["ASPTBLVEHTOKENID"].ToString());
                                dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt1.Rows[i]["ASPTBLVEHTOKENID1"].ToString());
                                dataGridView1.Rows[i].Cells[4].Value = Convert.ToInt64("0" + dt1.Rows[i]["CompCode"].ToString());
                                dataGridView1.Rows[i].Cells[5].Value = Convert.ToInt64("0" + dt1.Rows[i]["ITEMNAME"].ToString());
                                dataGridView1.Rows[i].Cells[6].Value = Convert.ToInt64("0" + dt1.Rows[i]["KM"].ToString());
                                if (dt1.Rows[i]["LITRES"].ToString() == "FULL")
                                {
                                    dataGridView1.Rows[i].Cells[7].ReadOnly = false;
                                    dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(dt1.Rows[i]["LITRES"].ToString());
                                }
                                else
                                {

                                    if (litreenable == true || Class.Users.IPADDRESS=="192.168.101.15")
                                    {
                                        dataGridView1.Rows[i].Cells[7].ReadOnly = false;
                                        dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(dt1.Rows[i]["LITRES"].ToString());
                                    }
                                    if (litreenable == false)
                                    {
                                        dataGridView1.Rows[i].Cells[7].ReadOnly = true;
                                        dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(dt1.Rows[i]["LITRES"].ToString());
                                    }

                                }
                                dataGridView1.Rows[i].Cells[8].Value = dt1.Rows[i]["NOTES"].ToString();


                            }
                        }

                    }


                }

              
            }
            catch (Exception EX)
            {

            }
            txtmidcard.Select();
        }

        public void Searchs()
        {
            GridLoad();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            //string sel2 = " SELECT A.ASPTBLVEHTOKENID,D.EMPNAME ,E.VEHICLENO,A.TOKENNO,F.COMPNAME,B.ITEMDESC,B.LITRES  FROM ASPTBLVEHTOKEN A  JOIN ASPTBLVEHTOKENDET B  ON  A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID JOIN ASPTBLEMP D ON D.ASPTBLEMPID=A.EMPNAME  JOIN ASPTBLVEHMAST E ON E.ASPTBLVEHMASTID = A.VEHICLENO JOIN GTCOMPMAST F ON F.GTCOMPMASTID = A.COMPCODE JOIN asptblusermas G ON G.userid = A.USERNAME WHERE  A.ASPTBLVEHTOKENID =" + textBox1.Text
            //DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLVEHTOKEN");
            //Report.CrystalReport5 rd = new Report.CrystalReport5();
            ////  rd.SetParameterValue("ASPTBLVEHTOKENID", txtfuelid.Text);
            //rd.SetDataSource(ds2.Tables[0]);
            //crystalReportViewer1.ReportSource = rd;//F.COMPCODE ='" + Class.Users.HCompcode + "' AND G.USERNAME ='" + Class.Users.HUserName+"' AND


        }

        private void CrystalReportViewer1_Load(object sender, EventArgs e)
        {

        }

        public void Deletes()
       {
            try
            {
                if (txtfuelid.Text != "")
                {
                    //string sel = "SELECT A.ASPTBLVEHTOKENID FROM ASPTBLVEHTOKEN A JOIN GTCOMPMAST b ON b.GTCOMPMASTID=A.COMPCODE  JOIN ASPTBLPETMAS c ON C.COMPCODE=A.COMPCODE and C.COMPCODE=B.GTCOMPMASTID WHERE A.ASPTBLVEHTOKENID=" + txtfuelid.Text;
                    //DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHTOKEN");
                    //DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                    //if (dt.Rows.Count > 0)
                    //{
                    //    MessageBox.Show("Child Record Found.Can not Delete.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //}
                    //else
                    //{
                    var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (confirmation == DialogResult.Yes)
                    {
                        string del = "DELETE  FROM ASPTBLVEHTOKEN   Where COMPCODE='" + Class.Users.COMPCODE + "' and  ASPTBLVEHTOKENID='" + txtfuelid.Text + "'";
                        Utility.ExecuteNonQuery(del);
                        string del1 = "DELETE   FROM  ASPTBLVEHTOKENDET     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  ASPTBLVEHTOKENID='" + txtfuelid.Text + "'";
                        Utility.ExecuteNonQuery(del1);
                        MessageBox.Show("Record Deleted Successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                }
                else
                {
                    MessageBox.Show("Invalid", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Question);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Txttokensearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; listView1.Items.Clear(); int item1 = 1;int r = 1;
                if (txttokensearch.Text.Length >= 1 && listfilter.Items.Count>1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[4].ToString().Contains(txttokensearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[5].ToString().Contains(txttokensearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[6].ToString().Contains(txttokensearch.Text.ToUpper()) || listfilter.Items[item0].SubItems[8].ToString().Contains(txttokensearch.Text.ToUpper()))
                        {


                            list.Text = item1.ToString(); //listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);

                            if (r % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            r++;
                            //list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
                            // list.SubItems.Add(listfilter.Items[item0].SubItems[10].Text);
                            listView1.Items.Add(list);


                        }
                        item0++; item1++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {
                   ListView ll = new ListView(); item0 = 1;
                    listView1.Items.Clear(); 
                    foreach (ListViewItem item in listfilter.Items)
                    {

                        if (item0 % 2 == 0)
                        {
                            item.BackColor = Color.White;
                        }
                        else
                        {
                            item.BackColor = Color.WhiteSmoke;
                        }
                        this.listView1.Items.Add((ListViewItem)item.Clone());
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

        private void DataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }
        Color hdrColor = SystemColors.Control;

        
        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            vechineno(); GridLoad();
        }

        private void Pdfs_Click(object sender, EventArgs e)
        {

        }

        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
            //{
            //    if (oneCell.Selected)

            //        // MessageBox.Show(oneCell.RowIndex);
            //        //
            //        dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
            //    // string sel1 = "DELETE  FROM ASPTBLVEHMASDET A  WHERE A.ASPTBLVEHMASDETID=" + dataGridView1.Rows[;
            //    // Utility.ExecuteNonQuery(sel1);
            //}

            
        }

     
        private void refreshToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            comboitem();
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            //try
            //{
            //    if (e.Button == MouseButtons.Right)
            //    {
            //        this.dataGridView1.Rows[e.RowIndex].Selected = true;
            //        this.rowIndex = e.RowIndex;
            //        this.dataGridView1.CurrentCell = this.dataGridView1.Rows[e.RowIndex].Cells[1];
                   
            //        //string sel1 = "DELETE  FROM ASPTBLVEHTOKENDET A  WHERE A.ASPTBLVEHTOKENDETID=" + this.dataGridView1.CurrentCell.Value.ToString();
            //        //Utility.ExecuteNonQuery(sel1); this.dataGridView1.ClearSelection();
            //    }
               
            //}
            //catch (Exception ex) { }
        }

      
        private void addRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.Add();
            }
            catch (Exception ex) { }
        }

        private void refreshGridToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtfuelid.Text != "")
                {
                    string sel1 = "SELECT B.ASPTBLVEHTOKENDETID,B.ASPTBLVEHTOKENID,B.ASPTBLVEHTOKENID1,B.COMPCODE, C.GTGENITEMMASTID AS ITEMNAME,B.KM,B.LITRES,B.NOTES FROM ASPTBLVEHTOKEN A JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID1=B.ASPTBLVEHTOKENID   JOIN GTGENITEMMAST C ON C.GTGENITEMMASTID = B.ITEMNAME WHERE A.ASPTBLVEHTOKENID1='" + txtfuelid.Text + "' AND A.COMPCODE='" + Class.Users.COMPCODE + "'";
                    DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKENDET");
                    DataTable dt1 = ds1.Tables["ASPTBLVEHTOKENDET"];
                    dataGridView1.DataSource = dt1;
                    
                }
                FuelToken_Load(sender,e); usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            }
            catch (Exception EX) { }
        }

        private void deleteRowToolStripgridfuelMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                foreach (DataGridViewCell oneCell in dataGridView1.SelectedCells)
                {
                    if (oneCell.Selected)
                    {
                        dataGridView1.Rows.RemoveAt(oneCell.RowIndex);
                        if (txtfuelid.Text != "")
                        {
                            var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (confirmation == DialogResult.Yes)
                            {
                                string del1 = "DELETE   FROM  ASPTBLVEHTOKENDET     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  ASPTBLVEHTOKENDETID='" + griddelrow + "'";
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

        private void refreshToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            combovechineno.Refresh(); bunkname(); 
            vechineno();  combovechineno.SelectedIndex = -1; combovechineno.Text = "";
            comboempname.Text = "";
            combovechiletype.Text = ""; 
        }

        private void comboempname_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                comboempname.BackColor = Color.White;
          
            }
        }

        private void contextMenuStrip2_Opening(object sender, CancelEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null)
                {
                    if (txtfuelid.Text != "")
                    {
                        griddelrow = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();

                    }


                }
            }catch(Exception E) { }
            
        }

        private void buttsearch_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear(); listfilter.Items.Clear(); int r = 1; dateTimePicker3.Value.ToShortDateString();
            try
            {
                string sel = "SELECT A.ASPTBLVEHTOKENID,A.ASPTBLVEHTOKENID1,A.FINYEAR,C.COMPCODE,A.TOKENNO,A.TOKENDATE,(SELECT  AA.VEHICLENO    FROM HRVEHMAST AA  JOIN HRVEHTYPEMAST Ab on AA.VEHICLETYPE=AB.HRVEHTYPEMASTID     JOIN GTCOMPMAST AC ON AC.GTCOMPMASTID=AA.COMPCODE  AND AA.ACTIVE='T'    WHERE AA.HRVEHMASTID=A.VEHICLENO  UNION ALL    SELECT BA.VEHICLENO  FROM ASPTBLVEHMAS BA  JOIN HRVEHTYPEMAST BB ON BA.VEHICLETYPE=BB.HRVEHTYPEMASTID   JOIN GTCOMPMAST BC ON BC.GTCOMPMASTID=BA.COMPCODE    AND BA.ACTIVE='T'      WHERE BA.ASPTBLVEHMASID=A.VEHICLENO) AS VEHICLENO,  A.VEHICLETYPE,CONCAT(E.FNAME, concat('-', F.MIDCARD)) AS EMPNAME, A.EMPNAME1,A.ACTIVE,A.TOKENCANCEL  FROM ASPTBLVEHTOKEN A     JOIN     GTCOMPMAST C ON C.GTCOMPMASTID=A.COMPCODE JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME   JOIN  HREMPLOYDETAILS F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID AND  E.IDCARDNO = F.IDCARD WHERE C.COMPCODE='" + Class.Users.HCompcode + "' AND A.TOKENDATE = to_date('" + dateTimePicker3.Text + "', 'dd-MM-yyyy') order by ASPTBLVEHTOKENID DESC";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHTOKEN");
                DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                int iGLCount = 1;
                if (dt.Rows.Count > 0)
                {
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(iGLCount.ToString());
                        list.SubItems.Add(dt.Rows[j]["ASPTBLVEHTOKENID"].ToString());
                        list.SubItems.Add(dt.Rows[j]["ASPTBLVEHTOKENID1"].ToString());
                        list.SubItems.Add(dt.Rows[j]["TOKENNO"].ToString());
                        list.SubItems.Add(dt.Rows[j]["EMPNAME"].ToString());
                        list.SubItems.Add(dt.Rows[j]["VEHICLENO"].ToString());
                        if (dt.Rows[j]["ACTIVE"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");
                        if (dt.Rows[j]["TOKENCANCEL"].ToString() == "T") list.SubItems.Add("T"); else list.SubItems.Add("F");
                        list.SubItems.Add(Convert.ToDateTime(dt.Rows[j]["TOKENDATE"].ToString()).ToString("dd-MM-yyyy"));
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        if (r % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        r++;
                        listView1.Items.Add(list);
                        iGLCount++;
                    }
                }
                else
                {
                    // MessageBox.Show("No Data Found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                lbltotal.Text = "Total Count    :" + listView1.Items.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void menuRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void txtmidcard_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //if (txtmidcard.Text == "")
                //{
                //    EMPNAME(); //comboempname.SelectedIndex = -1; comboempname.Text = "";
                //}
                if (txtfuelid.Text != "" && txtmidcard.Text == "")
                {
                    string sel0 = "SELECT A.ASPTBLVEHTOKENID,A.ASPTBLVEHTOKENID1,G.FINYR AS FINYEAR,C.COMPCODE,A.TOKENNO,A.TOKENDATE, H.BUNKNAME, (     SELECT  AA.VEHICLENO    FROM HRVEHMAST AA  JOIN HRVEHTYPEMAST Ab on AA.VEHICLETYPE=AB.HRVEHTYPEMASTID       JOIN GTCOMPMAST AC ON AC.GTCOMPMASTID=AA.COMPCODE   JOIN ASPTBLVEHTOKEN AD  ON AD.VEHICLENO=AA.HRVEHMASTID  AND AA.ACTIVE='T'  WHERE  AD.ASPTBLVEHTOKENID='" + txtfuelid.Text + "'   UNION ALL    SELECT BA.VEHICLENO  FROM ASPTBLVEHMAS BA JOIN HRVEHTYPEMAST BB ON BA.VEHICLETYPE=BB.HRVEHTYPEMASTID   JOIN GTCOMPMAST BC ON BC.GTCOMPMASTID=BA.COMPCODE  JOIN ASPTBLVEHTOKEN BD  ON BD.VEHICLENO=BA.ASPTBLVEHMASID   AND BA.ACTIVE='T'    WHERE BD.ASPTBLVEHTOKENID='" + txtfuelid.Text + "'    )AS VEHICLENO, B.VEHTYPE AS VEHICLETYPE,E.HREMPLOYMASTID AS asptblempid, CONCAT(E.FNAME, concat('-', F.MIDCARD)) AS EMPNAME,    A.ACTIVE,A.TOKENCANCEL,A.REMARKS    FROM ASPTBLVEHTOKEN A JOIN  HRVEHTYPEMAST B ON B.HRVEHTYPEMASTID=A.VEHICLETYPE  JOIN     GTCOMPMAST C ON C.GTCOMPMASTID=A.COMPCODE JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.EMPNAME  JOIN  HREMPLOYDETAILS F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID AND  E.IDCARDNO = F.IDCARD and F.IDACTIVE='YES' JOIN GTFINANCIALYEAR  G ON G.GTFINANCIALYEARID=A.FINYEAR JOIN ASPTBLPETMAS  H ON H.COMPCODE=C.GTCOMPMASTID AND H.COMPCODE=A.COMPCODE AND H.ASPTBLPETMASID=A.BUNKNAME WHERE A.ASPTBLVEHTOKENID='" + txtfuelid.Text + "'  ";
                    DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "ASPTBLVEHTOKEN");
                    DataTable dt = ds0.Tables["ASPTBLVEHTOKEN"];

                    //if (comboempname.Text == "")
                    //{
                    txtempname.DisplayMember = "EMPNAME";
                    txtempname.ValueMember = "asptblempid";
                    txtempname.DataSource = dt;
                }

                // txtempname.Text = dt.Rows[0]["EMPNAME"].ToString();

                if (txtfuelid.Text != "" && txtmidcard.Text != "")
                {
                    string selemp = "select a.hremploymastid as asptblempid,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as empname from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid AND B.IDACTIVE='YES' join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname WHERE B.MIDCARD='" + txtmidcard.Text + "'";
                    DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
                    DataTable dtemp = dsemp.Tables["hremploymast"];
                    //if (comboempname.Text == "")
                    //{
                    txtempname.DisplayMember = "EMPNAME";
                    txtempname.ValueMember = "asptblempid";
                    txtempname.DataSource = dtemp;
                    //}
                    //   txtempname.Text = dtemp.Rows[0]["EMPNAME"].ToString();
                }

                if (txtfuelid.Text == "" && txtmidcard.Text != "")
                {
                    string selemp = "select a.hremploymastid as asptblempid,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as empname from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid AND B.IDACTIVE='YES'  join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname WHERE B.MIDCARD='" + txtmidcard.Text + "'";
                    DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
                    DataTable dtemp = dsemp.Tables["hremploymast"];
                    //if (comboempname.Text == "")
                    //{
                    txtempname.DisplayMember = "EMPNAME";
                    txtempname.ValueMember = "asptblempid";
                    txtempname.DataSource = dtemp;
                    //}
                    // txtempname.Text = dtemp.Rows[0]["EMPNAME"].ToString();
                }
              
            }
            catch(Exception ex) { }
        }

        private void combovechineno_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combovechineno.Text != "System.Data.DataRowView")
            {

                vechineno(combovechineno.Text);

                label17.Refresh();
               
            }
        }

        private void txtmidcard_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void txtempname_KeyPress(object sender, KeyPressEventArgs e)
        {
           // e.Handled = !(char.IsLetter(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void combobunk_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboempname_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name == "LITRES")
            {

                e.Control.KeyPress += Control_KeyPress;
            }

            if (dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name == "KM")
            {

                e.Control.KeyPress += Control_KeyPress1;
            }
           
        }

        private void Control_KeyPress1(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == ',' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == '.' || e.KeyChar == 'F' || e.KeyChar == 'U' || e.KeyChar == 'L' || e.KeyChar == 'f' || e.KeyChar == 'u' || e.KeyChar == 'l' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }
        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            EMPNAME();
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
          //  if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { this.TreeButtons.Visible = false; txttokendate.Enabled = true; } else { this.TreeButtons.Visible = false; txttokendate.Enabled = false; }

        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show(); GlobalVariables.HeaderName.Text = "";
            this.Hide();
        }

        
        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtprekm_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtlastkm_TextChanged(object sender, EventArgs e)
        {
            lblInvalid.Text = "";

            if (Convert.ToInt64("0" + txtlastkm.Text) > Convert.ToInt64("0" + txtprekm.Text))
            {
                Int64 KM = Convert.ToInt64("0" + txtlastkm.Text) - Convert.ToInt64("0" + txtprekm.Text);

                
                
                    txtkm.Text = KM.ToString();
                    Int64 RKM = Convert.ToInt64("0" + txtloilrtotal.Text) + Convert.ToInt64("0" + txtkm.Text);
                    txtluboilremaining.Text = RKM.ToString();
                    Int64 C1 = Convert.ToInt64("0" + txtluboil.Text) - Convert.ToInt64("0" + txtluboilremaining.Text);
                label17.Refresh();
                    label17.Text = "Lubricant Oil: '" + C1.ToString() + "' KM Remaining...";
                
            }
 
            if (Convert.ToInt64("0" + txtluboilremaining.Text) >= Convert.ToInt64("0" + txtluboil.Text))
            {

                txtluboilremaining.Text = "0";
                label17.Refresh();
                label17.Text = "Lubricant Oil: KM Remaining...";
            }

            //else
            //{
            //    Int64 KM = Convert.ToInt64("0" + txtluboilremaining.Text) - Convert.ToInt64("0" + txtloilrtotal.Text);
            //    txtluboilremaining.Text = KM.ToString();
            //    lblInvalid.Text = "KM - Invalid"; lblInvalid.ForeColor = Color.Red;

            //}

        }

        private void txtlastkm_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= '0' && e.KeyChar <= '9' || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
            {
                e.Handled = false; //Do not reject the input
            }
            else
            {
                e.Handled = true; //Reject the input
            }

        }

        private void txtprekm_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }


        //private void Control_KeyPress(object sender, KeyPressEventArgs e)
        //{ 
        //   //e.Handled = !(char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back);

        //    //if (e.KeyChar >= '0' && e.KeyChar <= '9' && !(char.IsLetter(e.KeyChar)= 'FULL') || e.KeyChar == (char)Keys.Back) //The  character represents a backspace
        //    //{
        //    //    e.Handled = false; //Do not reject the input
        //    //}
        //    //else
        //    //{
        //    //    e.Handled = true; //Reject the input
        //    //}
        //}
    }
}

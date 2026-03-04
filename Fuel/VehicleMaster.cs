using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Pinnacle.Fuel
{
    public partial class VehicleMaster : Form,ToolStripAccess
    {
        private static VehicleMaster _instance;
        Models.Master mas = new Models.Master();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        byte[] bytes; int i = 1;
        public VehicleMaster()
        {
            InitializeComponent();
           
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            tabControl1.TabPages.Remove(tabPage6);           
            tabControl1.TabPages.Remove(tabPage1);
            usercheck(Class.Users.HCompcode, Class.Users.HUserName, Class.Users.ScreenName);
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; 
        }
        public void ReadOnlys()
        {

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
                        if (dt1.Rows[r]["READONLY"].ToString() == "T") { } else { this.contextMenuStrip1.Enabled = true; }
                        if (dt1.Rows[r]["TREEBUTTON"].ToString() == "T") { CheckActive1.Visible = true; } else { CheckActive1.Visible = false; }


                    }
                }
              
            }
            else
            {
                
                GlobalVariables.Toolstrip1.Enabled = false;
            }

        }

        public static VehicleMaster Instance
        {
            get { if (_instance == null) _instance = new VehicleMaster(); GlobalVariables.CurrentForm = _instance; return _instance; }
        }
   
        void EMPNAME()
        {

            DataTable dt = mas.EmpName();

            comboempname.DisplayMember = "EMPNAME";
            comboempname.ValueMember = "ASPTBLEMPID";
            comboempname.DataSource = dt;

        }
        void empty()
        {

            int i = 0;
            do
            {
                for (i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (dataGridView1.Rows.Count > 1);
            Party1();
            FuelType(); VehType(); combofinearLoad(); CompCode2();EMPNAME(); CompCodeLoad1();
            txtHRVehicleID.Text = "";
            dateTimeDOCDATE.CustomFormat = " "; txtvehname.Text = "";
            txtVehicleNo.Text = "";
            comboempname.Text = ""; 
            comboVehicleType.SelectedIndex = -1;
            txtCapacity.Text = "";
            txtMRFName.Text = "";
            dateTimeMFGDate.CustomFormat = "";
            comboNewUsed.SelectedIndex = -1;
            comboFuelType.SelectedIndex = -1;
            txtTankCapacity.Text = "";
            comboVehicleUsage.SelectedIndex = -1;
            txtEngineNo.Text = "";
            dateTimeEngineDate.CustomFormat = " ";
            txtchassisno.Text = "";
            dateTimeRoadTaxPermitDate.CustomFormat = "";
            txtPartyName.Text = "";
            txtAddress.Text = "";
            txtBrokerName.Text = "";
            comboMajorService.SelectedIndex = -1;
            txtNoofSeats.Text = "";
            txtNoofTyres.Text = "";
            txtNoofStepney.Text = "";
            txtNoofAlex.Text = "";
            txtMileageLitre.Text = "";
            txtVehicleIncharge.Text = "";
            txtRunningKM.Text = "";
            txtRunningPerKM.Text = "";
            txtTonPerKM.Text = "";
            txtInsuranceNo.Text = "";
            txtInsuranceCompany.Text = "";
            dateTimeInsuranceFrom.CustomFormat = "";
            dateTimeInsuranceTo.CustomFormat = "";
            txtInsuranceAmount.Text = "";
            txtInsuranceAmount1.Text = "";
            txtPolicyName.Text = "";
            txtAgentName.Text = "";
            txtReminderBeforeDays.Text = "";
            txtResponsibility.Text = "";
            txtLastServiceKM.Text = "";
            dateTimeLastServiceDate.CustomFormat = "";
            txtLastServiceDesc.Text = "";
            txtNextServiceKM.Text = "";
            dateTimeNextServiceDate.CustomFormat = "";
            txtsearch.Text = "";
            txtHRVehicleID1.Text = "";
            CheckActive1.Checked = false;txtmidcard.Text = "";
            panelgrid.Visible = true; txtVehicleNo.Focus(); tabControl1.SelectTab(tabPage3); txtsearch.Select();
           
        }

        private void VehicleMaster_Load(object sender, EventArgs e)
        {
            News(); GridLoad(); EMPNAME();  txtsearch.Select();
        }
        public void autono()
        {
            if (txtHRVehicleID.Text == "")
            {
                string sel1 = "SELECT MAX(TO_NUMBER(A.ASPTBLVEHMASID1)+1) AS ID FROM  ASPTBLVEHMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE A.COMPCODE=" + Class.Users.COMPCODE;
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHMAS");
                DataTable dt = ds.Tables["ASPTBLVEHMAS"];
                Int64 cnt = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString());
                if (cnt == 0)
                {
                    txtHRVehicleID1.Text = "1";
                }
                else
                {
                    txtHRVehicleID1.Text = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()).ToString();
                }
            }
        }
        void FuelType()
        {
            try
            {
                string sel1 = "SELECT  A.GTGENITEMMASTID,A.ITEMNAME FROM GTGENITEMMAST A WHERE A.ACTIVE='T' AND A.FT='T'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "GTGENITEMMAST");
                DataTable dt = ds.Tables["GTGENITEMMAST"];
                comboFuelType.ValueMember = "GTGENITEMMASTID";
                comboFuelType.DisplayMember = "ITEMNAME";
                comboFuelType.DataSource = dt;
            }
            catch (Exception EX) { }
        }
        void VehType()
        {
            try
            {
                string sel1 = " SELECT a.HRVEHTYPEMASTID,a.VEHTYPE  FROM HRVEHTYPEMAST a where A.ACTIVE='T' and A.SHORTCODE='TW' OR A.SHORTCODE='PVC'";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRVEHTYPEMAST");
                DataTable dt = ds.Tables["HRVEHTYPEMAST"];
                comboVehicleType.ValueMember = "HRVEHTYPEMASTID";
                comboVehicleType.DisplayMember = "VEHTYPE";
                comboVehicleType.DataSource = dt;
            }
            catch (Exception EX) { }
        }
        void Party1()
        {
            try
            {
                DataTable dt = mas.party();
                CompService.ValueMember = "GTCOMPMASTID";
                CompService.DisplayMember = "COMPNAME";
                CompService.DataSource = dt;
            }
            catch (Exception EX) { }

        }

        void CompCodeLoad1()
        {
            try
            {
                DataTable dt = mas.findcomcode(Class.Users.HCompcode);

                combocompcode.ValueMember = "GTCOMPMASTID";
                combocompcode.DisplayMember = "COMPCODE";
                combocompcode.DataSource = dt;

            }
            catch (Exception EX) { }

        }

        void CompCode2()
        {
            //try
            //{
            //    DataTable dt = mas.party();
            //    combo.ValueMember = "GTCOMPMASTID";
            //    CompService.DisplayMember = "COMPCODE";
            //    CompService.DataSource = dt;


            //}
            //catch (Exception EX) { }

        }
        void combofinearLoad()
        {
            try
            {
                DataTable dt = mas.finyear();
                combofinear.ValueMember = "gtfinancialyearid";
                combofinear.DisplayMember = "finyear";
                combofinear.DataSource = dt;
                Class.Users.Finyear = dt.Rows[0]["gtfinancialyear"].ToString();
            }
            catch (Exception EX) { }
        }
        //void Vehicletype()
        //{
        //    string sel1 = " SELECT a.HRVEHTYPEMASTID,a.VEHTYPE  FROM HRVEHTYPEMAST a where A.ACTIVE='T' and A.SHORTCODE='TW' OR A.SHORTCODE='PVC'";
        //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "HRVEHTYPEMAST");
        //    DataTable dt = ds.Tables["HRVEHTYPEMAST"];
        //    comboVehicleType.ValueMember = "HRVEHTYPEMASTID";
        //    comboVehicleType.DisplayMember = "VEHTYPE";
        //    comboVehicleType.DataSource = dt;
            
        //}
        bool ch = false;
        private bool Checks()
        {
            try
            {

                Models.Validate val = new Models.Validate();

                if (txtVehicleNo.Text == "")
                {
                    MessageBox.Show("Vechile Number is empty." + txtVehicleNo.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.txtVehicleNo.Focus(); return false;
                }
                if (comboempname.Text == "" || comboempname.Text == null)
                {
                    MessageBox.Show("Employee Name is empty." + txtVehicleNo.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.comboempname.Focus(); return false;
                }

                //if (val.IsStringNumberic(txtVehicleNo.Text) == false)
                //{
                //    MessageBox.Show("'VechileNo  is Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    comboempname.Select();
                //    return false;
                //}
                if (comboVehicleType.Text == "" && comboVehicleType.SelectedValue == null)
                {

                    MessageBox.Show("Vehicle Type is empty." + txtVehicleNo.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.comboVehicleType.Focus();
                    return false;
                }
                if (comboFuelType.Text == "")
                {
                  
                    MessageBox.Show("Fuel Type is empty." + txtVehicleNo.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.comboFuelType.Focus();
                    return false;
                }
                if (combofinear.Text == "")
                {
                   
                    MessageBox.Show("FinYear Type is empty." + txtVehicleNo.Text, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }


                //foreach (DataGridViewRow row in dataGridView1.Rows)
                //{
                //    try
                //    {
                //        Models.HRVehicleDet c1 = new Models.HRVehicleDet();
                //        c1.CompService = Convert.ToInt64("0" + row.Cells["CompService"].Value);
                //        c1.Days = Convert.ToInt64("0" + row.Cells["Days"].Value);
                //        c1.KM = Convert.ToInt64("0" + row.Cells["KM"].Value);
                //        if (c1.CompService < 1)
                //        {

                //            MessageBox.Show("CompService is empty");
                //            return false;
                //        }

                //        if (c1.KM < 1)
                //        {
                //            MessageBox.Show("KM is empty");
                //            return false;
                //        }

                //        if (c1.Days.ToString() == "")
                //        {
                //            MessageBox.Show("Days is empty");
                //            return false;
                //        }

                //        return true;
                //    }
                //    catch (Exception EX)
                //    {
                //        MessageBox.Show("pls Enter Correct formate" + EX.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error); return false;
                //    }
                //}
                return true;
            }
            catch (Exception EX)
            {
                MessageBox.Show("Error" + EX.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            return ch;

        }
        public void Saves()
        {
          
            if (Checks() == true)
            {

                try
                {
                    autono();
                    Models.HRVehicle v1 = new Models.HRVehicle();
                    DataTable dt = new DataTable();
                    string words = txtVehicleNo.Text;
                    words = Regex.Replace(words, @" ", "");
                    txtVehicleNo.Text = words;
                    var chk= "";var chk1 = "";


                    if (CheckActive.Checked == true)  chk = "T"; else chk = "F";
                    if (CheckActive1.Checked == true) chk1 = "T"; else chk1 = "F";

                    if (txtHRVehicleID.Text == "") { v1.HRVehicleID1 = Convert.ToInt64("0" + txtHRVehicleID1.Text); }
                    else { v1.HRVehicleID = Convert.ToInt64("0" + txtHRVehicleID.Text); v1.HRVehicleID1 = Convert.ToInt64("0" + txtHRVehicleID1.Text); }
                    v1.FinYear = Convert.ToInt64(combofinear.SelectedValue.ToString());
                    v1.Active = chk;
                    v1.Active1 = chk1;
                    v1.VCategory = comboVCategory.Text.ToString();
                    v1.DocDate = dateTimeDOCDATE.Value.ToString();
                    v1.VehicleNo = txtVehicleNo.Text.ToUpper();
                    v1.VehicleName = txtvehname.Text.ToUpper();
                    v1.VehicleType = Convert.ToInt64("0" + comboVehicleType.SelectedValue);
                    v1.Capacity = Convert.ToInt64("0" + txtCapacity.Text);
                    v1.MRFName = txtMRFName.Text.ToUpper();
                    v1.MFGDate = dateTimeMFGDate.Value.ToString();
                    v1.NewUsed = Convert.ToString(comboNewUsed.SelectedValue);
                    v1.FuelType = Convert.ToInt64("0" + comboFuelType.SelectedValue);
                    v1.TankCapacity = Convert.ToInt64("0" + txtTankCapacity.Text);
                    v1.VehicleUsage = Convert.ToInt64("0" + comboVehicleUsage.SelectedValue);
                    v1.EngineNo = txtEngineNo.Text.ToUpper();
                    v1.EngineDate = dateTimeEngineDate.Value.ToString();
                    v1.RoadTaxPermitDate = dateTimeRoadTaxPermitDate.Value.ToString();
                    v1.PartyName = comboempname.SelectedValue.ToString();
                    v1.Address = txtAddress.Text.ToUpper();
                    v1.BrokerName = txtBrokerName.Text.ToUpper();
                    v1.MajorService = Convert.ToString(comboMajorService.SelectedValue);
                    v1.NoofSeats = Convert.ToInt64("0" + txtNoofSeats.Text);
                    v1.NoofTyres = Convert.ToInt64("0" + txtNoofTyres.Text);
                    v1.NoofStepney = Convert.ToInt64("0" + txtNoofStepney.Text);
                    v1.NoofAlex = Convert.ToInt64("0" + txtNoofAlex.Text);
                    v1.MileageLitre = Convert.ToInt64("0" + txtMileageLitre.Text);
                    v1.VehicleIncharge = txtVehicleIncharge.Text.ToUpper();
                    v1.RunningKM = Convert.ToInt64("0" + txtRunningKM.Text);
                    v1.RunningPerKM = Convert.ToInt64("0" + txtRunningPerKM.Text);
                    v1.TonPerKM = Convert.ToInt64("0" + txtTonPerKM.Text);
                    v1.InsuranceNo = txtInsuranceNo.Text.ToUpper();
                    v1.InsuranceCompany = txtInsuranceCompany.Text.ToUpper();
                    v1.InsuranceFrom = dateTimeInsuranceFrom.Value.ToString();
                    v1.InsuranceTo = dateTimeInsuranceTo.Value.ToString();
                    v1.InsuranceAmount = Convert.ToInt64("0" + txtInsuranceAmount.Text);
                    v1.InsuranceAmount1 = Convert.ToInt64("0" + txtInsuranceAmount1.Text);
                    v1.PremiumAmount = Convert.ToInt64("0" + txtInsuranceAmount1.Text);
                    v1.PolicyName = txtPolicyName.Text.ToUpper();
                    v1.AgentName = txtAgentName.Text.ToUpper();
                    v1.ReminderBeforeDays = txtReminderBeforeDays.Text.ToUpper();
                    v1.Responsibility = txtResponsibility.Text.ToUpper();
                    v1.LastServiceKM = Convert.ToInt64("0" + txtLastServiceKM.Text);
                    v1.LastServiceDate = dateTimeLastServiceDate.Value.ToString();
                    v1.LastServiceDesc = txtLastServiceDesc.Text.ToUpper();
                    v1.NextServiceKM = Convert.ToInt64("0" + txtNextServiceKM.Text);
                    v1.NextServiceDate = dateTimeNextServiceDate.Value.ToString();
                    v1.HRVehicleID = Convert.ToInt64("0" + txtHRVehicleID.Text);
                    v1.CompCode = Class.Users.COMPCODE;
                    v1.UserName = Class.Users.USERID;
                    Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
                    Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
                    v1.IPAddress = Class.Users.IPADDRESS;
                    v1.Createdon = Convert.ToString(Class.Users.CREATED);
                    v1.ModifiedOn = Convert.ToString(Class.Users.CREATED);
                   // DataTable dt1 = v1.select(v1.HRVehicleID1, v1.VCategory, v1.Active, v1.Active1, v1.DocDate, v1.VehicleNo, v1.VehicleName, v1.VehicleType, v1.Capacity, v1.MRFName, v1.MFGDate, v1.NewUsed, v1.FuelType, v1.TankCapacity, v1.VehicleUsage, v1.EngineNo, v1.EngineDate, v1.ChassisNo, v1.RoadTaxPermitDate, v1.PartyName, v1.Address, v1.BrokerName, v1.MajorService, v1.NoofSeats, v1.NoofTyres, v1.NoofStepney, v1.NoofAlex, v1.MileageLitre, v1.VehicleIncharge, v1.RunningKM, v1.RunningPerKM, v1.TonPerKM, v1.InsuranceNo, v1.InsuranceCompany, v1.InsuranceFrom, v1.InsuranceTo, v1.InsuranceAmount, v1.InsuranceAmount1, v1.PremiumAmount, v1.PolicyName, v1.AgentName, v1.ReminderBeforeDays, v1.Responsibility, v1.LastServiceKM, v1.LastServiceDate, v1.LastServiceDesc, v1.NextServiceKM, v1.NextServiceDate, v1.HRVehicleID, v1.FinYear, v1.CompCode);

                    DataTable dt1 = v1.select(v1.HRVehicleID1, v1.VCategory,v1.Active, v1.Active1, v1.PartyName, v1.VehicleNo, v1.VehicleType,  v1.FuelType,v1.CompCode);
                  
                    if (dt1.Rows.Count != 0)
                    {
                        MessageBox.Show("Child Record Found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        empty();
                        return;
                    }
                    else if (dt1.Rows.Count != 0 && v1.HRVehicleID == 0 || v1.HRVehicleID == 0)
                    {

                        v1 = new Models.HRVehicle(v1.HRVehicleID, v1.HRVehicleID1,v1.Active, v1.Active1, v1.VCategory, v1.DocDate, v1.VehicleNo, v1.VehicleName, v1.VehicleType, v1.Capacity, v1.MRFName, v1.MFGDate, v1.NewUsed, v1.FuelType, v1.TankCapacity, v1.VehicleUsage, v1.EngineNo, v1.EngineDate, v1.ChassisNo, v1.RoadTaxPermitDate, v1.PartyName, v1.Address, v1.BrokerName, v1.MajorService, v1.NoofSeats, v1.NoofTyres, v1.NoofStepney, v1.NoofAlex, v1.MileageLitre, v1.VehicleIncharge, v1.RunningKM, v1.RunningPerKM, v1.TonPerKM, v1.InsuranceNo, v1.InsuranceCompany, v1.InsuranceFrom, v1.InsuranceTo, v1.InsuranceAmount, v1.InsuranceAmount1, v1.PremiumAmount, v1.PolicyName, v1.AgentName, v1.ReminderBeforeDays, v1.Responsibility, v1.LastServiceKM, v1.LastServiceDate, v1.LastServiceDesc, v1.NextServiceKM, v1.NextServiceDate, v1.CompCode, v1.UserName, v1.IPAddress, v1.Createdon, v1.ModifiedOn, v1.FinYear);
                        // dt = v1.maxid();
                        MessageBox.Show("Record Saved Successfully." + txtHRVehicleID1.Text, "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();

                    }
                    else
                    {
                        v1 = new Models.HRVehicle(v1.HRVehicleID1, v1.VCategory,v1.Active, v1.Active1, v1.DocDate, v1.VehicleNo, v1.VehicleName, v1.VehicleType, v1.Capacity, v1.MRFName, v1.MFGDate, v1.NewUsed, v1.FuelType, v1.TankCapacity, v1.VehicleUsage, v1.EngineNo, v1.EngineDate, v1.ChassisNo, v1.RoadTaxPermitDate, v1.PartyName, v1.Address, v1.BrokerName, v1.MajorService, v1.NoofSeats, v1.NoofTyres, v1.NoofStepney, v1.NoofAlex, v1.MileageLitre, v1.VehicleIncharge, v1.RunningKM, v1.RunningPerKM, v1.TonPerKM, v1.InsuranceNo, v1.InsuranceCompany, v1.InsuranceFrom, v1.InsuranceTo, v1.InsuranceAmount, v1.InsuranceAmount1, v1.PremiumAmount, v1.PolicyName, v1.AgentName, v1.ReminderBeforeDays, v1.Responsibility, v1.LastServiceKM, v1.LastServiceDate, v1.LastServiceDesc, v1.NextServiceKM, v1.NextServiceDate, v1.CompCode, v1.UserName, v1.IPAddress, v1.Createdon, v1.ModifiedOn, v1.FinYear, v1.HRVehicleID);
                        MessageBox.Show("Record Updated Successfully." + txtHRVehicleID1.Text, "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();
                    }
                  //  Models.HRVehicleDet v2 = new Models.HRVehicleDet();

                    //foreach (DataGridViewRow row in dataGridView1.Rows)
                    //{
                    //    if (txtHRVehicleID.Text == "") { v2.HRVehicleID = Convert.ToInt64("0" + dt.Rows[0]["ID"].ToString()); v2.HRVehicleID1 = Convert.ToInt64("0" + txtHRVehicleID1.Text); }
                    //    else { v2.HRVehicleID = Convert.ToInt64("0" + txtHRVehicleID.Text); v2.HRVehicleID1 = Convert.ToInt64("0" + txtHRVehicleID1.Text); }
                    //    v2.HRVehicleDetID = Convert.ToInt64("0" + row.Cells["ASPTBLVEHMASDETID"].Value);
                    //    v2.CompCode = Convert.ToInt64("0" + Class.Users.COMPCODE);
                    //    v2.CompService = Convert.ToInt64("0" + row.Cells["CompService"].Value);
                    //    v2.Days = Convert.ToInt64("0" + row.Cells["Days"].Value);
                    //    v2.KM = Convert.ToInt64("0" + row.Cells["KM"].Value);
                    //    v2.ServiceDate = Convert.ToString(row.Cells["ServiceDate"].Value);
                    //    v2.Notes = Convert.ToString(row.Cells["NOTES"].Value);
                    //    v2.HRVEHMASTDETROW = Convert.ToInt64("0" + row.Cells["SNo"].Value);
                    //    if (v2.CompService > 1 && v2.KM > 1 && v2.Days > 1)
                    //    {
                    //        DataTable dt2 = v2.select(v2.HRVehicleDetID, v2.HRVehicleID, v2.HRVehicleID1, v2.CompCode, v2.CompService, v2.KM, v2.Days, v2.ServiceDate, v2.Notes);
                    //        if (dt2.Rows.Count != 0)
                    //        {
                    //            MessageBox.Show("Child Record Found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //        }
                    //        else if (dt2.Rows.Count != 0 && v2.HRVehicleDetID == 0 || v2.HRVehicleDetID == 0)
                    //        {
                    //            v2 = new Models.HRVehicleDet(v2.HRVehicleID, v2.HRVehicleID1, v2.CompCode, v2.CompService, v2.KM, v2.Days, v2.ServiceDate, v2.Notes, v2.HRVEHMASTDETROW);
                    //        }
                    //        else
                    //        {
                    //            v2 = new Models.HRVehicleDet(v2.HRVehicleDetID, v2.HRVehicleID, v2.HRVehicleID1, v2.CompCode, v2.CompService, v2.KM, v2.Days, v2.ServiceDate, v2.Notes, v2.HRVEHMASTDETROW);
                    //        }
                    //    }
                    //}

                    //if (txtHRVehicleID.Text == "")
                    //{
                    //    MessageBox.Show("Record Saved Successfully." + txtHRVehicleID1.Text, "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //    GridLoad(); empty();
                    //}
                    //else
                    //{
                    //    MessageBox.Show("Record Updated Successfully." + txtHRVehicleID1.Text, "Update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //    GridLoad(); empty();
                    //}
                }
                catch (Exception EX)
                {
                   // MessageBox.Show("ERROR." + EX.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                this.dataGridView1.AllowUserToAddRows = true;
                MessageBox.Show("Pls Enter Mandatory Fields ", "Informmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

       public void GridLoad()
        {
            try
            {
                listviewveh.Items.Clear(); listfilter.Items.Clear();

                string sel1 = "SELECT A.ASPTBLVEHMASID,A.DOCDATE,A.VEHICLENO ,A.VEHICLENAME,C.VEHTYPE,CONCAT(E.fname ,concat('-',F.MIDCARD) ) as PARTYNAME,B.COMPCODE,A.ACTIVE,A.ACTIVE1 FROM ASPTBLVEHMAS A join gtcompmast b on A.COMPCODE = B.GTCOMPMASTID JOIN HRVEHTYPEMAST C ON C.HRVEHTYPEMASTID = A.VEHICLETYPE JOIN GTGENITEMMAST D ON D.GTGENITEMMASTID = A.FUELTYPE  JOIN  hremploymast E  ON E.HREMPLOYMASTID=A.PARTYNAME join hremploydetails F on E.hremploymastid = F.hremploymastid  AND E.IDCARDNO = F.IDCARD WHERE  B.COMPCODE = '" + Class.Users.HCompcode + "' ORDER BY A.ASPTBLVEHMASID DESC";//  
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHMAS");
                DataTable dt = ds.Tables["ASPTBLVEHMAS"];

                if (dt.Rows.Count > 0)
                {
                    i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["ASPTBLVEHMASID"].ToString());
                        list.SubItems.Add(myRow["DOCDATE"].ToString());
                        list.SubItems.Add(myRow["VEHICLENO"].ToString());
                        list.SubItems.Add(myRow["VEHICLENAME"].ToString()); 
                        list.SubItems.Add(myRow["VEHTYPE"].ToString());
                        list.SubItems.Add(myRow["PARTYNAME"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["ACTIVE"].ToString());
                        list.SubItems.Add(myRow["ACTIVE1"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        listviewveh.Items.Add(list);
                        i++;
                    }

                    lblcount.Text = "Total Rows Count:  " + listviewveh.Items.Count.ToString();
                }
                else
                {

                    lblcount.Text = "Total Rows Count:  " + listviewveh.Items.Count.ToString();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("GridLoad" + ex.Message);
            }

        }

        private void buttsearch_Click(object sender, EventArgs e)
        {
            listfilter.Items.Clear(); listviewveh.Items.Clear();
            GridLoad(); txtsearch.Select();
            //try
            //{
            //    listviewveh.Items.Clear();
            //    listfilter.Items.Clear();
            //    string sel1 = "SELECT A.ASPTBLVEHMASID,A.DOCDATE,A.VEHICLENO ,A.VEHICLENAME,C.VEHTYPE,CONCAT(E.fname ,concat('-',F.MIDCARD) ) as PARTYNAME,B.COMPCODE,A.ACTIVE,A.ACTIVE1 FROM ASPTBLVEHMAS A join gtcompmast b on A.COMPCODE = B.GTCOMPMASTID JOIN HRVEHTYPEMAST C ON C.HRVEHTYPEMASTID = A.VEHICLETYPE JOIN GTGENITEMMAST D ON D.GTGENITEMMASTID = A.FUELTYPE   JOIN  hremploymast E  ON E.HREMPLOYMASTID=A.PARTYNAME join hremploydetails F on E.hremploymastid = F.hremploymastid  AND E.IDCARDNO = F.IDCARD WHERE  B.COMPCODE = '" + Class.Users.HCompcode + "' ORDER BY A.ASPTBLVEHMASID DESC";

            //    //  string sel1 = "SELECT A.ASPTBLVEHMASID,A.DOCDATE,A.VEHICLENO ,A.VEHICLENAME,C.VEHTYPE,A.PARTYNAME,A.COMPCODE FROM ASPTBLVEHMAS a join gtcompmast b on A.COMPCODE = B.GTCOMPMASTID JOIN HRVEHTYPEMAST C ON C.HRVEHTYPEMASTID = A.VEHICLETYPE JOIN GTGENITEMMAST D ON D.GTGENITEMMASTID = A.FUELTYPE  WHERE  B.COMPCODE = '" + Class.Users.HCompcode + "'  order by A.ASPTBLVEHMASID desc";//OR A.VEHICLENO LIKE'%" + txtsearch.Text + "%' OR A.VEHICLENAME LIKE'%" + txtsearch.Text + "%' OR A.PARTYNAME LIKE'%" + txtsearch.Text + "%'
            //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHMAS");
            //    DataTable dt = ds.Tables["ASPTBLVEHMAS"];

            //    if (dt.Rows.Count > 0)
            //    {

            //        i = 1;
            //        foreach (DataRow myRow in dt.Rows)
            //        {
            //            ListViewItem list = new ListViewItem();
            //            list.SubItems.Add(i.ToString());

            //            list.SubItems.Add(myRow["ASPTBLVEHMASID"].ToString());
            //            list.SubItems.Add(myRow["DOCDATE"].ToString());
            //            list.SubItems.Add(myRow["VEHICLENO"].ToString());
            //            list.SubItems.Add(myRow["VEHICLENAME"].ToString());
            //            list.SubItems.Add(myRow["VEHTYPE"].ToString());
            //            list.SubItems.Add(myRow["PARTYNAME"].ToString());
            //            list.SubItems.Add(myRow["COMPCODE"].ToString());
            //            list.SubItems.Add(myRow["ACTIVE"].ToString());
            //            list.SubItems.Add(myRow["ACTIVE1"].ToString());
            //            this.listfilter.Items.Add((ListViewItem)list.Clone());
            //            listviewveh.Items.Add(list);
            //            i++;
            //        }
            //        txtsearch.Select();

            //    }
            //    else
            //    {

            //        lblcount.Text = "Total Rows Count:  " + listviewveh.Items.Count.ToString();
            //    }


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("GridLoad" + ex.Message);
            //}
        }


        public void Searchs()
        {
            tabControl1.TabPages.Remove(tabPage3);
            tabControl1.TabPages.Add(tabPage3);
           
            tabControl1.SelectTab(tabPage3); txtsearch.Select();

        }


        private void Listviewveh_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listviewveh.Items.Count > 0)
                {
                    Models.HRVehicle v1 = new Models.HRVehicle();
                    v1.HRVehicleID = Convert.ToInt32(listviewveh.SelectedItems[0].SubItems[2].Text);
                    string sel1 = "SELECT A.ASPTBLVEHMASID,A.ASPTBLVEHMASID1,A.ACTIVE, A.ACTIVE1,E.FINYR,A.VCATEGORY, A.DOCDATE,A.VEHICLENO ,A.VEHICLENAME,C.VEHTYPE," +
                        "A.CAPACITY ,A.MRFNAME ,A.MFGDATE,A.NEWUSED,D.ITEMNAME,A.TANKCAPACITY," +
                        "A.VEHICLEUSAGE,A.ENGINENO,A.ENGINEDATE, A.ROADTAXPERMITDATE ,f.HREMPLOYMASTID,CONCAT(F.fname ,concat('-',G.MIDCARD) ) as PARTYNAME,A.ADDRESS ,A.BROKERNAME,A.MAJORSERVICE ,A.NOOFSEATS,A.NOOFTYRES ," +
                        "A.NOOFSTEPNEY,A.NOOFALEX ,A.MILEAGELITRE,A.VEHICLEINCHARGE,A.RUNNINGKM, A.RUNNINGPERKM ,A.TONPERKM ,A.INSURANCENO ,A.INSURANCECOMPANY ," +
                        "A.INSURANCEFROM ,A.INSURANCETO ,A.INSURANCEAMOUNT ,A.INSURANCEAMOUNT1 , A.POLICYNAME ,A.AGENTNAME , A.REMINDERBEFOREDAYS,A.RESPONSIBILITY ,A.LASTSERVICEKM, A.LASTSERVICEDATE,  A.LASTSERVICEDESC ,A.NEXTSERVICEKM, A.NEXTSERVICEDATE FROM ASPTBLVEHMAS a join gtcompmast b on A.COMPCODE = B.GTCOMPMASTID JOIN HRVEHTYPEMAST C ON C.HRVEHTYPEMASTID = A.VEHICLETYPE JOIN GTGENITEMMAST D ON D.GTGENITEMMASTID = A.FUELTYPE   JOIN GTFINANCIALYEAR E ON E.GTFINANCIALYEARID=A.FINYEAR" +
                        " JOIN  hremploymast F  ON  F.HREMPLOYMASTID=A.PARTYNAME join hremploydetails G on F.hremploymastid = G.hremploymastid AND F.IDCARDNO = G.IDCARD  WHERE A.ASPTBLVEHMASID = '" + v1.HRVehicleID + "'  AND A.COMPCODE = '" + Class.Users.COMPCODE + "'";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHMAS");
                    DataTable dt = ds.Tables["ASPTBLVEHMAS"];
                    if (dt.Rows.Count >= 1)
                    {
                        txtHRVehicleID.Text = dt.Rows[0]["ASPTBLVEHMASID"].ToString();
                        txtHRVehicleID1.Text = dt.Rows[0]["ASPTBLVEHMASID1"].ToString();
                        if (dt.Rows[0]["ACTIVE"].ToString() == "T")
                        {
                            CheckActive.Checked = true;
                        }
                        else
                        {
                            CheckActive.Checked = false;
                        }
                        if (dt.Rows[0]["ACTIVE1"].ToString() == "T")
                        {
                            CheckActive1.Checked = true;
                        }
                        else
                        {
                            CheckActive1.Checked = false;
                        }
                        combofinear.Text = Convert.ToString(dt.Rows[0]["FINYR"].ToString());

                        dateTimeDOCDATE.Text = Convert.ToString(dt.Rows[0]["DOCDATE"].ToString());
                        txtVehicleNo.Text = Convert.ToString(dt.Rows[0]["VEHICLENO"].ToString());
                        txtvehname.Text = dt.Rows[0]["VEHICLENAME"].ToString();
                        comboVehicleType.Text = dt.Rows[0]["VEHTYPE"].ToString();
                        txtCapacity.Text = dt.Rows[0]["CAPACITY"].ToString();
                        txtMRFName.Text = dt.Rows[0]["MRFNAME"].ToString();
                        dateTimeMFGDate.Text = dt.Rows[0]["MFGDATE"].ToString();
                        comboNewUsed.Text = dt.Rows[0]["NEWUSED"].ToString();
                        comboFuelType.Text = dt.Rows[0]["ITEMNAME"].ToString();
                        txtTankCapacity.Text = dt.Rows[0]["TANKCAPACITY"].ToString();
                        comboVehicleUsage.Text = dt.Rows[0]["VEHICLEUSAGE"].ToString();
                        txtEngineNo.Text = dt.Rows[0]["ENGINENO"].ToString();
                        dateTimeEngineDate.Text = dt.Rows[0]["ENGINEDATE"].ToString();
                        dateTimeRoadTaxPermitDate.Text = dt.Rows[0]["ROADTAXPERMITDATE"].ToString();
                        
                         comboempname.SelectedValue = dt.Rows[0]["HREMPLOYMASTID"].ToString();
                        
                        txtAddress.Text = dt.Rows[0]["ADDRESS"].ToString();
                        txtBrokerName.Text = dt.Rows[0]["BROKERNAME"].ToString();
                        comboMajorService.Text = dt.Rows[0]["MAJORSERVICE"].ToString();
                        txtNoofSeats.Text = dt.Rows[0]["NOOFSEATS"].ToString();
                        txtNoofTyres.Text = dt.Rows[0]["NOOFTYRES"].ToString();
                        txtNoofStepney.Text = dt.Rows[0]["NOOFSTEPNEY"].ToString();
                        txtNoofAlex.Text = dt.Rows[0]["NOOFALEX"].ToString();
                        txtMileageLitre.Text = dt.Rows[0]["MILEAGELITRE"].ToString();
                        txtVehicleIncharge.Text = dt.Rows[0]["VEHICLEINCHARGE"].ToString();
                        txtRunningKM.Text = dt.Rows[0]["RUNNINGKM"].ToString();
                        txtRunningPerKM.Text = dt.Rows[0]["RUNNINGPERKM"].ToString();
                        txtTonPerKM.Text = dt.Rows[0]["TONPERKM"].ToString();
                        txtInsuranceNo.Text = dt.Rows[0]["INSURANCENO"].ToString();
                        txtInsuranceCompany.Text = dt.Rows[0]["INSURANCECOMPANY"].ToString();
                        dateTimeInsuranceFrom.Text = dt.Rows[0]["INSURANCEFROM"].ToString();
                        dateTimeInsuranceTo.CustomFormat = dt.Rows[0]["INSURANCETO"].ToString();
                        txtInsuranceAmount.Text = dt.Rows[0]["INSURANCEAMOUNT"].ToString();
                        txtInsuranceAmount1.Text = dt.Rows[0]["INSURANCEAMOUNT1"].ToString();
                        txtPolicyName.Text = dt.Rows[0]["POLICYNAME"].ToString();
                        txtAgentName.Text = dt.Rows[0]["AGENTNAME"].ToString();
                        txtReminderBeforeDays.Text = dt.Rows[0]["REMINDERBEFOREDAYS"].ToString();
                        txtResponsibility.Text = dt.Rows[0]["RESPONSIBILITY"].ToString();
                        txtLastServiceKM.Text = dt.Rows[0]["LASTSERVICEKM"].ToString();
                        dateTimeLastServiceDate.Text = dt.Rows[0]["LASTSERVICEDATE"].ToString();
                        txtLastServiceDesc.Text = dt.Rows[0]["LASTSERVICEDESC"].ToString();
                        txtNextServiceKM.Text = dt.Rows[0]["NEXTSERVICEKM"].ToString();
                        dateTimeNextServiceDate.Text = dt.Rows[0]["NEXTSERVICEDATE"].ToString();
                        Party1();
                        //string sel2 = "SELECT A.ASPTBLVEHMASDETID,A.ASPTBLVEHMASID,A.ASPTBLVEHMASID1,A.COMPCODE,A.CompService, A.DAYS,A.KM,A.SERVICEDATE,A.NOTES  FROM ASPTBLVEHMASDET A JOIN ASPTBLVEHMAS B ON A.ASPTBLVEHMASID=B.ASPTBLVEHMASID AND A.COMPCODE=B.COMPCODE  JOIN GTCOMPMAST C ON C.GTCOMPMASTID=A.CompService  WHERE  B.ASPTBLVEHMASID = '" + txtHRVehicleID.Text + "'  ORDER BY 1";//and  A.COMPCODE='" + Class.Users.COMPCODE + "'
                        //DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLVEHMASDET");
                        //DataTable dt2 = ds2.Tables["ASPTBLVEHMASDET"];
                        //dataGridView1.DataSource = dt2;
                        //for (int i = 0; i < dt2.Rows.Count; i++)
                        // {
                        //if (Convert.ToInt64(dataGridView1.Rows[i].Cells[1].Value) > 0)
                        //{

                        // dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt2.Rows[i]["ASPTBLVEHMASDETID"].ToString());
                        //dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt2.Rows[i]["ASPTBLVEHMASID"].ToString());
                        //dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt2.Rows[i]["ASPTBLVEHMASID1"].ToString());
                        //dataGridView1.Rows[i].Cells[4].Value = Convert.ToInt64("0" + dt2.Rows[i]["COMPCODE"].ToString());
                        //dataGridView1.Rows[i].Cells[5].Value = Convert.ToInt64("0" + dt2.Rows[i]["PARTY"].ToString());
                        //dataGridView1.Rows[i].Cells[6].Value = Convert.ToInt64("0" + dt2.Rows[i]["DAYS"].ToString());
                        //dataGridView1.Rows[i].Cells[7].Value = Convert.ToInt64("0" + dt2.Rows[i]["KM"].ToString());
                        //dataGridView1.Rows[i].Cells[8].Value = Convert.ToString(dt2.Rows[i]["SERVICEDATE"].ToString());
                        //dataGridView1.Rows[i].Cells[9].Value = Convert.ToString(dt2.Rows[i]["NOTES"].ToString());

                        // }
                        //}
                        lblcount.Text = "Total Rows ount:  " + listviewveh.Items.Count.ToString();

                    }
                    //  tabControl1.TabPages.Remove(tabPage3);
                    tabControl1.SelectTab(tabPage0);
                }
                else
                {
                    MessageBox.Show("No Data Found or CompCode not Match. ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Listviewveh_ItemActivate" + ex.Message);
            }
            txtmidcard.Text = "";
        }


        private void DataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells["SNo"].Value = (e.RowIndex + 1).ToString();
        }

        private void Txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0; i = 1;
                if (txtsearch.Text.Length > 0 && listfilter.Items.Count > 0)
                {
                    listviewveh.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[3].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[4].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[5].ToString().Contains(txtsearch.Text) || listfilter.Items[item0].SubItems[7].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = i.ToString();
                            //list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[4].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[5].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[6].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[7].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[8].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[9].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[10].Text);
                            if (i % 2 == 0)
                            {
                                list.BackColor = Color.White;
                            }
                            else
                            {
                                list.BackColor = Color.WhiteSmoke;
                            }
                            i++;
                            listviewveh.Items.Add(list);
                        }
                        item0++;i++;
                        lblcount.Text = "TotalCount:  " + listviewveh.Items.Count.ToString();
                    }

                }
                else
                {
                    ListView ll = new ListView();
                    listviewveh.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {


                       

                        if (i % 2 == 0)
                        {
                            item.BackColor = Color.White;
                        }
                        else
                        {
                            item.BackColor = Color.WhiteSmoke;
                        }
                        i++;
                        this.listviewveh.Items.Add((ListViewItem)item.Clone());
                        item0++;
                    }
                    lblcount.Text = "TotalCount:  " + listviewveh.Items.Count.ToString();
                }




                //if (txtsearch.Text.Length >= 1 && listfilter.Items.Count > 1)
                //{
                //    listviewveh.Items.Clear();
                //    try
                //    {
                //        string sel1 = "SELECT A.ASPTBLVEHMASID,A.DOCDATE,A.VEHICLENO ,A.VEHICLENAME,C.VEHTYPE,A.PARTYNAME,A.COMPCODE FROM ASPTBLVEHMAS a join gtcompmast b on A.COMPCODE = B.GTCOMPMASTID JOIN HRVEHTYPEMAST C ON C.HRVEHTYPEMASTID = A.VEHICLETYPE JOIN GTGENITEMMAST D ON D.GTGENITEMMASTID = A.FUELTYPE  WHERE  A.COMPCODE = '" + Class.Users.COMPCODE + "' AND substr(a.DOCDATE,1,10) LIKE'%" + txtsearch.Text + "%' OR  A.VEHICLENO LIKE'%" + txtsearch.Text + "%' OR a.VEHICLENAME LIKE'%" + txtsearch.Text + "%' OR C.VEHTYPE LIKE'%" + txtsearch.Text + "%' OR a.PARTYNAME LIKE'%" + txtsearch.Text + "%' order by a.ASPTBLVEHMASID desc";
                //        DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHMAS");
                //        DataTable dt = ds.Tables["ASPTBLVEHMAS"];
                //        if (dt.Rows.Count > 0)
                //        {
                //            i = 1;
                //            foreach (DataRow myRow in dt.Rows)
                //            {
                //                ListViewItem list = new ListViewItem();
                //                list.SubItems.Add(i.ToString());
                //                list.SubItems.Add(myRow["ASPTBLVEHMASID"].ToString());
                //                list.SubItems.Add(myRow["DOCDATE"].ToString());
                //                list.SubItems.Add(myRow["VEHICLENO"].ToString());
                //                list.SubItems.Add(myRow["VEHICLENAME"].ToString());
                //                list.SubItems.Add(myRow["VEHTYPE"].ToString());
                //                list.SubItems.Add(myRow["PARTYNAME"].ToString());
                //                list.SubItems.Add(myRow["COMPCODE"].ToString());
                //                listviewveh.Items.Add(list);
                //                i++;
                //            }
                //            lblcount.Text = "TotalCount:  " + listviewveh.Items.Count.ToString();
                //        }
                //        else
                //        {

                //            MessageBox.Show("No Data Found");
                //            lblcount.Text = "Total Rows Count:  " + listviewveh.Items.Count.ToString();


                //        }


                //    }
                //    catch (Exception ex)
                //    {
                //        MessageBox.Show(ex.ToString());
                //    }

                //}
                //else
                //{
                //    if (listfilter.Items.Count <= 1)
                //    {
                //        listviewveh.Items.Clear();
                //        GridLoad();
                //    }
                //}
            }
            catch (Exception ex)
            {
                //MessageBox.Show("---" + ex.ToString());
            }

        }



        public void Deletes()
        {

            try
            {
                if (txtHRVehicleID.Text != "")
                {
                    string sel = "SELECT A.ASPTBLVEHTOKENID FROM ASPTBLVEHTOKEN A JOIN GTCOMPMAST b ON b.GTCOMPMASTID=A.COMPCODE  JOIN ASPTBLVEHMAS c ON c.ASPTBLVEHMASID = A.VEHICLENO  and C.COMPCODE=A.COMPCODE and C.COMPCODE=B.GTCOMPMASTID WHERE C.ASPTBLVEHMASID=" + txtHRVehicleID.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHTOKEN");
                    DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("Child Record Found.Can not Delete.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string del = "DELETE  FROM ASPTBLVEHMAS   Where COMPCODE='" + Class.Users.COMPCODE + "' and  ASPTBLVEHMASID='" + txtHRVehicleID.Text + "'";
                        Utility.ExecuteNonQuery(del);
                        //string del2 = "DELETE  FROM HRVEHMAST   Where COMPCODE='" + Class.Users.COMPCODE + "' and  ASPTBLVEHMASID='" + txtHRVehicleID.Text + "'";
                        //Utility.ExecuteNonQuery(del2);
                        //string del3 = "DELETE   FROM  HRVEHMASTDET     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  ASPTBLVEHMASID='" + txtHRVehicleID.Text + "'";
                        //Utility.ExecuteNonQuery(del3);
                        string del4 = "DELETE   FROM  ASPTBLVEHMASDET     Where COMPCODE='" + Class.Users.COMPCODE + "'  AND  ASPTBLVEHMASID='" + txtHRVehicleID.Text + "'";
                        Utility.ExecuteNonQuery(del4);
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
                MessageBox.Show("Deletes_Click" + ex.Message);
            }
            //try
            //{
            //    if (txtHRVehicleID.Text != "")
            //    {
            //        string del = "delete from   ASPTBLVEHMAS where COMPCODE='" + Class.Users.COMPCODE + "' and  ASPTBLVEHMASID1='" + txtHRVehicleID.Text + "'";
            //        Utility.ExecuteNonQuery(del);
            //        string del1 = "delete from   ASPTBLVEHMASDET where  ASPTBLVEHMASID='" + txtHRVehicleID.Text + "'";
            //        Utility.ExecuteNonQuery(del1);
            //        MessageBox.Show("Record Deleted Successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        GridLoad(); empty();
            //    }
            //    else
            //    {
            //        MessageBox.Show("Invalid", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Question);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void Imports_Click(object sender, EventArgs e)
        {

        }

        private void Pdfs_Click(object sender, EventArgs e)
        {

        }

        private void DownLoads_Click(object sender, EventArgs e)
        {

        }

        private void ChangePasswords_Click(object sender, EventArgs e)
        {

        }

        private void Logins_Click(object sender, EventArgs e)
        {

        }

        private void GlobalSearchs_Click(object sender, EventArgs e)
        {

        }

        private void TreeButtons_Click(object sender, EventArgs e)
        {

        }



        public void News()
        {
            empty(); EMPNAME();
        }

      

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtHRVehicleID.Text != "")
                {
                    string sel2 = "SELECT A.ASPTBLVEHMASDETID,A.ASPTBLVEHMASID,A.ASPTBLVEHMASID1,A.COMPCODE,A.CompService, A.DAYS,A.KM,A.SERVICEDATE,A.NOTES  FROM ASPTBLVEHMASDET A JOIN ASPTBLVEHMAS B ON A.ASPTBLVEHMASID=B.ASPTBLVEHMASID AND A.COMPCODE=B.COMPCODE  JOIN GTCOMPMAST C ON C.GTCOMPMASTID=A.COMPSERVICE WHERE  B.ASPTBLVEHMASID = '" + txtHRVehicleID.Text + "' ";
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "ASPTBLVEHMASDET");
                    DataTable dt2 = ds2.Tables["ASPTBLVEHMASDET"];
                    dataGridView1.DataSource = dt2;
                    //for (int i = 0; i < dt2.Rows.Count; i++)
                    //{
                    //    if (Convert.ToInt64(dataGridView1.Rows[i].Cells[1].Value) > 0)
                    //    {

                    //        dataGridView1.Rows[i].Cells[1].Value = Convert.ToInt64("0" + dt2.Rows[i]["ASPTBLVEHMASDETID"].ToString());
                    //        dataGridView1.Rows[i].Cells[2].Value = Convert.ToInt64("0" + dt2.Rows[i]["ASPTBLVEHMASID"].ToString());
                    //        dataGridView1.Rows[i].Cells[3].Value = Convert.ToInt64("0" + dt2.Rows[i]["COMPCODE"].ToString());
                    //        dataGridView1.Rows[i].Cells[4].Value = Convert.ToInt64("0" + dt2.Rows[i]["DAYS"].ToString());
                    //        dataGridView1.Rows[i].Cells[5].Value = Convert.ToInt64("0" + dt2.Rows[i]["KM"].ToString());
                    //        dataGridView1.Rows[i].Cells[6].Value = Convert.ToString(dt2.Rows[i]["SERVICEDATE"].ToString());
                    //        dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(dt2.Rows[i]["NOTES"].ToString());

                    //    }
                    //}
                }

            }
            catch (Exception EX)
            {
                MessageBox.Show("refreshToolStripMenuItem_Click" + EX.Message.ToString());
            }

        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //txtVehicleNo.Focus();
           // txtEngineNo.Focus();
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])//your specific tabname
            {
                //tabControl1.TabPages.Remove(tabPage3);
                txtsearch.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage0"])//your specific tabname
            {
                //  tabControl1.TabPages.Remove(tabPage3);
                txtVehicleNo.Select();
            }
            //if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage6"])//your specific tabname
            //{
            //    tabControl1.TabPages.Remove(tabPage3);
            //}
        }

       
       
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Party1();
        }

        private void vechileTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VehType();
        }

        private void fuelTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FuelType();
        }
        private int rowIndex = 0;

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    this.dataGridView1.Rows[e.RowIndex].Selected = true;
                    this.rowIndex = e.RowIndex;
                    this.dataGridView1.CurrentCell = this.dataGridView1.Rows[e.RowIndex].Cells[1];
         
                }
               
            }
            catch (Exception ex) { }
        }
        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var confirmation = MessageBox.Show("Do You want Delete this Record ?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmation == DialogResult.Yes)
                {
                    if (!this.dataGridView1.Rows[this.rowIndex].IsNewRow)
                    {
                        this.dataGridView1.Rows.RemoveAt(this.rowIndex);
                    }
                }
            }
            catch (Exception EX)
            {
               // MessageBox.Show("dataGridView1_CellContentClick" + EX.Message.ToString());
            }

        }
      //  DataTable dtadd = new DataTable();
        private void addRowToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //dtadd.Columns.Add("SNo", typeof(string));
            //dtadd.Columns.Add("ASPTBLVEHMASDETID", typeof(string));
            //dtadd.Columns.Add("ASPTBLVEHMASID", typeof(string));
            //dtadd.Columns.Add("ASPTBLVEHMASID1", typeof(string));            
            //dtadd.Columns.Add("COMPCODE", typeof(string));
            //dtadd.Columns.Add("CompService", typeof(string));
            //dtadd.Columns.Add("DAYS", typeof(string));
            //dtadd.Columns.Add("KM", typeof(string));
            //dtadd.Columns.Add("SERVICEDATE", typeof(string));
            //dtadd.Columns.Add("Notes", typeof(string));
            //dataGridView1.DataSource = dtadd;
            try
            {
                dataGridView1.Rows.Add();
            }
            catch(Exception ex) { }
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void refreshToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            listfilter.Items.Clear();GridLoad();
        }

        private void butclear_Click(object sender, EventArgs e)
        {
            //listfilter.Items.Clear(); listviewveh.Items.Clear();
            //    GridLoad();txtsearch.Select();
        }

        

        private void txtVehicleNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void txtVehicleNo_TextChanged(object sender, EventArgs e)
        {
           
            if (txtVehicleNo.Text.Length >= 8 && txtHRVehicleID.Text == "")
            {
                string sel = " SELECT DISTINCT X.VEHICLENO FROM (select A.VEHICLENO from ASPTBLVEHMAS A  where  A.VEHICLENO='" + txtVehicleNo.Text + "'  UNION ALL  select B.VEHICLENO from HRVEHMAST B  where  B.VEHICLENO='" + txtVehicleNo.Text + "')X ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHMAS");
                DataTable dt = ds.Tables["ASPTBLVEHMAS"];

                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("Already Record Found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtVehicleNo.Text = "";
                    txtVehicleNo.Select();return;
                }
            }

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dateTimeDOCDATE_ValueChanged(object sender, EventArgs e)
        {

        }

        private void txtmidcard_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //  comboempname.DataSource = null;
              
                if (txtHRVehicleID.Text != "" && txtmidcard.Text == "")
                {
                    string sel1 = "SELECT E.HREMPLOYMASTID AS asptblempid, E.FNAME AS EMPNAME   FROM ASPTBLVEHMAS A JOIN  HRVEHTYPEMAST B ON B.HRVEHTYPEMASTID=A.VEHICLETYPE  JOIN     GTCOMPMAST C ON C.GTCOMPMASTID=A.COMPCODE JOIN HREMPLOYMAST E ON E.HREMPLOYMASTID=A.PARTYNAME     JOIN  HREMPLOYDETAILS F ON F.HREMPLOYMASTID=E.HREMPLOYMASTID AND  E.IDCARDNO = F.IDCARD AND F.IDACTIVE='YES'     WHERE  A.ASPTBLVEHMASID = '" + txtHRVehicleID.Text + "'  ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHMAS");
                    DataTable dt = ds.Tables["ASPTBLVEHMAS"];
                    comboempname.DisplayMember = "EMPNAME";
                    comboempname.ValueMember = "asptblempid";
                    comboempname.DataSource = dt;
                    // comboempname.Text = dt.Rows[0]["EMPNAME"].ToString();
                }
                if (txtHRVehicleID.Text != "" && txtmidcard.Text != "")
                {
                    string selemp = "select a.hremploymastid as asptblempid,a.fname as EMPNAME from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid AND b.IDACTIVE='YES' join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname WHERE  B.MIDCARD='" + txtmidcard.Text + "' ";
                    DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
                    DataTable dtemp = dsemp.Tables["hremploymast"];

                    comboempname.DisplayMember = "EMPNAME";
                    comboempname.ValueMember = "asptblempid";
                    comboempname.DataSource = dtemp;
                }
                if (txtHRVehicleID.Text == "" && txtmidcard.Text != "")
                {
                    string selemp = "select a.hremploymastid as asptblempid,a.fname as EMPNAME from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid AND b.IDACTIVE='YES' join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname WHERE   B.MIDCARD='" + txtmidcard.Text + "' ";
                    DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
                    DataTable dtemp = dsemp.Tables["hremploymast"];

                    comboempname.DisplayMember = "EMPNAME";
                    comboempname.ValueMember = "asptblempid";
                    comboempname.DataSource = dtemp;
                    

                }
                if (txtmidcard.Text == "")
                {
                    comboempname.DataSource = null;
                    EMPNAME();
                }

            }
            catch (Exception ex) { }
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }
        void findempname(Int64 empid)
        {
            string selemp = "select a.hremploymastid as asptblempid,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as EMPNAME from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid AND b.IDACTIVE='YES' join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname WHERE C.COMPCODE='" + Class.Users.HCompcode + "' AND A.hremploymastid='" + empid + "' ";
            DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
            DataTable dtemp = dsemp.Tables["hremploymast"];

            comboempname.DisplayMember = "EMPNAME";
            comboempname.ValueMember = "asptblempid";
            comboempname.DataSource = dtemp;
        }
        private void comboempname_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void listviewveh_SelectedIndexChanged(object sender, EventArgs e)
        {

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

        private void butclear_Click_1(object sender, EventArgs e)
        {

        }

        public void Prints()
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

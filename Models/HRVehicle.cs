using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
  public class HRVehicle
    {
        public Int64 HRVehicleID;
        public Int64 HRVehicleID1;
        public string DocDate;
        public string VehicleNo;
        public string VehicleName;
        public Int64 VehicleType;
        public Int64 Capacity;
        public string MRFName;
        public string MFGDate;
        public string NewUsed;
        public Int64 FuelType;
        public Int64 TankCapacity;
        public Int64 VehicleUsage;

        public string EngineNo;
        public string EngineDate;
        public string ChassisNo;
        public string RoadTaxPermitDate;
        public string PartyName;
        public string Address;
        public string BrokerName;
        public string MajorService;

        public Int64 NoofSeats;
        public Int64 NoofTyres;
        public Int64 NoofStepney;
        public Int64 NoofAlex;
        public Int64 MileageLitre;
        public string VehicleIncharge;
        public Int64 RunningKM;
        public Int64 RunningPerKM;
        public Int64 TonPerKM;

        public string InsuranceNo;
        public string InsuranceCompany;
        public string InsuranceFrom;
        public string InsuranceTo;
        public Int64 InsuranceAmount;
        public Int64 InsuranceAmount1;
        public Int64 PremiumAmount;
        public string PolicyName;
        public string AgentName;       
        public string ReminderBeforeDays;
        public string Responsibility;

        public Int64 LastServiceKM;
        public string LastServiceDate;
        public string LastServiceDesc;
        public Int64 NextServiceKM;
        public string NextServiceDate;
        public Int64 CompCode;
        public Int64 UserName;
        public string IPAddress;
        public string Createdon;
        public string ModifiedOn;
        public Int64 FinYear;
        public string Active;
        public string Active1;
        public string VCategory;
        public HRVehicle() { }

        internal DataTable select(long hRVehicleID1, string vCategory, string active, string active1,  string vehicleNo, string PartyName, long vehicleType,long fuelType,long compCode)
        {
            string sel = "select ASPTBLVEHMASID from ASPTBLVEHMAS  where ASPTBLVEHMASID1='" + hRVehicleID1 + "' and VCategory='" + vCategory + "' and ACTIVE='" + active + "'  and ACTIVE1='" + active1 + "' and PARTYNAME='" + PartyName + "' and VEHICLENO='" + vehicleNo + "'   and VEHICLETYPE='" + vehicleType + "' and FUELTYPE='" + fuelType + "' AND  COMPCODE='" + compCode + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHMAS");
            DataTable dt = ds.Tables["ASPTBLVEHMAS"];
            return dt;
        }
        internal DataTable select(long hRVehicleID1, string vCategory, string active, string active1, string docDate, string vehicleNo, string vehicleName, long vehicleType, long capacity, string mRFName, string mFGDate, string newUsed, long fuelType, long tankCapacity, long vehicleUsage, string engineNo, string engineDate, string chassisNo, string roadTaxPermitDate, string partyName, string address, string brokerName, string majorService, long noofSeats, long noofTyres, long noofStepney, long noofAlex, long mileageLitre, string vehicleIncharge, long runningKM, long runningPerKM, long tonPerKM, string insuranceNo, string insuranceCompany, string insuranceFrom, string insuranceTo, long insuranceAmount, long insuranceAmount1, long premiumAmount ,string policyName, string agentName, string reminderBeforeDays, string responsibility, long lastServiceKM, string lastServiceDate, string lastServiceDesc, long nextServiceKM, string nextServiceDate, long hRVehicleID, long finYear, long compCode)
        {
            string sel = "select ASPTBLVEHMASID from ASPTBLVEHMAS  where ASPTBLVEHMASID1='" + hRVehicleID1 + "' and VCategory='" + vCategory + "' and ACTIVE='" + active + "'  and ACTIVE1='" + active1 + "'and DOCDATE='" + docDate + "' and VEHICLENO='" + vehicleNo + "' and VEHICLENAME='" + vehicleName + "' and VEHICLETYPE='" + vehicleType + "' and CAPACITY='" + capacity + "' and MRFNAME='" + mRFName + "' and MFGDATE='" + mFGDate + "' and NEWUSED='" + newUsed + "' and FUELTYPE='" + fuelType + "' and TANKCAPACITY='" + tankCapacity + "' and VEHICLEUSAGE='" + vehicleUsage + "' and ENGINENO='" + engineNo + "' and ENGINEDATE='" + engineDate + "' and CHASSISNO='" + ChassisNo + "' and ROADTAXPERMITDATE='" + roadTaxPermitDate + "'  and PARTYNAME='" + partyName + "' and ADDRESS='" + address + "'  and BROKERNAME='" + brokerName + "' and MAJORSERVICE='" + majorService + "' and NOOFSEATS='" + noofSeats + "' and NOOFTYRES='" + noofTyres + "' and NOOFSTEPNEY='" + NoofStepney + "' and NOOFALEX='" + noofAlex + "'  and MILEAGELITRE='" + mileageLitre + "' and VEHICLEINCHARGE='" + vehicleIncharge + "' and RUNNINGKM='" + runningKM + "' and  RUNNINGPERKM='" + runningPerKM + "'  and TONPERKM='" + tonPerKM + "'  and INSURANCENO='" + insuranceNo + "'  and INSURANCECOMPANY='" + insuranceCompany + "' and INSURANCEFROM='" + insuranceFrom + "' and INSURANCETO='" + insuranceTo + "'  and INSURANCEAMOUNT='" + insuranceAmount + "'  and INSURANCEAMOUNT1='" + insuranceAmount1 + "' and PREMIUMAMT='" + premiumAmount + "'and POLICYNAME='" + policyName + "'  and AGENTNAME='" + agentName + "'  and REMINDERBEFOREDAYS='" + reminderBeforeDays + "' and RESPONSIBILITY='" + responsibility + "'  and LASTSERVICEKM='" + lastServiceKM + "' and LASTSERVICEDATE='" + lastServiceDate + "' and  LASTSERVICEDESC='" + lastServiceDesc + "'  and NEXTSERVICEKM='" + nextServiceKM + "' and NEXTSERVICEDATE='" + nextServiceDate + "' and ASPTBLVEHMASID1='" + hRVehicleID + "' and FINYEAR='" + finYear + "' AND  COMPCODE='" + compCode + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHMAS");
            DataTable dt = ds.Tables["ASPTBLVEHMAS"];
            return dt;
        }

        //internal DataTable select(string vehicleNo, string vehicleName, long vehicleType,long fuelType,string partyName,long finYear, long compCode)
        //{
        //    string sel = "select ASPTBLVEHMASID from ASPTBLVEHMAS  where VEHICLENO='" + vehicleNo + "' and VEHICLENAME='" + vehicleName + "' and VEHICLETYPE='" + vehicleType + "'  and FUELTYPE='" + fuelType + "'  and PARTYNAME='" + partyName + "'  and FINYEAR='" + finYear + "' AND  COMPCODE='" + compCode + "'";
        //    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHMAS");
        //    DataTable dt = ds.Tables["ASPTBLVEHMAS"];
        //    return dt;
        //}
        public HRVehicle(long hRVehicleID1, string vCategory, string active, string active1, string docDate, string vehicleNo, string vehicleName, long vehicleType, long capacity, string mRFName, string mFGDate, string newUsed, long fuelType, long tankCapacity, long vehicleUsage, string engineNo, string engineDate, string chassisNo, string roadTaxPermitDate, string partyName, string address, string brokerName, string majorService, long noofSeats, long noofTyres, long noofStepney, long noofAlex, long mileageLitre, string vehicleIncharge, long runningKM, long runningPerKM, long tonPerKM, string insuranceNo, string insuranceCompany, string insuranceFrom, string insuranceTo, long insuranceAmount, long insuranceAmount1, long premiumAmount, string policyName, string agentName, string reminderBeforeDays, string responsibility, long lastServiceKM, string lastServiceDate, string lastServiceDesc, long nextServiceKM, string nextServiceDate, long compCode, long userName, string iPAddress, string createdon, string modifiedOn, long finYear, long hRVehicleID)
        {
            HRVehicleID1 = hRVehicleID1;
            VCategory = vCategory;
            Active = active;
            Active1 = active1;
            DocDate = docDate;
            VehicleNo = vehicleNo;
            VehicleName = vehicleName;
            VehicleType = vehicleType;
            Capacity = capacity;
            MRFName = mRFName;
            MFGDate = mFGDate;
            NewUsed = newUsed;
            FuelType = fuelType;
            TankCapacity = tankCapacity;
            VehicleUsage = vehicleUsage;
            EngineNo = engineNo;
            EngineDate = engineDate;
            ChassisNo = chassisNo;
            RoadTaxPermitDate = roadTaxPermitDate;
            PartyName = partyName;
            Address = address;
            BrokerName = brokerName;
            MajorService = majorService;
            NoofSeats = noofSeats;
            NoofTyres = noofTyres;
            NoofStepney = noofStepney;
            NoofAlex = noofAlex;
            MileageLitre = mileageLitre;
            VehicleIncharge = vehicleIncharge;
            RunningKM = runningKM;
            RunningPerKM = runningPerKM;
            TonPerKM = tonPerKM;
            InsuranceNo = insuranceNo;
            InsuranceCompany = insuranceCompany;
            InsuranceFrom = insuranceFrom;
            InsuranceTo = insuranceTo;
            InsuranceAmount = insuranceAmount;
            InsuranceAmount1 = insuranceAmount1;
            PremiumAmount = premiumAmount;
            PolicyName = policyName;
            AgentName = agentName;
            ReminderBeforeDays = reminderBeforeDays;
            Responsibility = responsibility;
            LastServiceKM = lastServiceKM;
            LastServiceDate = lastServiceDate;
            LastServiceDesc = lastServiceDesc;
            NextServiceKM = nextServiceKM;
            NextServiceDate = nextServiceDate;
            CompCode = compCode;
            UserName = userName;
            IPAddress = iPAddress;
            Createdon = createdon;
            ModifiedOn = modifiedOn;
            FinYear = finYear;
            HRVehicleID = hRVehicleID;
            string up = "update ASPTBLVEHMAS set ASPTBLVEHMASID1='" + hRVehicleID1 + "', VCategory='" + vCategory + "' , ACTIVE='" + active + "',ACTIVE1='" + active1 + "',DOCDATE='" + DocDate + "',VEHICLENO='" + VehicleNo + "' ,VEHICLENAME='" + vehicleName + "',VEHICLETYPE='" + VehicleType + "',CAPACITY='" + Capacity + "' ,MRFNAME='" + MRFName + "' ,MFGDATE='" + MFGDate + "',NEWUSED='" + NewUsed + "',FUELTYPE='" + FuelType + "',TANKCAPACITY='" + TankCapacity + "',VEHICLEUSAGE='" + VehicleUsage + "',ENGINENO='" + EngineNo + "',ENGINEDATE='" + EngineDate + "', ROADTAXPERMITDATE='" + RoadTaxPermitDate + "' ,PARTYNAME='" + PartyName + "',ADDRESS='" + Address + "' ,BROKERNAME='" + BrokerName + "',MAJORSERVICE='" + MajorService + "' ,NOOFSEATS='" + NoofSeats + "',NOOFTYRES='" + NoofTyres + "' ,NOOFSTEPNEY='" + NoofStepney + "',NOOFALEX='" + NoofAlex + "' ,MILEAGELITRE='" + MileageLitre + "',VEHICLEINCHARGE='" + VehicleIncharge + "',RUNNINGKM='" + RunningKM + "', RUNNINGPERKM='" + RunningPerKM + "' ,TONPERKM='" + TonPerKM + "' ,INSURANCENO='" + InsuranceNo + "' ,INSURANCECOMPANY='" + InsuranceCompany + "' ,INSURANCEFROM='" + InsuranceFrom + "' ,INSURANCETO='" + InsuranceTo + "' ,INSURANCEAMOUNT='" + InsuranceAmount + "' ,INSURANCEAMOUNT1='" + InsuranceAmount1 + "' , PREMIUMAMT='" + PremiumAmount + "',POLICYNAME='" + PolicyName + "' ,AGENTNAME='" + AgentName + "' , REMINDERBEFOREDAYS='" + ReminderBeforeDays + "',RESPONSIBILITY='" + Responsibility + "' ,LASTSERVICEKM='" + LastServiceKM + "', LASTSERVICEDATE='" + LastServiceDate + "',  LASTSERVICEDESC='" + LastServiceDesc + "' , NEXTSERVICEKM='" + NextServiceKM + "', NEXTSERVICEDATE='" + NextServiceDate + "' ,CompCode='" + CompCode + "' ,UserName='" + UserName + "',IPADDRESS='" + IPAddress + "',CREATEDON='" + Createdon + "',MODIFIEDON='" + ModifiedOn + "',FINYEAR='" + finYear + "' WHERE ASPTBLVEHMASID='" + HRVehicleID + "'  AND COMPCODE='" + CompCode + "'";
            Utility.ExecuteNonQuery(up);
            //string up1 = " UPDATE   HRVEHMAST SET ASPTBLVEHMASID='" + hRVehicleID + "',ASPTBLVEHMASID1='" + hRVehicleID1 + "', VCategory='" + vCategory + "' ,TANKCAPACITY='" + tankCapacity + "', VEHICLEUSAGE='" + VehicleUsage + "', FUELTYPE='" + FuelType + "', NEWUSED='" + NewUsed + "', MFGDATE=to_date('" + Convert.ToDateTime(MFGDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'), MRFNAME='" + MRFName + "', CAPACITY='" + Capacity + "',VEHICLETYPE='" + VehicleType + "', VEHICLENAME='" + VehicleName + "', VEHICLENO='" + VehicleNo + "', DOCDATE=to_date('" + Convert.ToDateTime(DocDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'), COMPCODE='" + CompCode + "',IPADD='" + IPAddress + "', USERID='" + UserName + "', MODIFIED_BY=to_date('" + Convert.ToDateTime(ModifiedOn).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),TONPERKM='" + TonPerKM + "', RUNPERKMS='" + RunningPerKM + "', RUNKMS='" + RunningKM + "',VEHINCHARGE='" + VehicleIncharge + "', MILAGELITRE='" + MileageLitre + "', NOOFALEX='" + NoofAlex + "', NOOFSTEP='" + NoofStepney + "', NOOFTYRE='" + NoofTyres + "',NOOFSEAT='" + NoofSeats + "', MAJORSERVICE='" + MajorService + "', ADDRESS='" + Address + "', PARTYNAME='" + PartyName + "', ROADTAXPERDT=to_date('" + Convert.ToDateTime(RoadTaxPermitDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),CHASSISNO='" + ChassisNo + "', ENGINENO='" + EngineNo + "', BROKERNAME='" + BrokerName + "', RESPONSIBILITY='" + Responsibility + "', REMBEFDAY='" + ReminderBeforeDays + "', AGENTNAME='" + AgentName + "', POLICYNAME='" + PolicyName + "', INSUREDAMT='" + InsuranceAmount + "', INSURANCETILL=to_date('" + Convert.ToDateTime(InsuranceTo).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'), INSURFORMDT=to_date('" + Convert.ToDateTime(InsuranceFrom).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),INSCOMPANY='" + InsuranceCompany + "', INSURNO='" + InsuranceNo + "', PREMIUMAMT='" + PremiumAmount + "' WHERE  ASPTBLVEHMASID='" + HRVehicleID + "' and COMPCODE='" + CompCode + "'";
            //Utility.ExecuteNonQuery(up1);
        }
        //public DataTable maxid()
        //{
        //    string sel1 = "SELECT MAX(A.ASPTBLVEHMASID) ID FROM  ASPTBLVEHMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE A.COMPCODE=" + Class.Users.COMPCODE;
        //    DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHMAS");
        //    DataTable dt = ds.Tables["ASPTBLVEHMAS"];
        //    return dt;
        //}

        public HRVehicle(long hRVehicleID, long hRVehicleID1, string active, string active1, string vCategory, string docDate, string vehicleNo, string vehicleName, long vehicleType, long capacity, string mRFName, string mFGDate, string newUsed, long fuelType, long tankCapacity, long vehicleUsage, string engineNo, string engineDate, string chassisNo, string roadTaxPermitDate, string partyName, string address, string brokerName, string majorService, long noofSeats, long noofTyres, long noofStepney, long noofAlex, long mileageLitre, string vehicleIncharge, long runningKM, long runningPerKM, long tonPerKM, string insuranceNo, string insuranceCompany, string insuranceFrom, string insuranceTo, long insuranceAmount, long insuranceAmount1,long premiumAmount, string policyName, string agentName, string reminderBeforeDays, string responsibility, long lastServiceKM, string lastServiceDate, string lastServiceDesc, long nextServiceKM, string nextServiceDate, long compCode, long userName, string iPAddress, string createdon, string modifiedOn, long finYear)
        {
            HRVehicleID = hRVehicleID;
            HRVehicleID1 = hRVehicleID1;
            Active = active;
            Active1 = active1;
            VCategory = vCategory;           
            DocDate = docDate;
            VehicleNo = vehicleNo;
            VehicleName = vehicleName;
            VehicleType = vehicleType;
            Capacity = capacity;
            MRFName = mRFName;
            MFGDate = mFGDate;
            NewUsed = newUsed;
            FuelType = fuelType;
            TankCapacity = tankCapacity;
            VehicleUsage = vehicleUsage;
            EngineNo = engineNo;
            EngineDate = engineDate;
            ChassisNo = chassisNo;
            RoadTaxPermitDate = roadTaxPermitDate;
            PartyName = partyName;
            Address = address;
            BrokerName = brokerName;
            MajorService = majorService;
            NoofSeats = noofSeats;
            NoofTyres = noofTyres;
            NoofStepney = noofStepney;
            NoofAlex = noofAlex;
            MileageLitre = mileageLitre;
            VehicleIncharge = vehicleIncharge;
            RunningKM = runningKM;
            RunningPerKM = runningPerKM;
            TonPerKM = tonPerKM;
            InsuranceNo = insuranceNo;
            InsuranceCompany = insuranceCompany;
            InsuranceFrom = insuranceFrom;
            InsuranceTo = insuranceTo;
            InsuranceAmount = insuranceAmount;
            InsuranceAmount1 = insuranceAmount1;
            PremiumAmount = premiumAmount;
            PolicyName = policyName;
            AgentName = agentName;
            ReminderBeforeDays = reminderBeforeDays;
            Responsibility = responsibility;
            LastServiceKM = lastServiceKM;
            LastServiceDate = lastServiceDate;
            LastServiceDesc = lastServiceDesc;
            NextServiceKM = nextServiceKM;
            NextServiceDate = nextServiceDate;
            CompCode = compCode;
            UserName = userName;
            IPAddress = iPAddress;
            Createdon = createdon;
            ModifiedOn = modifiedOn;
            FinYear = finYear;

            string ins = "insert into ASPTBLVEHMAS(ASPTBLVEHMASID1,VCategory,ACTIVE,ACTIVE1,DOCDATE,VEHICLENO ,VEHICLENAME,VEHICLETYPE,CAPACITY ,MRFNAME ,MFGDATE,NEWUSED,FUELTYPE,TANKCAPACITY,VEHICLEUSAGE,ENGINENO,ENGINEDATE,CHASSISNO,ROADTAXPERMITDATE ,PARTYNAME,ADDRESS ,BROKERNAME,MAJORSERVICE ,NOOFSEATS,NOOFTYRES ,NOOFSTEPNEY,NOOFALEX ,MILEAGELITRE,VEHICLEINCHARGE,RUNNINGKM, RUNNINGPERKM ,TONPERKM ,INSURANCENO ,INSURANCECOMPANY ,INSURANCEFROM ,INSURANCETO ,INSURANCEAMOUNT ,INSURANCEAMOUNT1 , POLICYNAME ,AGENTNAME , REMINDERBEFOREDAYS,RESPONSIBILITY ,LASTSERVICEKM, LASTSERVICEDATE,  LASTSERVICEDESC , NEXTSERVICEKM, NEXTSERVICEDATE,COMPCODE,USERNAME,IPADDRESS,CREATEDON,MODIFIEDON,FINYEAR,PREMIUMAMT)" +
                "values('" + hRVehicleID1 + "','" + vCategory + "','" + Active + "','" + Active1 + "','" + DocDate + "','" + VehicleNo + "','" + VehicleName + "','" + VehicleType + "','" + Capacity + "','" + MRFName + "','" + MFGDate + "','" + NewUsed + "','" + FuelType + "','" + TankCapacity + "','" + VehicleUsage + "','" + EngineNo + "','" + EngineDate + "','" + ChassisNo + "','" + RoadTaxPermitDate + "','" + PartyName + "','" + Address + "','" + BrokerName + "','" + MajorService + "','" + NoofSeats + "','" + NoofTyres + "','" + NoofStepney + "','" + NoofAlex + "','" + MileageLitre + "','" + VehicleIncharge + "','" + RunningKM + "','" + RunningPerKM + "','" + TonPerKM + "','" + InsuranceNo + "','" + InsuranceCompany + "','" + InsuranceFrom + "','" + InsuranceTo + "','" + InsuranceAmount + "','" + InsuranceAmount1 + "','" + PolicyName + "','" + AgentName + "','" + ReminderBeforeDays + "','" + Responsibility + "','" + lastServiceKM + "','" + LastServiceDate + "','" + LastServiceDesc + "','" + NextServiceKM + "','" + NextServiceDate + "','" + CompCode + "','" + UserName + "','" + IPAddress + "','" + Createdon + "','" + ModifiedOn + "','" + finYear + "','" + PremiumAmount + "')";
            Utility.ExecuteNonQuery(ins);
           // DataTable dt = maxid();
            //string ins1 = "INSERT INTO  HRVEHMAST(ASPTBLVEHMASID, ASPTBLVEHMASID1,VCategory, TANKCAPACITY, VEHICLEUSAGE, FUELTYPE, NEWUSED, MFGDATE, " +
            //    "MRFNAME, CAPACITY, VEHICLETYPE, VEHICLENAME, VEHICLENO,DOCDATE, COMPCODE, IPADD, USERID, MODIFIED_BY, MODIFIED_ON, CREATED_BY, CREATED_ON, TONPERKM, RUNPERKMS," +
            //    " RUNKMS, VEHINCHARGE, MILAGELITRE, NOOFALEX, NOOFSTEP, NOOFTYRE, NOOFSEAT, MAJORSERVICE, ADDRESS, PARTYNAME, ROADTAXPERDT, CHASSISNO, ENGINENO, BROKERNAME," +
            //    " RESPONSIBILITY, REMBEFDAY, AGENTNAME, POLICYNAME, INSUREDAMT, INSURANCETILL, INSURFORMDT, INSCOMPANY, INSURNO,PREMIUMAMT)" +
            //    "VALUES('" + dt.Rows[0]["ID"].ToString() + "','" + hRVehicleID1 + "','" + vCategory + "','" + tankCapacity + "','" + VehicleUsage + "','" + FuelType + "','" + NewUsed + "',to_date('" + Convert.ToDateTime(MFGDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')," +
            //    "'" + MRFName + "','" + Capacity + "','" + VehicleType + "','" + VehicleName + "','" + VehicleNo + "',to_date('" + Convert.ToDateTime(DocDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')," +
            //    "'" + CompCode + "','" + IPAddress + "','" + UserName + "',to_date('" + Convert.ToDateTime(ModifiedOn).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')," +
            //    "to_date('" + Convert.ToDateTime(Createdon).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),to_date('" + Convert.ToDateTime(Createdon).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')," +
            //    "to_date('" + Convert.ToDateTime(Createdon).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),'" + TonPerKM + "','" + RunningPerKM + "'," +
            //    "'" + RunningKM + "','" + VehicleIncharge + "','" + MileageLitre + "','" + NoofAlex + "','" + NoofStepney + "','" + NoofTyres + "','" + NoofSeats + "'," +
            //    "'" + MajorService + "','" + Address + "','" + PartyName + "',to_date('" + Convert.ToDateTime(RoadTaxPermitDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')," +
            //    "'" + ChassisNo + "','" + EngineNo + "','" + BrokerName + "','" + Responsibility + "','" + ReminderBeforeDays + "','" + AgentName + "','" + PolicyName + "'," +
            //    "'" + InsuranceAmount + "',to_date('" + Convert.ToDateTime(InsuranceTo).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),to_date('" + Convert.ToDateTime(InsuranceFrom).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')," +
            //    "'" + InsuranceCompany + "','" + InsuranceNo + "','" + PremiumAmount + "')";
            //Utility.ExecuteNonQuery(ins1);
        }
        
    }
    public class HRVehicleDet
    {
        public Int64 HRVehicleDetID;
        public Int64 HRVehicleID;
        public Int64 HRVehicleID1;
        public Int64 CompCode;
        public Int64 CompService;
        public Int64 Days;
        public Int64 KM;
        public String ServiceDate;
        public string Notes;
        public Int64 HRVEHMASTDETROW;
        public HRVehicleDet()
        {
            
        }
        public HRVehicleDet(long hRVehicleID, long hRVehicleID1, long compCode, long compService, long kM, long days, string serviceDate, string notes, long hRVEHMASTDETROW)
        {

            HRVehicleID = hRVehicleID;
            HRVehicleID1 = hRVehicleID1;
            CompCode = compCode;
            CompService = compService;
            KM = kM;
            Days = days;
            ServiceDate = serviceDate;
            Notes = notes;
            HRVEHMASTDETROW = hRVEHMASTDETROW;
            string ins = "insert into ASPTBLVEHMASDET(ASPTBLVEHMASID,ASPTBLVEHMASID1,COMPCODE,CompService,KM,DAYS,SERVICEDATE,NOTES)values('" + HRVehicleID + "','" + HRVehicleID1 + "','" + CompCode + "','" + CompService + "','" + KM + "','" + Days + "',to_date('" + Convert.ToDateTime(ServiceDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),'" + Notes + "')";
            Utility.ExecuteNonQuery(ins);

           // string ins1 = "INSERT INTO HRVEHMASTDET(HRVEHMASTID,ASPTBLVEHMASID, ASPTBLVEHMASID1, COMPCODE, COMPSERVICE,  KILOMETERS, DAYS,SERVICEDATE,NOTES,HRVEHMASTDETROW) VALUES('" + HRVehicleID + "','" + HRVehicleID + "','" + HRVehicleID1 + "','" + CompCode + "','" + CompService + "','" + KM + "','" + Days + "',to_date('" + Convert.ToDateTime(ServiceDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),'" + Notes + "','"+ hRVEHMASTDETROW + "')";
           //Utility.ExecuteNonQuery(ins1);
        }
        public HRVehicleDet(long hRVehicleDetID, long hRVehicleID, long hRVehicleID1, long compCode, long compService, long kM, long days, string serviceDate, string notes,long hRVEHMASTDETROW)
        {
            HRVehicleDetID = hRVehicleDetID;
            HRVehicleID = hRVehicleID;
            HRVehicleID1 = hRVehicleID1;
            CompService = compService;
            CompCode = compCode;
            KM = kM;
            Days = days;
            ServiceDate = serviceDate;           
            Notes = notes;
            HRVEHMASTDETROW = hRVEHMASTDETROW;
            string up = "update  ASPTBLVEHMASDET SET ASPTBLVEHMASID='" + HRVehicleID + "', ASPTBLVEHMASID1='" + HRVehicleID1 + "' , COMPCODE='" + CompCode + "' , CompService='" + compService + "' ,KM='" + KM + "',DAYS='" + Days + "',SERVICEDATE=to_date('" + Convert.ToDateTime(ServiceDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),NOTES='" + Notes + "' WHERE ASPTBLVEHMASDETID='" + HRVehicleDetID + "' AND COMPCODE='" + CompCode + "'";
            Utility.ExecuteNonQuery(up);
            //string up1 = "UPDATE HRVEHMASTDET SET HRVEHMASTID='" + HRVehicleID + "',ASPTBLVEHMASID='" + HRVehicleID + "',ASPTBLVEHMASID1='" + HRVehicleID1 + "',COMPCODE='" + CompCode + "' ,COMPSERVICE='" + CompService + "' ,  KILOMETERS='" + KM + "' , DAYS='" + Days + "' ,SERVICEDATE=to_date('" + Convert.ToDateTime(ServiceDate).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),NOTES='" + Notes + "',HRVEHMASTDETROW='" + hRVEHMASTDETROW + "' WHERE HRVEHMASTID='" + HRVehicleDetID + "' AND COMPCODE='" + CompCode + "'";
            //Utility.ExecuteNonQuery(up1);
        }

        internal DataTable select(long hRVehicleDetID, long hRVehicleID, long hRVehicleID1, long compCode, long compService, long kM, long days, string serviceDate, string notes)
        {
            string sel = "SELECT ASPTBLVEHMASDETID FROM ASPTBLVEHMASDET WHERE ASPTBLVEHMASID='" + hRVehicleDetID + "' AND ASPTBLVEHMASID1='" + hRVehicleID1 + "' AND COMPCODE='" + CompCode + "' AND CompService='" + CompService + "' AND KM='" + KM + "' AND DAYS='" + Days + "' AND SERVICEDATE='" + ServiceDate + "' AND NOTES='" + Notes + "' AND ASPTBLVEHMASDETID='" + HRVehicleDetID + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLVEHMASDET");
            DataTable dt = ds.Tables["ASPTBLVEHMASDET"];
            return dt;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    class FuelRateMasterModel
    {
       
     public FuelRateMasterModel()
        { }

        public Int64 ASPTBLFUELRATEMASID { get; set; }
         public Int64 ASPTBLFUELRATEID { get; set; }
        public Int64 FINYEAR { get; set; }
        public Int64 COMPCODE { get; set; }
        public Int64 BUNKNAME { get; set; }
        public DateTime FUELDATE { get; set; }
        public string FUELTOKEN { get; set; }
        public string MONTHNAME { get; set; }
        public Int64 USERNAME { get; set; }
        public DateTime MODIFIED { get; set; }
        public DateTime CREATEDON { get; set; }
        public String IPADDRESS { get; set; }

        internal DataTable select(long fINYEAR, long cOMPCODE, DateTime fUELDATE,long bUNKNAME)
        {
            string sel1 = " SELECT A.ASPTBLFUELRATEMASID  FROM  ASPTBLFUELRATEMAS A JOIN ASPTBLFUELRATEMASDET B ON A.ASPTBLFUELRATEMASID=B.ASPTBLFUELRATEMASID WHERE A.FINYEAR=" + fINYEAR + "   AND A.COMPCODE=" + cOMPCODE + "  AND A.FUELDATE=to_date('" + Convert.ToDateTime(fUELDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy') and  A.BUNKNAME='" + bUNKNAME + "' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLFUELRATEMAS");
            DataTable dt = ds.Tables["ASPTBLFUELRATEMAS"];
            return dt;
        }
        public FuelRateMasterModel(long aSPTBLFUELRATEID, long fINYEAR, long cOMPCODE, DateTime fUELDATE, long uSERNAME, DateTime mODIFIED, string iPADDRESS, string fUELTOKEN, long bUNKNAME, long aSPTBLFUELRATEMASID,string mONTHNAME)
        {
            ASPTBLFUELRATEID = aSPTBLFUELRATEID;
            FINYEAR = fINYEAR;
            COMPCODE = cOMPCODE;
            FUELDATE = fUELDATE;
            USERNAME = uSERNAME;
            MODIFIED = mODIFIED;
            IPADDRESS = iPADDRESS;
            FUELTOKEN = fUELTOKEN;
            BUNKNAME = bUNKNAME;
            ASPTBLFUELRATEMASID = aSPTBLFUELRATEMASID;
            MONTHNAME = mONTHNAME;
            string UP = "UPDATE ASPTBLFUELRATEMAS SET ASPTBLFUELRATEID='" + ASPTBLFUELRATEID + "',FINYEAR=" + FINYEAR + ",COMPCODE=" + COMPCODE + ",FUELDATE=to_date('" + Convert.ToDateTime(FUELDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'), USERNAME='" + USERNAME + "',MODIFIED=to_date('" + Convert.ToDateTime(MODIFIED) + "','DD-MM-YYYY HH24:MI:SS'), IPADDRESS='" + IPADDRESS + "',  FUELTOKEN='" + FUELTOKEN + "' ,BUNKNAME='" + BUNKNAME + "',MONTHNAME='" + MONTHNAME + "' WHERE ASPTBLFUELRATEMASID='" + ASPTBLFUELRATEMASID + "' and COMPCODE='" + COMPCODE + "' ";
            Utility.ExecuteNonQuery(UP);
        }

        public FuelRateMasterModel(long aSPTBLFUELRATEID, long fINYEAR, long cOMPCODE, DateTime fUELDATE, long uSERNAME, DateTime mODIFIED, DateTime cREATEDON, string iPADDRESS, string fUELTOKEN,long bUNKNAME,string mONTHNAME)
        {
            ASPTBLFUELRATEID = aSPTBLFUELRATEID;
            FINYEAR = fINYEAR;
            COMPCODE = cOMPCODE;
            FUELDATE = fUELDATE;
            USERNAME = uSERNAME;
            MODIFIED = mODIFIED;
            CREATEDON = cREATEDON;
            IPADDRESS = iPADDRESS;
            FUELTOKEN = fUELTOKEN; BUNKNAME = bUNKNAME;
            MONTHNAME = mONTHNAME;
            string ins = "INSERT INTO ASPTBLFUELRATEMAS(ASPTBLFUELRATEID,FINYEAR,COMPCODE,FUELDATE,USERNAME,MODIFIED,CREATEDON,IPADDRESS,FUELTOKEN,BUNKNAME,MONTHNAME)VALUES('" + ASPTBLFUELRATEID + "','" + FINYEAR + "','" + COMPCODE + "',to_date('" + Convert.ToDateTime(FUELDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),'" + USERNAME + "',to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),to_date('" + Convert.ToDateTime(Class.Users.CREATED).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),'" + IPADDRESS + "','" + FUELTOKEN + "','" + BUNKNAME + "','" + MONTHNAME + "')";
            Utility.ExecuteNonQuery(ins);
        }

      
       
    }
    class FuelRateMasterModelDet
    {
        public Int64 ASPTBLFUELRATEMASDETID { get; set; }
        public Int64 ASPTBLFUELRATEMASID { get; set; }
        public long ASPTBLFUELRATEID { get; set; }
        public Int64 COMPCODE { get; set; }
        public Int64 BUNKNAME { get; set; }
        public Int64 ITEMNAME { get; set; }
        public Decimal FUELRATE { get; set; }
       
        public string FUELRATE2 { get; set; }
        public DateTime FUELDATE { get; set; }

        public FuelRateMasterModelDet()
        {
        }

        public FuelRateMasterModelDet(long aSPTBLFUELRATEMASID, long aSPTBLFUELRATEID, long cOMPCODE, long iTEMNAME, string fUELRATE2, DateTime fUELDATE,long bUNKNAME)
        {
            ASPTBLFUELRATEMASID = aSPTBLFUELRATEMASID;
            ASPTBLFUELRATEID = aSPTBLFUELRATEID;
            COMPCODE = cOMPCODE;
            ITEMNAME = iTEMNAME;
            FUELRATE2 = fUELRATE2;
            FUELDATE = fUELDATE;
            BUNKNAME = bUNKNAME;
            string ins = "INSERT INTO ASPTBLFUELRATEMASDET(ASPTBLFUELRATEMASID,ASPTBLFUELRATEID, COMPCODE ,  ITEMNAME,  FUELRATE2,FUELDATE,BUNKNAME)" +
               "VALUES(" + ASPTBLFUELRATEMASID + ",'" + ASPTBLFUELRATEID + "','" + COMPCODE + "'," + ITEMNAME + ",'" +FUELRATE2 + "',to_date('" + Convert.ToDateTime(FUELDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),'" + BUNKNAME + "')";
            Utility.ExecuteNonQuery(ins);
        }

        public FuelRateMasterModelDet(long aSPTBLFUELRATEMASID, long aSPTBLFUELRATEID, long cOMPCODE, long iTEMNAME, string fUELRATE2, DateTime fUELDATE, long bUNKNAME, long aSPTBLFUELRATEMASDETID)
        {
            ASPTBLFUELRATEMASID = aSPTBLFUELRATEMASID;
            ASPTBLFUELRATEID = aSPTBLFUELRATEID;
            COMPCODE = cOMPCODE;
            ITEMNAME = iTEMNAME;
            FUELRATE2 = fUELRATE2;
            FUELDATE = fUELDATE; BUNKNAME = bUNKNAME;
            ASPTBLFUELRATEMASDETID = aSPTBLFUELRATEMASDETID;
            string UP = "UPDATE  ASPTBLFUELRATEMASDET SET ASPTBLFUELRATEMASID=" + ASPTBLFUELRATEMASID + ",  ASPTBLFUELRATEID='" + ASPTBLFUELRATEID + "', COMPCODE='" + COMPCODE + "', ITEMNAME=" + ITEMNAME + " ,  FUELRATE2='" + FUELRATE2 + "',FUELDATE=to_date('" + Convert.ToDateTime(FUELDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'), BUNKNAME ='" + BUNKNAME + "'  WHERE ASPTBLFUELRATEMASDETID='" + ASPTBLFUELRATEMASDETID + "' AND COMPCODE='" + COMPCODE + "' ";
            Utility.ExecuteNonQuery(UP);

        }

        
        internal DataTable select(long cOMPCODE, long iTEMNAME, string fUELRATE2, DateTime fUELDATE, long bUNKNAME)
        {
            string sel1 = " SELECT B.ASPTBLFUELRATEMASDETID  FROM   ASPTBLFUELRATEMASDET B WHERE B.COMPCODE=" + cOMPCODE + " AND B.ITEMNAME=" + iTEMNAME + "  AND B.FUELRATE2=" + fUELRATE2 + " AND B.FUELDATE=to_date('" + Convert.ToDateTime(fUELDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy') and BUNKNAME='" + bUNKNAME + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLFUELRATEMASDET");
            DataTable dt = ds.Tables["ASPTBLFUELRATEMASDET"];
            return dt;
        }
    }
}

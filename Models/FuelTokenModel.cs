using System;
using System.Data;

namespace Pinnacle.Models
{
    class FuelTokenModel
    {
      
        public Int64 ASPTBLVEHTOKENID { get; set; }
        public Int64 ASPTBLVEHTOKENID1 { get; set; }
        public Int64 FINYEAR { get; set; }
        public Int64 COMPCODE { get; set; }
        public string TOKENNO { get; set; }
        public DateTime TOKENDATE { get; set; }
        public Int64 VEHICLENO { get; set; }
        public Int64 VEHICLETYPE { get; set; }
        public Int64 EMPNAME { get; set; }
        public string EMPNAME1 { get; set; }
        public Int64 BUNKNAME { get; set; }
        public string ACTIVE { get; set; }
        public string TOKENCANCEL { get; set; }
        public Int64 TOTALLITRES { get; set; }
        public string REMARKS { get; set; }
        public Byte[] TOKENIMAGE { get; set; }
        public Int64 USERNAME { get; set; }
        public DateTime MODIFIED { get; set; }
        public DateTime CREATEDON { get; set; }
        public string IPADDRESS { get; set; }
        public string FUELTOKEN { get; set; }
        public string VCategory { get; set; }
        public Int64 PREKM { get; set; }
        public Int64 LASTKM { get; set; }
        public Int64 TOTALKM { get; set; }
        public Int64 LOILKM { get; set; }
        public Int64 LOILRKM { get; set; }
        public FuelTokenModel(long aSPTBLVEHTOKENID1, long fINYEAR,string vCategory, long bUNKNAME, long cOMPCODE, string tOKENNO, DateTime tOKENDATE, long vEHICLENO, long vEHICLETYPE, long eMPNAME, string eMPNAME1, string aCTIVE, string tOKENCANCEL, long tOTALLITRES, string rEMARKS, long uSERNAME, DateTime mODIFIED, DateTime cREATEDON, string iPADDRESS,string fUELTOKEN, Int64 pREKM, Int64 lASTKM, Int64 tOTALKM, Int64 lOILKM, Int64 lOILRKM)
        {
            ASPTBLVEHTOKENID1 = aSPTBLVEHTOKENID1;
            FINYEAR = fINYEAR;
            VCategory = vCategory;
            BUNKNAME = bUNKNAME;
            COMPCODE = cOMPCODE;
            TOKENNO = tOKENNO;
            TOKENDATE = tOKENDATE;
            VEHICLENO = vEHICLENO;
            VEHICLETYPE = vEHICLETYPE;
            EMPNAME = eMPNAME;
            EMPNAME1 = eMPNAME1;
            ACTIVE = aCTIVE;
            TOKENCANCEL = tOKENCANCEL;
            TOTALLITRES = tOTALLITRES;
            REMARKS = rEMARKS;
            USERNAME = uSERNAME;
            MODIFIED = mODIFIED;
            CREATEDON = cREATEDON;
            IPADDRESS = iPADDRESS;
            FUELTOKEN = fUELTOKEN;
            PREKM = pREKM; LASTKM = lASTKM; TOTALKM = tOTALKM;
            LOILKM = lOILKM;
            LOILRKM = lOILRKM;


            string ins = "INSERT INTO ASPTBLVEHTOKEN(ASPTBLVEHTOKENID1,FINYEAR,VCategory,BUNKNAME,COMPCODE,  TOKENNO,  TOKENDATE ,  VEHICLENO,  VEHICLETYPE ,  EMPNAME,EMPNAME1, ACTIVE,  TOKENCANCEL,  TOTALLITRES ,  REMARKS ,  USERNAME ,  MODIFIED,  CREATEDON ,  IPADDRESS,FUELTOKEN,PREKM ,LASTKM,TOTALKM,LOILKM,LOILRKM)VALUES('" + ASPTBLVEHTOKENID1 + "','" + FINYEAR + "','" + VCategory + "','" + BUNKNAME + "' ,'" + COMPCODE + "',  '" + TOKENNO + "', to_date('" + Convert.ToDateTime(TOKENDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),  " + VEHICLENO + ",  " + VEHICLETYPE + " ,  '" + EMPNAME + "', '" + EMPNAME1 + "','" + ACTIVE + "',  '" + TOKENCANCEL + "'," + TOTALLITRES + " ,  '" + REMARKS + "' ,  " + USERNAME + " ,'" + MODIFIED + "','" + CREATEDON + "',  '" + IPADDRESS + "','" + FUELTOKEN + "','" + pREKM + "','" + lASTKM + "' ,'" + tOTALKM + "','" + lOILKM + "','" + lOILRKM + "')";
            Utility.ExecuteNonQuery(ins);
        }

        public FuelTokenModel(long aSPTBLVEHTOKENID1, long fINYEAR,string vCategory, long bUNKNAME, long cOMPCODE, string tOKENNO, DateTime tOKENDATE, long vEHICLENO, long vEHICLETYPE, long eMPNAME, string eMPNAME1, string aCTIVE, string tOKENCANCEL, long tOTALLITRES, string rEMARKS, long uSERNAME, DateTime mODIFIED, string iPADDRESS,string fUELTOKEN, Int64 pREKM, Int64 lASTKM, Int64 tOTALKM, Int64 lOILKM, Int64 lOILRKM,long aSPTBLVEHTOKENID)
        {
            ASPTBLVEHTOKENID = aSPTBLVEHTOKENID;
            ASPTBLVEHTOKENID1 = aSPTBLVEHTOKENID1;
            FINYEAR = fINYEAR;
            VCategory = vCategory;
            BUNKNAME = bUNKNAME;
            COMPCODE = cOMPCODE;
            TOKENNO = tOKENNO;
            TOKENDATE = tOKENDATE;
            VEHICLENO = vEHICLENO;
            VEHICLETYPE = vEHICLETYPE;
            EMPNAME = eMPNAME;
            EMPNAME1 = eMPNAME1;
            ACTIVE = aCTIVE;
            TOKENCANCEL = tOKENCANCEL;
            TOTALLITRES = tOTALLITRES;
            REMARKS = rEMARKS;
            USERNAME = uSERNAME;
            MODIFIED = mODIFIED;
            IPADDRESS = iPADDRESS;
            FUELTOKEN = fUELTOKEN;
            PREKM = pREKM; LASTKM = lASTKM; TOTALKM = tOTALKM;
            LOILKM = lOILKM;
            LOILRKM = lOILRKM;
            string UP = "UPDATE ASPTBLVEHTOKEN SET ASPTBLVEHTOKENID1='" + aSPTBLVEHTOKENID1 + "',FINYEAR=" + FINYEAR + ",VCategory='"+VCategory+"', BUNKNAME=" + BUNKNAME + ",COMPCODE=" + COMPCODE + ",  TOKENNO='" + TOKENNO + "',  TOKENDATE=to_date('" + Convert.ToDateTime(TOKENDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy'),  VEHICLENO=" + VEHICLENO + ",  VEHICLETYPE=" + VEHICLETYPE + " ,  EMPNAME='" + EMPNAME + "',EMPNAME1='" + EMPNAME1 + "',  ACTIVE='" + ACTIVE + "',  TOKENCANCEL='" + TOKENCANCEL + "',TOTALLITRES=" + TOTALLITRES + " ,  REMARKS='" + REMARKS + "' ,  USERNAME=" + USERNAME + " ,MODIFIED='"+MODIFIED+"',IPADDRESS='" + IPADDRESS + "',FUELTOKEN='" + FUELTOKEN + "' , PREKM = '" + pREKM + "' , LASTKM = '" + lASTKM + "' ,  TOTALKM = '" + tOTALKM + "',  LOILKM = '" + lOILKM + "',  LOILRKM = '" + lOILRKM + "' WHERE ASPTBLVEHTOKENID='" + ASPTBLVEHTOKENID + "' and COMPCODE='" + COMPCODE + "'";
            Utility.ExecuteNonQuery(UP);

        }

        public FuelTokenModel()
        {
        }

        internal DataTable select(long aSPTBLVEHTOKENID, long aSPTBLVEHTOKENID1, long fINYEAR, long bUNKNAME, long cOMPCODE, string tOKENNO,DateTime tOKENDATE, long vEHICLENO, long vEHICLETYPE, long eMPNAME, string eMPNAME1, string aCTIVE, string tOKENCANCEL, long tOTALLITRES,string rEMARKS, Int64 pREKM, Int64 lASTKM, Int64 tOTALKM, Int64 lOILKM, Int64 lOILRKM)
        {
            string sel1 = " SELECT A.ASPTBLVEHTOKENID  FROM  ASPTBLVEHTOKEN A JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENID WHERE A.ASPTBLVEHTOKENID='" + aSPTBLVEHTOKENID + "' AND  A.ASPTBLVEHTOKENID1='" + aSPTBLVEHTOKENID1 + "' AND A.FINYEAR=" + fINYEAR + "  AND A.BUNKNAME=" + bUNKNAME + " AND A.COMPCODE=" + cOMPCODE + " AND A.TOKENNO='" + tOKENNO + "' AND A.TOKENDATE=to_date('" + Convert.ToDateTime(tOKENDATE).ToString("dd-MM-yyyy") + "', 'dd-MM-yyyy')  AND A.VEHICLENO=" + vEHICLENO + "  AND A.VEHICLETYPE=" + vEHICLETYPE + "  AND A.EMPNAME=" + eMPNAME + " AND A.EMPNAME1='" + eMPNAME1 + "' AND A.ACTIVE='" + aCTIVE + "' AND A.TOKENCANCEL='" + tOKENCANCEL + "'  AND A.TOTALLITRES='" + tOTALLITRES + "' AND A.REMARKS='" + rEMARKS + "'  AND PREKM = '" + pREKM + "' AND LASTKM = '" + lASTKM + "' AND  TOTALKM = '" + tOTALKM + "' AND  LOILRKM = '" + lOILKM + "' AND  LOILRKM = '" + lOILRKM + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKEN");
            DataTable dt = ds.Tables["ASPTBLVEHTOKEN"];
            return dt;
        }
    }
    class FuelTokenModelDet
    {
        
        public Int64 ASPTBLVEHTOKENDETID { get; set; }
        public Int64 ASPTBLVEHTOKENID { get; set; }
        public Int64 ITEMNAME { get; set; }
        public string ITEMDESC { get; set; }
        public Int64 KM { get; set; }
        public string LITRES { get; set; }
           
        public string NOTES { get; set; }
        public Int64 COMPCODE { get; set; }
        public Int64 ASPTBLVEHTOKENID1 { get; set; }

       
        public FuelTokenModelDet(long aSPTBLVEHTOKENID, long iTEMNAME, string iTEMDESC, long kM, string lITRES, string nOTES, long cOMPCODE, long aSPTBLVEHTOKENID1)
        {
            ASPTBLVEHTOKENID = aSPTBLVEHTOKENID;
            ITEMNAME = iTEMNAME;
            ITEMDESC = iTEMDESC;
            KM = kM;
            LITRES = lITRES;
            NOTES = nOTES;
            COMPCODE = cOMPCODE;
            ASPTBLVEHTOKENID1 = aSPTBLVEHTOKENID1;
            string ins = "INSERT INTO ASPTBLVEHTOKENDET(ASPTBLVEHTOKENID,  ITEMNAME,ITEMDESC, KM ,  LITRES, NOTES,COMPCODE,ASPTBLVEHTOKENID1)VALUES(" + ASPTBLVEHTOKENID + "," + ITEMNAME + ",'" + ITEMDESC + "'," + KM + ",'" + LITRES + "','" + NOTES + "','" + COMPCODE + "','" + aSPTBLVEHTOKENID1 + "' )";
            Utility.ExecuteNonQuery(ins);
        }

        public FuelTokenModelDet(long aSPTBLVEHTOKENID, long iTEMNAME, string iTEMDESC, long kM, string lITRES, string nOTES, long cOMPCODE, long aSPTBLVEHTOKENID1, long aSPTBLVEHTOKENDETID)
        {
            ASPTBLVEHTOKENID = aSPTBLVEHTOKENID;
            ITEMNAME = iTEMNAME;
            ITEMDESC = iTEMDESC;
            KM = kM;
            LITRES = lITRES;
            NOTES = nOTES;
            COMPCODE = cOMPCODE;
            ASPTBLVEHTOKENID1 = aSPTBLVEHTOKENID1;
            ASPTBLVEHTOKENDETID = aSPTBLVEHTOKENDETID;
           
            string UP = "UPDATE  ASPTBLVEHTOKENDET SET ASPTBLVEHTOKENID=" + ASPTBLVEHTOKENID + ",  ITEMNAME=" + ITEMNAME + ", ITEMDESC='" + ITEMDESC + "',  KM=" + KM + " ,  LITRES='" + LITRES + "',NOTES='" + NOTES + "' ,COMPCODE='" + COMPCODE + "',ASPTBLVEHTOKENID1='" + aSPTBLVEHTOKENID1 + "' WHERE ASPTBLVEHTOKENDETID='" + ASPTBLVEHTOKENDETID + "' AND COMPCODE='" + COMPCODE + "' ";
            Utility.ExecuteNonQuery(UP);
        }

        public FuelTokenModelDet()
        {
        }

        internal DataTable select(long aSPTBLVEHTOKENDETID, long aSPTBLVEHTOKENID, long iTEMNAME, long kM, string lITRES, string nOTES, long cOMPCODE, long aSPTBLVEHTOKENID1)
        {
            string sel1 = " SELECT A.ASPTBLVEHTOKENID  FROM  ASPTBLVEHTOKEN A JOIN ASPTBLVEHTOKENDET B ON A.ASPTBLVEHTOKENID=B.ASPTBLVEHTOKENDETID WHERE B.ASPTBLVEHTOKENDETID=" + aSPTBLVEHTOKENDETID + "  AND B.ASPTBLVEHTOKENID1=" + aSPTBLVEHTOKENID1 + "  and B.ITEMNAME=" + iTEMNAME + "AND B.KM=" + kM + " AND B.LITRES='" + lITRES + "' AND B.NOTES='" + nOTES + "' AND B.COMPCODE='" + cOMPCODE + "' AND B.ASPTBLVEHTOKENID=" + aSPTBLVEHTOKENID + " ";
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLVEHTOKENDET");
            DataTable dt = ds.Tables["ASPTBLVEHTOKENDET"];
            return dt;
        }
    }
}
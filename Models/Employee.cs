using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    class Employee
    {

        public Int64 ASPTBLEMPID { get; set; }
        public Int64 ASPTBLEMPID1 { get; set; }
        public Int64 COMPCODE { get; set; }
        public Int64 COMNAME { get; set; }
        public string EMPNAME { get; set; }
        public string LASTNAME { get; set; }
        public string ADDRESS { get; set; }
        public string GENDER { get; set; }
        public string DATEOFBIRTH { get; set; }
        public Int64 DEPARTMENT { get; set; }
        public string DATEOFJOIN { get; set; }
        public Int64 IDCARDNO { get; set; }
        public string CONTACT { get; set; }
        public string BLOODGROUP { get; set; }
        public string EMPLOYEETYPE { get; set; }
        public string ACTIVE { get; set; }
        public Int64 USERNAME { get; set; }
        public string IPADDRESS { get; set; }       
        public string CREATEDON { get; set; }
        public string MODIFIEDON { get; set; }
        public byte[] bytes { get; set; }
        public string ObjectId { get; set; }
        public string eformId { get; set; }
        public string fileContent { get; set; }
        public string fileName { get; set; }
        public Int64 SALARY { get; set; }
        public Int64 imagebyte { get; set; }
        internal DataTable select(long cOMPCODE, string eMPNAME, string lASTNAME, string aDDRESS, string gENDER, string dATEOFBIRTH, long dEPARTMENT, string dATEOFJOIN, long iDCARDNO, string cONTACT, string bLOODGROUP, string aCTIVE, long uSERNAME, string eMPLOYEETYPE,long sALARY,Int64 iMAGEBYTE)
        {
            string sel = "select asptblempid from asptblemp where  compcode=" + cOMPCODE + " and empname ='" + eMPNAME + "' and lastname='" + lASTNAME + "' and address ='" + aDDRESS + "' and  gender='" + gENDER + "' and dateofbirth='" + dATEOFBIRTH + "' and department=" + dEPARTMENT + " and dateofjoin ='" + dATEOFJOIN + "' and idcardno =" + iDCARDNO + "  and contact ='" + cONTACT + "' and bloodgroup ='" + bLOODGROUP + "' and active ='" + aCTIVE + "' and username=" + uSERNAME + " and employeetype='" + eMPLOYEETYPE + "' AND  SALARY='" + sALARY+ "' AND  imagebyte='" + imagebyte + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblemp");
            DataTable dt = ds.Tables["asptblemp"];
            return dt;
        }
        //internal DataTable select(long aSPTBLEMPID1, long cOMPCODE, string eMPNAME, string lASTNAME, string aDDRESS, string gENDER, string dATEOFBIRTH, long dEPARTMENT,string dATEOFJOIN,long iDCARDNO, string cONTACT, string bLOODGROUP, string aCTIVE, long uSERNAME,string eMPLOYEETYPE,Int64 sALARY)
        //{
        //    string sel = "select asptblempid from asptblemp where asptblempid1='" + aSPTBLEMPID1 + "' and compcode=" + cOMPCODE + " and empname ='" + eMPNAME + "' and lastname='" + lASTNAME + "' and address ='" + aDDRESS + "' and  gender='" + gENDER + "' and dateofbirth='" + dATEOFBIRTH + "' and department=" + dEPARTMENT + " and dateofjoin ='" + dATEOFJOIN + "' and idcardno =" + iDCARDNO + "  and contact ='" + cONTACT + "' and bloodgroup ='" + bLOODGROUP + "' and active ='" + aCTIVE + "' and username=" + uSERNAME + " and employeetype='" + eMPLOYEETYPE + "' AND  SALARY=" + sALARY ;
        //    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblemp");
        //    DataTable dt = ds.Tables["asptblemp"];
        //    return dt;
        //}
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace Pinnacle.Models
{
    public class SecInventryModel
    {


        public Int64 InventryID;
        public string QrCode;
        public string Category;
        public string SystemDate;
        public string SystemTime;        
        public Int64 CompCode;
        public Int64 UserName;
        public string GateName;
        public DateTime Modified;
        public DateTime CreatedOn;
        public string IpAddress;
        public Int64 FinYear;
        public string GateDcNo;
        public string AGFIN;
        public string AGFOUT;
        public string AGFMIN;
        public string AGFMOUT;
        public string FLFIN;
        public string FLFOUT;
        public string FLFDIN;
        public string FLFDOUT;

        public string AGFCIN;
        public string AGFCOUT;
        public string AGFPIN;
        public string AGFPOUT;
        public string AGFKIN;
        public string AGFKOUT;
        public string AGFMGIIIN;
        public string AGFMGIIOUT;
        public string BSAIN;
        public string BSAOUT;
        public string BVKIN;
        public string BVKOUT;
        public SecInventryModel()
        {
            

        }
        public SecInventryModel(Int64 finYear, string gateDcNo, string aGFIN, string aGFOUT, string aGFMIN, string aGFMOUT, string fLFIN, string fLFOUT, string fLFDIN, string fLFDOUT, string qrCode, string category, string systemDate, string systemTime, Int64 compCode, Int64 userName, string gatename, DateTime modified, DateTime createdOn, string ipAddress)
        {

            QrCode = qrCode;
            Category = category;
            SystemDate = systemDate;
            SystemTime = systemTime;
            CompCode = compCode;
            UserName = userName;
            GateName = gatename;
            Modified = modified;
            CreatedOn = createdOn;
            IpAddress = ipAddress;
            FinYear = finYear;
            GateDcNo = gateDcNo;
            AGFIN = aGFIN;
            AGFOUT = aGFOUT;
            AGFMIN = aGFMIN;
            AGFMOUT = aGFMOUT;
            FLFIN = fLFIN;
            FLFOUT = fLFOUT;
            FLFDIN = fLFDIN;
            FLFDOUT = fLFDOUT;           
            string ins = "INSERT INTO ASPINVENTRY(FINYEAR,GATEDCNO,AGFIN,AGFOUT,AGFMIN,AGFMOUT,FLFIN,FLFOUT,FLFDIN,FLFDOUT,QrCode,Category,SystemDate,SystemTime,CompCode,userName,GateName,Modified,CreatedOn,IpAddress)  VALUES(" + FinYear + ",'" + GateDcNo + "','" + AGFIN + "', '" + AGFOUT + "','" + AGFMIN + "','" + AGFMOUT + "','" + FLFIN + "','" + FLFOUT + "','" + FLFDIN + "','" + FLFDOUT + "','" + QrCode + "','" + Category + "', '" + SystemDate + "','" + SystemTime + "','" + CompCode + "','" + UserName + "','" + GateName + "' ,to_date('" + Convert.ToDateTime(Modified).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),to_date('" + Convert.ToDateTime(CreatedOn.ToString()) + "', 'dd/MM/yyyy hh24:MI:SS'),'" + IpAddress + "')";
            Utility.ExecuteNonQuery(ins);
        }
        public SecInventryModel(Int64 finYear, string gateDcNo, string aGFIN, string aGFOUT, string aGFMIN, string aGFMOUT, string fLFIN, string fLFOUT, string fLFDIN, string fLFDOUT, string qrCode, string category, string systemDate, string systemTime, Int64 compCode, Int64 userName, string gatename, DateTime modified, DateTime createdOn, string ipAddress, string bSAIN, string bSAOUT, string bVKIN, string bVKOUT)
        {

            QrCode = qrCode;
            Category = category;
            SystemDate = systemDate;
            SystemTime = systemTime;
            CompCode = compCode;
            UserName = userName;
            GateName = gatename;
            Modified = modified;
            CreatedOn = createdOn;
            IpAddress = ipAddress;
            FinYear = finYear;
            GateDcNo = gateDcNo;
            AGFIN = aGFIN;
            AGFOUT = aGFOUT;
            AGFMIN = aGFMIN;
            AGFMOUT = aGFMOUT;
            FLFIN = fLFIN;
            FLFOUT = fLFOUT;
            FLFDIN = fLFDIN;
            FLFDOUT = fLFDOUT;
            BSAIN = bSAIN; BSAOUT= bSAOUT; BVKIN = bVKIN; BVKOUT = bVKOUT;
            string ins = "INSERT INTO ASPINVENTRY(FINYEAR,GATEDCNO,AGFIN,AGFOUT,AGFMIN,AGFMOUT,FLFIN,FLFOUT,FLFDIN,FLFDOUT,QrCode,Category,SystemDate,SystemTime,CompCode,userName,GateName,Modified,CreatedOn,IpAddress, BSAIN,BSAOUT,BVKIN,BVKOUT)  VALUES(" + FinYear + ",'" + GateDcNo + "','" + AGFIN + "', '" + AGFOUT + "','" + AGFMIN + "','" + AGFMOUT + "','" + FLFIN + "','" + FLFOUT + "','" + FLFDIN + "','" + FLFDOUT + "','" + QrCode + "','" + Category + "', '" + SystemDate + "','" + SystemTime + "','" + CompCode + "','" + UserName + "','" + GateName + "' ,to_date('" + Convert.ToDateTime(Modified).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),to_date('" + Convert.ToDateTime(CreatedOn.ToString()) + "', 'dd/MM/yyyy hh24:MI:SS'),'" + IpAddress + "','" + BSAIN + "','" + BSAOUT + "','" + BVKIN + "','" + BVKOUT + "' )";
            Utility.ExecuteNonQuery(ins);
        }
        public SecInventryModel(Int64 finYear, string gateDcNo, string aGFIN, string aGFOUT, string aGFMIN, string aGFMOUT, string fLFIN, string fLFOUT, string fLFDIN, string fLFDOUT, string qrCode, string category, string systemDate, string systemTime, Int64 compCode, Int64 userName, string gatename, DateTime modified, string ipAddress, string bSAIN, string bSAOUT, string bVKIN, string bVKOUT, Int64 inventryID)
        {
            QrCode = qrCode;
            Category = category;
            SystemDate = systemDate;
            SystemTime = systemTime;
            CompCode = compCode;
            UserName = userName;
            GateName = gatename;
            Modified = modified;
            IpAddress = ipAddress;
            FinYear = finYear;
            GateDcNo = gateDcNo;
            AGFIN = aGFIN;
            AGFOUT = aGFOUT;
            AGFMIN = aGFMIN;
            AGFMOUT = aGFMOUT;
            FLFIN = fLFIN;
            FLFOUT = fLFOUT;
            FLFDIN = fLFDIN;
            FLFDOUT = fLFDOUT;
            InventryID = inventryID;
            BSAIN = bSAIN; BSAOUT = bSAOUT; BVKIN = bVKIN; BVKOUT = bVKOUT;
            string up = "UPDATE  ASPINVENTRY A SET A.FINYEAR=" + FinYear + ",GATEDCNO='" + GateDcNo + "',AGFIN='"+ AGFIN+ "', AGFOUT='" + AGFOUT + "', aGFMIN='" + AGFMIN + "', AGFMOUT='" + AGFMOUT + "',FLFIN='" + FLFIN + "', FLFOUT='" + FLFOUT + "',    FLFDIN ='" + FLFDIN + "',FLFDOUT ='" + FLFDOUT + "', A.QrCode='" + QrCode + "' ,Category='" + category + "', A.SystemDate='" + systemDate + "' ,  A.SystemTime=to_char('" + SystemTime.ToString() + "') , A.CompCode='" + CompCode + "' ,A.UserName='" + UserName + "' ,A.GateName='" + GateName + "' , A.Modified=to_date('" + Convert.ToDateTime(Modified).ToString() + "', 'dd/MM/yyyy hh24:MI:SS'),A.IpAddress='" + IpAddress + "',BSAIN ='"+BSAIN+ "', BSAOUT='" + BSAOUT + "', BVKIN ='" + BVKIN + "', BVKOUT='" + BVKOUT + "', WHERE A.InventryID='" + InventryID + "'";
            Utility.ExecuteNonQuery(up);
        }



        internal DataTable select()
        {
            string sel = "SELECT MAX(A.INVENTRYID) AS INVENTRYID FROM ASPINVENTRY A   ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPINVENTRY");
            DataTable dt = ds.Tables["ASPINVENTRY"];
            return dt;
        }
        internal DataTable findcompcode()
        {
            string sel1 = "SELECT MAX(A.INVENTRYID) INVENTRYID FROM  ASPINVENTRY A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID JOIN gtfinancialyear C ON C.GTFINANCIALYEARID=A.FINYEAR  and C.CURRENTFINYR='T' WHERE A.COMPCODE=" + Class.Users.COMPCODE;
            DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPINVENTRY");
            DataTable dt = ds.Tables["ASPINVENTRY"];
            return dt;
        }
        internal DataTable select1()
        {
            string sel = "SELECT MAX(A.INVENTRYID) AS INVENTRYID FROM ASPINVENTRY A  ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPINVENTRY");
            DataTable dt = ds.Tables["ASPINVENTRY"];
            return dt;
        }
        internal DataTable select(Int64 InventryID)
        {
            string sel = "SELECT A.INVENTRYID,A.GATEDCNO, A.AGFIN, A.SYSTEMDATE,A.SYSTEMTIME FROM ASPINVENTRY A where a.InventryID=" + InventryID;
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPINVENTRY");
            DataTable dt = ds.Tables["ASPINVENTRY"];
            return dt;
        }
        internal DataTable select(string InventryID,Int64 compcode,Int64 finyear)
        {
            string sel = "SELECT A.INVENTRYID,A.GATEDCNO,A.SYSTEMDATE,A.SYSTEMTIME FROM ASPINVENTRY A join GTCOMPMAST b on a.compcode=b.gtcompmastid where a.INVENTRYID='" + InventryID + "' and a.compcode='" + compcode + "'and a.finyear='" + finyear + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPINVENTRY");
            DataTable dt = ds.Tables["ASPINVENTRY"];
            return dt;
        }
    }
}
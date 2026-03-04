using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Collections;
using System.Text;
namespace Pinnacle.Models
{
    
    public class CityMasterModel
    {
       
        private string p1;
        private string p2;
        private int p3;
        private int p4;
        private string p5;
        private string p6;
        private DateTime p7;
        private string p8;
        private DateTime p9;
        private string p10;
       
        public string CityId { get; set; }
        public string City { get; set; }  
        public int State { get; set; }       
        public int Country { get; set; }
        public string chk { get; set; }
        public bool Active { get; set; }
        public string UserName { get; set; }
        public DateTime Modified { get; set; }   
        public DateTime CreatedOn { get; set; }
        public string Createdby { get; set; }
        public string IpAddress { get; set; }
        public  CityMasterModel()
        {

           
        }
      
        public DataTable select(string p1, string p2, int p3, int p4, string p5)
        {
           
            //this.p1 = CityId;
            //this.p2 = City;
            //this.p3 = State;
            //this.p4 = Country;
            //this.p5 = chk;
            //this.p6 = UserName;
            //this.p7 = IpAddress;
            //this.p8 = CreatedOn;
            string sel = "SELECT A.TBLCITMASID FROM TBLCITMAS A WHERE   A.city='" + p2 + "' AND A.state='" + p3 + "' AND  A.COUNTRY='" + p4 + "' AND A.ACTIVE='" + p5 + "'  ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "TBLCITMAS");
            DataTable dt = ds.Tables["TBLCITMAS"];
            return dt;
        }
        public CityMasterModel(string p2, int p3, int p4, string p5, string p6, DateTime p7, string p8, DateTime p9, string p10)
        {
            //this.p2 = City;
            //this.p3 = State;
            //this.p4 = Country;
            //this.p5 = chk;
            //this.p6 = UserName;
            //this.p7 = IpAddress;
            //this.p8 = CreatedOn;
            //this.p9 = Createdby;
            //this.p10 = Modified;
            string ins = "INSERT INTO TBLCITMAS(CITY,STATE,COUNTRY,ACTIVE,USERNAME,MODIFIEDON,CREATEDBY,CREATEDON,IPADD)  VALUES('" + p2 + "','" + p3 + "','" + p4 + "','" + p5 + "','" + p6 + "',to_date('" + Convert.ToDateTime(p7).ToString() + "','dd-MM-yyyy hh24:mi:ss'),'" + p8 + "',to_date('" + Convert.ToDateTime(p9).ToString() + "','dd-MM-yyyy hh24:mi:ss'),'dd-MM-yyyy hh24:mi:ss') ,'" + p10 + "')";
            Utility.ExecuteNonQuery(ins);
        }
        public CityMasterModel(string p1, string p2, int p3, int p4, string p5, string p6, DateTime p7, string p8, DateTime p9, string p10)
        {
           
            //this.p1 = CityId;
            //this.p2 = City;
            //this.p3 = State;
            //this.p4 = Country;
            //this.p5 = chk;
            //this.p6 = UserName;
            //this.p7 = IpAddress;
            //this.p8 = CreatedOn;
            //this.p9 = Createdby;
            //this.p10 = Modified;
            string up = "UPDATE  TBLCITMAS A SET   A.CITY='" + p2 + "' , A.STATE='" + p3 + "' ,  A.COUNTRY='" + p4 + "' , A.ACTIVE='" + p5 + "' , A.USERNAME='" + p6 + "' , A.MODIFIEDON=to_date('" + Convert.ToDateTime(p7).ToString() + "','dd-MM-yyyy hh24:mi:ss') , A.CREATEDBY='" + p8 + "',A.CREATEDON=to_date('" + Convert.ToDateTime(p9).ToString() + "','dd-MM-yyyy hh24:mi:ss'),A.IPADD='" + p10 + "' WHERE A.TBLCITMASID='" + p1 + "'";
            Utility.ExecuteNonQuery(up);                                           
        }
       
        //public DataTable selectstate()
        //{
        //    string sel = "SELECT A.STATEID,A.STATE FROM TBLSTAMAS A WHERE A.ACTIVE='True' OR A.ACTIVE='T'";
        //    DataSet ds = Utility.ExecuteSelectQuery(sel, "TBLSTAMAS");
        //    DataTable dt = ds.Tables["TBLSTAMAS"];
        //    return dt;
        //}
      
    
        public DataTable selectstate()
        {
            string sel = " SELECT A.TBLSTAMASID,A.STATE FROM TBLSTAMAS A JOIN TBLCOUMAS B ON B.TBLCOUMASID=A.COUNTRY    WHERE A.ACTIVE='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "TBLSTAMAS");
            DataTable dt = ds.Tables["TBLSTAMAS"];
            return dt;
        }
       
        public DataTable selectcity1(string s)
        {

            string sel = "SELECT B.TBLCOUMASID,B.COUNTRY FROM TBLSTAMAS A JOIN TBLCOUMAS B ON A.COUNTRY=B.TBLCOUMASID WHERE  A.TBLSTAMASID ='" + s + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "TBLCOUMAS");
            DataTable dt = ds.Tables["TBLCOUMAS"];
            return dt;
        }
        public DataTable selectcity2(string s)
        {

            string sel = "SELECT A.TBLSTAMASID FROM TBLSTAMAS A WHERE A.ACTIVE='T' AND  A.STATE ='" + s + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "TBLSTAMAS");
            DataTable dt = ds.Tables["TBLSTAMAS"];
            return dt;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Collections;
namespace Pinnacle.Models
{
    public class Companay
    {
        public Companay()
        {
            // TODO: Complete member initialization
        }


        private int p1 { get; set; }
        private int p2 { get; set; }
        private string p3 { get; set; }
        private string p4 { get; set; }
        private string p5 { get; set; }
        private int p6 { get; set; }
        private int p7 { get; set; }
        private int p8 { get; set; }
        private int p9 { get; set; }
        private string p10 { get; set; }
        private string p11 { get; set; }
        private DateTime dateTime { get; set; }
        private Int64 p12 { get; set; }
        private Int64 p13 { get; set; }
        private string p14 {get;set ;}
        private string p15 { get; set; }
        private string p16 { get ; set ; }
        private bool p17 { get ; set ; }
        private string p18 { get; set; }
        private string p19 { get; set; }
        private string p20 { get; set; }
        private string p21 { get ; set ; }
        private string p22 { get ; set ; }

        public int CompId { get { return this.p1; } set { this.p1 = value; } }
        public int SequenceId { get { return this.p2; } set { this.p2 = value; } }
        public string CompCode { get { return this.p3; } set { this.p3 = value; } }
        public string CompName { get { return this.p4; } set { this.p4 = value; } }
        public string Address { get { return this.p5; } set { this.p5 = value; } }
        public int City { get { return this.p6; } set { this.p6 = value; } }
        public int State { get { return this.p7; } set { this.p7 = value; } }
        public int Country { get { return this.p8; } set { this.p8 = value; } }
        public int PinCode { get { return this.p9; } set { this.p9 = value; } }
        public string PanNo { get { return this.p10; } set { this.p10 = value; } }
        public string CstNo { get { return this.p11; } set { this.p11 = value; } }
        public DateTime GstDate { get { return this.dateTime; } set { this.dateTime = value; } }
        public Int64 PhoneNo { get { return this.p12; } set { this.p12 = value; } }
        public Int64 FaxNo { get { return this.p13; } set { this.p13 = value; } }
        public string Email { get { return this.p14; } set { this.p14 = value; } }
        public string WebSite { get { return this.p15; } set { this.p15 = value; } }
        public string Active { get { return this.p16; } set { this.p16 = value; } }
        public bool chk { get { return this.p17; } set { this.p17 = value; } }
        public string UserName { get { return this.p18; } set { this.p18 = value; } }
        public string CreatedOn { get { return this.p19; } set { this.p19 = value; } }
        public string CreatedBy { get { return this.p20; } set { this.p20 = value; } }
        public string ModifiedOn { get { return this.p21; } set { this.p21 = value; } }
        public string IpAddress { get { return this.p22; } set { this.p22 = value; } }


        public Companay(int p2, string p3, string p4, string p5, int p6, int p7, int p8, int p9, string p10, string p11, DateTime dateTime, Int64 p12, Int64 p13, string p14, string p15, string p16, string p18, string p19, string p20, string p21, string p22)
        {
         
            string ins = "insert into gtcompmast (sequenceid ,  compcode,   compname,  address,   city,  state ,   country,   pincode,  " +
            " panno,    cstno,    gstdate ,    phoneno ,   faxno ,  email,  website,   active,  username,   createdon,   createdby,   modifiedon,   ipaddress)" +
            "Values('" + p2 + "','" + p3 + "','" + p4 + "','" + p5 + "','" + p6 + "','" + p7 + "','" + p8 + "','" + p9 + "','" + p10 + "','" + p11 + "','" + Convert.ToDateTime(dateTime).ToString("dd-MMM-yyyy") + "','" + p12 + "','" + p13 + "','" + p14 + "','" + p15 + "','" + p16 + "','" + p18 + "','" + p19 + "','" + p20 + "','" + p21 + "','" + p22 + "')";
            Utility.ExecuteNonQuery(ins);
        }
        public Companay(int p1, int p2, string p3, string p4, string p5, int p6, int p7, int p8, int p9, string p10, string p11, DateTime dateTime, Int64 p12, Int64 p13, string p14, string p15, string p16, string p18, string p19, string p20, string p21, string p22)
        {
         
            string up = "update  gtcompmast  set sequenceid='" + p2 + "' ,  compcode='" + p3 + "' ,   compname='" + p4 + "' ,  address='" + p5 + "' ,   city='" + p6 + "' ,  " +
                "State='" + p7 + "'  ,   country='" + p8 + "' ,   pincode='" + p9 + "' ,   panno='" + p10 + "' ,    cstno='" + p11 + "' ,    gstdate='" + Convert.ToDateTime(dateTime).ToString("dd-MMM-yyyy") + "' ,    phoneno='" + p12 + "' , " +
                "Faxno='" + p13 + "'  ,  email='" + p14 + "' ,  website='" + p15 + "' ,   active='" + p16 + "' ,  username='" + p18 + "' ,   createdon='" + p19 + "' ,   createdby='" + p20 + "' ," +
                "Modifiedon='" + p21 + "' ,   ipaddress='" + p22 + "' where gtcompmastid='" + p1 + "'";
            Utility.ExecuteNonQuery(up);
        }





        internal DataTable select(int p1, int p2, string p3, string p4, string p5, int p6, int p7, int p8, int p9, string p10, string p11, DateTime dateTime, Int64 p12, Int64 p13, string p14, string p15, string p16, string p18, string p19, string p20, string p21, string p22)
        {
            string sel = "select gtcompmastid from gtcompmast where sequenceid='" + p2 + "'  and  compcode='" + p3 + "'  and   compname='" + p4 + "' and  address='" + p5 + "'  and   city='" + p6 + "' and  " +
                "State='" + p7 + "'  and  country='" + p8 + "'  and   pincode='" + p9 + "' and  panno='" + p10 + "'  and  cstno='" + p11 + "'  and   gstdate='" + Convert.ToDateTime(dateTime).ToString("dd-MMM-yyyy") + "'  and phoneno='" + p12 + "'  and " +
                "Faxno='" + p13 + "' and  email='" + p14 + "' and  website='" + p15 + "' and   active='" + p16 + "' and  username='" + p18 + "'  and  createdon='" + p19 + "' and createdby='" + p20 + "' and " +
                "Modifiedon='" + p21 + "' and  ipaddress='" + p22 + "' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }
        public DataTable selectcity()
        {
            string sel = " select a.tblcitmasid,a.city from tblcitmas a  where a.active='T'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "tblcitmas");
            DataTable dt = ds.Tables["tblcitmas"];
            return dt;
        }
        public DataTable selectcity(int id)
        {
            string sel = " select a.tblstamasid,a.state from tblstamas a join tblcitmas b on b.state=a.tblstamasid and b.country=a.country    where a.active='T' and b.tblcitmasid='" + id + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "tblstamas");
            DataTable dt = ds.Tables["tblstamas"];
            return dt;
        }
        public DataTable selectcountry(int s)
        {

            string sel = " select distinct b.tblcoumasid,b.country from tblstamas a join tblcoumas b on a.country=b.tblcoumasid  where  a.tblstamasid='" + s + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "tblstamas");
            DataTable dt = ds.Tables["tblstamas"];
            return dt;
        }

        public DataTable selectdate(int s)
        {

            string sel = " SELECT DISTINCT A.GSTDATE FROM gtcompmast A  WHERE  A.gtcompmastID='" + s + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }
    }
}

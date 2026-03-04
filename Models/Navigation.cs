using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
namespace Pinnacle.Models
{
    class Navigation
    {

       

          
            public Int64 MenuID { get; set; }
            public Int64 MenuNameID { get; set; }
            public string MenuName { get; set; }
            public string NavURL { get; set; }
            public Int64 ParentMenuID { get; set; }
            public Int64 CompCode { get; set; }
            public Int64 UserName { get; set; }
            public bool Chk { get; set; }
            public string Active { get; set; }
            public Navigation()
            {


            }
        //public string Encrypt(string str)
        //{
        //    string EncrptKey = "2013;[pnuLIT)WebCodeExpert";
        //    byte[] byKey = { };
        //    byte[] IV = { 18, 52, 86, 120, 144, 171, 205, 239 };
        //    byKey = System.Text.Encoding.UTF8.GetBytes(EncrptKey.Substring(0, 8));
        //    DESCryptoServiceProvider des = new DESCryptoServiceProvider();
        //    byte[] inputByteArray = Encoding.UTF8.GetBytes(str);
        //    MemoryStream ms = new MemoryStream();
        //    CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write);
        //    cs.Write(inputByteArray, 0, inputByteArray.Length);
        //    cs.FlushFinalBlock();
        //    return Convert.ToBase64String(ms.ToArray());
        //}

        //public string Decrypt(string str)
        //{
        //    str = str.Replace(" ", "+");
        //    string DecryptKey = "2013;[pnuLIT)WebCodeExpert";
        //    byte[] byKey = { };
        //    byte[] IV = { 18, 52, 86, 120, 144, 171, 205, 239 };
        //    byte[] inputByteArray = new byte[str.Length];

        //    byKey = System.Text.Encoding.UTF8.GetBytes(DecryptKey.Substring(0, 8));
        //    DESCryptoServiceProvider des = new DESCryptoServiceProvider();
        //    inputByteArray = Convert.FromBase64String(str.Replace(" ", "+"));
        //    MemoryStream ms = new MemoryStream();
        //    CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write);
        //    cs.Write(inputByteArray, 0, inputByteArray.Length);
        //    cs.FlushFinalBlock();
        //    System.Text.Encoding encoding = System.Text.Encoding.UTF8;
        //    return encoding.GetString(ms.ToArray());
        //}

        public Navigation(string menuName, string navURL, Int64 parentMenuID, string active, Int64 menuNameID, Int64 compCode, Int64 userName)
        {
            MenuName = menuName;
            NavURL = navURL;
            ParentMenuID = parentMenuID;
            Active = active;
            MenuNameID = menuNameID;
            CompCode = compCode;
            UserName = userName;
            string ins = "insert into  asptblnavigation  (menuname,navurl,parentmenuid,active,menunameid,compcode,username)values( '" + MenuName + "' , '" + NavURL + "' , " + ParentMenuID + " , '" + Active + "' , " + MenuNameID + " , " + CompCode + " , " + UserName + "  )";
            Utility.ExecuteNonQuery(ins);
        }
        public Navigation(string menuName, string navURL, Int64 parentMenuID, string active, Int64 menuNameID, Int64 compCode, Int64 userName, Int64 menuID)
        {
            MenuName = menuName;
            NavURL = navURL;
            ParentMenuID = parentMenuID;
            Active = active;
            MenuNameID = menuNameID;
            CompCode = compCode;
            UserName = userName;
            MenuID = menuID; 
            string up = "update asptblnavigation  set   menuname='" + MenuName + "' , navurl='" + NavURL + "' , parentmenuid=" + ParentMenuID + " , active='" + Active + "' , menunameid=" + MenuNameID + " ,compcode=" + CompCode + ",username=" + UserName + " where menuid=" + MenuID;
            Utility.ExecuteNonQuery(up);
        }

      
      
        
        internal System.Data.DataTable Ordertreefilter(Int64 s, Int64 ss)
            {


                string sel = "select   a.menuid,a.menuname,a.navurl,a.parentmenuid from asptbluserrights a where a.compcode='" + s + "' and a.username='" + ss + "' and a.active='T'   and a.parentmenuid >0   order by 2 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
                DataTable dt1 = ds.Tables["asptbluserrights"];
                return dt1;
            }
            internal System.Data.DataTable Imagefilter1()
            {
                string sel = " select a.tblvistorinfoid, a.phoneno,a.empname ,a.imagepath,a.intime  from tblvistorinfo   a   where a.sactive='F' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "tblvistorinfo");
                DataTable dt1 = ds.Tables["tblvistorinfo"];
                return dt1;
            }
        public DataTable findmenuid(Int64 s)
        {
            string sel = "   select b.gtcompmastid, b.compcode, d.userid,d.username,c.menunameid,c.menuname from asptblnavigation a join  gtcompmast b on a.compcode=b.gtcompmastid join asptblmenuname c on  a.menunameid = c.menunameid     join asptblusermas d on d.userid = a.username where a.menuid=" + s;
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            DataTable dt = ds.Tables["asptbluserrights"];
            return dt;
        }
        //public DataTable findmenuid(Int64 s)
        //{
        //    string sel = "    SELECT B.GTCOMPMASTID,B.COMPCODE,C.USERID,C.USERNAME FROM ASPTBLNAVIGATION A JOIN GTCOMPMAST B ON A.COMPCODE = B.GTCOMPMASTID  JOIN asptblusermas C ON C.USERID = A.USERNAME    WHERE A.MENUID=" + s;
        //    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLNAVIGATION");
        //    DataTable dt = ds.Tables["ASPTBLNAVIGATION"];
        //    return dt;
        //}

        internal DataTable select(string menuName,string navURL , long parentMenuID, string active, long menuNameID, long compCode, long userName)
        {
          
            MenuName = menuName;
            NavURL = navURL;
            ParentMenuID = parentMenuID;
            Active = active;
            MenuNameID = menuNameID;
            CompCode = compCode;
            UserName = userName;


            string sel = "select  a.menuid from asptblnavigation a where   a.menuname='" + MenuName + "' and a.navurl='" + NavURL + "'  and a.parentmenuid=" + ParentMenuID + " and a.active='" + Active + "' and a.menunameid=" + MenuNameID + " and a.compcode=" + CompCode + " and a.username=" + UserName;
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblnavigation");
            DataTable dt1 = ds.Tables["asptblnavigation"];
            return dt1;
        }

        internal System.Data.DataTable Imagefilter1(string d)
            {

                string sel = " select a.asptblvisinfoid, a.phone,a.firstname,a.imagepath,a.startdate  from asptblvisinfo a  where a.active='T' and a.imagetitle='" + d + "' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblvisinfo");
                DataTable dt1 = ds.Tables["asptblvisinfo"];
                return dt1;
            }
            public DataTable treefilter(Int64 s, Int64 ss)
            {
                string sel = "select a.menuid,a.menuname,a.navurl,a.parentmenuid from asptbluserrights a where a.compcode='" + s + "' and a.username='" + ss + "' and a.active='T' order by 4,2 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
                DataTable dt1 = ds.Tables["asptbluserrights"];
                return dt1;

            }
            public DataTable treefilter()
            {
                string sel = "select a.menuid,ifnull(a.parentmenuid,0) as parentmenuid,a.menuname,ifnull(b.formname,'') as formname,ifnull(a.formid,0) as formid from mas_menu a left join mas_forms b on b.formid = a.formid where a.active ='Y' order by a.menuid,parentmenuid asc";
            // string sel = "SELECT A.MENUID,A.MENUNAME,A.NAVURL,A.PARENTMENUID FROM ASPTBLUSERRIGHTS A WHERE A.COMPCODE='" + s + "' AND A.USERNAME='" + ss + "' AND A.ACTIVE='T' ORDER BY 4,2 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "Mas_Menu");
                DataTable dt1 = ds.Tables["Mas_Menu"];
            Class.Users.Query = "";
            Class.Users.Query = sel;
            return dt1;

            }
            public DataTable UserFilter(Int64 s, Int64 ss)
            {

                string sel = "select a.menuid,a.menuname,a.navurl,a.parentmenuid from asptblmenuname a where a.compcode='" + s + "' and a.username='" + ss + "' and a.active='T' order by 3";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblmenuname");
                DataTable dt1 = ds.Tables["asptblmenuname"];
            Class.Users.Query = "";
            Class.Users.Query = sel;
            return dt1;

            }
            public DataTable headerdropdowns(Int64 s, Int64 ss)
            {
                string sel = "select distinct a.menuname,a.navurl,a.active,a.saves,a.prints,a.readonly,a.search,a.deletes from asptbluserrights a where a.compcode='" + s + "' and a.username='" + ss + "' and a.active='T'  and a.parentmenuid >0  order by 1 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
                DataTable dt1 = ds.Tables["asptbluserrights"];
            Class.Users.Query = "";
            Class.Users.Query = sel;
            return dt1;
            }



        public DataTable select()
        {
            string sel = "  select a.menuid as MENUID1 ,a.menuname as menuname1,a.navurl,a.parentmenuid,a.active,a.menunameid as menunameid1,b.GTCOMPMASTID ,c.USERID from asptblnavigation a  join gtcompmast b on a.compcode = b.gtcompmastid  join asptblusermas c on c.userid = a.username    join asptblmenuname d on d.menunameid=a.menunameid  order by a.menuid asc ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblnavigation");
            DataTable dt1 = ds.Tables["asptblnavigation"];
            Class.Users.Query = "";
            Class.Users.Query = sel;
            return dt1;
        }
        public DataTable select(string s, string ss)
        {
           string sel = "select a.menuid as MENUID1 ,d.menuname as menuname1,a.navurl,a.parentmenuid,a.active,a.menunameid as menunameid1,b.GTCOMPMASTID ,c.USERID from asptblnavigation a  join gtcompmast b on a.compcode = b.gtcompmastid  join asptblusermas c on c.userid = a.username   join asptblmenuname d on d.menunameid=a.menunameid  where b.compcode = '" + s + "' and c.username  = '" + ss + "'  order by a.menuid asc ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblnavigation");
            DataTable dt1 = ds.Tables["asptblnavigation"];
            Class.Users.Query = "";
            Class.Users.Query = sel;
            return dt1;
        }
        public DataTable select(string s, string ss,string sss)
        {
            string sel = "select a.menuid as menuid1 ,d.menuname as menuname1 ,a.navurl,a.parentmenuid,a.active,a.menunameid as menunameid1,b.gtcompmastid ,c.userid from asptblnavigation a  join gtcompmast b on a.compcode = b.gtcompmastid  join asptblusermas c on c.userid = a.username   join asptblmenuname d on d.menunameid=a.menunameid  where b.compcode = '" + s + "' and c.username  = '" + ss + "' and d.menuname LIKE'%" + sss + "%'   order by 2 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblnavigation");
            DataTable dt1 = ds.Tables["asptblnavigation"];
            Class.Users.Query = "";
            Class.Users.Query = sel;
            return dt1;
        }
    }
}

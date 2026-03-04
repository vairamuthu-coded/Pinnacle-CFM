using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Reflection;
using System.Data;
using System.Text;
using System.Security.Cryptography;
using System.IO;
using Pinnacle.Models;
namespace Pinnacle.Models
{

    
    public class Menu
    {
       
        public Menu()
        {
        }
        public string msg { get; set; }
        public string MENUNAMEID { get; set; }
        public string MENUNAME { get; set; }
        public string ACTIVE { get; set; }
        public bool chk { get; set; }
        public string PARENTMENUID = "0";
      
        public DataTable select(string MENUNAMEID, string MENUNAME, string ACTIVE, string PARENTMENUID)
        {
            this.MENUNAMEID = MENUNAMEID;
            this.MENUNAME = MENUNAME;
            this.ACTIVE = ACTIVE;
            this.PARENTMENUID = PARENTMENUID;
            string sel = "select * from asptblmenuname a where a.menunameid='" + MENUNAMEID + "' and  a.menuname='" + MENUNAME + "' and a.active='" + ACTIVE + "' and a.parentmenuid='" + PARENTMENUID + "'   ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblmenuname");
            DataTable dt = ds.Tables["asptblmenuname"];
            return dt;
        }
        public Menu(string MENUNAME, string ACTIVE, string MENUNAMEID, string PARENTMENUID)
        {
            this.MENUNAMEID = MENUNAMEID;
            this.MENUNAME = MENUNAME;
            this.ACTIVE = ACTIVE;
            this.PARENTMENUID = PARENTMENUID;
            string up = "update  asptblmenuname a set   a.menuname='" + MENUNAME + "' , a.active='" + ACTIVE + "' , a.parentmenuid='" + PARENTMENUID + "' where a.menunameid='" + MENUNAMEID + "'";
            Utility.ExecuteNonQuery(up);

        }

        public Menu(string MENUNAME, string ACTIVE, string PARENTMENUID)
        {
            this.MENUNAME = MENUNAME;
            this.ACTIVE = ACTIVE;
            this.PARENTMENUID = PARENTMENUID;
            string ins = "insert into asptblmenuname a (a.menuname,a.active,a.parentmenuid) values('" + MENUNAME.ToUpper() + "','" + ACTIVE + "', '" + PARENTMENUID + "')";
            Utility.ExecuteNonQuery(ins);
        }
        public DataTable select()
        {
            string sel = " select a.menunameid,a.menuname,a.active,a.parentmenuid from asptblmenuname a where a.active='T' order by 1 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblmenuname");
            DataTable dt = ds.Tables["asptblmenuname"];
            return dt;
        }

    }

    public class SITEMENU
    {
    
        
           
            private string p1;
            private string p2;
            private int p3;
            private string p4;
            private string p5;
            private string p6;
            private string p7;
            private string p8;
            public string MenuID { get; set; }
            public string MenuNameID { get; set; }
            public string MenuName { get; set; }
            public string NavURL { get; set; }
            public int ParentMenuID { get; set; }
            public string CompCode { get; set; }
            public string UserName { get; set; }
            public bool Chk { get; set; }
            public string Active { get; set; }
            public SITEMENU()
            {


            }
            public string Encrypt(string str)
            {
                string EncrptKey = "2013;[pnuLIT)WebCodeExpert";
                byte[] byKey = { };
                byte[] IV = { 18, 52, 86, 120, 144, 171, 205, 239 };
                byKey = System.Text.Encoding.UTF8.GetBytes(EncrptKey.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                byte[] inputByteArray = Encoding.UTF8.GetBytes(str);
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write);
                cs.Write(inputByteArray, 0, inputByteArray.Length);
                cs.FlushFinalBlock();
                return Convert.ToBase64String(ms.ToArray());
            }

            public string Decrypt(string str)
            {
                str = str.Replace(" ", "+");
                string DecryptKey = "2013;[pnuLIT)WebCodeExpert";
                byte[] byKey = { };
                byte[] IV = { 18, 52, 86, 120, 144, 171, 205, 239 };
                byte[] inputByteArray = new byte[str.Length];

                byKey = System.Text.Encoding.UTF8.GetBytes(DecryptKey.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                inputByteArray = Convert.FromBase64String(str.Replace(" ", "+"));
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write);
                cs.Write(inputByteArray, 0, inputByteArray.Length);
                cs.FlushFinalBlock();
                System.Text.Encoding encoding = System.Text.Encoding.UTF8;
                return encoding.GetString(ms.ToArray());
            }
            public SITEMENU(string p1, string p2, int p3, string p4, string p5, string p6, string p7)
            {

                this.p1 = MenuName;
                this.p2 = NavURL;
                this.p3 = ParentMenuID;
                this.p4 = Active;
                this.p5 = MenuNameID;
                this.p6 = CompCode;
                this.p7 = UserName;
                string ins = "insert into  SITEMENU A (A.MENUNAME,A.NAVURL,A.PARENTMENUID,A.ACTIVE,A.MENUNAMEID,A.COMPCODE,A.USERNAME)values( '" + p1 + "' , '" + p2 + "' , '" + p3 + "' , '" + p4 + "' , '" + p5 + "' , '" + p6 + "' , '" + p7 + "'  )";
                Utility.ExecuteNonQuery(ins);
            }

            public SITEMENU(string p1, string p2, int p3, string p4, string p5, string p6, string p7, string p8)
            {

                this.p1 = MenuName;
                this.p2 = NavURL;
                this.p3 = ParentMenuID;
                this.p4 = Active;
                this.p5 = MenuNameID;

                this.p6 = CompCode;
                this.p7 = UserName;
                this.p8 = MenuID;
                string up = "update SITEMENU A set   A.MENUNAME='" + p1 + "' , A.NAVURL='" + p2 + "' , A.PARENTMENUID='" + p3 + "' , A.ACTIVE='" + p4 + "' , A.MENUNAMEID='" + p5 + "' ,A.COMPCODE='" + p6 + "',A.USERNAME='" + p7 + "' where A.MENUID='" + p8 + "' ";
                Utility.ExecuteNonQuery(up);
            }




            internal DataTable select(string p1, string p2, int p3, string p4, string p5, string p6, string p7, string p8)
            {
                this.p1 = MenuName;
                this.p2 = NavURL;
                this.p3 = ParentMenuID;
                this.p4 = Active;
                this.p5 = MenuNameID;
                this.p6 = CompCode;
                this.p7 = UserName;
                this.p8 = MenuID;
                string sel = "SELECT  A.MENUID,A.MENUNAME,A.NAVURL,A.PARENTMENUID,A.ACTIVE,A.MENUNAMEID FROM SITEMENU A WHERE  A.MENUNAME='" + p1 + "' AND A.NAVURL='" + p2 + "' AND A.PARENTMENUID='" + p3 + "' AND A.ACTIVE='" + p4 + "' AND A.MENUNAMEID='" + p5 + "' AND A.COMPCODE='" + p6 + "' AND A.USERNAME='" + p7 + "' AND A.MENUID='" + p8 + "'  ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "SITEMENU");
                DataTable dt1 = ds.Tables["SITEMENU"];
                return dt1;
            }
            internal System.Data.DataTable Ordertreefilter(int s, int ss)
            {


                string sel = "SELECT   A.MENUID,A.MENUNAME,A.NAVURL,A.PARENTMENUID FROM ASPTBLUSERRIGHTS A WHERE A.COMPCODE='" + s + "' AND A.USERNAME='" + ss + "' AND A.ACTIVE='T'   AND A.PARENTMENUID >0   ORDER BY 2 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLUSERRIGHTS");
                DataTable dt1 = ds.Tables["ASPTBLUSERRIGHTS"];
                return dt1;
            }
            internal System.Data.DataTable Imagefilter1()
            {
                string sel = " SELECT A.TBLVISTORINFOID, A.PHONENO,A.EMPNAME ,A.IMAGEPATH,A.INTIME  FROM TBLVISTORINFO   A   WHERE A.SACTIVE='F' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "TBLVISTORINFO");
                DataTable dt1 = ds.Tables["TBLVISTORINFO"];
                return dt1;
            }
            internal System.Data.DataTable Imagefilter1(string d)
            {

                string sel = " SELECT A.asptblvisinfoID, A.PHONE,A.FIRSTNAME,A.IMAGEPATH,A.STARTDATE  FROM asptblvisinfo A  where A.ACTIVE='T' AND A.IMAGETITLE='" + d + "' ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblvisinfo");
                DataTable dt1 = ds.Tables["asptblvisinfo"];
                return dt1;
            }
            public DataTable treefilter(int s, int ss)
            {
                string sel = "SELECT A.MENUID,A.MENUNAME,A.NAVURL,A.PARENTMENUID FROM ASPTBLUSERRIGHTS A WHERE A.COMPCODE='" + s + "' AND A.USERNAME='" + ss + "' AND A.ACTIVE='T' ORDER BY 4,2 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLUSERRIGHTS");
                DataTable dt1 = ds.Tables["ASPTBLUSERRIGHTS"];
                return dt1;

            }
            public DataTable treefilter()
            {
                string sel = "SELECT A.MenuID,IFNULL(A.ParentMenuID,0) As ParentMenuID,A.MenuName,IFNULL(B.FormName,'') As FormName,IFNULL(A.FormID,0) As FormID FROM Mas_Menu A LEFT JOIN Mas_Forms B ON B.FormID = A.FormID WHERE A.Active ='Y' ORDER BY A.MenuID,ParentMenuID ASC";
                // string sel = "SELECT A.MENUID,A.MENUNAME,A.NAVURL,A.PARENTMENUID FROM ASPTBLUSERRIGHTS A WHERE A.COMPCODE='" + s + "' AND A.USERNAME='" + ss + "' AND A.ACTIVE='T' ORDER BY 4,2 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "Mas_Menu");
                DataTable dt1 = ds.Tables["Mas_Menu"];
                return dt1;

            }
            public DataTable UserFilter(int s, int ss)
            {

                string sel = "SELECT A.MENUID,A.MENUNAME,A.NAVURL,A.PARENTMENUID FROM SITEMENU A WHERE A.COMPCODE='" + s + "' AND A.USERNAME='" + ss + "' AND A.ACTIVE='T' ORDER BY 3";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "SITEMENU");
                DataTable dt1 = ds.Tables["SITEMENU"];
                return dt1;

            }
            public DataTable headerdropdowns(int s, int ss)
            {
                string sel = "SELECT DISTINCT A.MENUNAME,A.NAVURL,A.ACTIVE,A.SAVES,A.PRINTS,A.READONLY,A.SEARCH,A.DELETES FROM ASPTBLUSERRIGHTS A WHERE A.COMPCODE='" + s + "' AND A.USERNAME='" + ss + "' AND A.ACTIVE='T'  AND A.PARENTMENUID >0  ORDER BY 4,1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLUSERRIGHTS");
                DataTable dt1 = ds.Tables["ASPTBLUSERRIGHTS"];
                return dt1;
            }


        }

    }
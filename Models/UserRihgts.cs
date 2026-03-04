using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Collections;
using System.Text;
using System.IO;
using Pinnacle.Models;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace Pinnacle.Models
{
    
    public class UserRights
    {

        public  UserRights()
        {

        }

        public Int64 UserRightsID;
        public Int64 MenuID;
        public string MenuName;
        public string NavURL;
        public string FormName;
        public Int64 FormID;
        public Int64 ParentMenuID;
        public string Active;
        public bool Activechk;
        public string News;
        public bool Newchk;
        public string Save;
        public bool Savechk;
        public string Print;
        public bool Printschk;
        public string ReadOnly;
        public bool ReadOnlychk;
        public string Search;
        public bool Searchchk;
        public string Delete;
        public bool Deletechk;
        public string TreeButton;
        public bool TreeButtonchk;
        public string GlobalSearch;
        public bool GlobalSearchchk;
        public string Login;
        public bool Loginchk;
        public string ChangePassword;
        public bool ChangePasswordchk;
        public string ChangeSkin;
        public bool ChangeSkinchk;
        public string DownLoad;
        public bool DownLoadchk;
        public string Contact;
        public bool Contactchk;
        public string Pdf;
        public bool Pdfchk;
        public string Imports;
        public bool Importschk;

        public Int64 CompCode;
        public Int64 UserName;
        public int Sno;
        public string IpAddress;
        public string Createdby;
        public DateTime Createdon;
        public DateTime Modified;
        

        public UserRights(Int64 menuID, string menuName, string navURL, Int64 parentMenuID, string active, string news, string save, string print, string readOnly, string search, string delete, Int64 compCode, Int64 userName, string treeButton, string globalSearch, string login, string changePassword, string changeSkin, string downLoad, string contact, string pdf, string imports,int sno,string ipaddress)
        {
            MenuID = menuID;
            MenuName = menuName;
            NavURL = navURL;
            ParentMenuID = parentMenuID;
            Active = active;
            News = news;
            Save = save;
            Print = print;
            ReadOnly = readOnly;
            Search = search;
            Delete = delete;
            CompCode = compCode;
            UserName = userName;
            TreeButton = treeButton;
            GlobalSearch = globalSearch;
            Login = login;
            ChangePassword = changePassword;
            ChangeSkin = changeSkin;
            DownLoad = downLoad;
            Contact = contact;
            Pdf = pdf;
            Imports = imports;
            Sno = sno;
            IpAddress = ipaddress;
            Class.Users.Query = "insert into asptbluserrights(menuid ,  menuname,  navurl,  parentmenuid ,  active,news, saves, prints,  readonly ,  search ,  deletes,compcode,username,sno, treebutton,globalsearch,login,changepassword, changeskin,download,contact,pdf,imports,sno,ipaddress ) values('" + MenuID + "','" + MenuName + "','" + NavURL + "','" + ParentMenuID + "','" + Active + "','" + News + "','" + Save + "','" + Print + "','" + ReadOnly + "','" + Search + "','" + Delete + "','" + CompCode + "','" + UserName + "','" + Sno + "', '" + TreeButton + "','" + GlobalSearch + "','" + Login + "','" + ChangePassword + "','" + ChangeSkin + "','" + DownLoad + "','" + Contact + "','" + Pdf + "','" + Imports + "'," + Sno + ",'" + IpAddress + "');";
            Utility.ExecuteNonQuery(Class.Users.Query);
        }
        public UserRights(Int64 menuID, string menuName, string navURL, Int64 parentMenuID, string active, string news, string save, string print, string readOnly, string search, string delete, Int64 compCode, Int64 userName, string treeButton, string globalSearch, string login, string changePassword, string changeSkin, string downLoad, string contact, string pdf, string imports,DateTime createdon, DateTime modified, int sno, string ipaddress)
        {
            MenuID = menuID;
            MenuName = menuName;
            NavURL = navURL;
            ParentMenuID = parentMenuID;
            Active = active;
            News = news;
            Save = save;
            Print = print;
            ReadOnly = readOnly;
            Search = search;
            Delete = delete;
            CompCode = compCode;
            UserName = userName;
            TreeButton = treeButton;
            GlobalSearch = globalSearch;
            Login = login;
            ChangePassword = changePassword;
            ChangeSkin = changeSkin;
            DownLoad = downLoad;
            Contact = contact;
            Pdf = pdf;
            Imports = imports;
            Createdon = createdon;
            Modified = modified;
            Sno = sno;
            IpAddress = ipaddress;
            Class.Users.Query = "insert into asptbluserrights(menuid ,  menuname,  navurl,  parentmenuid ,  active,news, saves, prints,  readonly ,  search ,  deletes,compcode,username,treebutton,globalsearch,login,changepassword, changeskin,download,contact,pdf,imports,createdon,modifiedon,sno,ipaddress ) values(" + MenuID + ",'" + MenuName + "','" + NavURL + "'," + ParentMenuID + ",'" + Active + "','" + News + "','" + Save + "','" + Print + "','" + ReadOnly + "','" + Search + "','" + Delete + "'," + CompCode + "," + UserName + ",'" + TreeButton + "','" + GlobalSearch + "','" + Login + "','" + ChangePassword + "','" + ChangeSkin + "','" + DownLoad + "','" + Contact + "','" + Pdf + "','" + Imports + "','" + Convert.ToDateTime(Createdon).ToString() + "','" + Convert.ToDateTime(Modified).ToString() + "'," + Sno + ",'" + IpAddress + "');";
            Utility.ExecuteNonQuery(Class.Users.Query);
        }
        public UserRights(Int64 menuID, string menuName, string navURL, Int64 parentMenuID, string active, string news, string save, string print, string readOnly, string search, string delete, Int64 compCode, Int64 userName, string treeButton, string globalSearch, string login, string changePassword, string changeSkin, string downLoad, string contact, string pdf, string imports, int userRightsID)
        {
            MenuID = menuID;
            MenuName = menuName;
            NavURL = navURL;
            ParentMenuID = parentMenuID;
            Active = active;
            News = news;
            Save = save;
            Print = print;
            ReadOnly = readOnly;
            Search = search;
            Delete = delete;
            CompCode = compCode;
            UserName = userName;
            TreeButton = treeButton;
            GlobalSearch = globalSearch;
            Login = login;
            ChangePassword = changePassword;
            ChangeSkin = changeSkin;
            DownLoad = downLoad;
            Contact = contact;
            Pdf = pdf;
            Imports = imports;
            UserRightsID = userRightsID;
            Class.Users.Query = "update  asptbluserrights set menuid='" + MenuID + "' ,  menuname='" + MenuName + "' ,  navurl='" + NavURL + "' ,  parentmenuid='" + ParentMenuID + "'  ,  active='" + Active + "', news='" + News + "' , saves='" + Save + "' ,  print64s='" + Print + "' ,  readonly='" + ReadOnly + "'  ,  search='" + Search + "'  ,  deletes='" + Delete + "' , compcode='" + CompCode + "' , username='" + UserName + "' ,sno='" + Sno + "', treebutton='" + TreeButton + "',globalsearch='" + GlobalSearch + "',login='" + Login + "',changepassword='" + ChangePassword + "', changeskin='" + ChangeSkin + "',download='" + DownLoad + "',contact='" + Contact + "',pdf='" + Pdf + "',imports='" + Imports + "' where  userrightsid='" + UserRightsID + "'; ";
            Utility.ExecuteNonQuery(Class.Users.Query);
        }


        public UserRights(long menuID, string menuName, string navURL, long parentMenuID, string active, string news, string save, string print, string readOnly, string search, string delete, long compCode, long userName, string treeButton, string globalSearch, string login, string changePassword, string changeSkin, string downLoad, string contact, string pdf, string imports, DateTime modified, long userRightsID)
        {
            MenuID = menuID;
            MenuName = menuName;
            NavURL = navURL;
            ParentMenuID = parentMenuID;
            Active = active;
            News = news;
            Save = save;
            Print = print;
            ReadOnly = readOnly;
            Search = search;
            Delete = delete;
            CompCode = compCode;
            UserName = userName;
            TreeButton = treeButton;
            GlobalSearch = globalSearch;
            Login = login;
            ChangePassword = changePassword;
            ChangeSkin = changeSkin;
            DownLoad = downLoad;
            Contact = contact;
            Pdf = pdf;
            Imports = imports;
            Modified = modified;
            UserRightsID = userRightsID; 
            string up = "update  asptbluserrights set menuid='" + MenuID + "' ,  menuname='" + MenuName + "' ,  navurl='" + NavURL + "' ,  parentmenuid='" + ParentMenuID + "'  ,  active='" + Active + "', news='" + News + "' , saves='" + Save + "' ,  prints='" + Print + "' ,  readonly='" + ReadOnly + "'  ,  search='" + Search + "'  ,  deletes='" + Delete + "' , compcode='" + CompCode + "' , username='" + UserName + "' ,sno='" + Sno + "', treebutton='" + TreeButton + "',globalsearch='" + GlobalSearch + "',login='" + Login + "',changepassword='" + ChangePassword + "', changeskin='" + ChangeSkin + "',download='" + DownLoad + "',contact='" + Contact + "',pdf='" + Pdf + "',imports='" + Imports + "', modifiedon='" + Modified.ToString()+"' where  userrightsid=" + UserRightsID+";";
            Utility.ExecuteNonQuery(up);

        }

        public DataTable treefilter()
        {
            Class.Users.Query = "select a.menuid,e.alisname as menuname,a.navurl,a.parentmenuid from asptbluserrights a  join gtcompmast b on a.compcode=b.compcode  and a.active='T'   join asptblnavigation d on d.menuid=a.menuid   join asptblmenuname  e on e.menunameid=d.menunameid  order by 4,2; ";
            DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "asptbluserrights");
            DataTable dt1 = ds.Tables["asptbluserrights"];
            return dt1;

        }
        public DataTable treefilter(string s, string ss)
        {
            Class.Users.Query = "select a.menuid,e.aliasname as menuname,a.navurl,a.parentmenuid from asptbluserrights a  join gtcompmast b on a.compcode=b.compcode   join asptblnavigation d on d.menuid=a.menuid   join asptblmenuname  e on e.menunameid=d.menunameid where b.compcode='" + s + "' and a.username='" + ss + "' and a.active='T' order by 4,2; ";
            DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "asptbluserrights");
            DataTable dt1 = ds.Tables["asptbluserrights"];
            return dt1;

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

        public DataTable headerdropdowns()
        {
            //string sel = "select 0 as menuid, '----select-----'  as menuname from dual union all select  a.menuid,  d.menuname   from  asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username join asptblmenuname d on d.menunameid = a.menuid and c.userid=a.username  where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.active='T'  and a.parentmenuid>1 order by 2";// and a.parentmenuid = 1
            //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            //DataTable dt1 = ds.Tables["asptbluserrights"];

            Class.Users.Query = "select 0 as menuid, '----select-----'  as menuname from dual union all SELECT a.menuid, e.aliasname as  menuname FROM ASPTBLUSERRIGHTS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID=A.COMPCODE JOIN ASPTBLUSERMAS C ON C.USERID=A.USERNAME JOIN ASPTBLNAVIGATION D ON D.MENUID = A.MENUID AND C.USERID=A.USERNAME JOIN ASPTBLMENUNAME E ON E.MENUNAME=D.MENUNAME WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND C.USERNAME='" + Class.Users.HUserName + "' AND A.ACTIVE='T' and a.parentmenuid>1  ORDER BY 1 ";
            DataSet ds0 = Utility.ExecuteSelectQuery(Class.Users.Query, "ASPTBLUSERRIGHTS");
            DataTable dt1 = ds0.Tables["ASPTBLUSERRIGHTS"];
            return dt1;
        }
        public DataTable TreeView(string s, string ss, int id)
        {
            //string sel = "select 0 as menuid, '----select-----'  as menuname from dual union all select  a.menuid,  d.menuname   from  asptbluserrights a  join gtcompmast b on b.gtcompmastid=a.compcode join asptblusermas c on c.userid=a.username join asptblmenuname d on d.menunameid = a.menuid and c.userid=a.username  where  b.compcode='" + Class.Users.HCompcode + "'      and c.username='" + Class.Users.HUserName + "'  and  a.active='T'  and a.parentmenuid>1 order by 2";// and a.parentmenuid = 1
            //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            //DataTable dt1 = ds.Tables["asptbluserrights"];

            Class.Users.Query = "SELECT a.menuid, e.aliasname as  menuname ,a.parentmenuid FROM ASPTBLUSERRIGHTS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID=A.COMPCODE JOIN ASPTBLUSERMAS C ON C.USERID=A.USERNAME JOIN ASPTBLNAVIGATION D ON D.MENUID = A.MENUID AND C.USERID=A.USERNAME JOIN ASPTBLMENUNAME E ON E.MENUNAME=D.MENUNAME WHERE B.COMPCODE='" + s + "' AND C.USERNAME='" +ss + "' and  a.parentmenuid='" + id + "' AND A.ACTIVE='T'   ORDER BY 1 ";
            DataSet ds0 = Utility.ExecuteSelectQuery(Class.Users.Query, "ASPTBLUSERRIGHTS");
            DataTable dt1 = ds0.Tables["ASPTBLUSERRIGHTS"];
            return dt1;
        }
        public DataTable headerdropdowns(string s, string ss)
        {
            Class.Users.Query = " SELECT A.USERRIGHTSID,D.MENUID,E.aliasname as  MENUNAME,D.NAVURL ,A.parentmenuid,A.ACTIVE,A.SAVES,A.PRINTS,A.READONLY,A.SEARCH,A.DELETES,A.NEWS,A.TREEBUTTON,A.GLOBALSEARCH,A.LOGIN,A.CHANGEPASSWORD,A.CHANGESKIN,A.DOWNLOAD,A.CONTACT,A.PDF,A.IMPORTS,B.COMPCODE,C.USERNAME FROM ASPTBLUSERRIGHTS A JOIN GTCOMPMAST B ON   B.gtcompmastid = A.COMPCODE JOIN asptblusermas C ON C.USERID = A.USERNAME JOIN ASPTBLNAVIGATION D ON D.MENUID=A.MENUID    JOIN asptblmenuname e on E.MENUNAMEID=D.MENUNAMEID   WHERE B.COMPCODE = '" + s + "' AND C.USERNAME = '" + ss + "' order by 1";
            DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "ASPTBLUSERRIGHTS");
            DataTable dt1 = ds.Tables["ASPTBLUSERRIGHTS"];

            //string sel = " select a.userrightsid,a.menuid as menuid1,E.menuname as menuname1,d.navurl ,a.parentmenuid,a.active,a.news as new1,a.saves as save1,a.prints as print1,a.readonly,a.search,a.deletes as delete1,a.treebutton,a.globalsearch,a.login,a.changepassword,a.changeskin,a.download,a.contact,a.pdf,a.imports as import1,b.gtcompmastid as compcode, c.userid as username from asptbluserrights a join gtcompmast b on   b.gtcompmastid = a.compcode join asptblusermas c on c.userid = a.username join asptblnavigation d on d.menuid = a.menuid and c.userid=a.username join asptblmenuname e on e.menunameid=d.menunameid   where b.compcode = '" + s + "' and c.username = '" + ss + "' order by 1,2";
            //DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            //DataTable dt1 = ds.Tables["asptbluserrights"];
            return dt1;
        }
        //public DataTable headerdropdownsAlias(string s, string ss, string sss)
        //{
        //    try
        //    {
        //        Class.Users.Query = "SELECT  E.MENUNAME,e.aliasname, A.NAVURL,A.ACTIVE,A.NEWS, A.SAVES,A.PRINTS,A.READONLY,A.SEARCH,A.DELETES,A.TREEBUTTON,A.GLOBALSEARCH,A.LOGIN,A.CHANGEPASSWORD,  A.CHANGESKIN,A.DOWNLOAD,A.PDF,A.IMPORTS,A.COMPCODE,A.USERNAME FROM ASPTBLUSERRIGHTS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID=A.COMPCODE JOIN ASPTBLUSERMAS C ON C.USERID=A.USERNAME JOIN ASPTBLNAVIGATION D ON D.MENUID = A.MENUID AND C.USERID=A.USERNAME JOIN ASPTBLMENUNAME E ON E.MENUNAMEID=D.MENUNAMEID WHERE B.COMPCODE='" + s + "' AND C.USERNAME='" + ss + "' AND A.ACTIVE='T' AND E.aliasname='" + sss + "'    ORDER BY 1 ";
        //        DataSet ds0 = Utility.ExecuteSelectQuery(Class.Users.Query, "ASPTBLUSERRIGHTS");
        //        DataTable dt0 = ds0.Tables["ASPTBLUSERRIGHTS"];
        //        if (dt0 == null)
        //        {

        //            GlobalVariables.Toolstrip1.Visible = false;

        //        }
        //        return dt0;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}
        public DataTable headerdropdowns(string s, string ss, string sss)
        {
            try
            {
                Class.Users.Query = "SELECT  e.aliasname as MENUNAME, A.NAVURL,A.ACTIVE,A.NEWS, A.SAVES,A.PRINTS,A.READONLY,A.SEARCH,A.DELETES,A.TREEBUTTON,A.GLOBALSEARCH,A.LOGIN,A.CHANGEPASSWORD,  A.CHANGESKIN,A.DOWNLOAD,A.PDF,A.IMPORTS,A.COMPCODE,A.USERNAME FROM ASPTBLUSERRIGHTS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID=A.COMPCODE JOIN ASPTBLUSERMAS C ON C.USERID=A.USERNAME JOIN ASPTBLNAVIGATION D ON D.MENUID = A.MENUID AND C.USERID=A.USERNAME JOIN ASPTBLMENUNAME E ON E.MENUNAMEID=D.MENUNAMEID WHERE B.COMPCODE='" + s + "' AND C.USERNAME='" + ss + "' AND A.ACTIVE='T' AND E.aliasname='" + sss + "'    ORDER BY 1 ";
                DataSet ds0 = Utility.ExecuteSelectQuery(Class.Users.Query, "ASPTBLUSERRIGHTS");
                DataTable dt0 = ds0.Tables["ASPTBLUSERRIGHTS"];
                if (dt0 == null)
                {

                    GlobalVariables.Toolstrip1.Visible = false;

                }
                return dt0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable headerdropdowns(string s, string ss, Int64 sss)
        {
            try
            {
                //change-12-10
               // string f = "SELECT A.menuid,a.parentmenuid, e.MENUNAME FROM ASPTBLUSERRIGHTS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID=A.COMPCODE JOIN ASPTBLUSERMAS C ON C.USERID=A.USERNAME JOIN ASPTBLNAVIGATION D ON D.MENUID = A.MENUID AND C.USERID=A.USERNAME JOIN ASPTBLMENUNAME E ON E.MENUNAMEID=D.MENUNAMEID  WHERE B.COMPCODE='" + s + "' AND C.USERNAME='" + ss + "' AND A.ACTIVE='T' AND a.PARENTMENUID='" + sss + "'    ORDER BY 3 ";

                string f = "SELECT A.menuid,a.parentmenuid, e.aliasname as MENUNAME FROM ASPTBLUSERRIGHTS A JOIN GTCOMPMAST B ON B.GTCOMPMASTID=A.COMPCODE JOIN ASPTBLUSERMAS C ON C.USERID=A.USERNAME JOIN ASPTBLNAVIGATION D ON D.MENUID = A.MENUID AND C.USERID=A.USERNAME JOIN ASPTBLMENUNAME E ON E.MENUNAMEID=D.MENUNAMEID  WHERE B.COMPCODE='" + s + "' AND C.USERNAME='" + ss + "' AND A.ACTIVE='T' AND a.PARENTMENUID='" + sss + "'    ORDER BY 3 ";
                DataSet ds0 = Utility.ExecuteSelectQuery(f, "ASPTBLUSERRIGHTS");
                DataTable dt0 = ds0.Tables["ASPTBLUSERRIGHTS"];

                return dt0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable userid()
        {
            DataTable dt1 = new DataTable();
            //if (Class.Users.HCompcode == "SCTS")
            //{
            //    string sel = "  select distinct b.userid,b.username   from  asptblschoolmas a join  asptblusermas b on a.asptblschoolmasid= b.compcode  ";
            //    DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
            //    dt1 = ds1.Tables["asptblusermas"];
            //}
            //else
            //{
                string sel = "  select distinct b.userid,b.username   from  gtcompmast a join  asptblusermas b on a.gtcompmastid= b.compcode  ";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                 dt1 = ds1.Tables["asptblusermas"];
           // }
            return dt1;
        }
        public DataTable userid(string id)
        {
            DataTable dt1 = new DataTable();
            //if (Class.Users.HCompcode == "SCTS")
            //{
            //    string sel = "  select distinct b.userid,b.username   from  asptblschoolmas a join  asptblusermas b on a.asptblschoolmasid= b.compcode where a.compcode='" + id + "'  ";
            //    DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
            //    dt1 = ds1.Tables["asptblusermas"];
            //}
            //else
            //{
                string sel = "  select distinct b.userid,b.username   from  gtcompmast a join  asptblusermas b on a.gtcompmastid= b.compcode where a.compcode='" + id + "'  ";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                dt1 = ds1.Tables["asptblusermas"];
            //}
            return dt1;
        }
      
        public DataTable userid(string id,string idd)
        {
            DataTable dt1 = new DataTable();
            //if (Class.Users.HCompcode == "SCTS")
            //{
            //    string sel = "select distinct b.userid,b.username   from  asptblschoolmas a join  asptblusermas b on a.asptblschoolmasid= b.compcode where a.compcode='" + id + "' and b.username='" + idd + "' ";
            //    DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
            //    dt1 = ds1.Tables["asptblusermas"];
            //}
            //else
            //{
                string sel = "select distinct b.userid,b.username   from  gtcompmast a join  asptblusermas b on a.gtcompmastid= b.compcode where a.compcode='" + id + "' and b.username='" + idd + "' ";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                 dt1 = ds1.Tables["asptblusermas"];
            //}
            return dt1;
        }
        public DataTable passid(string id)
        {

            string sel = "select  distinct b.pasword from  gtcompmast a join  asptblusermas b on a.gtcompmastid= b.compcode where a.username='" + id + "'  ";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            DataTable dt1 = ds1.Tables["asptbluserrights"];
            return dt1;
        }

        internal DataTable select(Int64 menuID, string menuName,string navurl, Int64 parentMenuID,string active, string news, string save, string print, string readOnly, string search, string delete, Int64 compCode, Int64 userName, string treeButton, string globalSearch, string login, string changePassword, string changeSkin, string downLoad, string contact, string Pdf, string Imports)
        {
            string sel = " select   userrightsid  from  asptbluserrights  where  menuid=" + menuID + " and menuname='" + menuName + "' and navurl='" + navurl + "' and parentmenuid=" + parentMenuID + " and  active='" + active + "'   and  news='" + news + "' and  saves='" + save + "'  and  prints='" + print + "'  and readonly='" + readOnly + "'  and  search='" + search + "'   and  deletes='" + delete + "' and  compcode=" + compCode + "  and username=" + userName + " and treebutton='" + treeButton + "' and globalsearch='" + globalSearch + "' and login='" + Login + "' and changepassword='" + changePassword + "' and changeskin='" + changeSkin + "' and download='" + downLoad + "' and contact='" + contact + "' and pdf='" + Pdf + "' and imports='" + Imports + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            DataTable dt = ds.Tables["asptbluserrights"];
            return dt;
        }
        internal DataTable select(Int64 menuID, string menuName,Int64 parentMenuID,Int64 compCode, Int64 userName)
        {
            string sel = " select userrightsid  from  asptbluserrights  where  menuid=" + menuID + " and menuname='" + menuName + "' and parentmenuid=" + parentMenuID + " and  compcode=" + compCode + "  and username=" + userName;
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbluserrights");
            DataTable dt = ds.Tables["asptbluserrights"];
            return dt;
        }

        //internal DataTable select(Int64 menuID, string menuName, Int64 parentMenuID, string active, string news, string save, string print, string search, string delete,  string treeButton, string globalSearch, string login, string changePassword, string changeSkin, string downLoad, string contact, string Pdf, string Imports)
        //{
        //    string sel = " SELECT   USERRIGHTSID,COMPCODE,MENUNAME  FROM  ASPTBLUSERRIGHTS  WHERE  MenuID=" + menuID + " AND MenuName='" + menuName + "' AND ParentMenuID=" + parentMenuID + " AND  ACTIVE='" + active + "'   AND  NEWS='" + news + "' AND  SAVES='" + save + "'  AND  PRINTS='" + print + "' AND  SEARCH='" + search + "'   AND  DELETES='" + delete + "'  AND TREEBUTTON='" + treeButton + "' AND GLOBALSEARCH='" + globalSearch + "' AND LOGIN='" + Login + "' AND CHANGEPASSWORD='" + changePassword + "' AND CHANGESKIN='" + changeSkin + "' AND DOWNLOAD='" + downLoad + "' AND CONTACT='" + contact + "' AND Pdf='" + Pdf + "' AND Imports='" + Imports + "'";
        //    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLUSERRIGHTS");
        //    DataTable dt = ds.Tables["ASPTBLUSERRIGHTS"];
        //    return dt;
        //}

    }
}
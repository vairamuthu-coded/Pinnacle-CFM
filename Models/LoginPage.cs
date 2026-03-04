using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Collections;
using System.Text;
using System.IO;
namespace Pinnacle.Models
{
    public class LoginPage
    {

        public LoginPage()
        {
            // TODO: Complete member initialization
        }

        public int LoginID { get; set; }
        public int CompCode { get; set; }
        public int UserName { get; set; }
        public string Password { get; set; }
        public string Active { get; set; }
        public bool Activechk { get; set; }
        public string IpAddress { get; set; }
        public string CreatedOn { get; set; }
        public LoginPage(int CompCode, int UserName, string Password, string Active)
        {
            string ins = "INSERT INTO AVALOGIN(COMPCODE, USERNAME,PASSWORD,ACTIVE) VALUES('" + CompCode + "'  , " + UserName + "  ,  '" + Password + "'  , '" + Active + "'  ) ";
            Utility.ExecuteNonQuery(ins);

        }

        public LoginPage(int CompCode, int UserName, string Password, string Active, int LoginID)
        {
            string up = "UPDATE AVALOGIN SET  COMPCODE ='" + CompCode + "'  , USERNAME =" + UserName + "  ,  PASSWORD='" + Password + "'  , ACTIVE='" + Active + "' WHERE  LOGINID='" + LoginID + "' ";
            Utility.ExecuteNonQuery(up);
        }
        internal System.Data.DataTable Select(int CompCode, int UserName, string Password, string Active, int LoginID)
        {
            string sel = "SELECT LOGINID COMPCODE, USERNAME,PASSWORD,ACTIVE FROM avalogin  where LOGINID='" + LoginID + "'  AND  COMPCODE ='" + CompCode + "'  AND USERNAME =" + UserName + "  AND  PASSWORD='" + Password + "'  AND ACTIVE='" + Active + "'  ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "avalogin");
            DataTable dt1 = ds.Tables["avalogin"];
            return dt1;
        }

    }
}
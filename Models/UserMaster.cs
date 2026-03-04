using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Pinnacle.TreeView;

namespace Pinnacle.Models
{
    

        public class UserMaster
        {
      
        public Int64 Userid { get; set; }
            public Int64 FinYear { get; set; }
            public Int64 CompCode { get; set; }
            public string EmpNo { get; set; }
        public Int64 EMPNAME { get; set; }
        public Int64 Dept { get; set; }
            public string UserName { get; set; }
            public string GateName { get; set; }
            public string Password { get; set; }
        public string NewPassword { get; set; }
        public string Active { get; set; }
            public string Ipaddress { get; set; }
            public string Createdon { get; set; }
        public string SessionTime { get; set; }
        public UserMaster(Int64 finYear, Int64 compCode, Int64 eMPNAME, Int64 dept, string userName, string gateName, string password, string active, string ipaddress, string createdon, string sessiontime)
            {
                FinYear = finYear;
                CompCode = compCode;
                EMPNAME = eMPNAME;
                Dept = dept;
                UserName = userName;
                GateName = gateName;
                Password = password;
                Active = active;
                Ipaddress = ipaddress;
                Createdon = createdon;
            SessionTime = sessiontime;
                string ins = "insert into asptblusermas (finyear,  compcode,  empname ,  dept ,  username ,gatename,  pasword ,  active ,  ipaddres,createdon,SessionTime)values('" + FinYear + "','" + CompCode + "'," + EMPNAME + ",'" + Dept + "','" + UserName + "','" + GateName + "','" + Password + "','" + Active + "','" + Ipaddress + "','" + createdon + "','" + SessionTime + "')";
                Utility.ExecuteNonQuery(ins);
            }
        public UserMaster(Int64 finYear, Int64 compCode, Int64 eMPNAME, Int64 dept, string userName, string gateName, string password, string newPassword, string active, string ipaddress, string createdon,string sessiontime)
        {
            FinYear = finYear;
            CompCode = compCode;
            EMPNAME = eMPNAME;
            Dept = dept;
            UserName = userName;
            GateName = gateName;
            Password = password;
            NewPassword = newPassword;
            Active = active;
            Ipaddress = ipaddress;
            Createdon = createdon; SessionTime = sessiontime;
            string ins = "insert into asptblusermas (finyear,  compcode,  empname ,  dept ,  username ,gatename,  pasword ,newpassword,  active ,  ipaddres,createdon)values('" + FinYear + "','" + CompCode + "'," + EMPNAME + ",'" + Dept + "','" + UserName + "','" + GateName + "','" + Password + "','" + NewPassword + "','" + Active + "','" + Ipaddress + "','" + createdon + "','" + SessionTime + "')";
            Utility.ExecuteNonQuery(ins);
        }
        public UserMaster(Int64 finYear, Int64 compCode, Int64 eMPNAME, Int64 dept, string userName, string gateName, string password, string active, string ipaddress, string createdon, string sessiontime, Int64 userid)
        {
            FinYear = finYear;
            CompCode = compCode;
            EMPNAME = eMPNAME;
            Dept = dept;
            UserName = userName;
            GateName = gateName;
            Password = password;
            Active = active;
            Ipaddress = ipaddress;
            Createdon = createdon;
            Userid = userid; SessionTime = sessiontime;
            string up = "update asptblusermas set finyear='" + FinYear + "',  compcode='" + CompCode + "',   empname=" + EMPNAME + ",   dept='" + Dept + "',   username='" + UserName + "', gatename='" + GateName + "',  pasword='" + Password + "',  active='" + Active + "',   ipaddres='" + Ipaddress + "', createdon='" + Createdon + "' ,SessionTime='" + SessionTime + "' where userid='" + Userid + "'";
            Utility.ExecuteNonQuery(up);
        }

        public UserMaster()
        {
        }

        public UserMaster(long compCode, long userid, string newPassword, string ipaddress, string createdon)
        {
            CompCode = compCode;
            Userid = userid;
            NewPassword = newPassword;
            Ipaddress = ipaddress;
            Createdon = createdon;
            string ins = "insert into asptblusermas (finyear,  compcode,  empname ,  dept ,  username ,gatename,  pasword ,  active ,  ipaddres,createdon)values('" + FinYear + "','" + CompCode + "'," + EMPNAME + ",'" + Dept + "','" + UserName + "','" + GateName + "','" + Password + "','" + Active + "','" + Ipaddress + "','" + createdon + "')";
            Utility.ExecuteNonQuery(ins);
        }

        internal DataTable Select(Int64 compCode, Int64 eMPNAME, Int64 dept,  string userName, string gateName, string password, string active,string sessiontime)
            {
                string sel = "select  username  from asptblusermas where   compcode='" + compCode + "' AND EMPNAME='" + eMPNAME + "' AND DEPT='" + dept + "' and  username='" + userName + "'  and  gatename='" + gateName + "'  and  pasword='" + password + "'  and  active='" + active + "' and  SessionTime='" + sessiontime + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                DataTable dt = ds.Tables["asptblusermas"];
                return dt;
            }
            public DataTable select()
            {
                string sel = " select a.userid,a.empno,a.username,a.gatename,a.active,a.ipaddres,a.createdon from asptblusermas a  order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
                DataTable dt = ds.Tables["asptblusermas"];
                return dt;
            }

      
    }
}

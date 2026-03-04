using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    class Department
    {
      

        public Int64 DepartmentId { get; set; }      
        public String DepartmentName { get; set; }
        public string Active { get; set; }
        public Int64 UserName { get; set; }
        public String IpAddress { get; set; }
        public String Createdon { get; set; }
        public String Createdby { get; set; }
        public String Modifiedon { get; set; }
        internal DataTable select(string departmentName, string active)
        {
            string sel = "select ASPTBLDEPID from ASPTBLDEP WHERE DEPARTMENT='" + DepartmentName + "' AND ACTIVE='" + Active + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel,"ASPTBLDEP");
            DataTable dt = ds.Tables["ASPTBLDEP"];
            return dt;
        }
        internal DataTable select()
        {
            string sel = "select gtdeptdesgmastid,DEPARTMENT,ACTIVE from gtdeptdesgmast ORDER BY 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtdeptdesgmast");
            DataTable dt = ds.Tables["gtdeptdesgmast"];
            return dt;
        }
        public Department(string departmentName, string active, Int64 userName, string ipAddress, string createdby,string modifiedon, long departmentId)
        {
            DepartmentName = departmentName;
            Active = active;
            UserName = userName;
            IpAddress = ipAddress;
            Modifiedon = modifiedon;          
            Createdby = createdby;
            DepartmentId = departmentId;
            string up = "UPDATE ASPTBLDEP A SET A.DEPARTMENT='" + DepartmentName + "',A.ACTIVE='" + Active + "',A.USERNAME=" + UserName + ",A.IPADDRESS='" + IpAddress + "',A.CREATEDBY='" + Createdby + "',A.MODIFIEDON='" + Modifiedon + "' WHERE A.ASPTBLDEPID=" + DepartmentId;
            Utility.ExecuteNonQuery(up);
        }
        public Department(string departmentName, string active, Int64 userName, string ipAddress,string createdon, string createdby,string modifiedon)
        {
            DepartmentName = departmentName;
            Active = active;
            UserName = userName;
            IpAddress = ipAddress;
            Modifiedon = modifiedon;
            Createdon = createdon;
            Createdby = createdby;
            string ins = "INSERT INTO ASPTBLDEP A(A.DEPARTMENT,A.ACTIVE,A.USERNAME,A.IPADDRESS,A.CREATEDON,A.CREATEDBY,A.MODIFIEDON)VALUES('" + DepartmentName+"','"+Active+"',"+UserName+",'"+IpAddress+"','" + Createdon + "','" + Createdby + "','" + Modifiedon+"')";
            Utility.ExecuteNonQuery(ins);
        }

        public Department()
        {
        }
    }
}

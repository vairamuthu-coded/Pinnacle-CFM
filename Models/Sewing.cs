using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models
{
    class Sewing
    {
        public DataTable section()
        {
            string sel = "select  asptblsecmasid  ,section  from   asptblsecmas";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblsecmas");
            DataTable dt = new DataTable();
            dt = ds.Tables["asptblsecmas"];
            return dt;
        }
        public DataTable StyleGroup()
        {
            string sel = "select c.ASPTBLSTYMASID,c.STYLENAME from asptblstylegroupmas a join asptblstylegroupdet b on a.asptblstylegroupmasid=b.asptblstylegroupmasid join asptblstymas c on c.asptblstymasid=b.stylename and c.active='T'   order by 2 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblstymas");
            DataTable dt = ds.Tables["asptblstymas"];
            return dt;
        }
        public DataTable StyleGroup(string s)
        {
            string sel = "select c.ASPTBLSTYMASID,c.STYLENAME from asptblstylegroupmas a join asptblstylegroupdet b on a.asptblstylegroupmasid=b.asptblstylegroupmasid join asptblstymas c on c.asptblstymasid=b.stylename and c.active='T'  where a.stylegroup='" + s + "' order by 2 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblstymas");
            DataTable dt = ds.Tables["asptblstymas"];
            return dt;
        }
        public DataTable Buyer()
        {
            string sel = " select distinct a.asptblbuymasid,a.buyername from  asptblbuymas  a where a.active='T'  order by 1 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblbuymas");
            DataTable dt = ds.Tables["asptblbuymas"];
            return dt;
        }
        public DataTable Group()
        {

            string sel = " select distinct a.asptblstylegroupmasid,a.stylegroup from  asptblstylegroupmas  a where a.active='T' order by 1 ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblstylegroupmas");
            DataTable dt = ds.Tables["asptblstylegroupmas"];
            return dt;
        }
        public DataTable Line()
        {
            string sel = "select A.asptbllinemasid,A.lineno from asptbllinemas a where a.active='T' ORDER BY 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllinemas");
            DataTable dt = ds.Tables["asptbllinemas"];
            return dt;
        }
        public DataTable shift()
        {
            string sel = "select A.asptblshitypeid,A.shiftno from asptblshitype a where a.active='T' ORDER BY 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblshitype");
            DataTable dt = ds.Tables["asptblshitype"];
            return dt;
        }
    }
}

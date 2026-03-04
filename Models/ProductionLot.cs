using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{

    class ProductionLot:CommonClass
    {
        public Int64 asptblprolotid;
        public Int64 asptblprolot1id;
        public string shortcode;
        public string finyear;
        public string wdate;
        public string prodno;
        public Int64 compcode;
        public Int64 orderqty;
        public Int64 stylename;
        public Int64 buyer;
        public string pono;
        public string lotno;
        public string bundle;
        public string size;
        public Int64 processname;
        public string processtype;
        public string productioncancel;
        public string issuetype;
        public string stitching;
        public string inward;
        public string delivery;
        public string restitching;
        public string rechecking;
        public string panelmistake;
        public string active;
        public Int64 compcode1;
        public Int64 username;
        public string createdon;
        public string createdby;
        public string modified;
        public string modifiedby;
        public string ipaddress;
        public string inward1;
        public string rework1;
        public string delivery1;
        public string notes;
        public override string HideButton(string pono, string tbl)
        {
            Class.Users.Query = "select distinct a.pono from " + tbl + " a  where a.pono='" + pono + "'";
            Class.Users.dt2 = select(Class.Users.Query, tbl);
            if (Class.Users.dt2.Rows.Count > 0)
            {
                pono = Class.Users.dt2.Rows[0]["pono"].ToString();
            }
            else
            {
                pono = "";
            }
            return pono;
        }
        internal DataTable select(long asptblprolotid, long asptblprolot1id,long compcode, string finyear, string shortcode, string wdate, string prodno,  long buyer, string pono, long orderqty, long stylename, string lotno, string bundle, string size, long processname, string processtype, string productioncancel, string active,string issuetype, string restitching, string rechecking, string inward, string delivery)
        {
            Class.Users.Query = "select distinct asptblprolotid   from  asptblprolot   where  asptblprolotid='" + asptblprolotid + "' and asptblprolot1id='" + asptblprolot1id + "'  and compcode='" + compcode + "' and finyear='" + finyear + "' and shortcode='"+shortcode+"' and wdate='"+wdate+"' and prodno='"+prodno+"' and buyer='"+buyer+"'  and  pono='" + pono + "'  and  orderqty='" + orderqty + "' and stylename='" + stylename + "' and bundle='" + bundle + "' and lotno = '" + lotno + "' and size='"+size+ "' and processname='"+processname+ "' and processtype='"+processtype+ "' and productioncancel='"+ productioncancel + "' and active='"+active+ "' and issuetype='"+issuetype+ "'  and restitching='" + restitching + "' and rechecking='" + rechecking + "' and inward='" + inward + "' and delivery='" + delivery + "'";// and orderqty = '" + pp.orderqty + "'
            DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "asptblprolot");
            System.Data.DataTable dt = ds.Tables["asptblprolot"];
            return dt;
        }
    }
        public class ProductionLotDet
        {
            public Int64 asptblprolotdetid;
            public Int64 asptblpurdet1id;
        public Int64 barcode;
        public Int64 asptblpurdetid;
            public Int64 asptblpurid;    
            public string colorname;
            public string sizename;
            public string pono;
            public Int64 orderqty;
            public Int64 comqty;
            public Int64 lotqty;
            public Int64 balqty;
            public string notes;
            public string process;
        public string processcheck;
    }
  
}

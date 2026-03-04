using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Text;

namespace Pinnacle.Models
{
    public class PurchasesModel:CommonClass
    {
        public Int64 asptblpurid;
        public Int64 asptblpur1id;
        public string shortcode;
        public string finyear;
        public string podate;
        public Int64 compcode;
        public Int64 compname;
        public Int64 orderqty;
        public Int64 excessqty;
        public Int64 buyer;
        public Int64 bundle;
        public string processtype;
        public string pono;
        public Int64 sizegroup;
        public Int64 stylename;
        public Int64 colorname;
        public string sizename;
        public Int64 processname;
        public string lotno;
        public string pocancel;
        public String issuetype;
        public string panelmistake;
        public Int64 barcode;
        public string active;
        public Int64 compcode1;
        public Int64 username;
        public string createdon;
        public string createdby;
        public string modified;
        public string modified1;
        public string modifiedby;
        public string ipaddress;
        public string orderno;
        public string styleref;
        public override string HideButton(string pono,string tbl)
        {            
            Class.Users.Query = "select distinct a.pono from "+ tbl + " a  where a.pono='" + pono + "'";
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

    }

        public class PurchasesModeldetail:PurchasesModel
        {
            public Int64 asptblpurdetid;
            public Int64 asptblpurid;
            public Int64 asptblpur1id;
            public Int64 compcode;
            public string pono;
            public string colorname;
        public string colorname1;
        public Int64 portion;
            public string sizename;
            public Int64 orderqty;
            public string Remarks;       
    }
        public class PurchasesModel1detail : PurchasesModeldetail
        {
            public Int64 asptblpurdet1id;
            public Int64 asptblpurdetid;           
            public Int64 portion;                    
            public Int64 pcs;
            public Int64 sno;
            public string Remarks;
        public string cutting;
        public string stitching;
        public string checking;
        public string restitching; 
        public string rechecking;
        public string inward;
        public string delivery;
       

        public PurchasesModel1detail()
            {
            
            }

        public PurchasesModel1detail(long barcode, long Asptblpurid, long Compcode, string Pono, string Colorname, long Portion, string Sizename,long Orderqty,string Restitching,string Rechecking,string inward,string delivery,string panelmistake,string modified)
        {
            Class.Users.Query = "update  asptblpurdet1  set barcode='"+barcode+"', panelmistake='" + panelmistake+"', colorname='" + Colorname + "', portion='" + Portion + "', sizename='" + Sizename + "', orderqty='" + Orderqty + "' , restitching='" + Restitching + "' , rechecking='" + Rechecking + "',inward='" + inward + "' , delivery='" + delivery + "', panelmistake='" + panelmistake + "', processcheck='F',ISSUETYPE='CUTTING',modified=date_format('" + modified + "','%Y-%m-%d') where  asptblpurid='" + Asptblpurid + "' AND  compcode='" + Compcode + "'  AND pono='" + Pono + "'  AND colorname='" + Colorname + "' AND sizename='" + Sizename + "'";
            Utility.ExecuteNonQuery(Class.Users.Query);
        }

        internal System.Data.DataTable Select(long barcode, long asptblpurdetid, long asptblpurid, long asptblpur1id, long compcode, string pono, string colorname, long portion, string sizename, long sno,long orderqty)
        {
            this.asptblpurdetid = asptblpurdetid;
            this.asptblpurid = asptblpurid;
            this.asptblpur1id = asptblpur1id;
            this.compcode = compcode;
            this.pono = pono;
            this.colorname = colorname;
            this.portion = portion;
            this.sizename = sizename;
            this.sno = sno;
            this.orderqty = orderqty;
            this.barcode = barcode;
            Class.Users.Query = "select distinct asptblpurdet1id   from  asptblpurdet1   where  barcode='" + barcode + "' and asptblpurdetid='" + asptblpurdetid + "' and asptblpurid='" + asptblpurid + "' AND   asptblpur1id='" + asptblpur1id + "' and compcode='" + compcode + "' and  pono='" + pono + "'  and  colorname='" + colorname + "' and sizename='" + sizename + "' and sno='" + sno + "' and orderqty = '" + orderqty + "'  ";// and orderqty = '" + pp.orderqty + "'
            DataSet ds = Utility.ExecuteSelectQuery(Class.Users.Query, "asptblpurdet1");
            System.Data.DataTable dt = ds.Tables["asptblpurdet1"];
            return  dt;
        }
        //public DataTable PurchasesModel1detail(long asptblpurdetid, long asptblpurid, long asptblpur1id, long compcode, string pono, string colorname, long portion, string sizename, long sno)
        //{
        //    this.asptblpurdetid = asptblpurdetid;
        //    this.asptblpurid = asptblpurid;
        //    this.asptblpur1id = asptblpur1id;
        //    this.compcode = compcode;restitching
        //    this.pono = pono;
        //    this.colorname = colorname;
        //    this.portion = portion;
        //    this.sizename = sizename;
        //    this.sno = sno;
        //    Class.Users.Query = "select distinct asptblpurdet1id   from  asptblpurdet1   where   asptblpurid='" + p3.asptblpurid + "' AND   asptblpur1id='" + p3.asptblpur1id + "' and compcode='" + combocompcode.SelectedValue + "' and  pono='" + p3.pono + "'  and  colorname='" + p3.colorname + "' and sizename='" + p3.sizename + "' and sno='" + p3.sno + "' ";// and orderqty = '" + pp.orderqty + "'
        //    Class.Users.ds = Utility.ExecuteSelectQuery(Class.Users.Query, "asptblpurdet1");
        //    Class.Users.dt = Class.Users.ds.Tables["asptblpurdet1"];
        //    return Class.Users.dt;
        //}

        public PurchasesModel1detail(long barcode, long asptblpurdetid, long asptblpurid, long asptblpur1id, long compcode, string pono, string colorname, long portion, string sizename, long orderqty,string colorname1, string finyear, string cutting, string stitching, string checking,string restitching,string rechecking, long sno,string inward,string delivery,string panelmistake,string modified)
        {
            this.asptblpurdetid = asptblpurdetid;
            this.asptblpurid = asptblpurid;
            this.asptblpur1id = asptblpur1id;
            this.compcode = compcode;
            this.pono = pono;
            this.colorname = colorname;           
            this.portion = portion;
            this.sizename = sizename;
            this.orderqty = orderqty;
            this.finyear = finyear;
            this.colorname1 = colorname1;
            this.cutting = cutting;
            this.stitching = stitching;
            this.checking = checking;
            this.restitching = restitching;
            this.rechecking = rechecking;
            this.sno = sno;
            this.inward = inward;
            this.delivery = delivery;
            this.panelmistake = panelmistake;
            this.barcode = barcode;
            this.modified = modified;
            Class.Users.Query = "insert into asptblpurdet1(asptblpurdetid,asptblpurid,asptblpur1id,compcode,pono,colorname,portion,sizename,orderqty,colorname1,finyear,cutting,stitching,checking,restitching,rechecking,sno,inward,delivery,panelmistake,barcode,processcheck,ISSUETYPE,modified) values('" + asptblpurdetid + "','" + asptblpurid + "', '" + asptblpur1id + "' ,'" + compcode + "' ,'" + pono + "' , '" + colorname + "','" + portion + "','" + sizename + "','" + orderqty + "','" + colorname1 + "','" + Class.Users.Finyear + "','F','F','F','" + restitching + "','" + rechecking + "','" + sno + "','" + inward + "','" + delivery + "','" + panelmistake + "','" + barcode + "','F','CUTTING',date_format('" + modified + "','%Y-%m-%d'))";
            Utility.ExecuteNonQuery(Class.Users.Query);
        }

        public PurchasesModel1detail(long barcode, long asptblpurid,long asptblpurdetid, long compcode, string pono, string modified,long asptblpurdet1id)
        {
            this.barcode = barcode;
            this.asptblpurid = asptblpurid;
            this.asptblpurdetid = asptblpurdetid;
            this.compcode = compcode;
            this.pono = pono;
            this.modified = modified;
            this.asptblpurdet1id = asptblpurdet1id;
            Class.Users.Query = "update  asptblpurdet1  set barcode='" + barcode + "' ,asptblpurid='" + asptblpurid + "',pono='" + pono + "',asptblpurdetid='" + asptblpurdetid + "',modified=date_format('" + modified + "','%Y-%m-%d') where   asptblpurdet1id='" + asptblpurdet1id + "'";
            Utility.ExecuteNonQuery(Class.Users.Query);
        }
    }
    
}

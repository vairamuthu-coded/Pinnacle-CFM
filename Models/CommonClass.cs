using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Controls;

namespace Pinnacle.Models
{
    public abstract class BaseClassEvent
    {

        public abstract DataTable Tables(string s,string ss);
    }

    public class CommonClass
    {
        public virtual string HideButton(string pono, string tbl)
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
        public virtual string query { get; set; }
        public virtual DataSet ds { get; set; }
        public virtual DataTable dt { get; set; }

        public virtual DataTable select(string qry, string tbl)
        {
            this.query = qry;
            ds = Utility.ExecuteSelectQuery(query, tbl);
            dt = ds.Tables[tbl];
            return dt;
        }
        public virtual DataSet select1(string qry, string tbl)
        {
            this.query = qry;
            ds = Utility.ExecuteSelectQuery(query, tbl);
            return ds;
        }
        public virtual DataTable autonumberload1(string y, string com1, string scr, string tbl)
        {

            Class.Users.Query = "select count(a." + tbl + "1id)+1 as id,max(a.barcode)+1 as barcode from " + tbl + " a join gtcompmast b on a.compcode = b.gtcompmastid  where a.finyear='" + y + "' and b.compcode='" + com1 + "'";
            this.ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
            this.dt = this.ds.Tables[0];
            if (dt.Rows[0]["id"].ToString()=="")
            {
                Class.Users.Query = "select max(a.sequenceno)+1 as id,a.barcode, a.shortcode,a.finyear,b.compcode,b.compname from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptblnavigation c on c.menuid=a.screen where a.finyear='" + y + "' and b.compcode='" + com1 + "' AND c.menuname='" + scr + "' group by a.barcode,a.shortcode,a.finyear,b.compcode,b.compname";
                this.ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                this.dt = this.ds.Tables[0];
            }


            return dt;
        }
        public virtual DataTable autonumberload(string y, string com1, string scr,string tbl)
        {
            Class.Users.Query = "";
            Class.Users.Query = "select count(a." + tbl + "1id)+1 as id from " + tbl + " a join gtcompmast b on a.compcode = b.gtcompmastid  where a.finyear='" + y + "' and b.compcode='" + com1 + "' ";
                    this.ds=Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
            this.dt = this.ds.Tables[0];
            if (dt.Rows.Count <= 0)
            {
                Class.Users.Query = "";
                Class.Users.Query = "select max(a.sequenceno)+1 as id,a.shortcode,a.finyear,b.compcode,b.compname from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptblnavigation c on c.menuid=a.screen where a.finyear='" + y + "' and b.compcode='" + com1 + "' AND c.menuname='" + scr + "' group by a.barcode,a.shortcode,a.finyear,b.compcode,b.compname";
                this.ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
                this.dt = this.ds.Tables[0];
            }
           

            return dt;
        }
        public virtual DataTable shortcode1(string y, string com1, string scr, string tbl)
        {
            Class.Users.Query = "";
            Class.Users.Query = "select distinct a.shortcode,a.barcode from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptblnavigation c on c.menuid=a.screen where a.finyear='" + y + "' and b.compcode='" + com1 + "' AND c.menuname='" + scr + "'";
            this.ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
            this.dt = this.ds.Tables[0];
            return dt;
        }
        public virtual DataTable shortcode(string y, string com1, string scr, string tbl)
        {
            Class.Users.Query = "";
            Class.Users.Query = "select distinct a.shortcode,b.compcode,b.COMPNAME,a.barcode from asptblautogeneratemas a join gtcompmast b on a.compcode = b.gtcompmastid join asptblnavigation c on c.menuid=a.screen where a.finyear='" + y+ "' and b.compcode='"+com1+ "' AND c.menuname='" + scr+"'";
            this.ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
            this.dt = this.ds.Tables[0];
            return dt;
        }

        public virtual string findFinyear(string tbl)
        {
            // Class.Users.Query= "select A.gtfinancialyearid, A.FINYEAR from gtfinancialyear A JOIN GTCOMPMAST B ON A.COMPCODE = B.GTCOMPMASTID WHERE B.COMPCODE = '" + Class.Users.HCompcode + "' AND A.CURRENTFINYR = 'T'";
            //this.ds = Utility.ExecuteSelectQuery(Class.Users.Query, tbl);
            //this.dt = this.ds.Tables[0];
            return dt.Rows[0]["finyear"].ToString();
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models.LYLA
{
    internal class CutPanelReturnModel
    {
        public Int64 asptblcutpanretid;
        public Int64 asptblcutpanret1id;
        public Int64 asptblpurdetid;
        public Int64 asptblpurid;
        public string shortcode;
        public string panelno;
        public string finyear;
        public string cutpaneldate;
        public Int64 compcode;
        public Int64 compname;
        public string pono;
        public Int64 orderqty;
        public Int64 buyer;
        public string lotno;
        public string bundle;  
        public Int64 stylename;
        public string issuetype;
        public string panelmistake;
        public Int64 processname;
        public string Remarks;
        public string cutting;
        public string delivery;
        public string stitching;
        public string checking;
        public string restitching;
        public string rechecking;
        public string   defecttype;
        public string notes;
        public Int64 compcode1;
        public Int64 username;
        public string createdon;
        public string createdby;
        public string modified;
        public string modifiedby;
        public string ipaddress;
    }
    internal class CutPanelReturndetModel :CutPanelReturnModel
    {
        public Int64 asptblcutpanretdetid;          
        public Int64 asptblpurdet1id;
        public Int64 barcode;
        public string colorname;       
        public string sizename;
        public Int64 pcs;
    }
}

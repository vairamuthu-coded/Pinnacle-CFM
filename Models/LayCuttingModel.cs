using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    class LayCuttingModel
    {
        public Int64 asptbllayid;
        public Int64 asptbllayid1;
        public string docid;
        public string shortcode;
        public string finyear;
        public string laydate;
        public Int64 compcode;
        public string prodno;
        public Int64 buyer;
        public string pono;
        public string layno;
        public Int64 compname;
        public Int64 orderqty;   
        public Int64 stylename;
        public string lotno;
        public string markerno;
        public Int64 tableno;
        public string laycancel;
        public string active;
        public Int64 compcode1;
        public Int64 username;
        public string createdon;
        public string createdby;
        public string modified;
        public string modifiedby;
        public string ipaddress;


        public class LayCuttingModeldet
        {
            public Int64 asptbllaydetid;
            public Int64 asptbllayid;
            public Int64 asptbllayid1;
            public Int64 compcode;
            public string prodno;
            public string pono; 
            public string fabric;
            public string colorname;
            public string sizename;
            public string markerno;
            public Int64 orderqty;
            public string uom;
        }
    }
}

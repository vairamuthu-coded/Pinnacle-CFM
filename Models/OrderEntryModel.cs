using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    class OrderEntryModel
    {
        public Int64 asptblordentryid;
        public string asptblordentryid1;
        public string finyear;
        public string orderdate;
        public Int64 compcode;
        public Int64 compname;
        public Int64 buyer;
        public Int64 buyingagent;
        public string orderno;
        public Int64 orderqty;
        public string pono;
        public string stylerefno;
        public Int64 stylename;
        public Int64 sizegroup;
        public Int64 merchandiser;
        public Int64 currency;
        public Int64 currencyvalue;
        public string ordercancel;
        public string active;
        public Int64 compcode1;
        public Int64 username;
        public string createdon;
        public string createdby;
        public string modified;
        public string modifiedby;
        public string ipaddress;

    }
    public class OrderEntryModeldetail
    {
        public Int64 asptblordentrydetid;
        public Int64 asptblordentryid;
        public string asptblordentryid1;
        public Int64 compcode;
        public string orderno;
        public string fabric;
        public string colorname;
        public Int64 qty;
        public string notes;
        public Int64 indexno;
    }
    public class OrderEntryModelsubdetail
    {


        public Int64 asptblordentrysubdetid;
        public Int64 asptblordentrydetid;
        public Int64 asptblordentryid;
        public string asptblordentryid1;
        public Int64 compcode;
        public string orderno;
        public string fabric;
        public string colorname;
        public string sizename;
        public Int64 qty;
        public Int64 excessqty;       
        public Int64 shipqty;
        public Int64 indexno;
    }
}

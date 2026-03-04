using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{

    public class ShiftModel
    {
        public Int64 asptblshimasid;
        public string asptblshimasid1;
        public string finyear;
        public string shiftdate;
        public Int64 compcode;
        public Int64 linegroup;
        public Int64 shiftname;
        public Int64 shiftno;
        public string shiftstart;
        public string shiftend;
        public Int64 breaktime;
        public string breakstart;
        public string breakend;
        public string otminutes;
        public string breakminuts;
        public string shiftcancel;
        public string active;
        public Int64 compcode1;
        public Int64 username;
        public DateTime createdon;
        public string createdby;
        public string modified;
        public string modifiedby;
        public string ipaddress;
        public class ShiftModeldet
        {
            public Int64 asptblshidetid;
            public Int64 asptblshimasid;
            public string asptblshimasid1;
            public Int64 compcode;
            public Int64 shiftname;
            public string weakday;
            public string linename;        
        }
    }
}

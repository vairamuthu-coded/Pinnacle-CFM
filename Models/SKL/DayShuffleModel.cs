using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models
{
    class DayShuffleModel
    {
        public Int64 hrdayshfid { get; set; }
        public Int64 hrdayshfid1;
        public string finyear;
        public Int64 compcode;
        public Int64 compname;
        public string docid;
        public string date;
        public string month;
        public string active;
        public Int64 compcode1;
        public Int64 username;
        public string ipAddress;
        public string createdby;
        public string modifiedby;
        public DateTime createdon;
        public string modified;
    }
    public class DayShuffleDetailModel
    {

        public Int64 hrdayshfdetid; 
        public Int64 hrdayshfid;
        public Int64 hrdayshfid1;
        public Int64 compcode;
        public string ddate;
        public string suftime;
        public string notes;
    }
}

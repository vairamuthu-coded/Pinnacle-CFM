using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models
{
    class BuyerShufflemodel
    {
        public Int64 HrSufTimeid { get; set; }
        public Int64 HrSufTimeid1;
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
    public class HrSufTimeDetmodel
    {

        public Int64 HrSufTimeDetid;
        public Int64 HrSufTimeid;
        public Int64 HrSufTimeid1;
        public Int64 compcode;   
        public string ddate;
        public string suftime;
        public string notes;
    }
}

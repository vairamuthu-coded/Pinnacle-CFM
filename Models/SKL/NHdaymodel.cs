using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models
{
    public class NHdaymodel
    {
        //public NHdaymodel()
        //{
        //}

        public Int64 hrnhmastid;
        public Int64 hrnhmastid1;
        public string finyear;
        public Int64 compcode;
        public Int64 compname;
        public string docid;
        public string date;
        public string active;       
        public Int64 compcode1;
        public Int64 username;
        public string ipAddress;
        public string createdby;
        public string modifiedby;
        public DateTime createdon;
        public string modified;

        public class hrnhdetails
        {

            public Int64 hrnhdetailsid;
            public Int64 hrnhmastid;
            public Int64 hrnhmastid1;
            public Int64 compcode;
            public Int64 Rownumber;
            public string nhdate;
            public string month;
            public string day;
            public string year;
            public string Reason;
            public Int64 holidaycategory;
            public string ApplicableDetails;
            public string WorkingDay;
            public string notes;
        }
       

    }
}

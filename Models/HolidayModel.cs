using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    public class HolidayModel
    {
        public Int64 asptblholmasid;
        public Int64 asptblholmasid1;
        public string finyear;
        public Int64 compcode;
        public Int64 compname;
        public string month;
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

        public class HolidayModeldetails
        {

            public Int64 asptblholdetid;
            public Int64 asptblholmasid;
            public Int64 asptblholmasid1;
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

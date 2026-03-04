using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    public class ProcessGroupModel
    {
        public Int64 asptblprocessgroupid;
        public string processgroup;
        public string active;
        public Int64 compcode;
        public Int64 username;
        public string createdby;
        public DateTime createdon;
        public string modifiedby;
        public DateTime modified;
        public string ipAddress;
    }

    public class ProcessGroupDetModel: ProcessGroupModel
    {
        public Int64 asptblprocessgroupdetid;
        public Int64 processname;        
        public decimal rate;
        public string notes;
    }
}
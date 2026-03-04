using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
   public class stylegroupModel
    {

        public Int64 asptblstygrpmasid;
        public string stylegroup;
        public string active;
        public Int64 compcode;
        public Int64 username;
        public string createdby;
        public DateTime createdon;
        public string modifiedby;
        public DateTime modified;
        public string ipAddress;
    }
    internal class stylegroupModeldet
    {
        public Int64 asptblstygrpdetid;
        public Int64 asptblstygrpmasid;
        public Int64 stylename;
        public string stylegroup;
        public string notes;



    }
}

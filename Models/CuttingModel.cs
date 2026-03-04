using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models
{
  
    class CuttingModel
    {
        public Int64 asptblcutid;
        public Int64 asptblcutid1;

        public Int64 asptblbunid;
        public Int64 asptblbunid1;

        public string docid;
        public string shortcode;
        public string finyear;
        public string docdate;
        public Int64 compcode;
        public Int64 buyer;
        public string pono;
        public string layno;
        public Int64 compname;
        public Int64 orderqty;
        public Int64 stylename;
        public string lotno;
        public string cutno;
        public Int64 shiftno;
        public string cutcancel;
        public string active;
        public Int64 compcode1;
        public Int64 username;
        public string createdon;
        public string createdby;
        public string modified;
        public string modifiedby;
        public string ipaddress;
        public class CuttingModeldet
        {
            public Int64 asptblcutdetid;
            public Int64 asptblcutid;
            public Int64 asptblcutid1;

            public Int64 asptblbundetid;
            public Int64 asptblbunid;
            public Int64 asptblbunid1;
            
            public Int64 compcode;
            public string pono;
            public string fabric;
            public string colorname;
            public string sizename;
            public string markerno;
            public Int64 orderqty;
            public string notes;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models.Bank
{
   public class Ifsc
    {
        public Int64 asptblifscmasid { get; set; }
        public string ifsc { get; set; }
        public string branch { get; set; }
        public string bankname { get; set; }
        public string active { get; set; }
        public Int64 compcode { get; set; }
        public Int64 username { get; set; }
        public string createdby { get; set; }
        public string ipaddress { get; set; }
        public string modifiedby { get; set; }
        public string createdon { get; set; }
        public Int64 paymenttype { get; set; }
    }
}

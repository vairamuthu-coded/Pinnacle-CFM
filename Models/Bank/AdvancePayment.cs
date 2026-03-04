using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models.Bank
{
  public  class AdvancePayment
    {
     public Int64  asptbladvpaydetid { get; set; }
        public Int64 asptbladvpaymasid { get; set; }
        public Int64  compcode { get; set; }
        public Int64 department { get; set; }
        public Int64 partyname { get; set; }
        public string invoicetype { get; set; }
        public string invoice { get; set; }
        public Int64 invoicebyte { get; set; }
        public Int64 proforminvoicebyte { get; set; }
        public Int64 quatationbyte { get; set; }
        public Int64 otherbyte { get; set; }
        public Int64 powobyte { get; set; }

    }
}

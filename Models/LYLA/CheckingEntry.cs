using MessagingToolkit.QRCode.Helper;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Pinnacle.Models.LYLA
{

   

    public class CheckingEntry
    {
        public Int64 asptblpurid;
        public Int64 asptblchkid;
        public Int64 asptblchk1id;
        public string shortcode;
        public string finyear;     
        public string date;
        public string checkno;
        public string prodno;
        public Int64 compcode;
        public Int64 compname;
        public Int64 orderqty;
        public Int64 buyer;
        public string pono;
        public string bundle;
        public Int64 processname;
        public string processtype;
        public string issuetype;
        public string checking;
        public Int64 stylename;
        public Int64 colorname;
        public string sizename;
        public string lotno;
        public string notes;
        public string pocancel;
        public string active;
        public string inward;
        public string delivery;
        public string restitching;
        public string rechecking;
        public Int64 compcode1;
        public Int64 username;
        public string createdon;
        public string createdby;
        public string modified;
        public string modifiedby;
        public string ipaddress;
    }
    public class CheckingEntrydetail:CheckingEntry
    {
        public Int64 asptblchkdetid;
        public Int64 asptblpurdet1id;
        public Int64 barcode;
        public Int64 asptblpurdetid;
        public Int64 asptblpurid;       
        public string colorname;
        public string sizename;       
        public Int64 orderqty;
        public Int64 comqty;
        public Int64 lotqty;
        public Int64 balqty;
        public string notes;
        public string process;
        public string processcheck;
        
    }
   
}

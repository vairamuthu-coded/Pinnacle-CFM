using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
   public class TaxTemplateModel
    {
        public Int64 asptbltaxtempmasid;
        public string finyear;
        public string taxgroup;
        public string taxdesc;
        public string active;
        public Int64 compcode;
        public Int64 compname;
        public Int64 username;
        public Int64 createdby;
        public string createdon;
        public Int64 modifiedby;
        public string modifiedon;
        public string ipaddress;
    }
    public class asptbltaxcomModel: TaxTemplateModel
    {
        public Int64 asptbltaxcomid;   
    }

    public class asptbltaxtempdetModel: TaxTemplateModel
    {
        public Int64 asptbltaxtempdetid;
        public Int64 taxname;
        public string taxtype;
        public string tax;
        public string taxid;
        public string formula;        
    }
}

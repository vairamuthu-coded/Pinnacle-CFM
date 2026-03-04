using System;
using System.Collections.Generic;
using System.Data;
namespace Pinnacle.Models
{
    internal class SizeGroupModel
    {
        public Int64 asptblsizgrpid;        
        public string sizegroup;
        public string active;
        public Int64 compcode;
        public Int64 username;    
        public string createdby;    
        public DateTime createdon;
        public string modifiedby;
        public DateTime modified;
        public string ipAddress;
    }
    internal class SizeGroupDetModel
    {
        public Int64 asptblsizgrpdetid;
        public Int64 asptblsizgrpid;      
        public Int64 sizename;
        public string sizegroup;
        public string notes;

       

    }
}
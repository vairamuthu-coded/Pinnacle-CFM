using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    class LineGroupModel
    {
       

            public Int64 asptbllingrpmasid;
            public string linegroup;
            public string active;
            public Int64 compcode;
            public Int64 username;
            public string createdby;
            public DateTime createdon;
            public string modifiedby;
            public DateTime modified;
            public string ipAddress;
        }
        internal class LineGroupdetModel
    {
            public Int64 asptbllingrpdetid;
            public Int64 asptbllingrpmasid;
            public Int64 linename;
            public string linegroup;
            public string notes;



        }
    }

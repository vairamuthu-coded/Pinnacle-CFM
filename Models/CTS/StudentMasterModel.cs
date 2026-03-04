using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models.CTS
{
    class StudentMasterModel
    {
        public Int64 asptblstudentmasid { get; set; }
        public Int64 asptblvotemasid { get; set; }
        public Int64 asptblstudentmasid1 { get; set; }
        public Int64 compcode { get; set; }
        public Int64 compname { get; set; }
        public string studentname { get; set; }
        public string lastname { get; set; }
        public string address { get; set; }
        public string gender { get; set; }
        public string dateofbirth { get; set; }
        public string ROLLNO { get; set; }
        public string dateofjoin { get; set; }
        public Int64 standard { get; set; }
        public string standard1 { get; set; }
        public string bloodgroup { get; set; }
        public Int64 SECTION { get; set; }
        public string contact { get; set; }
        public Int64 BLOCKNO { get; set; }
        public Byte[] STUDENTIMAGE { get; set; }
        public Byte[] STUDENTVOTELOG { get; set; }
        public string active { get; set; }
        public Int64 STUDENTIMAGEBYTES { get; set; }
        public Int64 STUDENTVOTEBYTES { get; set; }
        public Int64 username { get; set; }
        public string ipaddress { get; set; }
        public string createdby { get; set; }
        public string createdon { get; set; }
        public string modifiedon { get; set; }
        public string election { get; set; }
        public int votecount { get; set; }
        public string votedate { get; set; }
        public Int64 electionpost { get; set; }
        public string electiondate { get; set; }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle
{
    public class SMSResponse
    {
        public string ErrorCode { get; set; }
        public string ErrorDescription { get; set; }
        public List<MessageData> Data { get; set; }
    }
    public class MessageData
    {
        public string MessageErrorCode { get; set; }
        public string MessageErrorDescription { get; set; }
        public string MobileNumber { get; set; }
        public string MessageId { get; set; }
        public string Custom { get; set; }
    }
}

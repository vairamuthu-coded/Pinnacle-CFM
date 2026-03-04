using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
    public class Serial
    {
        public static string PortNo { get; set; }
        public static string PortIP { get; set; }
        public static string PortName { get; set; }
        public static Int32 PortWeight { get; set; }
        public static string PortType { get; set; }
        public static Int32 BaudRate = 9600;
        public static Int32 DataBits = 7;
        public static string StopBits = "1";
        public static string Parity = "Even";
        public static string FlowControl = "None";
        public static bool ServerConnected = false;
        public static string MessageClient = "0";
    }
}

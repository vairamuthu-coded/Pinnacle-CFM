using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Collections;
namespace Pinnacle.Models
{
    public class Party
    {
        public Party()
        {
           
        }


       
        public Int64 PartyID { get; set; }
     
        public string partycode { get; set; }
        public string partyname { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Country { get; set; }
        public Int64 PinCode { get; set; }
        public string PanNo { get; set; }
    public string TinNo {  get; set; }
    public string GstNo {  get; set; }
     
        public string GstDate { get; set; }
        public string PhoneNo { get; set; }
        public string accno { get; set; }
        public string Email { get; set; }
        public string WebSite { get; set; }
    public string contactname { get; set; }
    public string Active { get; set; }
        public string bankname { get; set; }
        public string branch { get; set; }
        public string ifsc { get; set; }
        public string accountholdername { get; set; }
        public string CodeNo { get; set; }
        public string UserName { get; set; }
        public string CreatedOn { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedOn { get; set; }
        public string IpAddress {  get; set; }


       
    }
}

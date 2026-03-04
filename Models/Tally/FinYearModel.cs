using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pinnacle.Models.Tally
{
    public class FinYearModel
    {
        public Int64 Asptblfinid { get; set; }
        public string Finyear { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public Int64 TotalDays { get; set; }
        public Int64 CompCode { get; set; }
        public Int64 UserName { get; set; }
        public string  Modifiedon  { get; set; }
        public string Createdon  { get; set; }
        public string CreaedBy  { get; set; }
        public string IPAddress  { get; set; }
        public string Finyr { get; set; }
        public string CurrentFinyr  { get; set; }
    }
    public class FinYearModel1 : FinYearModel
    {
        public Int64 Asptblfindetid { get; set; }
        public Int64 Asptblfindet1id { get; set; }
        public DateTime PeriodStartDate { get; set; }
        public DateTime PeriodEndDate { get; set; }
        public Int64 PeriodCode { get; set; }
        public string PeriodDescription { get; set; }
        public string SalCalDesc { get; set; }
        public string Quarter { get; set; }
        //Asptblfindetid,Asptblfinid,Finyear,Compcode,PeriodStartDate,PeriodEndDate,PeriodCode,PeriodDescription,
        //SalCalDesc,Quarter,TotalDays,Notes
        public string Notes { get; set; }

    }
    public class FinYearModel2 : FinYearModel1
    {
        public Int64 asptblfinwekdetid { get; set; }
        public Int64 asptblfinwekdet1id { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        //asptblfinwekdetid,Asptblfinid,Finyear,CompCode,StartDate,AddDays,EndDate,TotalDays,SalaryDate,PayMonth,PayPeriod,WeeklyHolDay,WorkingDay,Notes
        public Int64 AddDays { get; set; }
       
        public DateTime SalaryDate { get; set; }
        public string PayMonth { get; set; }
        public string PayPeriod { get; set; }
        public Int64 WeeklyHolDay { get; set; }
        public Int64 WorkingDay { get; set; }

    }
    public class FinYearModel3 : FinYearModel2
    {
        public Int64 asptblfinmondetid { get; set; }
        public Int64 asptblfinmondet1id { get; set; }

        //asptblfinmondetid,Asptblfinid,Finyear,CompCode,StateDate,EndDate,PayPeriod,SalaryDate,PayPeriodDays,
        //WeeklyHolDay,Notes


        public Int64 PayPeriodDays { get; set; }
        
    }
}
